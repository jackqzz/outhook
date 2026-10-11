#!/usr/bin/env python3
"""
ipiap.com 卡密批量充值 WebUI 服务端

静态 HTML 无法直接调 ipiap.com (Cloudflare + CORS), 所以:
  本进程持有真实 Chrome (持久化 profile) 会话, 浏览器页面内发 fetch;
  WebUI (index.html) 只是批处理控制面板, 走本地 JSON API。

接口:
  GET  /            -> index.html
  GET  /mail.html   -> Outlook 收信面板
  GET  /api/state   -> {browser:{status,detail}, batch:{...}}   前端轮询
  POST /api/batch   -> {mode:"check"|"recharge", items:[{cdk,org}...]}
  POST /api/stop    -> 停止当前批次 (跑完当前条目)
  POST /api/mail/login    -> {refresh_token, client_id?} -> {access_token}
  POST /api/mail/messages -> {access_token, folder?, top?} -> 邮件列表
  POST /api/mail/message  -> {access_token, id} -> 完整邮件(含 body)

邮件链: MSA refresh_token(M.C547...) + client_id 换 access_token
(scope=https://outlook.office.com/IMAP.AccessAsUser.All, 附带 Mail.ReadWrite),
再走 Outlook REST v2.0 /api/v2.0/me/... 读信。微软接口无 CF, 直接 urllib。

用法:
  ./server.py                 # http://127.0.0.1:8321
  ./server.py --port 9000 --headed
"""

import argparse
import http.server
import json
import os
import queue
import threading
import time
import urllib.error
import urllib.parse
import urllib.request

import recharge  # 同目录: open_site / api

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
INDEX_HTML = os.path.join(BASE_DIR, "index.html")

POLL_INTERVAL = 3.0      # step4 轮询间隔
ITEM_TIMEOUT = 300.0     # 单条 step4 最长等待
ITEM_GAP = 2.0           # 批次内每条之间的间隔, 防限流
RETRY_DELAY = 5.0        # 限流退避起始秒数
APPLE_DEFAULT_HOST = "https://kekoukele1822.applelll.com"
APPLE_FAIL_STATES = {"failed", "internal_error", "cancelled", "timeout"}
MOBI_DEFAULT_HOST = "https://9527.mobi"
MOBI_OK_STATUSES = {"success"}
MOBI_FAIL_STATUSES = {"failed", "manual_review"}


# ------------------------------------------------------------ browser --

class Browser:
    """worker 线程独占 playwright; 其余线程只读 status。"""

    def __init__(self):
        self.pw = None
        self.ctx = None
        self.page = None
        self.status = "starting"     # starting/ready/cf_blocked/dead
        self.detail = ""
        self.lock = threading.Lock() # 保护 status/detail

    def set(self, status, detail=""):
        with self.lock:
            self.status, self.detail = status, detail

    def state(self):
        with self.lock:
            return {"status": self.status, "detail": self.detail}

    def start(self, force_headed=False):
        from playwright.sync_api import sync_playwright
        self.set("starting", "启动浏览器 ...")
        try:
            self.pw = sync_playwright().start()
        except Exception as e:
            self.set("dead", f"playwright 启动失败: {e}")
            return
        ctx, page = None, None
        if not force_headed:
            ctx, page = recharge.open_site(self.pw, headless=True)
            if not page:
                self.set("starting", "headless 未过 CF, 改有头模式 ...")
        if not page:
            ctx, page = recharge.open_site(self.pw, headless=False)
        if not page:
            self.set("cf_blocked", "Cloudflare 90 秒未放行")
            self._shutdown_pw()
            return
        self.ctx, self.page = ctx, page
        self.set("ready")

    def ensure(self):
        """页面死掉时重启浏览器 (headed, 让人能过 CF)。"""
        try:
            self.page.evaluate("1")
            if self.status == "ready":
                return True
        except Exception:
            pass
        try:
            self.ctx.close()
        except Exception:
            pass
        self._shutdown_pw()
        self.start(force_headed=True)
        return self.status == "ready"

    def _shutdown_pw(self):
        try:
            self.pw.stop()
        except Exception:
            pass
        self.pw = self.ctx = self.page = None

    def stop(self):
        try:
            self.ctx.close()
        except Exception:
            pass
        self._shutdown_pw()
        self.set("dead", "已停止")


# ------------------------------------------------------------ batch --

class StepError(Exception):
    pass


def _post_once(page, step, payload):
    r = recharge.api(page, step, payload)
    if r["status"] != 200:
        raise StepError(f"HTTP {r['status']}: {r['body'][:200]}")
    try:
        resp = json.loads(r["body"])
    except json.JSONDecodeError:
        raise StepError(f"响应不是 JSON: {r['body'][:200]}")
    if resp.get("code") != 0:
        raise StepError(f"{resp.get('msg')} (code={resp.get('code')})")
    return resp.get("data") or {}


def _rate_limited(e):
    # code=1001 是通用错误码, 只有"频繁/稍后"这种才值得重试
    return "频繁" in str(e) or "稍后" in str(e)


def post(page, step, payload, note=None, tries=4):
    """带限流退避的 post。note(msg) 用于把重试状态显示到 UI。"""
    delay = RETRY_DELAY
    for i in range(tries):
        try:
            return _post_once(page, step, payload)
        except StepError as e:
            if i == tries - 1 or not _rate_limited(e):
                raise
            if note:
                note(f"限流, {int(delay)}s 后重试 ({i+1}/{tries-1}) ...")
            time.sleep(delay)
            delay += RETRY_DELAY


class Batch:
    def __init__(self, mode, items, site="ipiap",
                 host=APPLE_DEFAULT_HOST, product_id=None):
        self.mode = mode
        self.items = items   # [{idx,cdk,org,state,product,orderNo,message}]
        self.site = site
        self.host = host.rstrip("/")
        self.product_id = product_id
        self.finished = False
        self.stop_requested = False
        self.lock = threading.Lock()

    def snapshot(self):
        with self.lock:
            return {"mode": self.mode, "site": self.site,
                    "finished": self.finished,
                    "items": [dict(i) for i in self.items]}


JOB_QUEUE = queue.Queue()
BROWSER = Browser()
CURRENT = {"batch": None}
FORCE_HEADED = False


# ------------------------------------------------------------ mail api --

MS_TOKEN_URL = ("https://login.microsoftonline.com/consumers"
                "/oauth2/v2.0/token")
OAPI = "https://outlook.office.com/api/v2.0"
MAIL_SCOPE = ("https://outlook.office.com/IMAP.AccessAsUser.All "
              "offline_access")
DEFAULT_CLIENT_ID = "9e5f94bc-e8a4-4e73-b8be-63364c29d753"
FOLDERS = {"inbox": "inbox", "junkemail": "junkemail",
           "sentitems": "sentitems", "drafts": "drafts"}


def mail_login(client_id, refresh_token):
    """RT -> AT。返回 dict: 成功含 access_token, 失败含 error。"""
    data = urllib.parse.urlencode({
        "client_id": client_id or DEFAULT_CLIENT_ID,
        "grant_type": "refresh_token",
        "refresh_token": refresh_token,
        "scope": MAIL_SCOPE}).encode()
    req = urllib.request.Request(
        MS_TOKEN_URL, data=data, method="POST",
        headers={"content-type": "application/x-www-form-urlencoded"})
    try:
        r = json.loads(urllib.request.urlopen(req, timeout=30).read())
        return {"access_token": r["access_token"],
                "expires_in": r.get("expires_in")}
    except urllib.error.HTTPError as e:
        try:
            err = json.loads(e.read())
            msg = err.get("error_description") or err.get("error") or str(e)
        except Exception:
            msg = f"HTTP {e.code}"
        return {"error": msg.split("\r\n")[0][:300]}
    except Exception as e:
        return {"error": f"网络错误: {e}"}


def mail_get(access_token, path):
    req = urllib.request.Request(
        OAPI + path, headers={"authorization": "Bearer " + access_token})
    try:
        return json.loads(urllib.request.urlopen(req, timeout=30).read())
    except urllib.error.HTTPError as e:
        try:
            err = json.loads(e.read())
            msg = (err.get("error", {}) or {}).get("message") or str(e)
        except Exception:
            msg = f"HTTP {e.code}"
        return {"error": msg[:300]}
    except Exception as e:
        return {"error": f"网络错误: {e}"}


# ------------------------------------------------------------ applelll --

def apple_ref(host):
    """与前端 getRef() 一致: xxx.applelll.com -> xxx, 否则用整个 host。"""
    h = host.split("//", 1)[-1].split("/")[0].split(":")[0]
    if h.endswith(".applelll.com"):
        return h.split(".")[0]
    return h


def apple_headers():
    """防重放头: 每请求一个全新 nonce (服务端会 REPLAY_DETECTED)。"""
    import base64
    n = base64.urlsafe_b64encode(os.urandom(18)).decode().rstrip("=")
    return {"content-type": "application/json",
            "X-ASP-Timestamp": str(int(time.time())),
            "X-ASP-Nonce": n, "Idempotency-Key": n,
            "user-agent": recharge.UA}


def _apple_call(host, path, payload=None, tries=3):
    """GET(path不带body) 或 POST(path+payload)。返回 dict, 失败抛 StepError。"""
    for i in range(tries):
        if payload is None:
            req = urllib.request.Request(
                host + path, headers={"user-agent": recharge.UA})
        else:
            req = urllib.request.Request(
                host + path, data=json.dumps(payload).encode(),
                headers=apple_headers(), method="POST")
        try:
            return json.loads(urllib.request.urlopen(req, timeout=30).read())
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "ignore")
            try:
                d = json.loads(body)
                if d.get("code") == "REPLAY_DETECTED" and i < tries - 1:
                    continue          # 换新 nonce 重试
                msg = d.get("detail") or d.get("message") or str(d)[:200]
            except Exception:
                msg = body[:200] or f"HTTP {e.code}"
            if "频繁" in msg or "稍后" in msg or e.code == 429:
                if i < tries - 1:
                    time.sleep(RETRY_DELAY * (i + 1))
                    continue
            raise StepError(f"{msg} (HTTP {e.code})")
        except Exception as e:
            if i == tries - 1:
                raise StepError(f"网络错误: {e}")
            time.sleep(1)


def run_item_apple(item, batch):
    """applelll 兑换中心: 直连 HTTP, 不需要浏览器。
    check 模式 = status-batch 查卡状态(不消耗); recharge = 提交+轮询。"""
    host, ref = batch.host, apple_ref(batch.host)
    item["state"] = "running"
    item["message"] = "查询卡密状态 ..."
    if batch.mode == "check":
        r = _apple_call(host, "/api/store/redeem/status-batch",
                        {"ref": ref, "cdks": [item["cdk"]]})
        res = (r.get("results") or [{}])[0]
        item["product"] = res.get("product_name") or ""
        item["orderNo"] = res.get("redeem_account_id") or ""
        if not res.get("found"):
            item["state"] = "done"
            item["message"] = "系统无此卡兑换记录 (未被兑换过)"
            return
        st, ts = res.get("status"), res.get("task_status")
        att = res.get("attempt")
        if st == "redeemed" or ts == "success":
            item["state"] = "done"
            item["message"] = f"已被兑换 ({res.get('message') or 'redeemed'})"
        elif st == "active":
            # active = 卡在系统中且未被兑换 = 可用; attempts>0 说明上次回滚过
            item["state"] = "done"
            item["message"] = ("卡密可兑换 (active"
                               + (f", 历史尝试{att}次已回滚" if att else "")
                               + ")")
        elif ts in APPLE_FAIL_STATES:
            item["state"] = "error"
            item["message"] = (f"兑换失败: {res.get('message') or ts} "
                               f"(status={st} attempts={att})")
        else:
            item["state"] = "running"
            item["message"] = (f"status={st} task={ts} "
                               f"{res.get('message') or ''} "
                               f"attempts={att}")
        if item["state"] == "running":
            item["state"] = "done"   # check 模式不等待, 只报当前态
            item["message"] = "兑换中 " + item["message"]
        return

    # recharge 模式: 提交兑换
    item["message"] = "提交兑换 ..."
    r = _apple_call(host, "/api/store/redeem",
                    {"ref": ref, "cdk": item["cdk"],
                     "account_id": item["org"],
                     "product_id": batch.product_id})
    item["orderNo"] = r.get("account_id") or item["org"]
    item["message"] = "已受理, 等待到账 ..."

    # 轮询: 2.5s 间隔 (与官方前端一致), status=active 表示失败回滚
    deadline = time.time() + ITEM_TIMEOUT
    while True:
        d = _apple_call(host, "/api/store/redeem/status?ref="
                        + urllib.parse.quote(ref) + "&cdk="
                        + urllib.parse.quote(item["cdk"]))
        st, ts = d.get("status"), d.get("task_status")
        item["product"] = d.get("product_name") or item["product"]
        if st == "redeemed" or ts == "success":
            item["state"] = "done"
            item["message"] = "充值成功"
            return
        if ts in APPLE_FAIL_STATES or st == "active":
            raise StepError(f"充值失败: {d.get('message') or ts or st} "
                            f"(卡密已回滚可重试)")
        if ts == "possible_success":
            item["message"] = "疑似成功, 确认中 ..."
        else:
            item["message"] = f"兑换中 ({ts or st}) ..."
        if time.time() > deadline:
            raise StepError("处理超时, 可在站点查兑换状态")
        time.sleep(2.5)


# ------------------------------------------------------------ 9527.mobi --

def _mobi_call(host, path, payload=None, tries=3):
    """9527.mobi: GET/POST JSON。{ok:false,message} 或非 2xx 抛 StepError。"""
    for i in range(tries):
        if payload is None:
            req = urllib.request.Request(
                host + path,
                headers={"user-agent": recharge.UA,
                         "accept": "application/json"})
        else:
            req = urllib.request.Request(
                host + path, data=json.dumps(payload).encode(),
                headers={"user-agent": recharge.UA,
                         "content-type": "application/json",
                         "accept": "application/json"}, method="POST")
        try:
            r = json.loads(urllib.request.urlopen(req, timeout=30).read())
            if r.get("ok") is False:
                raise StepError(str(r.get("message") or r)[:200])
            return r
        except StepError:
            raise
        except urllib.error.HTTPError as e:
            body = e.read().decode("utf-8", "ignore")
            try:
                msg = json.loads(body).get("message") or body[:200]
            except Exception:
                msg = body[:200] or f"HTTP {e.code}"
            if ("频繁" in str(msg) or "稍后" in str(msg) or e.code == 429) \
                    and i < tries - 1:
                time.sleep(RETRY_DELAY * (i + 1))
                continue
            raise StepError(f"{msg} (HTTP {e.code})")
        except Exception as e:
            if i == tries - 1:
                raise StepError(f"网络错误: {e}")
            time.sleep(1)


def run_item_mobi(item, batch):
    """9527.mobi: precheck 验卡(不消耗) -> orders 下单 -> orders/{token} 轮询。
    仅支持 inputMode=organization_id 的卡 (Claude OrgID 充值)。"""
    host = batch.host
    item["state"] = "running"
    item["message"] = "识别卡密 ..."
    pre = _mobi_call(host, "/api/precheck", {"cardCode": item["cdk"]})
    pname = pre.get("productName") or ""
    mode = pre.get("inputMode") or ""
    item["product"] = pname
    if batch.mode == "check":
        item["state"] = "done"
        item["message"] = f"卡密有效: {pname or '?'} (mode={mode})"
        return
    # recharge: 目前只接 org-id 模式; 要 SK/session/ChatGPT 的卡直接报不支持
    if pre.get("appbuyClaudePreviewRequired") or pre.get("requiresSessionKey"):
        raise StepError(f"{pname}: 该卡要求 SessionKey 验证, 暂不支持(卡未消耗)")
    if mode != "organization_id":
        raise StepError(f"{pname}: inputMode={mode}, 仅支持 organization_id 卡")
    item["message"] = "创建订单 ..."
    r = _mobi_call(host, "/api/orders", {
        "cardCode": item["cdk"],
        "precheckToken": pre.get("precheckToken") or "",
        "sessionText": "", "organizationId": item["org"],
        "claudeSessionKey": "", "claudeAccountPreviewToken": "",
        "chatgptAccountId": ""})
    url = r.get("orderUrl") or ""
    token = url.rstrip("/").rsplit("/", 1)[-1] if url else \
        (r.get("orderToken") or r.get("token") or "")
    if not token:
        raise StepError(f"下单后未拿到订单 token: {str(r)[:150]}")
    item["orderNo"] = token
    item["message"] = "订单已创建, 等待充值 ..."
    deadline = time.time() + ITEM_TIMEOUT
    while True:
        d = _mobi_call(host, "/api/orders/" + urllib.parse.quote(token))
        o = d.get("order") or {}
        st = o.get("status")
        if st in MOBI_OK_STATUSES:
            item["state"] = "done"
            item["message"] = o.get("statusLabel") or "充值成功"
            return
        if st in MOBI_FAIL_STATUSES:
            raise StepError(f"{o.get('statusLabel') or st}: "
                            f"{o.get('reason') or o.get('statusDescription')}")
        item["message"] = f"{o.get('statusLabel') or st} " \
                          f"({o.get('statusDescription') or ''})"
        if o.get("isFinal"):
            raise StepError(f"订单终态 status={st}")
        if time.time() > deadline:
            raise StepError(f"处理超时, 订单页 /order/{token} 可查")
        time.sleep(5)   # 官方前端轮询 5s 起步


def run_item(page, item, mode):
    item["state"] = "running"
    note = lambda m: item.__setitem__("message", m)
    item["message"] = "校验卡密 ..."
    info = post(page, "step1", {"serialNumber": item["cdk"]}, note)
    item["product"] = (f"{info.get('productName')} · "
                       f"{info.get('platform')} · 到期{info.get('expiredTime')}")
    item["message"] = "校验组织 ..."
    post(page, "step2", {"organizationId": item["org"]}, note)
    if mode == "check":
        item["state"] = "done"
        item["message"] = "校验通过"
        return
    item["message"] = "下单中 ..."
    order = post(page, "step3", {"organizationId": item["org"],
                                 "serialNumber": item["cdk"]}, note)
    item["orderNo"] = order["orderNo"]
    deadline = time.time() + ITEM_TIMEOUT
    while True:
        st = post(page, "step4", {"orderNo": item["orderNo"]}, note)
        status = st.get("orderStatus")
        if status == 1:
            if time.time() > deadline:
                raise StepError(f"订单 {item['orderNo']} 处理超时")
            item["message"] = f"订单 {item['orderNo']} 处理中 ..."
            time.sleep(POLL_INTERVAL)
            continue
        if status == 2:
            item["state"] = "done"
            item["message"] = f"充值成功 ({st.get('reason') or 'ok'})"
            return
        raise StepError(f"充值失败 status={status}: {st.get('reason')}")


def worker():
    # playwright 绑定创建它的线程: 浏览器必须在本线程启动和使用
    BROWSER.start(FORCE_HEADED)
    while True:
        batch = JOB_QUEUE.get()
        CURRENT["batch"] = batch
        relaunched = False
        for item in batch.items:
            if batch.stop_requested:
                item["state"] = "stopped"
                item["message"] = "已停止"
                continue
            if batch.site in ("applelll", "mobi"):
                # 纯 HTTP, 不需要浏览器
                try:
                    (run_item_apple if batch.site == "applelll"
                     else run_item_mobi)(item, batch)
                except Exception as e:
                    item["state"] = "error"
                    item["message"] = str(e)
            else:
                if BROWSER.status != "ready" and not BROWSER.ensure():
                    item["state"] = "error"
                    item["message"] = "浏览器不可用: " + BROWSER.detail
                    continue
                try:
                    run_item(BROWSER.page, item, batch.mode)
                except Exception as e:
                    msg = str(e)
                    if "Execution context" in msg or "closed" in msg.lower():
                        if not relaunched and BROWSER.ensure():
                            relaunched = True
                            try:
                                run_item(BROWSER.page, item, batch.mode)
                                continue
                            except Exception as e2:
                                msg = str(e2)
                    item["state"] = "error"
                    item["message"] = msg
            # 条目间隔, 防触发 "操作过于频繁"
            if not batch.stop_requested:
                time.sleep(ITEM_GAP)
        with batch.lock:
            batch.finished = True


# ------------------------------------------------------------ http --

class Handler(http.server.BaseHTTPRequestHandler):
    def log_message(self, *a):
        pass

    def _json(self, obj, code=200):
        body = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("content-type", "application/json; charset=utf-8")
        self.send_header("content-length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _body(self):
        return json.loads(self.rfile.read(int(self.headers.get("content-length", 0))) or b"{}")

    def do_GET(self):
        if self.path == "/" or self.path == "/index.html":
            data = open(INDEX_HTML, "rb").read()
            self.send_response(200)
            self.send_header("content-type", "text/html; charset=utf-8")
            self.send_header("content-length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        elif self.path == "/mail" or self.path == "/mail.html":
            data = open(os.path.join(BASE_DIR, "mail.html"), "rb").read()
            self.send_response(200)
            self.send_header("content-type", "text/html; charset=utf-8")
            self.send_header("content-length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        elif self.path == "/picker" or self.path == "/picker.html":
            data = open(os.path.join(BASE_DIR, "picker.html"), "rb").read()
            self.send_response(200)
            self.send_header("content-type", "text/html; charset=utf-8")
            self.send_header("content-length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        elif self.path == "/api/state":
            b = CURRENT["batch"]
            self._json({"browser": BROWSER.state(),
                        "batch": b.snapshot() if b else None})
        elif self.path.startswith("/api/apple/products"):
            # 代理 applelll 的 site 接口, 给前端产品下拉用
            qs = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            host = (qs.get("host") or [APPLE_DEFAULT_HOST])[0].rstrip("/")
            try:
                d = _apple_call(host, "/api/store/site?ref="
                                + urllib.parse.quote(apple_ref(host)))
                self._json({"products": [
                    {"product_id": p.get("product_id"),
                     "name": p.get("name"), "platform": p.get("platform"),
                     "stock": p.get("stock")}
                    for p in (d.get("products") or [])]})
            except Exception as e:
                self._json({"error": str(e)})
        else:
            self.send_error(404)

    def do_POST(self):
        if self.path == "/api/batch":
            try:
                req = self._body()
            except Exception:
                return self._json({"error": "请求体不是 JSON"}, 400)
            mode = req.get("mode")
            items = req.get("items") or []
            site = req.get("site") or "ipiap"
            if mode not in ("check", "recharge"):
                return self._json({"error": "mode 必须是 check 或 recharge"}, 400)
            if site not in ("ipiap", "applelll", "mobi"):
                return self._json({"error": "site 必须是 ipiap/applelll/mobi"}, 400)
            if not items:
                return self._json({"error": "items 为空"}, 400)
            if len(items) > 200:
                return self._json({"error": "单批最多 200 条"}, 400)
            host = (req.get("host") or APPLE_DEFAULT_HOST).strip()
            product_id = req.get("product_id")
            if site == "applelll" and mode == "recharge" and not product_id:
                return self._json({"error": "applelll 充值需要 product_id"}, 400)
            cur = CURRENT["batch"]
            if cur and not cur.snapshot()["finished"]:
                return self._json({"error": "已有批次在运行"}, 409)
            batch = Batch(mode, [
                {"idx": i, "cdk": it["cdk"].strip(), "org": it["org"].strip(),
                 "acct": str(it.get("acct") or "")[:80],
                 "state": "pending", "product": "", "orderNo": "", "message": "等待中"}
                for i, it in enumerate(items)
            ], site=site, host=host, product_id=product_id)
            JOB_QUEUE.put(batch)
            return self._json({"ok": True})
        if self.path == "/api/stop":
            cur = CURRENT["batch"]
            if cur:
                cur.stop_requested = True
            return self._json({"ok": True})
        if self.path == "/api/mail/login":
            req = self._body()
            rt = (req.get("refresh_token") or "").strip()
            if not rt:
                return self._json({"error": "缺少 refresh_token"}, 400)
            return self._json(mail_login(req.get("client_id"), rt))
        if self.path == "/api/mail/messages":
            req = self._body()
            at = (req.get("access_token") or "").strip()
            if not at:
                return self._json({"error": "缺少 access_token"}, 400)
            folder = FOLDERS.get(req.get("folder") or "inbox", "inbox")
            top = min(int(req.get("top") or 20), 50)
            sel = ("$select=Id,Subject,From,ReceivedDateTime,"
                   "BodyPreview,IsRead,HasAttachments")
            return self._json(mail_get(
                at, f"/me/MailFolders/{folder}/messages?$top={top}&{sel}"))
        if self.path == "/api/mail/message":
            req = self._body()
            at = (req.get("access_token") or "").strip()
            mid = req.get("id") or ""
            if not at or not mid:
                return self._json({"error": "缺少 access_token 或 id"}, 400)
            return self._json(mail_get(
                at, "/me/messages/" + urllib.parse.quote(mid, safe="")))
        self.send_error(404)


def main():
    global FORCE_HEADED
    ap = argparse.ArgumentParser(description="ipiap 批量充值 WebUI")
    ap.add_argument("--port", type=int, default=8321)
    ap.add_argument("--headed", action="store_true", help="强制有头浏览器")
    args = ap.parse_args()
    FORCE_HEADED = args.headed

    threading.Thread(target=worker, daemon=True).start()

    srv = http.server.ThreadingHTTPServer(("127.0.0.1", args.port), Handler)
    print(f"WebUI: http://127.0.0.1:{args.port}  (Ctrl-C 退出)")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        BROWSER.stop()


if __name__ == "__main__":
    main()
