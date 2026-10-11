#!/usr/bin/env python3
"""
ipiap.com 卡密充值自动化工具

原理 (逆向自 https://ipiap.com 前端抓包, 全部接口无鉴权):
  step1  POST /app/recharge/step1  {serialNumber}                —— 校验卡密, 返回产品信息
  step2  POST /app/recharge/step2  {organizationId}              —— 校验目标组织
  step3  POST /app/recharge/step3  {organizationId,serialNumber} —— 下单, 返回 orderNo (消耗卡密!)
  step4  POST /app/recharge/step4  {orderNo}                     —— 轮询订单状态
         orderStatus: 1=处理中, 2=成功, 其他=失败(看 reason)

站点有 Cloudflare JS 挑战, TLS 指纹不过会被 403, 所以用真实 Chrome
(持久化 profile, 本目录 .chrome-profile) 打开站点后, 在页面上下文里
发 fetch —— 浏览器自己过 CF, 请求带全天然指纹, 无需 cf_clearance。

依赖: playwright (本目录 .venv 已装; uv pip install playwright)
用法:
  ./recharge.py --cdk XXXX-XXXX-XXXX-XXXX --org <organization-uuid>
  ./recharge.py --cdk XXX --org YYY --check        # 只做 step1/step2 校验, 不下单
  ./recharge.py --cdk XXX --org YYY --headed       # 显示浏览器窗口 (CF 难过时调试用)
"""

import argparse
import json
import os
import sys
import time

PROFILE_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           ".chrome-profile")
SITE = "https://ipiap.com/"
# cf_clearance 跟 UA 绑定, headless 的 HeadlessChrome UA 会让 cookie 失效,
# 统一成正常 Chrome UA (与本地 Chrome 154 一致)
UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
      "AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/154.0.0.0 Safari/537.36")


def log(msg):
    print(msg, flush=True)


def api(page, step, payload):
    """在页面上下文里 POST 一个 recharge 步骤, 返回 {status, body}。
    SPA 还在跳转时 evaluate 会撞 navigation, 重试几次。"""
    for attempt in range(5):
        try:
            return page.evaluate(
                """async ([step, payload]) => {
                    const res = await fetch('/app/recharge/' + step, {
                        method: 'POST',
                        headers: {'content-type': 'application/json',
                                  'x-language': 'zh'},
                        body: JSON.stringify(payload),
                    });
                    return {status: res.status, body: await res.text()};
                }""",
                [step, payload])
        except Exception as e:
            if "Execution context" not in str(e) or attempt == 4:
                raise
            time.sleep(1)


def post(page, step, payload, tries=4):
    """api() 的业务封装: 校验 code==0, 返回 data。
    命中 "频繁/稍后" 限流时自动退避重试。"""
    delay = 5.0
    for i in range(tries):
        r = api(page, step, payload)
        if r["status"] != 200:
            sys.exit(f"[{step}] HTTP {r['status']}: {r['body'][:300]}")
        try:
            resp = json.loads(r["body"])
        except json.JSONDecodeError:
            sys.exit(f"[{step}] 响应不是 JSON: {r['body'][:300]}")
        if resp.get("code") == 0:
            return resp.get("data") or {}
        msg = str(resp.get("msg"))
        if ("频繁" in msg or "稍后" in msg) and i < tries - 1:
            print(f"[{step}] 限流, {int(delay)}s 后重试 ({i+1}/{tries-1}) ...",
                  flush=True)
            time.sleep(delay)
            delay += 5.0
            continue
        sys.exit(f"[{step}] 接口报错 code={resp.get('code')}: {msg}")


def open_site(pw, headless):
    """启动持久化 Chrome, 过 Cloudflare, 返回 (ctx, page)。"""
    ctx = pw.chromium.launch_persistent_context(
        PROFILE_DIR, channel="chrome", headless=headless,
        viewport={"width": 1280, "height": 900},
        user_agent=UA,
        args=["--disable-blink-features=AutomationControlled",
              "--window-position=40,40"],
    )
    page = ctx.pages[0] if ctx.pages else ctx.new_page()
    page.goto(SITE, wait_until="domcontentloaded", timeout=60000)
    for _ in range(90):
        t = page.title() or ""
        if t and "moment" not in t.lower() and "请稍候" not in t \
                and "challenge" not in t.lower():
            try:
                page.wait_for_load_state("load", timeout=15000)
            except Exception:
                pass
            # 标题正常不代表 fetch 能过 CF (headless UA 会拿 HTML 挑战),
            # 用一个假卡密探一下 step1: 200=上下文可用, 403=还要再过挑战
            try:
                probe = api(page, "step1",
                            {"serialNumber": "0000-0000-0000-0000"})
                if probe["status"] == 200:
                    return ctx, page
            except Exception:
                pass
            ctx.close()
            return None, None
        for f in page.frames:
            if "challenges.cloudflare.com" in f.url or "turnstile" in f.url:
                try:
                    f.locator("body").click(position={"x": 25, "y": 25},
                                            force=True, timeout=800)
                except Exception:
                    pass
        page.wait_for_timeout(1000)
    ctx.close()
    return None, None


def get_page(pw, force_headed):
    """先 headless 试, CF 不过再 headed (首次跑一般需要 headed)。"""
    from playwright.sync_api import TimeoutError as _  # noqa: F401
    if not force_headed:
        ctx, page = open_site(pw, headless=True)
        if page:
            log("[cf] headless 模式已通过 Cloudflare")
            return ctx, page
        log("[cf] headless 未过, 改用有头模式 ...")
    ctx, page = open_site(pw, headless=False)
    if not page:
        sys.exit("Cloudflare 90 秒内未放行, 可重试或人工点窗口里的验证框")
    return ctx, page


def main():
    ap = argparse.ArgumentParser(description="ipiap.com 卡密充值")
    ap.add_argument("--cdk", required=True, help="充值卡密 (serialNumber)")
    ap.add_argument("--org", required=True, help="Organization ID")
    ap.add_argument("--check", action="store_true",
                    help="只校验卡密和组织 (step1+step2), 不下单")
    ap.add_argument("--headed", action="store_true",
                    help="直接显示浏览器窗口 (跳过 headless 尝试)")
    ap.add_argument("--interval", type=float, default=3.0,
                    help="step4 轮询间隔秒数 (默认 3)")
    ap.add_argument("--timeout", type=float, default=300.0,
                    help="step4 最长等待秒数 (默认 300)")
    args = ap.parse_args()

    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        ctx, page = get_page(pw, args.headed)
        try:
            # step1: 校验卡密
            info = post(page, "step1", {"serialNumber": args.cdk})
            log(f"[step1] 卡密有效: {info.get('productName')} "
                f"(platform={info.get('platform')}, "
                f"到期 {info.get('expiredTime')})")

            # step2: 校验组织
            org = post(page, "step2", {"organizationId": args.org})
            log(f"[step2] 组织有效: platform={org.get('platform')}")

            if args.check:
                log("[check] 校验通过, 未下单。")
                return

            # step3: 下单
            order = post(page, "step3", {"organizationId": args.org,
                                         "serialNumber": args.cdk})
            order_no = order["orderNo"]
            log(f"[step3] 已下单: orderNo={order_no}")

            # step4: 轮询
            deadline = time.time() + args.timeout
            while True:
                st = post(page, "step4", {"orderNo": order_no})
                status = st.get("orderStatus")
                if status == 1:
                    if time.time() > deadline:
                        sys.exit(f"[step4] 超时 ({args.timeout}s), "
                                 f"订单仍在处理: {order_no}")
                    log(f"[step4] 处理中, {args.interval}s 后重试 ...")
                    time.sleep(args.interval)
                    continue
                if status == 2:
                    log(f"[step4] 充值成功: {order_no} "
                        f"({st.get('reason') or 'ok'})")
                    return
                sys.exit(f"[step4] 充值失败 status={status}: "
                         f"{st.get('reason')}")
        finally:
            ctx.close()


if __name__ == "__main__":
    main()
