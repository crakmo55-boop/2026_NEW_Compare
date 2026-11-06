# -*- coding: utf-8 -*-
"""라이브 PF 화면에서 실제로 Compare 팝업을 띄워, 팝업이 호출하는 API 주소를 찾는다.

라이브날 Compare API 주소가 예상(LIVE_SEARCH)과 다를 때 쓴다. 라이브는 로그인이 필요 없다.
  python run_keyspec_live.py --discover
결과로 나온 base 주소를 --api-base 로 넘겨 다시 실행하면 된다.
"""
import asyncio
import re

PF_URL = "https://www.samsung.com/{site}/smartphones/all-smartphones/"


async def _discover(site, headless):
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        b = await p.chromium.launch(headless=headless)
        page = await (await b.new_context(viewport={"width": 1600, "height": 1000})).new_page()
        reqs = []
        page.on("request", lambda r: reqs.append(r.url) if r.resource_type in ("xhr", "fetch") else None)
        await page.goto(PF_URL.format(site=site), wait_until="domcontentloaded", timeout=90000)
        await page.wait_for_timeout(10000)
        try:                                            # 쿠키 배너 (필수만 허용)
            l = page.locator("#truste-consent-required").first
            if await l.is_visible(timeout=2000):
                await l.click()
        except Exception:
            pass
        # Compare 체크박스 2개 선택 (화면 클릭이 레이어에 막히는 경우가 있어 DOM 클릭)
        n = await page.evaluate("""() => {
            const ins = [...document.querySelectorAll('input.js-pfv2-compare-cta')].filter(i => !i.checked).slice(0, 2);
            ins.forEach(i => i.click());
            return ins.length; }""")
        await page.wait_for_timeout(2000)
        mark = len(reqs)
        clicked = await page.evaluate("""() => { const b = document.querySelector('button.js-pfv2-open-compare-layer');
            if (!b) return false; b.click(); return true; }""")
        await page.wait_for_timeout(10000)
        after = reqs[mark:]
        await b.close()
    hits = [u for u in after if re.search(r"compare", u, re.I)]
    return {"selected": n, "button": clicked, "compare_requests": hits, "all_requests": after}


def discover(site="uk", headless=True):
    r = asyncio.run(_discover(site, headless))
    print(f"[discover] Compare 체크 {r['selected']}개, Compare 버튼 클릭: {r['button']}")
    if not r["compare_requests"]:
        print("[discover] 팝업이 compare 관련 API 를 호출하지 않았습니다. (새 Compare 미배포이거나 화면 구조가 다름)")
        for u in r["all_requests"][:15]:
            print("   요청:", u[:200])
        return None
    for u in r["compare_requests"]:
        print("[discover] 팝업 요청:", u[:220])
    new = [u for u in r["compare_requests"] if "/master/compare" in u]
    if not new:
        # 2026-10-22 라이브: 기존 Compare 팝업이 .../product/spec/compare 를 호출 (새 Compare 아님)
        print("[discover] 새 Compare API(master/compare) 호출이 없습니다 → 라이브 팝업은 아직 기존 Compare 입니다."
              " 새 Compare 배포 후 다시 확인하세요.")
        return None
    base = new[0].split("?")[0].rsplit("/master/compare", 1)[0]
    print("[discover] 새 Compare API 확인. 사용할 --api-base:", base)
    return base
