# -*- coding: utf-8 -*-
"""[라이브] PF Compare 팝업 화면 기능 검증 - 89개국(91개 사이트코드)용.

카테고리별 PF 에서 제품 2~3개의 Compare 체크박스를 고르고 Compare 버튼 → 팝업
  1) 팝업 열기 : 팝업이 뜨고, 고른 제품이 팝업에 나오는지
  2) Key Specs : 'Key Specs' 섹션과 제품마다 값이 있는지
                 (고른 제품 중 하나라도 Key Specs 가 없으면 섹션 전체가 안 나오는 게 규칙 → '확인필요')
  3) 차이점    : 'Show me the differences' 켜기 → 같은 값 행 숨김 / 다른 값 행 노출 / 파란 글자 → 끄면 원래대로
  4) All Specs : 카테고리 버튼(첫·마지막)을 누르면 내용이 접히고, 다시 누르면 펼쳐지는지
  5) 앵커      : Key Specs 번호 → 하단 디스클레이머 이동, 화살표 → 번호 복귀 (팝업 안 스크롤로 판정, 기본 3개)
  6) 칩        : 제품마다 컬러칩 / 용량칩을 누르면 모델코드·이미지가 바뀌는지 (확인 후 원래 칩으로 되돌림)
  7) 닫기      : 팝업 X 버튼으로 닫히는지
  8) CTA       : Buy(기본은 링크 응답만) / Learn more → 제품 페이지 이동

사용 예
  python run_popup_ui_live.py --sites uk                # 한 나라
  python run_popup_ui_live.py --sites uk de fr pl
  python run_popup_ui_live.py --parallel 3 --headless   # 91개 사이트 전체, 3개 나라씩 동시에
  python run_popup_ui_live.py --resume                  # 마지막 실행 이어서
결과: compare/result/run_ui_live_<날짜시간>/PopupUI_LIVE_<날짜시간>.xlsx
      (사이트별 중간 결과 sites/, 실패 캡처 screenshots/)
"""
import argparse
import asyncio
import datetime as dt
import json
import random
import sys
from collections import Counter, OrderedDict
from pathlib import Path
from urllib.parse import urljoin

from openpyxl import Workbook
from openpyxl.styles import PatternFill
from playwright.async_api import async_playwright

from config import (
    BASE_URL,
    CATEGORY_SPECS,
    FILL_COLORS,
    HERE,
    MAPPING_LIST_FILE,
    TIMEOUTS,
)
from run_c_landing_live import (
    FILL,
    JS_LAYERS,
    JS_STATE,
    _sheet,
    check_all_specs,
    check_differences,
    dismiss_popups,
    load_live_sites,
    shot,
)
from run_pd_compare_live import PF_PAGES, check_chips

# ------------------------------------------------------------------ 브라우저 안에서 쓰는 JS
JS_PF_CHECKS = r"""() => [...document.querySelectorAll('input.js-pfv2-compare-cta')]
  .map((i, k) => ({k, checked: i.checked, disabled: i.disabled}))"""

JS_PF_SET = r"""(want) => {
  const ins = [...document.querySelectorAll('input.js-pfv2-compare-cta')];
  ins.forEach((i, k) => { if (i.checked && !want.includes(k)) i.click(); });
  want.forEach(k => { if (ins[k] && !ins[k].checked) ins[k].click(); });
  return ins.filter(i => i.checked).length;
}"""

JS_PF_OPEN = r"""() => {
  const b = document.querySelector('button.js-pfv2-open-compare-layer');
  if (!b) return null;
  const codes = b.getAttribute('data-modelcode') || '';
  b.click();
  return codes;
}"""

JS_POPUP = r"""() => {
  const vis = e => { if (!e) return false; const s = getComputedStyle(e), r = e.getBoundingClientRect();
                     return s.display !== 'none' && s.visibility !== 'hidden' && r.height > 100; };
  const r = document.querySelector('.pdd43-master-compare--popup');
  const layer = vis(document.getElementById('pdd43MasterCompare'));
  return {layer, root: !!r,
          show: !!(r && r.classList.contains('pdd43-master-compare--show') && layer),
          skus: r ? [...r.querySelectorAll('li.pdd43-master-compare__card-item .pdd43-master-compare__sku-text')]
                      .map(e => e.innerText.trim()) : []};
}"""

JS_POP_COLUMN = r"""(n) => {
  const root = document.querySelector('.pdd43-master-compare');
  const card = root && root.querySelector(`li.pdd43-master-compare__card-item.pdd43-master-compare__column--${n}`);
  if (!card) return null;
  const txt = s => { const e = card.querySelector(s); return e ? e.innerText.replace(/\s+/g, ' ').trim() : ''; };
  const buy = card.querySelector('.pdd43-master-compare__cta a.js-cta-buy-now');
  const lm = [...card.querySelectorAll('.pdd43-master-compare__cta a')].find(a => a !== buy && /cta--underline/.test(a.className));
  return {name: txt('.pdd43-master-compare__name-text') || txt('.menu__select-field-text'),
          sku: txt('.pdd43-master-compare__sku-text'),
          buy: buy ? {text: buy.innerText.trim(), link: buy.getAttribute('data-link_info') || buy.getAttribute('data-config_info') || ''} : null,
          learn: lm ? {text: lm.innerText.trim(), href: lm.getAttribute('href')} : null};
}"""

JS_SCROLL_POS = r"""() => {
  const l = document.getElementById('pdd43MasterCompare');
  if (l) {
    const ok = e => e && /(auto|scroll)/.test(getComputedStyle(e).overflowY) && e.scrollHeight > e.clientHeight + 5;
    let c = [...l.querySelectorAll('.scrollbar__wrap, .layer-popup__contents')].find(ok);
    if (!c) c = [...l.querySelectorAll('div')].find(ok);
    if (c) return Math.round(c.scrollTop);
  }
  return Math.round(window.scrollY);
}"""

JS_IN_VIEW_POP = r"""(sel) => {
  const els = [...document.querySelectorAll(sel)];
  const h = window.innerHeight;
  const hit = els.find(e => { const r = e.getBoundingClientRect(); return r.height > 0 && r.bottom > 0 && r.top < h && e.offsetParent; });
  const r = hit ? hit.getBoundingClientRect() : null;
  return {count: els.length, inView: !!hit, top: r ? Math.round(r.top) : null};
}"""


async def settle(page, min_ms, max_ms=6000, step=300):
    """번호/화살표 클릭 후 팝업 안 스크롤이 멈출 때까지 대기."""
    await page.wait_for_timeout(min_ms)
    last, same, waited = None, 0, min_ms
    while waited < max_ms:
        y = await page.evaluate(JS_SCROLL_POS)
        same = same + 1 if y == last else 0
        if same >= 2:
            return
        last = y
        await page.wait_for_timeout(step)
        waited += step


# ------------------------------------------------------------------ 팝업 띄우기
async def open_popup(page, pf_url, want, opts):
    """PF 를 열고 want(체크박스 순번)를 체크한 뒤 Compare 버튼 → 팝업."""
    resp = await page.goto(pf_url, wait_until="domcontentloaded", timeout=TIMEOUTS["page_load"])
    await page.wait_for_timeout(opts["wait"] * 1000)
    await dismiss_popups(page)
    if not resp or resp.status >= 400:
        return {"status": "NA", "reason": f"PF 없음 (HTTP {resp.status if resp else ''})"}
    checks = await page.evaluate(JS_PF_CHECKS)
    if not checks:
        return {"status": "NA", "reason": "PF 에 Compare 체크박스가 없음"}
    if want is None:
        avail = [c["k"] for c in checks if not c["disabled"]]
        if len(avail) < 2:
            return {"status": "SKIP", "reason": f"비교할 수 있는 제품이 {len(avail)}개뿐"}
        n = min(opts["select"], len(avail))
        want = sorted(random.Random(opts["seed"]).sample(avail, n)) if opts["seed"] is not None else avail[:n]
    await page.evaluate(JS_PF_SET, want)
    await page.wait_for_timeout(2000)
    codes = await page.evaluate(JS_PF_OPEN)
    if codes is None:
        return {"status": "FAIL", "reason": "체크 후 Compare 버튼(js-pfv2-open-compare-layer)이 없음", "want": want}
    try:
        await page.wait_for_function("""() => {
            const r = document.querySelector('.pdd43-master-compare--popup.pdd43-master-compare--show');
            const l = document.getElementById('pdd43MasterCompare');
            return !!(r && l && l.getBoundingClientRect().height > 100); }""", timeout=30000)
    except Exception:
        pass
    await page.wait_for_timeout(opts["popup_wait"] * 1000)
    pop = await page.evaluate(JS_POPUP)
    return {"status": "OK", "want": want, "codes": [c for c in codes.split(",") if c], "popup": pop}


async def ensure_open(page, pf_url, want, opts):
    pop = await page.evaluate(JS_POPUP)
    if pop["show"]:
        return True
    r = await open_popup(page, pf_url, want, opts)
    return r["status"] == "OK" and r["popup"]["show"]


# ------------------------------------------------------------------ 앵커 (팝업 안 스크롤)
async def check_popup_anchors(page, add, snap, nums, anchor_wait):
    bottom = await page.evaluate("""() => [...document.querySelectorAll('.pdd43-master-compare li[id^=masterCompareBottomDisclaimer]')]
        .map(e => +e.id.replace('masterCompareBottomDisclaimer', ''))""")
    missing = [n for n in nums if n not in bottom]
    if missing:
        add("앵커: 하단 디스클레이머 존재", "FAIL", f"하단에 없는 번호: {missing}", expected=str(nums), actual=str(bottom))
    for n in nums:
        num_sel = f".pdd43-master-compare a#masterCompareBottomDisclaimerGoBack{n}"
        disc_sel = f".pdd43-master-compare #masterCompareBottomDisclaimer{n}"
        num = page.locator(num_sel).locator("visible=true").first
        try:
            await num.scroll_into_view_if_needed(timeout=5000)
            await page.wait_for_timeout(500)
            before = await page.evaluate(JS_SCROLL_POS)
            await num.click(timeout=5000)
        except Exception as e:
            add(f"앵커 {n}: 번호 → 디스클레이머", "FAIL", f"번호 클릭 실패: {type(e).__name__}", screenshot=await snap(f"n{n}_click"))
            continue
        await settle(page, anchor_wait)
        v, pos = await page.evaluate(JS_IN_VIEW_POP, disc_sel), await page.evaluate(JS_SCROLL_POS)
        ok = v["inView"] and pos > before
        add(f"앵커 {n}: 번호 → 디스클레이머", "PASS" if ok else "FAIL", expected=f"하단 디스클레이머 {n} 이 팝업 화면에 보임",
            actual=f"보임={v['inView']}, 위치 top={v['top']}, 팝업 스크롤 {before}→{pos}",
            screenshot="" if ok else await snap(f"n{n}_down"))
        arrow = page.locator(f"{disc_sel} a[href='#masterCompareBottomDisclaimerGoBack{n}']").first
        if not await arrow.count():
            add(f"앵커 {n}: 화살표 → 번호", "FAIL", "디스클레이머에 화살표(Go to feature)가 없음", screenshot=await snap(f"n{n}_noarrow"))
            continue
        try:
            if not ok:
                await arrow.scroll_into_view_if_needed(timeout=5000)
                await page.wait_for_timeout(500)
            before2 = await page.evaluate(JS_SCROLL_POS)
            await arrow.click(timeout=5000)
        except Exception as e:
            add(f"앵커 {n}: 화살표 → 번호", "FAIL", f"화살표 클릭 실패: {type(e).__name__}", screenshot=await snap(f"n{n}_arrow"))
            continue
        await settle(page, anchor_wait)
        v2, pos2 = await page.evaluate(JS_IN_VIEW_POP, num_sel), await page.evaluate(JS_SCROLL_POS)
        ok2 = v2["inView"] and pos2 < before2
        add(f"앵커 {n}: 화살표 → 번호", "PASS" if ok2 else "FAIL", expected=f"상단 번호 {n} 이 팝업 화면에 보임",
            actual=f"보임={v2['inView']}, 위치 top={v2['top']}, 팝업 스크롤 {before2}→{pos2}",
            screenshot="" if ok2 else await snap(f"n{n}_up"))


# ------------------------------------------------------------------ 닫기
async def check_close(page, add, snap):
    btn = page.locator("#pdd43MasterCompare button.layer-popup__close").first
    if not await btn.count():
        add("닫기: X 버튼", "FAIL", "팝업 닫기 버튼이 없음", screenshot=await snap("close_nobtn"))
        return
    await btn.click(timeout=5000)
    await page.wait_for_timeout(1500)
    pop = await page.evaluate(JS_POPUP)
    ok = not pop["layer"] and not pop["show"]
    add("닫기: X 버튼", "PASS" if ok else "FAIL", expected="팝업이 닫힘",
        actual="닫힘" if ok else f"레이어 보임={pop['layer']}, 팝업 표시={pop['show']}",
        screenshot="" if ok else await snap("close"))


# ------------------------------------------------------------------ Buy / Learn more
def clicks_learn(n, opts):
    return not opts["learn_click_cols"] or n in opts["learn_click_cols"]


async def http_check(page, link):
    r = await page.context.request.get(urljoin(page.url, link), timeout=30000, max_redirects=5)
    return r.status, r.url


async def check_popup_links(page, add, snap, cols, opts, info):
    for n in cols:
        c = await page.evaluate(JS_POP_COLUMN, n)
        if not c or not c["sku"]:
            continue
        info[n] = c
        who = f"{n}열 {c['name']}"
        if not c["buy"]:
            add(f"CTA: Buy ({who})", "확인불가", "Buy 버튼이 없음 (판매 안 하는 모델일 수 있음)")
        elif not opts["click_buy"]:
            link = c["buy"]["link"]
            if not link:
                add(f"CTA: Buy ({who})", "확인필요", f"버튼 '{c['buy']['text']}' 에 구매 링크 정보가 없음 (클릭해야 확인 가능)")
            else:
                try:
                    st, url = await http_check(page, link)
                    add(f"CTA: Buy ({who})", "PASS" if st < 400 else "FAIL",
                        f"버튼 '{c['buy']['text']}' (클릭하지 않고 링크만 확인)", expected=f"구매 링크 {link} 응답",
                        actual=f"HTTP {st} {url}")
                except Exception as e:
                    add(f"CTA: Buy ({who})", "FAIL", f"구매 링크 확인 실패: {type(e).__name__}", expected=link)
        if not c["learn"]:
            add(f"CTA: Learn more ({who})", "FAIL", "Learn more 버튼이 없음", screenshot=await snap(f"cta_nolearn{n}"))
        elif not clicks_learn(n, opts):
            href = c["learn"]["href"]
            try:
                st, url = await http_check(page, href)
                add(f"CTA: Learn more ({who})", "PASS" if st < 400 else "FAIL",
                    f"클릭하지 않고 링크만 확인 (실제 클릭은 {', '.join(map(str, opts['learn_click_cols']))}열)",
                    expected=f"{href} 응답", actual=f"HTTP {st} {url}")
            except Exception as e:
                add(f"CTA: Learn more ({who})", "FAIL", f"링크 확인 실패: {type(e).__name__}", expected=href)


async def check_popup_clicks(page, add, snap, pf_url, want, opts, info):
    async def fresh(n):
        return await page.evaluate(JS_POP_COLUMN, n) or info[n]

    for n, c in info.items():
        who = f"{n}열 {c['name']}"
        card = f"#pdd43MasterCompare li.pdd43-master-compare__card-item.pdd43-master-compare__column--{n}"
        if c["buy"] and opts["click_buy"]:
            if not await ensure_open(page, pf_url, want, opts):
                add(f"CTA: Buy ({who})", "오류", "Buy 확인용으로 팝업을 다시 띄우지 못함", screenshot=await snap(f"cta_reopen{n}"))
                return
            c = await fresh(n)
            who = f"{n}열 {c['name']}"
            buy = page.locator(f"{card} .pdd43-master-compare__cta a.js-cta-buy-now").first
            url0, layers0, pages0 = page.url, await page.evaluate(JS_LAYERS), len(page.context.pages)
            try:
                await buy.scroll_into_view_if_needed()
                await buy.click()
                await page.wait_for_timeout(6000)
                new_page = len(page.context.pages) > pages0
                moved = page.url != url0
                layer = (await page.evaluate(JS_LAYERS)) > layers0 if not moved else False
                link = c["buy"]["link"]
                if moved:
                    ok = not link or link.rstrip("/").split("?")[0] in page.url
                    add(f"CTA: Buy ({who})", "PASS" if ok else "FAIL", f"버튼 '{c['buy']['text']}'",
                        expected=f"구매 페이지 {link} 로 이동", actual=page.url, screenshot="" if ok else await snap(f"cta_buy{n}"))
                elif new_page or layer:
                    add(f"CTA: Buy ({who})", "PASS", f"버튼 '{c['buy']['text']}' - {'새 창' if new_page else '구매 레이어'} 노출",
                        expected=f"구매 페이지 {link} 이동 또는 구매 레이어", actual="새 창 열림" if new_page else "레이어 노출",
                        screenshot=await snap(f"cta_buy{n}_layer"))
                    for extra in page.context.pages[pages0:]:
                        await extra.close()
                else:
                    add(f"CTA: Buy ({who})", "FAIL", f"버튼 '{c['buy']['text']}' 눌러도 반응 없음",
                        expected=f"구매 페이지 {link} 이동 또는 구매 레이어", actual="이동/레이어 없음", screenshot=await snap(f"cta_buy{n}"))
            except Exception as e:
                add(f"CTA: Buy ({who})", "FAIL", f"클릭 실패: {type(e).__name__}", screenshot=await snap(f"cta_buy{n}"))
        if c["learn"] and clicks_learn(n, opts):
            if not await ensure_open(page, pf_url, want, opts):
                add(f"CTA: Learn more ({who})", "오류", "Learn more 확인용으로 팝업을 다시 띄우지 못함",
                    screenshot=await snap(f"cta_reopen{n}"))
                return
            c = await fresh(n)
            who = f"{n}열 {c['name']}"
            if not c["learn"]:
                add(f"CTA: Learn more ({who})", "FAIL", "Learn more 버튼이 없음", screenshot=await snap(f"cta_nolearn{n}"))
                continue
            lm = page.locator(f"{card} .pdd43-master-compare__cta a[href='{c['learn']['href']}']").first
            target = c["learn"]["href"].split("#")[0].rstrip("/")
            url0 = page.url
            try:
                await lm.scroll_into_view_if_needed()
                await lm.click()
                await page.wait_for_url(lambda u: u.split("#")[0].rstrip("/").endswith(target) and u != url0,
                                        timeout=30000, wait_until="commit")
            except Exception:
                pass
            await page.wait_for_timeout(1500)
            ok = page.url.split("#")[0].rstrip("/").endswith(target)
            add(f"CTA: Learn more ({who})", "PASS" if ok else "FAIL", "버튼 클릭" if ok else "제품 페이지로 이동 안 됨",
                expected=f"{c['learn']['href']} 로 이동", actual=page.url, screenshot="" if ok else await snap(f"cta_learn{n}"))


# ------------------------------------------------------------------ 카테고리 1개 검증
async def check_category(ctx, host, site, spec, opts, shots_dir):
    group, label, pf_path = spec
    rows = []
    tag = f"{site}_{pf_path.split('/')[0]}"
    models = {"v": ""}

    def add(check, result, detail="", expected="", actual="", screenshot=""):
        rows.append({"site": site, "group": group, "page": label, "models": models["v"], "check": check,
                     "expected": expected, "actual": actual, "result": result, "detail": detail, "screenshot": screenshot})

    async def snap(name):
        return await shot(page, shots_dir, f"{tag}_{name}")

    pf_url = f"{host}/{site}/{pf_path}"
    page = await ctx.new_page()
    try:
        r = await open_popup(page, pf_url, None, opts)
        if r["status"] != "OK":
            res = {"NA": "해당 없음", "SKIP": "확인불가"}.get(r["status"], r["status"])
            add("팝업 열기", res, r["reason"], actual=pf_url, screenshot=await snap("open") if res == "FAIL" else "")
            return rows
        want, pop = r["want"], r["popup"]
        models["v"] = ", ".join(pop["skus"] or r["codes"])
        if not pop["root"]:
            layers = await page.evaluate(JS_LAYERS)
            add("팝업 열기", "FAIL", "새 Compare 팝업(.pdd43-master-compare--popup)이 없음" +
                (" - 다른 레이어는 떠 있음 (기존 Compare 팝업일 수 있음)" if layers else " - 팝업이 안 뜸"),
                expected="새 Compare 팝업", actual=f"고른 제품 {', '.join(r['codes'])}", screenshot=await snap("nopopup"))
            return rows
        same = not r["codes"] or sorted(r["codes"]) == sorted(pop["skus"])
        ok = pop["show"] and len(pop["skus"]) >= 2 and same
        add("팝업 열기", "PASS" if ok else "FAIL", f"PF 체크박스 {', '.join(str(k + 1) for k in want)}번째 제품",
            expected=f"팝업 표시 + 고른 제품 {', '.join(r['codes'])}",
            actual=f"표시={pop['show']}, 팝업 제품 {', '.join(pop['skus'])}",
            screenshot=await snap("popup") if (not ok or opts["shots_all"]) else "")

        # Key Specs
        st = await page.evaluate(JS_STATE)
        if not st["featVisible"] or not st["cols"]:
            add("Key Specs 노출", "확인필요", "고른 제품 중 Key Specs 없는 제품이 있으면 섹션 전체가 안 나오는 게 정상 "
                "→ run_popup_keyspec_live.py 결과로 제품별 Key Specs 확인", expected="Key Specs 섹션 노출",
                actual="섹션 없음", screenshot=await snap("keyspecs"))
        else:
            empty = [c for c, v in st["cols"].items() if v["filled"] == 0]
            detail = ", ".join(f"{c}열 값 {v['filled']}/{v['filled'] + v['empty']}" for c, v in sorted(st["cols"].items()))
            add("Key Specs 노출", "PASS" if not empty else "FAIL", expected="제품마다 Key Specs 값 있음",
                actual=f"'{st['title']}' 항목 {st['rows']}개 / {detail}",
                screenshot=await snap("keyspecs") if (empty or opts["shots_all"]) else "")

        sections = []
        if opts["diff"]:
            sections.append(("차이점", lambda: check_differences(page, add, snap, opts["shots_all"])))
        if opts["allspecs"]:
            sections.append(("All Specs", lambda: check_all_specs(page, add, snap)))
        if opts["anchor"]:
            nums = st["nums"][: opts["max_anchors"]] if opts["max_anchors"] else st["nums"]
            if not nums:
                add("앵커: 번호 → 디스클레이머", "확인불가", "Key Specs 에 디스클레이머 번호가 없음")
            else:
                sections.append(("앵커", lambda: check_popup_anchors(page, add, snap, nums, opts["anchor_wait"])))
        cols = sorted(int(c) for c in st["cols"] if str(c).isdigit()) or list(range(1, len(pop["skus"]) + 1))
        if opts["chips"]:
            chip_cols = [c for c in opts["chip_cols"] if c in cols] or cols

            async def all_chips():
                for c in chip_cols:
                    await check_chips(page, add, snap, c, opts["chip_wait"])
            sections.append(("칩", all_chips))
        cta_info = {}
        if opts["cta"]:
            sections.append(("CTA", lambda: check_popup_links(page, add, snap, opts["cta_cols"] or cols, opts, cta_info)))
        if opts["close"]:
            sections.append(("닫기", lambda: check_close(page, add, snap)))
        if opts["cta"]:
            sections.append(("CTA", lambda: check_popup_clicks(page, add, snap, pf_url, want, opts, cta_info)))
        for name, run in sections:
            try:
                await run()
            except Exception as e:
                add(f"{name}: 오류", "오류", f"{type(e).__name__}: {e}", screenshot=await snap(f"{name}_error"))
    except Exception as e:
        add("팝업 오류", "오류", f"{type(e).__name__}: {e}", screenshot=await snap("error"))
    finally:
        await page.close()
    return rows


# ------------------------------------------------------------------ 리포트
SUMMARY_HDR = ["사이트", "그룹", "카테고리", "비교 제품", "팝업 열기", "Key Specs", "차이점 보기", "All Specs",
               "번호→디스클레이머", "화살표→번호", "컬러칩", "용량칩", "닫기", "Buy", "Learn more", "FAIL 수", "판정"]
SUMMARY_W = [8, 7, 15, 40, 12, 12, 14, 14, 14, 12, 16, 16, 8, 14, 14, 8, 8]


def summary_rows(rows):
    by = OrderedDict()
    for r in rows:
        by.setdefault((r["site"], r["group"], r["page"]), []).append(r)
    out = []
    for (site, group, label), rs in by.items():
        def res(prefix):
            sub = [x for x in rs if x["check"].startswith(prefix)]
            if not sub:
                return "-"
            c = Counter(x["result"] for x in sub)
            return "PASS" if set(c) == {"PASS"} else ", ".join(f"{k} {v}" for k, v in c.items())
        down = [x for x in rs if "번호 → 디스클레이머" in x["check"] and x["result"] != "확인불가"]
        up = [x for x in rs if "화살표 → 번호" in x["check"]]
        fails = sum(1 for x in rs if x["result"] in ("FAIL", "오류"))
        na = all(x["result"] == "해당 없음" for x in rs)
        out.append([site, group, label, rs[-1]["models"], res("팝업 열기"), res("Key Specs"), res("차이점"), res("All Specs"),
                    f"{sum(x['result'] == 'PASS' for x in down)}/{len(down)}" if down else "번호 없음",
                    f"{sum(x['result'] == 'PASS' for x in up)}/{len(up)}" if up else "-",
                    res("컬러칩"), res("용량칩"), res("닫기"), res("CTA: Buy"), res("CTA: Learn more"),
                    fails, "FAIL" if fails else ("해당 없음" if na else "PASS")])
    return out


def run_info(args):
    return [["실행 시각", dt.datetime.now().strftime("%Y-%m-%d %H:%M")], ["환경", f"live ({BASE_URL})"],
            ["사이트", " ".join(args.sites)], ["카테고리", ", ".join(p[1] for p in select_pf(args))],
            ["제품 선택", f"PF Compare 체크박스 {args.select}개 - " +
             (f"임의 (seed {args.seed})" if args.seed is not None else "앞에서부터")],
            ["검증", "팝업 열기, Key Specs" + ("" if args.no_diff else ", 차이점 보기") +
             ("" if args.no_allspecs else ", All Specs 접기/펼치기") +
             ("" if args.no_anchor else f", 번호/화살표(최대 {args.max_anchors or '전부'}개)") +
             ("" if args.no_chips else f", 컬러칩/용량칩({' '.join(map(str, args.chip_cols)) + '열' if args.chip_cols else '모든 열'})") +
             ("" if args.no_close else ", 닫기") +
             ("" if args.no_cta else f", Buy/Learn more({' '.join(map(str, args.cta_cols)) + '열' if args.cta_cols else '모든 열'})")],
            ["Learn more", "모든 열 클릭" if not args.learn_click_cols else
             f"{' '.join(map(str, args.learn_click_cols))}열만 클릭, 나머지 열은 클릭 없이 링크 응답만 확인"],
            ["대기", f"PF {args.wait}초, 팝업 {args.popup_wait}초, 동작 사이 {args.slow_mo}ms"],
            ["Buy", "클릭 후 이동/레이어 확인" if args.click_buy else "클릭하지 않고 구매 링크 응답만 확인"]]


def write_report(path, rows, args):
    wb = Workbook()
    ws = wb.active
    ws.title = "국가별 요약"
    cats = [p[1] for p in select_pf(args)]
    by_site = OrderedDict()
    for r in rows:
        by_site.setdefault(r["site"], OrderedDict()).setdefault(r["page"], []).append(r)
    srows = []
    for site, pages in by_site.items():
        vs = []
        for cn in cats:
            rs = pages.get(cn, [])
            if not rs:
                vs.append("-")
            elif all(x["result"] == "해당 없음" for x in rs):
                vs.append("해당 없음")
            elif any(x["result"] in ("FAIL", "오류") for x in rs):
                vs.append("FAIL")
            else:
                vs.append("PASS")
        fails = sum(1 for rs in pages.values() for x in rs if x["result"] in ("FAIL", "오류"))
        srows.append([site, *vs, fails, "FAIL" if fails else ("해당 없음" if set(vs) <= {"해당 없음", "-"} else "PASS")])
    _sheet(ws, ["사이트", *cats, "FAIL 수", "국가 판정"], srows, [9] + [13] * len(cats) + [8, 10])
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=len(cats) + 3):
        for c in row:
            color = FILL.get(str(c.value))
            if color:
                c.fill = PatternFill("solid", fgColor=color)
    ws = wb.create_sheet("카테고리별 요약")
    _sheet(ws, SUMMARY_HDR, summary_rows(rows), SUMMARY_W, result_col=16)
    ws = wb.create_sheet("상세")
    _sheet(ws, ["사이트", "그룹", "카테고리", "비교 제품", "검증", "기대", "실제", "결과", "비고", "캡처"],
           [[r["site"], r["group"], r["page"], r["models"], r["check"], r["expected"], r["actual"], r["result"],
             r["detail"], r["screenshot"]] for r in rows], [8, 7, 15, 30, 28, 30, 50, 9, 40, 40], result_col=7)
    ws = wb.create_sheet("실행 정보")
    info = run_info(args)
    info[2] = ["사이트", f"{len(args.sites)}개: " + " ".join(args.sites)]
    for r in info + [["동시 실행", f"사이트 {args.parallel}개 × 사이트마다 카테고리 탭 {args.tabs}개"],
                     ["해당 없음", "그 나라에 PF 가 없거나 PF 에 Compare 체크박스가 없는 경우 (FAIL 아님)"]]:
        ws.append(r)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 120
    wb.save(path)


def select_pf(args):
    ps = PF_PAGES
    if args.groups:
        ps = [p for p in ps if p[0].lower() in {g.lower() for g in args.groups}]
    if args.categories:
        ps = [p for p in ps if p[1].lower() in {c.lower() for c in args.categories}]
    return ps


def make_opts(args):
    return {"wait": args.wait, "popup_wait": args.popup_wait, "anchor_wait": args.anchor_wait,
            "max_anchors": args.max_anchors, "chip_wait": args.chip_wait, "select": args.select, "seed": args.seed,
            "shots_all": args.shots_all, "diff": not args.no_diff, "allspecs": not args.no_allspecs,
            "anchor": not args.no_anchor, "chips": not args.no_chips, "close": not args.no_close,
            "cta": not args.no_cta, "cta_cols": args.cta_cols, "learn_click_cols": args.learn_click_cols,
            "chip_cols": args.chip_cols,
            "click_buy": args.click_buy,
            "no_checkbox": "NA"}


# ------------------------------------------------------------------ main
async def run_site(browser, args, site, cats, opts, run_dir):
    """사이트 1개: 새 브라우저 컨텍스트에서 모든 카테고리 검증 → sites/<site>.json 저장."""
    ctx = await browser.new_context(viewport={"width": args.width, "height": args.height})
    shots_dir = run_dir / "screenshots"
    sem = asyncio.Semaphore(max(1, args.tabs))

    async def one(spec):
        async with sem:
            t0 = dt.datetime.now()
            r = await check_category(ctx, BASE_URL, site, spec, opts, shots_dir)
            f = sum(1 for x in r if x["result"] in ("FAIL", "오류"))
            na = bool(r) and all(x["result"] == "해당 없음" for x in r)
            print(f"  [{site}] {spec[1]}: " + ("해당 없음" if na else f"검증 {len(r)}건, FAIL/오류 {f}건")
                  + f" ({(dt.datetime.now() - t0).seconds}초)")
            return r

    try:
        rows = [x for r in await asyncio.gather(*(one(s) for s in cats)) for x in r]
    finally:
        await ctx.close()
    (run_dir / "sites" / f"{site}.json").write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")
    return rows


async def amain(args):
    out_root = Path(args.out_dir)
    if args.resume:
        runs = sorted(out_root.glob("run_ui_live_*")) if args.resume == "latest" else [Path(args.resume)]
        if not runs or not runs[-1].exists():
            sys.exit("이어서 실행할 이전 결과 폴더가 없습니다.")
        run_dir = runs[-1]
    else:
        run_dir = out_root / f"run_ui_live_{dt.datetime.now():%Y%m%d_%H%M%S}"
    (run_dir / "sites").mkdir(parents=True, exist_ok=True)
    if not args.sites:
        args.sites = load_live_sites(args.mapping_list)
    args.sites = [s for s in args.sites if s not in set(args.exclude_sites)]
    cats = select_pf(args)
    opts = make_opts(args)
    todo = [s for s in args.sites if not (run_dir / "sites" / f"{s}.json").exists()]
    print(f"사이트 {len(args.sites)}개 (이번에 실행 {len(todo)}개, 동시 {args.parallel}개), 카테고리 {len(cats)}개"
          f" (사이트마다 탭 {args.tabs}개 동시) → {run_dir}")
    print("  " + ("전부 확인 (--full)" if args.full else
                  f"빠른 확인: Learn more {' '.join(map(str, args.learn_click_cols))}열만 클릭, "
                  f"칩 {' '.join(map(str, args.chip_cols))}열만, 대기 PF {args.wait}초·팝업 {args.popup_wait}초"))
    sem = asyncio.Semaphore(max(1, args.parallel))
    async with async_playwright() as p:
        b = await p.chromium.launch(headless=args.headless, slow_mo=args.slow_mo)

        async def one(site):
            async with sem:
                t0 = dt.datetime.now()
                print(f"[{site}] 시작")
                try:
                    await run_site(b, args, site, cats, opts, run_dir)
                except SystemExit:
                    raise
                except Exception as e:
                    print(f"[{site}] 오류로 중단: {type(e).__name__}: {e}  (--resume 으로 다시 실행하면 이 사이트부터)")
                    return
                print(f"[{site}] 완료 ({(dt.datetime.now() - t0).seconds}초)")

        await asyncio.gather(*(one(s) for s in todo))
        await b.close()
    rows = []
    for s in args.sites:
        f = run_dir / "sites" / f"{s}.json"
        if f.exists():
            rows += json.loads(f.read_text(encoding="utf-8"))
    out = run_dir / f"PopupUI_LIVE_{run_dir.name.split('_', 3)[3]}.xlsx"
    write_report(out, rows, args)
    print("결과 파일:", out)


def add_common_args(ap):
    ap.add_argument("--groups", nargs="*", default=[], help="MX / PC / TV / Audio / DA 중 선택")
    ap.add_argument("--categories", nargs="*", default=[], help="카테고리 이름: " + ", ".join(p[1] for p in PF_PAGES))
    ap.add_argument("--select", type=int, default=3, choices=[2, 3], help="팝업에 넣을 제품 수 (기본 3)")
    ap.add_argument("--seed", type=int, default=None, help="제품을 임의로 고를 때 쓰는 숫자 (생략하면 PF 앞에서부터)")
    ap.add_argument("--max-anchors", type=int, default=3, help="확인할 디스클레이머 번호 수 (기본 3, 0 = 전부)")
    ap.add_argument("--wait", type=int, default=TIMEOUTS["pf_wait"], help="PF 로딩 대기(초)")
    ap.add_argument("--popup-wait", type=int, default=TIMEOUTS["popup_wait"], help="팝업이 뜬 뒤 데이터 로딩 대기(초)")
    ap.add_argument("--anchor-wait", type=int, default=TIMEOUTS["anchor_wait"], help="번호/화살표 클릭 후 대기(ms)")
    ap.add_argument("--chip-wait", type=int, default=TIMEOUTS["chip_switch"], help="칩 클릭 후 대기(ms)")
    ap.add_argument("--no-diff", action="store_true", help="Show me the differences 검증 생략")
    ap.add_argument("--no-allspecs", action="store_true", help="All Specs 접기/펼치기 검증 생략")
    ap.add_argument("--no-anchor", action="store_true", help="번호/화살표 검증 생략")
    ap.add_argument("--no-chips", action="store_true", help="컬러칩/용량칩 검증 생략")
    ap.add_argument("--no-close", action="store_true", help="팝업 닫기 검증 생략")
    ap.add_argument("--no-cta", action="store_true", help="Buy / Learn more 검증 생략")
    ap.add_argument("--cta-cols", nargs="*", type=int, default=[], help="CTA 를 확인할 제품 열 (기본: 모든 열)")
    ap.add_argument("--learn-click-cols", nargs="*", type=int, default=[],
                    help="Learn more 를 실제로 클릭할 열 (기본: 1열). 나머지 열은 클릭 없이 링크 응답만 확인")
    ap.add_argument("--chip-cols", nargs="*", type=int, default=[], help="컬러칩/용량칩을 확인할 열 (기본: 1열)")
    ap.add_argument("--slow-mo", type=int, default=TIMEOUTS["slow_mo"], help="브라우저 동작 사이 지연(ms)")
    ap.add_argument("--click-buy", action="store_true", help="라이브에서도 Buy 를 실제로 클릭 (장바구니에 담김)")
    ap.add_argument("--shots-all", action="store_true", help="PASS 인 화면도 캡처")
    ap.add_argument("--headless", action="store_true", help="브라우저 창 없이 실행")
    ap.add_argument("--width", type=int, default=1600)
    ap.add_argument("--height", type=int, default=1000)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--sites", nargs="*", default=[], help="사이트코드 (생략하면 Live_기준 시트의 전체 사이트)")
    ap.add_argument("--exclude-sites", nargs="*", default=[], help="제외할 사이트코드")
    ap.add_argument("--mapping-list", default=MAPPING_LIST_FILE, help="사이트 목록을 읽을 매핑 리스트 파일")
    ap.add_argument("--parallel", type=int, default=2, help="동시에 돌릴 사이트 수 (기본 2)")
    ap.add_argument("--resume", nargs="?", const="latest", default=None, help="이전 실행 이어서 (폴더 지정 가능)")
    ap.add_argument("--tabs", type=int, default=3, help="한 나라 안에서 동시에 확인할 카테고리 수 (기본 3)")
    ap.add_argument("--full", action="store_true",
                    help="시간 단축 없이 전부 확인: Learn more·칩 모든 열, 대기 PF 10초·팝업 5초·지연 100ms")
    add_common_args(ap)
    ap.set_defaults(learn_click_cols=[1], chip_cols=[1], wait=TIMEOUTS["pf_wait"], popup_wait=TIMEOUTS["popup_wait"], slow_mo=TIMEOUTS["slow_mo"])
    ap.add_argument("--out-dir", default=HERE / "result")
    args = ap.parse_args()
    if args.full:
        args.learn_click_cols, args.chip_cols, args.wait, args.popup_wait, args.slow_mo = [], [], 10, 5, 100
    asyncio.run(amain(args))


if __name__ == "__main__":
    main()
