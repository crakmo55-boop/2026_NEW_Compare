# -*- coding: utf-8 -*-
"""[라이브] PD(제품 상세) 페이지 Compare 영역 화면 검증 - 89개국(91개 사이트코드)용.

카테고리별 PF 에서 제품을 골라(임의 또는 지정) PD 에 들어가, 상단 탭의 'Compare' 를 눌렀을 때를 확인한다.
  1) 앵커     : 'Compare' 탭 → 하단 Compare 영역으로 이동해 화면에 보이는지
  2) Key Specs: Compare 영역에 고정으로 나오는 3개 제품마다 Key Specs 값이 있는지
  3) 차이점   : 'Show me the differences' 켜기 → 같은 값 행 숨김 / 다른 값 행 노출 / 파란 글자 → 끄면 원래대로
  3-1) All Specs: 카테고리 버튼(첫·마지막)을 누르면 내용이 접히고, 다시 누르면 펼쳐지는지
  4) 번호     : Key Specs 번호 클릭 → 하단 디스클레이머로 이동
  5) 화살표   : 디스클레이머 화살표 클릭 → 상단 번호로 복귀
  6) 칩       : 제품마다 컬러칩 / 용량칩을 눌렀을 때 모델코드와 제품 이미지가 바뀌는지
                (용량칩은 같은 색이면 이미지가 같을 수 있어 이미지가 그대로면 '확인필요')

사용 예
  python run_pd_compare_live.py --sites uk                    # 한 나라, 카테고리마다 임의 제품 1개
  python run_pd_compare_live.py --sites uk de fr --seed 7      # 여러 나라 (같은 seed 면 같은 순서로 고름)
  python run_pd_compare_live.py                                # 91개 사이트 전체 (--parallel 로 동시 실행)
  python run_pd_compare_live.py --resume                       # 마지막 실행 이어서
  python run_pd_compare_live.py --sites uk --pd-urls /uk/smartphones/galaxy-a/galaxy-a57-5g-awesome-navy-256gb-sm-a576bdbdeub/
결과: compare/result/run_live_<날짜시간>/PDCompare_LIVE_<날짜시간>.xlsx
      (사이트별 중간 결과 sites/, 실패 캡처 screenshots/, PD Compare 영역 HTML dom/)
"""
import argparse
import asyncio
import datetime as dt
import json
import random
import sys
from collections import Counter, OrderedDict
from pathlib import Path

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
    JS_IN_VIEW,
    JS_STATE,
    _sheet,
    check_all_specs,
    check_differences,
    dismiss_popups,
    load_live_sites,
    shot,
    wait_scroll_settle,
)

# (그룹, 화면 이름, PF 경로) - config의 CATEGORY_SPECS로부터 생성
PF_PAGES = [(c[0], c[2], c[3]) for c in CATEGORY_SPECS]

# PF 카드에서 PD 주소 모으기 (Learn more 링크, 없으면 카드 안 첫 제품 링크)
JS_PD_LINKS = r"""([site, cat]) => {
  const out = [];
  for (const card of document.querySelectorAll('.pd21-product-card__item')) {
    const as = [...card.querySelectorAll('a[href]')].map(a => ({a, h: a.getAttribute('href') || ''}))
        .filter(x => x.h.startsWith('/' + site + '/' + cat + '/') && !/\/buy\/|javascript|#/.test(x.h));
    const lm = as.find(x => /learn more/i.test(x.a.innerText || x.a.getAttribute('aria-label') || '')) || as[0];
    if (lm) out.push({href: lm.h, name: (card.querySelector('[class*=name]') || {}).innerText || ''});
  }
  return out.filter((v, i, s) => s.findIndex(x => x.href === v.href) === i);
}"""

# PD 상단 탭의 Compare 링크 - 주소(#pdd43MasterCompare)로 찾음 (현지어 사이트 대응)
JS_PD_COMPARE_TAB = r"""() => [...document.querySelectorAll('a[href*="MasterCompare"]')]
  .filter(e => e.offsetParent && !/Disclaimer/.test(e.getAttribute('href') || ''))
  .map(e => ({tag: e.tagName, href: e.getAttribute('href') || '', cls: e.className, text: (e.innerText || '').trim()}))"""

# Compare 탭 링크가 숨김 상태로만 있고 Compare 영역이 없는지
JS_PD_HIDDEN_TAB = r"""() => {
  const a = document.querySelector('a[href="#pdd43MasterCompare"]');
  if (!a || a.offsetParent) return '';
  return document.getElementById('pdd43MasterCompare') || document.querySelector('.pdd43-master-compare')
    ? 'Compare 탭이 숨김 상태 (영역은 있음)' : 'Compare 탭이 숨김 상태이고 Compare 영역도 없음 (PD Compare 미적용)';
}"""

# 제품 열 N 의 모델코드 / 대표 이미지 / 칩 목록
JS_CHIPS = r"""(n) => {
  const root = document.querySelector('.pdd43-master-compare');
  const card = root && root.querySelector(`li.pdd43-master-compare__card-item.pdd43-master-compare__column--${n}`);
  if (!card) return null;
  const img = card.querySelector('.pdd43-master-compare__image img.image__main') || card.querySelector('.pdd43-master-compare__image img');
  const sku = card.querySelector('.pdd43-master-compare__sku-text');
  const chips = [...card.querySelectorAll('button.js-compare-option-fmychip')].map((b, i) => ({
    i, color: /option-selector-v2__color/.test(b.className), model: b.getAttribute('data-modelcode') || '',
    label: (b.innerText || b.getAttribute('data-option-name') || '').replace(/\s+/g, ' ').trim(),
    disabled: b.disabled || /disabled/.test(b.className)}));
  return {sku: sku ? sku.innerText.trim() : '', img: img ? (img.currentSrc || img.getAttribute('src') || '') : '', chips};
}"""


async def collect_pds(ctx, host, site, args):
    """카테고리별 PF 에서 PD 주소를 고른다. 반환: [(그룹, 화면 이름, PD 주소)]"""
    if args.pd_urls:
        return [("지정", "지정 PD", u if u.startswith("http") else host + u) for u in args.pd_urls]
    rnd = random.Random(args.seed)
    pages = [p for p in PF_PAGES if not args.groups or p[0].lower() in {g.lower() for g in args.groups}]
    out = []
    page = await ctx.new_page()
    for group, label, pf in pages:
        reason = ""
        try:
            resp = await page.goto(f"{host}/{site}/{pf}", wait_until="domcontentloaded", timeout=TIMEOUTS["page_load"])
            await page.wait_for_timeout(5000)
            await dismiss_popups(page)
            if not resp or resp.status >= 400 or f"/{site}/" not in page.url:
                links, reason = [], f"이 나라에 PF 가 없음 ({resp.status if resp else '응답 없음'}, {page.url})"
            else:
                links = await page.evaluate(JS_PD_LINKS, [site, pf.split("/")[0]])
                reason = "" if links else "PF 에 제품 카드가 없음"
        except Exception as e:
            links, reason = [], f"PF 열기 실패: {type(e).__name__}"
        picks = rnd.sample(links, min(args.per_category, len(links))) if links else []
        print(f"  [{site}] [{label}] PF 제품 {len(links)}개 중 {len(picks)}개 선택: {[p['href'] for p in picks]}")
        if picks:
            out += [(group, label, host + p["href"]) for p in picks]
        else:
            out.append((group, label, None, reason))
    await page.close()
    return out


async def check_anchors(page, add, snap, nums, anchor_wait):
    """Key Specs 번호 → 하단 디스클레이머 → 화살표 → 번호 검증."""
    bottom = await page.evaluate("""() => [...document.querySelectorAll('[id^=masterCompareBottomDisclaimer]')]
        .map(e => e.id).filter(id => !id.includes('GoBack')).map(id => +id.replace('masterCompareBottomDisclaimer', ''))""")
    missing = [n for n in nums if n not in bottom]
    if missing:
        add("앵커: 하단 디스클레이머 존재", "FAIL", f"하단에 없는 번호: {missing}", expected=str(nums), actual=str(bottom))
    for n in nums:
        num_sel = f"a#masterCompareBottomDisclaimerGoBack{n}"
        disc_sel = f"#masterCompareBottomDisclaimer{n}"
        num = page.locator(num_sel).locator("visible=true").first
        try:
            await num.scroll_into_view_if_needed(timeout=5000)
            await page.wait_for_timeout(500)
            before = await page.evaluate("() => Math.round(window.scrollY)")
            await num.click(timeout=5000)
        except Exception as e:
            add(f"앵커 {n}: 번호 → 디스클레이머", "FAIL", f"번호 클릭 실패: {type(e).__name__}", screenshot=await snap(f"n{n}_click"))
            continue
        await wait_scroll_settle(page, anchor_wait)
        v = await page.evaluate(JS_IN_VIEW, disc_sel)
        ok = v["inView"] and v["scrollY"] > before
        add(f"앵커 {n}: 번호 → 디스클레이머", "PASS" if ok else "FAIL", expected=f"하단 디스클레이머 {n} 이 화면에 보임",
            actual=f"보임={v['inView']}, 스크롤 {before}→{v['scrollY']}", screenshot="" if ok else await snap(f"n{n}_down"))
        arrow = page.locator(f"{disc_sel} a[href='#masterCompareBottomDisclaimerGoBack{n}']").first
        if not await arrow.count():
            add(f"앵커 {n}: 화살표 → 번호", "FAIL", "디스클레이머에 화살표가 없음", screenshot=await snap(f"n{n}_noarrow"))
            continue
        try:
            if not ok:
                await arrow.scroll_into_view_if_needed(timeout=5000)
                await page.wait_for_timeout(500)
            before2 = await page.evaluate("() => Math.round(window.scrollY)")
            await arrow.click(timeout=5000)
        except Exception as e:
            add(f"앵커 {n}: 화살표 → 번호", "FAIL", f"화살표 클릭 실패: {type(e).__name__}", screenshot=await snap(f"n{n}_arrow"))
            continue
        await wait_scroll_settle(page, anchor_wait)
        v2 = await page.evaluate(JS_IN_VIEW, num_sel)
        ok2 = v2["inView"] and v2["scrollY"] < before2
        add(f"앵커 {n}: 화살표 → 번호", "PASS" if ok2 else "FAIL", expected=f"상단 번호 {n} 이 화면에 보임",
            actual=f"보임={v2['inView']}, 스크롤 {before2}→{v2['scrollY']}", screenshot="" if ok2 else await snap(f"n{n}_up"))


async def check_chips(page, add, snap, col, wait_ms):
    """제품 열 하나: 컬러칩 1개, 용량칩 1개를 눌러 모델코드·이미지 변경 확인 후 원래 칩으로 되돌린다."""
    st = await page.evaluate(JS_CHIPS, col)
    if not st:
        return
    card = page.locator(f"li.pdd43-master-compare__card-item.pdd43-master-compare__column--{col}").first
    origin = st["sku"]
    for kind, is_color in (("컬러칩", True), ("용량칩", False)):
        cur = await page.evaluate(JS_CHIPS, col)
        cands = [c for c in cur["chips"] if c["color"] == is_color and not c["disabled"]
                 and c["model"] and c["model"] != cur["sku"]]
        if not [c for c in cur["chips"] if c["color"] == is_color]:
            add(f"{kind}: {col}열", "확인불가", f"{kind} 없음 (옵션이 하나뿐인 제품)", actual=cur["sku"])
            continue
        if not cands:
            add(f"{kind}: {col}열", "확인불가", f"현재 모델 말고 고를 수 있는 {kind} 없음", actual=cur["sku"])
            continue
        c = cands[0]
        btn = card.locator("button.js-compare-option-fmychip").nth(c["i"])
        try:
            await btn.scroll_into_view_if_needed(timeout=5000)
            await btn.click(timeout=5000)
        except Exception as e:
            add(f"{kind}: {col}열", "FAIL", f"칩 클릭 실패: {type(e).__name__}", screenshot=await snap(f"chip{col}_{is_color}"))
            continue
        await page.wait_for_timeout(wait_ms)
        after = await page.evaluate(JS_CHIPS, col)
        sku_ok = after["sku"] == c["model"]
        img_changed = after["img"] != cur["img"]
        img_model = c["model"].lower() in after["img"].lower()
        label = f"'{c['label'] or c['model']}'"
        add(f"{kind}: {col}열 모델코드", "PASS" if sku_ok else "FAIL", f"{label} 선택",
            expected=c["model"], actual=f"{cur['sku']} → {after['sku']}", screenshot="" if sku_ok else await snap(f"chip{col}_{is_color}_sku"))
        if img_changed and (img_model or not is_color):
            res, note = "PASS", ""
        elif img_changed:
            res, note = "PASS", "이미지는 바뀌었으나 주소에 새 모델코드가 없음"
        elif not is_color:
            res, note = "확인필요", "용량만 바뀌어 같은 색이면 이미지가 같을 수 있음"
        else:
            res, note = "FAIL", "컬러를 바꿨는데 이미지가 그대로"
        add(f"{kind}: {col}열 이미지 변경", res, note, expected="제품 이미지가 선택한 옵션으로 바뀜",
            actual=f"{cur['img'].split('/')[-1][:60]} → {after['img'].split('/')[-1][:60]}",
            screenshot=await snap(f"chip{col}_{is_color}_img") if res != "PASS" else "")
    # 원래 모델로 되돌리기
    back = await page.evaluate(JS_CHIPS, col)
    if back and back["sku"] != origin:
        for c in back["chips"]:
            if c["model"] == origin:
                try:
                    await card.locator("button.js-compare-option-fmychip").nth(c["i"]).click(timeout=5000)
                    await page.wait_for_timeout(wait_ms)
                except Exception:
                    pass
                break


async def check_pd(ctx, spec, site, opts, out_dir, shots_dir):
    group, label, url = spec[:3]
    rows = []
    if url is None:
        return [{"site": site, "group": group, "page": label, "pd": "", "check": "PD 선택", "expected": "",
                 "actual": "", "result": "해당 없음", "detail": spec[3], "screenshot": ""}]
    slug = url.rstrip("/").split("/")[-1][:50]
    tag = f"{site}_{slug}"

    def add(check, result, detail="", expected="", actual="", screenshot=""):
        rows.append({"site": site, "group": group, "page": label, "pd": url, "check": check, "expected": expected,
                     "actual": actual, "result": result, "detail": detail, "screenshot": screenshot})

    async def snap(name):
        return await shot(page, shots_dir, f"{tag}_{name}")

    page = await ctx.new_page()
    try:
        resp = await page.goto(url, wait_until="domcontentloaded", timeout=TIMEOUTS["page_load"])
        await page.wait_for_timeout(opts["wait"] * 1000)
        await dismiss_popups(page)
        if not resp or resp.status >= 400:
            add("PD 열기", "FAIL", expected="200", actual=str(resp.status if resp else ""), screenshot=await snap("open"))
            return rows

        # 1) 상단 'Compare' 탭 → 하단 Compare 영역
        tabs = await page.evaluate(JS_PD_COMPARE_TAB)
        if not tabs:
            hidden = await page.evaluate(JS_PD_HIDDEN_TAB)
            add("앵커: Compare 탭 → Compare 영역", "FAIL",
                hidden or "PD 상단에 Compare 탭(#pdd43MasterCompare 링크)이 없음", screenshot=await snap("notab"))
        else:
            before = await page.evaluate("() => Math.round(window.scrollY)")
            tab = page.locator(f"a[href='{tabs[0]['href']}']").locator("visible=true").first
            clicked = False
            try:
                await tab.click(timeout=5000)
                clicked = True
            except Exception:
                pass
            await wait_scroll_settle(page, opts["anchor_wait"])
            v = await page.evaluate(JS_IN_VIEW, ".pdd43-master-compare")
            ok = clicked and v["count"] > 0 and v["inView"] and v["scrollY"] > before
            add("앵커: Compare 탭 → Compare 영역", "PASS" if ok else "FAIL",
                f"탭 '{tabs[0]['text']}' ({tabs[0]['href']})", expected="하단 Compare 영역이 화면에 보임",
                actual=f"영역 {v['count']}개, 보임={v['inView']}, 스크롤 {before}→{v['scrollY']}",
                screenshot=await snap("compare_area") if (not ok or opts["shots_all"]) else "")

        if opts["dump"]:
            html = await page.evaluate("() => { const r = document.querySelector('.pdd43-master-compare'); return r ? r.outerHTML : ''; }")
            if html:
                (out_dir / "dom").mkdir(parents=True, exist_ok=True)
                (out_dir / "dom" / f"pd_dom_{tag}.html").write_text(html, encoding="utf-8")

        st = await page.evaluate(JS_STATE)
        if not st.get("hasRoot"):
            add("Compare 영역", "FAIL", "PD 에 .pdd43-master-compare 영역이 없음 (컴포넌트가 다를 수 있음)",
                actual=page.url, screenshot=await snap("noroot"))
            return rows

        # 2) Key Specs (고정 3개 제품)
        products = " / ".join(st["names"] or st["codes"])
        ncols = len(st["cols"])
        if not st["featVisible"] or not st["cols"]:
            add("Key Specs 노출", "FAIL", f"비교 제품: {products}", expected="Key Specs 섹션 노출",
                actual="섹션 없음/미노출", screenshot=await snap("keyspecs"))
        else:
            empty = [c for c, v in st["cols"].items() if v["filled"] == 0]
            detail = ", ".join(f"{c}열 값 {v['filled']}/{v['filled'] + v['empty']}" for c, v in sorted(st["cols"].items()))
            ok = not empty and ncols >= opts["expect_cols"]
            add("Key Specs 노출", "PASS" if ok else "FAIL", f"비교 제품: {products}",
                expected=f"제품 {opts['expect_cols']}개 모두 Key Specs 값 있음",
                actual=f"제품 {ncols}개 / 항목 {st['rows']}개 / {detail}",
                screenshot=await snap("keyspecs") if (not ok or opts["shots_all"]) else "")

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
                sections.append(("앵커", lambda: check_anchors(page, add, snap, nums, opts["anchor_wait"])))
        if opts["chips"]:
            cols = sorted(int(c) for c in st["cols"] if str(c).isdigit()) or [1, 2, 3]

            async def all_chips():
                for c in cols:
                    await check_chips(page, add, snap, c, opts["chip_wait"])
            sections.append(("칩", all_chips))
        for name, run in sections:
            try:
                await run()
            except Exception as e:
                add(f"{name}: 오류", "오류", f"{type(e).__name__}: {e}", screenshot=await snap(f"{name}_error"))
    except Exception as e:
        add("PD 오류", "오류", f"{type(e).__name__}: {e}", screenshot=await snap("error"))
    finally:
        await page.close()
    return rows


def write_report(path, rows, args):
    wb = Workbook()
    ws = wb.active
    ws.title = "국가별 요약"
    cats = [p[1] for p in PF_PAGES if not args.groups or p[0].lower() in {g.lower() for g in args.groups}]
    if args.pd_urls:
        cats = ["지정 PD"]
    by_site = OrderedDict()
    for r in rows:
        by_site.setdefault(r["site"], OrderedDict()).setdefault(r["page"], []).append(r)
    srows = []
    for site, cs in by_site.items():
        vs = []
        for c in cats:
            rs = cs.get(c, [])
            vs.append("-" if not rs else "해당 없음" if all(x["result"] == "해당 없음" for x in rs)
                      else "FAIL" if any(x["result"] in ("FAIL", "오류") for x in rs) else "PASS")
        fails = sum(1 for rs in cs.values() for x in rs if x["result"] in ("FAIL", "오류"))
        srows.append([site, *vs, fails, "FAIL" if fails else ("해당 없음" if set(vs) <= {"해당 없음", "-"} else "PASS")])
    _sheet(ws, ["사이트", *cats, "FAIL 수", "국가 판정"], srows, [9] + [13] * len(cats) + [8, 10])
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=len(cats) + 3):
        for c in row:
            color = FILL.get(str(c.value))
            if color:
                c.fill = PatternFill("solid", fgColor=color)
    ws = wb.create_sheet("PD별 요약")
    by = OrderedDict()
    for r in rows:
        by.setdefault((r["site"], r["group"], r["page"], r["pd"]), []).append(r)
    out = []
    for (site, group, label, pd), rs in by.items():
        def res(prefix):
            sub = [x for x in rs if x["check"].startswith(prefix)]
            if not sub:
                return "-"
            c = Counter(x["result"] for x in sub)
            return "PASS" if set(c) == {"PASS"} else ", ".join(f"{k} {v}" for k, v in c.items())
        down = [x for x in rs if "번호 → 디스클레이머" in x["check"] and x["result"] != "확인불가"]
        up = [x for x in rs if "화살표 → 번호" in x["check"]]
        fails = sum(1 for x in rs if x["result"] in ("FAIL", "오류"))
        out.append([site, group, label, pd, res("앵커: Compare 탭"), res("Key Specs"), res("차이점"), res("All Specs"),
                    f"{sum(x['result'] == 'PASS' for x in down)}/{len(down)}" if down else "번호 없음",
                    f"{sum(x['result'] == 'PASS' for x in up)}/{len(up)}" if up else "-",
                    res("컬러칩"), res("용량칩"), fails,
                    "FAIL" if fails else ("해당 없음" if all(x["result"] == "해당 없음" for x in rs) else "PASS")])
    _sheet(ws, ["사이트", "그룹", "카테고리", "PD 주소", "Compare 탭 앵커", "Key Specs", "차이점 보기", "All Specs",
                "번호→디스클레이머", "화살표→번호", "컬러칩", "용량칩", "FAIL 수", "판정"],
           out, [8, 7, 13, 60, 14, 12, 14, 14, 14, 12, 16, 16, 8, 8], result_col=13)
    ws = wb.create_sheet("상세")
    _sheet(ws, ["사이트", "그룹", "카테고리", "PD 주소", "검증", "기대", "실제", "결과", "비고", "캡처"],
           [[r["site"], r["group"], r["page"], r["pd"], r["check"], r["expected"], r["actual"], r["result"],
             r["detail"], r["screenshot"]] for r in rows], [8, 7, 13, 45, 26, 30, 50, 9, 40, 40], result_col=7)
    ws = wb.create_sheet("실행 정보")
    for r in [["실행 시각", dt.datetime.now().strftime("%Y-%m-%d %H:%M")], ["환경", f"live ({BASE_URL})"],
              ["사이트", f"{len(args.sites)}개: " + " ".join(args.sites)],
              ["해당 없음", "그 나라에 해당 카테고리 PF 나 제품이 없는 경우 (FAIL 아님)"],
              ["PD 선택", "지정: " + ", ".join(args.pd_urls) if args.pd_urls else
               f"카테고리마다 임의 {args.per_category}개 (seed {args.seed})"],
              ["검증", "Compare 탭 앵커, Key Specs" + ("" if args.no_diff else ", 차이점 보기") + ("" if args.no_allspecs else ", All Specs 접기/펼치기") +
               ("" if args.no_anchor else ", 번호/화살표") + ("" if args.no_chips else ", 컬러칩/용량칩")]]:
        ws.append(r)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 120
    wb.save(path)


async def run_site(browser, args, site, opts, run_dir):
    """사이트 1개: PD 고르기 → PD 마다 검증 → sites/<site>.json 저장."""
    ctx = await browser.new_context(viewport={"width": args.width, "height": args.height})
    rows = []
    try:
        pds = await collect_pds(ctx, BASE_URL, site, args)
        for spec in pds:
            r = await check_pd(ctx, spec, site, opts, run_dir, run_dir / "screenshots")
            rows += r
            if spec[2] is None:
                print(f"  [{site}] {spec[1]}: 해당 없음 ({spec[3]})")
            else:
                print(f"  [{site}] {spec[1]}: 검증 {len(r)}건, FAIL/오류 {sum(1 for x in r if x['result'] in ('FAIL', '오류'))}건")
    finally:
        await ctx.close()
    (run_dir / "sites" / f"{site}.json").write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")
    return rows


async def amain(args):
    out_root = Path(args.out_dir)
    if args.resume:
        runs = sorted(out_root.glob("run_live_*")) if args.resume == "latest" else [Path(args.resume)]
        if not runs or not runs[-1].exists():
            sys.exit("이어서 실행할 이전 결과 폴더가 없습니다.")
        run_dir = runs[-1]
    else:
        run_dir = out_root / f"run_live_{dt.datetime.now():%Y%m%d_%H%M%S}"
    (run_dir / "sites").mkdir(parents=True, exist_ok=True)
    if not args.sites:
        args.sites = load_live_sites(args.mapping_list)
    args.sites = [s for s in args.sites if s not in set(args.exclude_sites)]
    opts = {"wait": args.wait, "anchor_wait": args.anchor_wait, "max_anchors": args.max_anchors,
            "shots_all": args.shots_all, "diff": not args.no_diff, "anchor": not args.no_anchor,
            "allspecs": not args.no_allspecs, "chips": not args.no_chips, "chip_wait": args.chip_wait, "expect_cols": args.expect_cols,
            "dump": not args.no_dump}
    todo = [s for s in args.sites if not (run_dir / "sites" / f"{s}.json").exists()]
    print(f"사이트 {len(args.sites)}개 (이번에 실행 {len(todo)}개, 동시 {args.parallel}개) → {run_dir}")
    sem = asyncio.Semaphore(max(1, args.parallel))
    async with async_playwright() as p:
        b = await p.chromium.launch(headless=args.headless, slow_mo=TIMEOUTS["slow_mo"])

        async def one(site):
            async with sem:
                t0 = dt.datetime.now()
                print(f"[{site}] 시작")
                try:
                    await run_site(b, args, site, opts, run_dir)
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
    out = run_dir / f"PDCompare_LIVE_{run_dir.name.split('_', 2)[2]}.xlsx"
    write_report(out, rows, args)
    print("결과 파일:", out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--sites", nargs="*", default=[], help="사이트코드 (생략하면 Live_기준 시트의 전체 사이트)")
    ap.add_argument("--exclude-sites", nargs="*", default=[], help="제외할 사이트코드")
    ap.add_argument("--mapping-list", default=MAPPING_LIST_FILE, help="사이트 목록을 읽을 매핑 리스트 파일")
    ap.add_argument("--groups", nargs="*", default=[], help="MX / PC / TV / Audio / DA 중 선택")
    ap.add_argument("--pd-urls", nargs="*", default=[], help="검증할 PD 주소 직접 지정 (/uk/... 또는 전체 주소)")
    ap.add_argument("--per-category", type=int, default=1, help="카테고리마다 임의로 고를 PD 수 (기본 1)")
    ap.add_argument("--seed", type=int, default=None, help="임의 선택 고정용 숫자")
    ap.add_argument("--parallel", type=int, default=2, help="동시에 돌릴 사이트 수 (기본 2)")
    ap.add_argument("--resume", nargs="?", const="latest", default=None, help="이전 실행 이어서 (폴더 지정 가능)")
    ap.add_argument("--expect-cols", type=int, default=3, help="Compare 영역에 나와야 하는 제품 수 (기본 3)")
    ap.add_argument("--max-anchors", type=int, default=3,
                    help="PD 하나당 확인할 디스클레이머 번호 수 (기본 3 = 앞에서부터 3개, 0 = 전부)")
    ap.add_argument("--wait", type=int, default=TIMEOUTS["pf_wait"], help="PD 로딩 대기(초)")
    ap.add_argument("--anchor-wait", type=int, default=TIMEOUTS["anchor_wait"], help="클릭 후 스크롤 대기(ms)")
    ap.add_argument("--chip-wait", type=int, default=TIMEOUTS["chip_switch"], help="칩 클릭 후 대기(ms)")
    ap.add_argument("--no-diff", action="store_true", help="Show me the differences 검증 생략")
    ap.add_argument("--no-allspecs", action="store_true", help="All Specs 접기/펼치기 검증 생략")
    ap.add_argument("--no-anchor", action="store_true", help="번호/화살표 검증 생략")
    ap.add_argument("--no-chips", action="store_true", help="컬러칩/용량칩 검증 생략")
    ap.add_argument("--no-dump", action="store_true", help="PD Compare 영역 HTML 저장 안 함")
    ap.add_argument("--shots-all", action="store_true", help="PASS 인 화면도 캡처")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--width", type=int, default=1600)
    ap.add_argument("--height", type=int, default=1000)
    ap.add_argument("--out-dir", default=HERE / "result")
    asyncio.run(amain(ap.parse_args()))


if __name__ == "__main__":
    main()
