# -*- coding: utf-8 -*-
"""[라이브] Compare models 페이지 화면 검증 - 89개국(91개 사이트코드)용.

카테고리별 Compare 페이지를 실제 브라우저로 열어 아래를 확인한다.
  1) 진입     : PF 상단 Compare 탭 클릭 → Compare 페이지로 이동하는지 (탭이 없는 TV/Audio/Monitor 는 고정 주소로 진입)
  2) Key Specs: 'Key Specs' 섹션이 보이고, 비교 중인 제품(열)마다 Key Specs 값이 있는지
  3) 앵커 이동: Key Specs 의 번호(1, 2, …)를 누르면 하단 디스클레이머의 같은 번호로 이동하는지
  4) 복귀     : 그 디스클레이머의 화살표(Go to feature)를 누르면 상단의 같은 번호로 돌아오는지
  5) 차이점   : 'Show me the differences' 를 켜면 값이 같은 행은 숨겨지고, 다른 행은 남아 다른 값이 파란 글자로 보이는지
                (끄면 숨겨진 행이 다시 보이는지)
  5-1) All Specs: 카테고리 버튼(첫·마지막)을 누르면 내용이 접히고, 다시 누르면 펼쳐지는지
  6) 드롭다운 : 제품명 드롭다운을 열어 다른 모델을 고르면 모델명·모델코드·Key Specs 값이 바뀌는지
  7) CTA      : 제품마다 Buy 를 누르면 구매 페이지로 이동(또는 구매 레이어 노출)하는지,
                Learn more 를 누르면 해당 제품 페이지로 이동하는지 (CTA 는 페이지를 벗어나므로 맨 마지막에 확인)

사용 예
  python run_c_landing_live.py --sites uk                 # 한 나라
  python run_c_landing_live.py --sites uk de fr pl        # 여러 나라
  python run_c_landing_live.py                            # 91개 사이트 전체 (오래 걸림 - --parallel 로 동시 실행)
  python run_c_landing_live.py --parallel 3 --headless    # 3개 나라씩 동시에, 창 없이
  python run_c_landing_live.py --resume                   # 마지막 실행 이어서
  python run_c_landing_live.py --groups MX --max-anchors 3 --no-cta
결과: compare/result/run_live_<날짜시간>/ComparePage_LIVE_<날짜시간>.xlsx
      (사이트별 중간 결과 sites/, 실패 캡처 screenshots/)
"""
import argparse
import asyncio
import datetime as dt
import json
import sys
from pathlib import Path
from urllib.parse import urljoin

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from playwright.async_api import async_playwright

from config import (
    BASE_URL,
    CATEGORY_SPECS,
    FILL_COLORS,
    HERE,
    MAPPING_LIST_FILE,
    TIMEOUTS,
)

# 하위 호환성 유지
FILL = FILL_COLORS
PAGES = CATEGORY_SPECS


def load_live_sites(path=MAPPING_LIST_FILE):
    """'Live_기준' 시트에서 DISPLAY_STATUS=Y 제품이 있는 사이트코드 (시트 순서)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = next(w for w in wb.worksheets if w.title.upper().startswith("LIVE"))
    rows = ws.iter_rows(values_only=True)
    hdr = [str(h or "").strip() for h in next(rows)]
    i_site, i_disp = hdr.index("SITE_CD"), hdr.index("DISPLAY_STATUS")
    sites = []
    for r in rows:
        s = str(r[i_site] or "").strip()
        if s and str(r[i_disp] or "").strip().upper() == "Y" and s not in sites:
            sites.append(s)
    wb.close()
    return sites


# PF 상단 탭 중 Compare 페이지로 가는 링크 (글자가 아니라 주소로 찾는다 → 현지어 사이트 대응)
JS_PF_COMPARE_TAB = r"""() => [...document.querySelectorAll('a.tab__item-title')]
  .filter(a => /\/compare\/?(\?|$)/.test((a.getAttribute('href') || '').split('#')[0]) && a.offsetParent)
  .map(a => a.getAttribute('href'))"""


# ------------------------------------------------------------------ 브라우저 안에서 쓰는 JS
JS_STATE = r"""() => {
  const root = document.querySelector('.pdd43-master-compare');
  if (!root) return {hasRoot: false};
  const vis = e => !!(e && e.offsetParent !== null && e.getBoundingClientRect().height > 0);
  const feat = root.querySelector('.pdd43-master-compare__feature');
  const head = feat && feat.querySelector('.pdd43-master-compare__sub-headline');
  const lists = feat ? [...feat.querySelectorAll('ul.pdd43-master-compare__feature-list')] : [];
  const cols = {};
  for (const ul of lists) {
    for (const li of ul.querySelectorAll('li.pdd43-master-compare__feature-item')) {
      const m = li.className.match(/feature-item--(\d+)/); const c = m ? m[1] : '?';
      const v = (li.querySelector('.pdd43-master-compare__feature-value') || {}).innerText || '';
      cols[c] = cols[c] || {filled: 0, empty: 0};
      if (v.trim() && v.trim() !== '-') cols[c].filled++; else cols[c].empty++;
    }
  }
  const names = [...root.querySelectorAll('[class*=product-name], [class*=select-name], .pdd43-master-compare__product-select button, [class*=__name]')]
      .filter(vis).map(e => e.innerText.trim()).filter(t => t && t.length < 80);
  const codes = [...root.querySelectorAll('[class*=model-code], [class*=sku]')].filter(vis).map(e => e.innerText.trim()).filter(Boolean);
  const nums = [...new Set([...(feat ? feat.querySelectorAll('a[id^=masterCompareBottomDisclaimerGoBack]') : [])]
      .map(a => +a.id.replace('masterCompareBottomDisclaimerGoBack', '')))].sort((a, b) => a - b);
  const bottom = [...root.querySelectorAll('li[id^=masterCompareBottomDisclaimer]')].map(li => +li.id.replace('masterCompareBottomDisclaimer', ''));
  return {hasRoot: true, url: location.href, title: head ? head.innerText.trim() : '', featVisible: vis(feat),
          rows: lists.length, cols, names: [...new Set(names)].slice(0, 6), codes: [...new Set(codes)].slice(0, 6),
          nums, bottom};
}"""

JS_IN_VIEW = r"""(sel) => {
  const els = [...document.querySelectorAll(sel)];
  const h = window.innerHeight;
  const hit = els.find(e => { const r = e.getBoundingClientRect(); return r.height > 0 && r.bottom > 0 && r.top < h; });
  return {count: els.length, inView: !!hit, top: hit ? Math.round(hit.getBoundingClientRect().top) : null, scrollY: Math.round(window.scrollY), hash: location.hash};
}"""

JS_DIFF_ROWS = r"""() => {
  const root = document.querySelector('.pdd43-master-compare');
  const shown = e => { if (!e) return false; const s = getComputedStyle(e), r = e.getBoundingClientRect();
                       return s.display !== 'none' && s.visibility !== 'hidden' && r.height > 1 && e.offsetParent !== null; };
  const isBlue = e => { const m = getComputedStyle(e).color.match(/\d+/g); if (!m) return false;
                        const [r, g, b] = m.map(Number); return b >= 120 && b > r + 40 && b > g + 10; };
  const norm = t => (t || '').replace(/\s+/g, ' ').trim();
  const rows = [];
  const add = (kind, rowEl, cells, title) => {
    const vals = cells.map(c => norm(c.innerText));
    if (!vals.length) return;
    const same = vals.every(v => v === vals[0]);
    const marks = rowEl.querySelectorAll('mark.pdd43-master-compare__highlight');
    rows.push({kind, title: norm(title).slice(0, 60), same, shown: shown(rowEl),
               blue: [...marks].filter(m => shown(m) && isBlue(m)).length, marks: marks.length,
               values: vals.map(v => v.slice(0, 40)).join(' | ')});
  };
  for (const ul of root.querySelectorAll('.pdd43-master-compare__feature ul.pdd43-master-compare__feature-list')) {
    const cells = [...ul.querySelectorAll('li.pdd43-master-compare__feature-item')]
        .map(li => li.querySelector('.pdd43-master-compare__feature-value') || li);
    const t = ul.querySelector('.pdd43-master-compare__feature-title-text');
    add('Key Specs', ul, cells, t ? t.innerText : '');
  }
  for (const d of root.querySelectorAll('.pdd43-master-compare__spec-detail')) {
    const cells = [...d.querySelectorAll('li.pdd43-master-compare__spec-value-item')];
    const t = d.querySelector('.pdd43-master-compare__spec-title');
    add('All Specs', d, cells, t ? t.innerText : '');
  }
  const btn = root.querySelector('.pdd43-master-compare__switch button');
  return {hasSwitch: !!btn, pressed: btn ? btn.getAttribute('aria-pressed') : null, rows};
}"""

JS_COLUMN = r"""(n) => {
  const root = document.querySelector('.pdd43-master-compare');
  const card = root && root.querySelector(`li.pdd43-master-compare__card-item.pdd43-master-compare__column--${n}`);
  if (!card) return null;
  const txt = s => { const e = card.querySelector(s); return e ? e.innerText.replace(/\s+/g, ' ').trim() : ''; };
  const ks = [...root.querySelectorAll(`.pdd43-master-compare__feature li.pdd43-master-compare__feature-item--${n}`)]
      .map(li => ((li.querySelector('.pdd43-master-compare__feature-value') || {}).innerText || '').trim());
  const buy = card.querySelector('.pdd43-master-compare__cta a.js-cta-buy-now');
  const lm = [...card.querySelectorAll('.pdd43-master-compare__cta a')].find(a => a !== buy && /cta--underline/.test(a.className));
  return {name: txt('.menu__select-field-text'), sku: txt('.pdd43-master-compare__sku-text'),
          ks: ks.join(' | '), ksFilled: ks.filter(v => v && v !== '-').length,
          buy: buy ? {text: buy.innerText.trim(), href: buy.getAttribute('href'), link: buy.getAttribute('data-link_info') || buy.getAttribute('data-config_info') || ''} : null,
          learn: lm ? {text: lm.innerText.trim(), href: lm.getAttribute('href')} : null};
}"""

JS_OPTIONS = r"""(n) => {
  const root = document.querySelector('.pdd43-master-compare');
  const card = root.querySelector(`li.pdd43-master-compare__card-item.pdd43-master-compare__column--${n}`);
  const taken = [...root.querySelectorAll('li.pdd43-master-compare__card-item .menu__select-field-text')].map(e => e.innerText.trim());
  return [...card.querySelectorAll('.menu__list-option')].map((b, i) => ({i, text: b.innerText.trim(),
           selected: b.getAttribute('aria-selected') === 'true'})).filter(o => o.text && !o.selected && !taken.includes(o.text));
}"""

JS_LAYERS = r"""() => [...document.querySelectorAll('[role=dialog], [class*=layer-popup], [class*=popup], [class*=modal]')]
  .filter(e => { const s = getComputedStyle(e), r = e.getBoundingClientRect();
                 return s.display !== 'none' && s.visibility !== 'hidden' && r.width > 200 && r.height > 150; }).length"""


# ------------------------------------------------------------------ 공통 대기 및 팝업 처리
async def wait_scroll_settle(page, min_ms, max_ms=6000, step=300):
    """번호/화살표 클릭 후 부드러운 스크롤이 끝날 때까지 대기."""
    await page.wait_for_timeout(min_ms)
    last, same, waited = None, 0, min_ms
    while waited < max_ms:
        y = await page.evaluate("() => Math.round(window.scrollY)")
        same = same + 1 if y == last else 0
        if same >= 2:
            return
        last = y
        await page.wait_for_timeout(step)
        waited += step


async def dismiss_popups(page):
    """쿠키 배너 거부/필수만 허용 처리."""
    for sel in ["#truste-consent-required", "#onetrust-reject-all-handler", "button[id*='reject' i]",
                "button:has-text('Reject')"]:
        try:
            l = page.locator(sel).first
            if await l.is_visible(timeout=1500):
                await l.click()
                await page.wait_for_timeout(800)
                return
        except Exception:
            pass


async def shot(page, shots_dir, name):
    shots_dir.mkdir(parents=True, exist_ok=True)
    p = shots_dir / f"{name}.png"
    try:
        await page.screenshot(path=str(p))
        return str(p)
    except Exception:
        return ""


# ------------------------------------------------------------------ 5) Show me the differences
async def check_differences(page, add, snap, shots_all=False):
    before = await page.evaluate(JS_DIFF_ROWS)
    if not before["hasSwitch"]:
        add("차이점: 스위치", "FAIL", "'Show me the differences' 스위치가 없음", screenshot=await snap("diff_noswitch"))
        return
    btn = page.locator(".pdd43-master-compare__switch button").first
    await btn.scroll_into_view_if_needed()
    await btn.click()
    await page.wait_for_timeout(1500)
    on = await page.evaluate(JS_DIFF_ROWS)
    if shots_all:
        await page.locator(".pdd43-master-compare__feature").first.scroll_into_view_if_needed()
        await page.wait_for_timeout(500)
        on_shot = await snap("diff_on")
    else:
        on_shot = ""
    add("차이점: 스위치 켜짐", "PASS" if on["pressed"] == "true" else "FAIL",
        expected="aria-pressed=true", actual=f"aria-pressed={on['pressed']}", screenshot=on_shot)

    same = [r for r in on["rows"] if r["same"]]
    diff = [r for r in on["rows"] if not r["same"]]
    same_shown = [r for r in same if r["shown"]]
    diff_hidden = [r for r in diff if not r["shown"]]
    add("차이점: 같은 값 행 숨김", "PASS" if not same_shown else "FAIL",
        expected=f"값이 같은 행 {len(same)}개 모두 숨김",
        actual=f"보이는 행 {len(same_shown)}개" + (": " + "; ".join(f"[{r['kind']}] {r['title']}" for r in same_shown[:8]) if same_shown else ""),
        screenshot=await snap("diff_same") if same_shown else "")
    add("차이점: 다른 값 행 노출", "PASS" if not diff_hidden else "FAIL",
        expected=f"값이 다른 행 {len(diff)}개 모두 보임",
        actual=f"숨겨진 행 {len(diff_hidden)}개" + (": " + "; ".join(f"[{r['kind']}] {r['title']}" for r in diff_hidden[:8]) if diff_hidden else ""),
        screenshot=await snap("diff_hidden") if diff_hidden else "")
    not_blue = [r for r in diff if r["shown"] and r["blue"] == 0]
    add("차이점: 다른 값 파란 글자", "PASS" if diff and not not_blue else ("확인불가" if not diff else "FAIL"),
        "비교 제품끼리 값이 다른 행이 없음" if not diff else "",
        expected="보이는 '다른 값' 행마다 파란 글자 값이 있음",
        actual=f"파란 글자 없는 행 {len(not_blue)}/{len(diff)}" +
               (": " + "; ".join(f"[{r['kind']}] {r['title']} ({r['values']})" for r in not_blue[:5]) if not_blue else ""),
        screenshot=await snap("diff_blue") if not_blue else "")

    await btn.scroll_into_view_if_needed()
    await btn.click()
    await page.wait_for_timeout(1500)
    off = await page.evaluate(JS_DIFF_ROWS)
    back = sum(1 for r in off["rows"] if r["shown"])
    want = sum(1 for r in before["rows"] if r["shown"])
    add("차이점: 끄면 원래대로", "PASS" if off["pressed"] == "false" and back >= want else "FAIL",
        expected=f"aria-pressed=false, 보이는 행 {want}개", actual=f"aria-pressed={off['pressed']}, 보이는 행 {back}개")


# ------------------------------------------------------------------ 5-1) All Specs 접기/펼치기
JS_ACCORDION = r"""() => {
  const root = document.querySelector('.pdd43-master-compare');
  if (!root) return [];
  return [...root.querySelectorAll('button.pdd43-master-compare__spec-button')].map((b, i) => {
    const cat = b.closest('.pdd43-master-compare__spec-category') || b.parentElement;
    const wraps = [...cat.querySelectorAll('.pdd43-master-compare__spec-detail-wrap')];
    const h = wraps.reduce((s, w) => s + w.getBoundingClientRect().height, 0);
    return {i, title: (b.getAttribute('data-title') || b.innerText || '').replace(/\s+/g, ' ').trim().slice(0, 50),
            expanded: b.getAttribute('aria-expanded'), rows: wraps.length, height: Math.round(h)};
  });
}"""


async def check_all_specs(page, add, snap, wait_ms=1200):
    accs = await page.evaluate(JS_ACCORDION)
    if not accs:
        add("All Specs: 접기/펼치기", "FAIL", "All Specs 카테고리 버튼이 없음", screenshot=await snap("allspecs_none"))
        return
    targets = [accs[0]] + ([accs[-1]] if len(accs) > 1 else [])
    btns = page.locator(".pdd43-master-compare button.pdd43-master-compare__spec-button")
    for a in targets:
        name = f"All Specs: '{a['title']}'"
        if not a["rows"]:
            add(name, "확인불가", "카테고리 안에 스펙 행이 없음")
            continue
        btn = btns.nth(a["i"])
        opened0 = a["expanded"] == "true"
        full = a["height"]
        for step, want_open in (("1번 누름", not opened0), ("2번 누름", opened0)):
            await btn.scroll_into_view_if_needed(timeout=5000)
            await btn.click(timeout=5000)
            await page.wait_for_timeout(wait_ms)
            s = (await page.evaluate(JS_ACCORDION))[a["i"]]
            full = max(full, s["height"])
            ok = (s["expanded"] == ("true" if want_open else "false")
                  and (s["height"] > 0 if want_open else s["height"] <= max(2, full * 0.1)))
            act = "펼치기" if want_open else "접기"
            add(f"{name} {act}", "PASS" if ok else "FAIL", f"처음 {'펼침' if opened0 else '접힘'} 상태에서 {step}",
                expected=f"aria-expanded={'true' if want_open else 'false'}, 내용 {'보임' if want_open else '숨김'}",
                actual=f"aria-expanded={s['expanded']}, 내용 높이 {s['height']}px (스펙 행 {s['rows']}개)",
                screenshot="" if ok else await snap(f"allspecs{a['i']}_{act}"))


# ------------------------------------------------------------------ 6) 제품명 드롭다운
async def check_dropdown(page, add, snap, col):
    before = await page.evaluate(JS_COLUMN, col)
    if not before:
        add(f"드롭다운: {col}열", "확인불가", f"{col}번째 제품 열이 없음")
        return
    card = page.locator(f"li.pdd43-master-compare__card-item.pdd43-master-compare__column--{col}").first
    field = card.locator("button.menu__select-field").first
    if not await field.count():
        add(f"드롭다운: {col}열 열기", "FAIL", "제품명 드롭다운 버튼이 없음", screenshot=await snap("dd_nobtn"))
        return
    await field.scroll_into_view_if_needed()
    await field.click()
    await page.wait_for_timeout(1000)
    expanded = await field.get_attribute("aria-expanded")
    opts = await page.evaluate(JS_OPTIONS, col)
    add(f"드롭다운: {col}열 열기", "PASS" if expanded == "true" and opts else "FAIL",
        expected="목록 펼쳐짐 + 고를 수 있는 다른 모델 있음",
        actual=f"aria-expanded={expanded}, 선택 가능 모델 {len(opts)}개",
        screenshot="" if (expanded == "true" and opts) else await snap("dd_open"))
    if not opts:
        return
    pick = opts[0]
    option = card.locator(".menu__list-option").nth(pick["i"])
    await option.scroll_into_view_if_needed()
    await option.click()
    await page.wait_for_timeout(5000)
    after = await page.evaluate(JS_COLUMN, col)
    name_ok = after and after["name"] == pick["text"]
    sku_ok = after and after["sku"] and after["sku"] != before["sku"]
    add(f"드롭다운: {col}열 모델 변경", "PASS" if name_ok and sku_ok else "FAIL",
        expected=f"모델명 '{pick['text']}', 모델코드 변경",
        actual=f"모델명 '{before['name']}' → '{after['name'] if after else ''}', "
               f"모델코드 {before['sku']} → {after['sku'] if after else ''}",
        screenshot="" if name_ok and sku_ok else await snap("dd_model"))
    st = await page.evaluate(JS_STATE)
    if not st.get("featVisible"):
        add(f"드롭다운: {col}열 Key Specs 변경", "확인필요",
            f"'{pick['text']}' 선택 후 Key Specs 섹션이 사라짐 - 이 모델에 Key Specs 가 없는지 확인 필요",
            expected="Key Specs 값이 새 모델 값으로 바뀜", actual="Key Specs 섹션 없음", screenshot=await snap("dd_noks"))
    else:
        changed = after and after["ks"] != before["ks"] and after["ksFilled"] > 0
        add(f"드롭다운: {col}열 Key Specs 변경", "PASS" if changed else "FAIL",
            expected="Key Specs 값이 새 모델 값으로 바뀜",
            actual=f"값 있는 항목 {before['ksFilled']} → {after['ksFilled'] if after else 0}, "
                   f"값 {'바뀜' if after and after['ks'] != before['ks'] else '그대로'}",
            screenshot="" if changed else await snap("dd_ks"))


# ------------------------------------------------------------------ 7) Buy / Learn more
async def check_cta(page, add, snap, cmp_url, cols, wait_s, click_buy=False):
    async def reload():
        await page.goto(cmp_url, wait_until="domcontentloaded", timeout=TIMEOUTS["page_load"])
        await page.wait_for_timeout(wait_s * 1000)
        await dismiss_popups(page)

    for n in cols:
        await reload()
        c = await page.evaluate(JS_COLUMN, n)
        if c and not c["sku"]:
            await fill_empty_columns(page, lambda *a, **k: None, snap)
            c = await page.evaluate(JS_COLUMN, n)
        if not c or not c["sku"]:
            continue
        who = f"{n}열 {c['name']}"
        if not c["learn"]:
            add(f"CTA: Learn more ({who})", "FAIL", "Learn more 버튼이 없음", screenshot=await snap(f"cta_nolearn{n}"))
        else:
            lm = page.locator(f"li.pdd43-master-compare__card-item.pdd43-master-compare__column--{n} "
                              f".pdd43-master-compare__cta a[href='{c['learn']['href']}']").first
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
            add(f"CTA: Learn more ({who})", "PASS" if ok else "FAIL", "" if ok else "제품 페이지로 이동 안 됨",
                expected=f"{c['learn']['href']} 로 이동", actual=page.url,
                screenshot="" if ok else await snap(f"cta_learn{n}"))
            await reload()

        if not c["buy"]:
            add(f"CTA: Buy ({who})", "확인불가", "Buy 버튼이 없음 (판매 안 하는 모델일 수 있음)")
            continue
        if not click_buy:
            link = c["buy"]["link"]
            if not link:
                add(f"CTA: Buy ({who})", "확인필요", f"버튼 '{c['buy']['text']}' 에 구매 링크 정보가 없음 (클릭해야 확인 가능)")
                continue
            url = urljoin(page.url, link)
            try:
                r = await page.context.request.get(url, timeout=30000, max_redirects=5)
                ok = r.status < 400
                add(f"CTA: Buy ({who})", "PASS" if ok else "FAIL", f"버튼 '{c['buy']['text']}' (클릭하지 않고 링크만 확인)",
                    expected=f"구매 링크 {link} 응답", actual=f"HTTP {r.status} {r.url}")
            except Exception as e:
                add(f"CTA: Buy ({who})", "FAIL", f"구매 링크 확인 실패: {type(e).__name__}", expected=link)
            continue
        buy = page.locator(f"li.pdd43-master-compare__card-item.pdd43-master-compare__column--{n} "
                           f".pdd43-master-compare__cta a.js-cta-buy-now").first
        url0, layers0, pages0 = page.url, await page.evaluate(JS_LAYERS), len(page.context.pages)
        try:
            await buy.scroll_into_view_if_needed()
            await buy.click()
        except Exception as e:
            add(f"CTA: Buy ({who})", "FAIL", f"클릭 실패: {type(e).__name__}", screenshot=await snap(f"cta_buy{n}"))
            continue
        await page.wait_for_timeout(6000)
        new_page = len(page.context.pages) > pages0
        moved = page.url != url0
        layer = (await page.evaluate(JS_LAYERS)) > layers0 if not moved else False
        link = c["buy"]["link"]
        if moved:
            ok = not link or link.rstrip("/").split("?")[0] in page.url
            add(f"CTA: Buy ({who})", "PASS" if ok else "FAIL", f"버튼 '{c['buy']['text']}'",
                expected=f"구매 페이지 {link} 로 이동", actual=page.url,
                screenshot="" if ok else await snap(f"cta_buy{n}"))
        elif new_page or layer:
            add(f"CTA: Buy ({who})", "PASS", f"버튼 '{c['buy']['text']}' - {'새 창' if new_page else '구매 레이어'} 노출",
                expected=f"구매 페이지 {link} 이동 또는 구매 레이어", actual="새 창 열림" if new_page else "레이어 노출",
                screenshot=await snap(f"cta_buy{n}_layer"))
            for extra in page.context.pages[pages0:]:
                await extra.close()
        else:
            add(f"CTA: Buy ({who})", "FAIL", f"버튼 '{c['buy']['text']}' 눌러도 반응 없음",
                expected=f"구매 페이지 {link} 이동 또는 구매 레이어", actual="이동/레이어 없음",
                screenshot=await snap(f"cta_buy{n}"))


# ------------------------------------------------------------------ 빈 제품 칸 채우기
async def fill_empty_columns(page, add, snap, max_cols=3, wait_ms=5000):
    filled = []
    for n in range(1, max_cols + 1):
        c = await page.evaluate(JS_COLUMN, n)
        if not c or c["sku"]:
            continue
        card = page.locator(f"li.pdd43-master-compare__card-item.pdd43-master-compare__column--{n}").first
        field = card.locator("button.menu__select-field").first
        if not await field.count():
            continue
        try:
            await field.scroll_into_view_if_needed()
            await field.click()
            await page.wait_for_timeout(1000)
            opts = await page.evaluate(JS_OPTIONS, n)
            if not opts:
                await page.keyboard.press("Escape")
                continue
            await card.locator(".menu__list-option").nth(opts[0]["i"]).click()
            await page.wait_for_timeout(wait_ms)
            filled.append(f"{n}열 '{opts[0]['text']}'")
        except Exception as e:
            add(f"빈 칸 채우기: {n}열", "오류", f"{type(e).__name__}: {e}", screenshot=await snap(f"fill{n}"))
    if filled:
        add("기본 제품 없음 → 드롭다운으로 선택", "PASS" if len(filled) >= 2 else "확인필요",
            "Compare 페이지에 기본 선택 제품이 없어 드롭다운 첫 모델로 채운 뒤 검증", actual=", ".join(filled))
    return len(filled)


# ------------------------------------------------------------------ 페이지 1개 검증
async def check_page(ctx, host, site, spec, opts, shots_dir):
    group, key, label, pf_path, cmp_path = spec
    rows = []
    tag = f"{site}_{key}"

    def add(check, result, detail="", expected="", actual="", screenshot=""):
        rows.append({"site": site, "group": group, "page": label, "check": check, "expected": expected,
                     "actual": actual, "result": result, "detail": detail, "screenshot": screenshot})

    page = await ctx.new_page()
    try:
        cmp_url = f"{host}/{site}/{cmp_path}"
        # 1) 진입
        if pf_path:
            pf_resp = await page.goto(f"{host}/{site}/{pf_path}", wait_until="domcontentloaded", timeout=TIMEOUTS["page_load"])
            await page.wait_for_timeout(5000)
            await dismiss_popups(page)
            pf_ok = pf_resp is not None and pf_resp.status < 400 and f"/{site}/" in page.url
            hrefs = await page.evaluate(JS_PF_COMPARE_TAB) if pf_ok else []
            if hrefs:
                cmp_url = hrefs[0] if hrefs[0].startswith("http") else host + hrefs[0]
                link = page.locator(f"a.tab__item-title[href='{hrefs[0]}']").locator("visible=true").first
                if not await link.count():
                    link = page.locator(f"a[href='{hrefs[0]}']").locator("visible=true").first
                await link.scroll_into_view_if_needed()
                await link.click()
                await page.wait_for_load_state("domcontentloaded", timeout=TIMEOUTS["page_load"])
                ok = "/compare" in page.url
                add("진입: PF Compare 탭", "PASS" if ok else "FAIL", expected="Compare 페이지로 이동",
                    actual=page.url, screenshot="" if ok else await shot(page, shots_dir, f"{tag}_entry"))
                if not ok:
                    await page.goto(cmp_url, wait_until="domcontentloaded", timeout=TIMEOUTS["page_load"])
            else:
                add("진입: PF Compare 탭", "해당 없음" if not pf_ok else "FAIL",
                    f"이 나라에 PF 가 없음 ({pf_resp.status if pf_resp else '응답 없음'}, {page.url})" if not pf_ok
                    else "PF 상단에 Compare 탭(/compare 링크)이 없음 → 고정 주소로 진입해 계속 확인",
                    expected="Compare 탭", actual="없음",
                    screenshot="" if not pf_ok else await shot(page, shots_dir, f"{tag}_notab"))
                cmp_resp = await page.goto(cmp_url, wait_until="domcontentloaded", timeout=TIMEOUTS["page_load"])
                if not cmp_resp or cmp_resp.status >= 400:
                    add("Compare 페이지", "해당 없음", f"이 나라에 Compare 페이지가 없음 ({cmp_resp.status if cmp_resp else ''})",
                        actual=cmp_url)
                    return rows
        else:
            resp = await page.goto(cmp_url, wait_until="domcontentloaded", timeout=TIMEOUTS["page_load"])
            if not resp or resp.status >= 400:
                add("진입: 고정 주소", "해당 없음", f"이 나라에 Compare 페이지가 없음 ({resp.status if resp else ''})",
                    actual=cmp_url)
                return rows
            add("진입: 고정 주소", "PASS", "PF 에 Compare 탭 없는 카테고리", expected="200", actual=str(resp.status))
        await page.wait_for_timeout(opts["wait"] * 1000)
        await dismiss_popups(page)

        st = await page.evaluate(JS_STATE)
        if not st.get("hasRoot"):
            add("Compare 페이지", "FAIL", "Compare models 영역(.pdd43-master-compare)이 없음 - 새 Compare 미배포 또는 다른 페이지",
                actual=page.url, screenshot=await shot(page, shots_dir, f"{tag}_noroot"))
            return rows

        empty = [n for n in (1, 2, 3) if (await page.evaluate(JS_COLUMN, n) or {"sku": "x"})["sku"] == ""]
        if empty:
            async def _snap(name):
                return await shot(page, shots_dir, f"{tag}_{name}")
            if await fill_empty_columns(page, add, _snap):
                st = await page.evaluate(JS_STATE)
        products = " / ".join(st["names"] or st["codes"])

        # 2) Key Specs
        if not st["featVisible"] or not st["cols"]:
            add("Key Specs 노출", "FAIL", f"비교 제품: {products}", expected="Key Specs 섹션 노출",
                actual="섹션 없음/미노출", screenshot=await shot(page, shots_dir, f"{tag}_keyspecs"))
        else:
            empty_cols = [c for c, v in st["cols"].items() if v["filled"] == 0]
            detail = ", ".join(f"{c}열 값 {v['filled']}/{v['filled'] + v['empty']}" for c, v in sorted(st["cols"].items()))
            ok = not empty_cols
            loc = page.locator(".pdd43-master-compare__feature").first
            await loc.scroll_into_view_if_needed()
            add("Key Specs 노출", "PASS" if ok else "FAIL", f"비교 제품: {products}",
                expected="제품마다 Key Specs 값 있음", actual=f"'{st['title']}' 항목 {st['rows']}개 / {detail}",
                screenshot=await shot(page, shots_dir, f"{tag}_keyspecs") if (not ok or opts["shots_all"]) else "")

        # 3)·4) 번호 → 하단 디스클레이머 → 화살표 → 번호
        nums = st["nums"][: opts["max_anchors"]] if opts["max_anchors"] else st["nums"]
        if not opts["anchor"]:
            nums = []
        elif not nums:
            add("앵커: 번호 → 디스클레이머", "확인불가", "Key Specs 에 디스클레이머 번호가 없음")
        missing = [n for n in st["nums"] if n not in st["bottom"]]
        if missing and opts["anchor"]:
            add("앵커: 하단 디스클레이머 존재", "FAIL", f"하단에 없는 번호: {missing}",
                expected=str(st["nums"]), actual=str(st["bottom"]))
        for n in nums:
            num_sel = f"a#masterCompareBottomDisclaimerGoBack{n}"
            disc_sel = f"#masterCompareBottomDisclaimer{n}"
            arrow_sel = f"{disc_sel} a.pdd43-master-compare__disclaimer-number-cta-top"
            num = page.locator(num_sel).locator("visible=true").first
            try:
                await num.scroll_into_view_if_needed(timeout=5000)
                await page.wait_for_timeout(500)
                before = await page.evaluate("() => Math.round(window.scrollY)")
                await num.click(timeout=5000)
            except Exception as e:
                add(f"앵커 {n}: 번호 → 디스클레이머", "FAIL", f"번호 {n} 클릭 실패: {type(e).__name__}",
                    screenshot=await shot(page, shots_dir, f"{tag}_n{n}_click"))
                continue
            await wait_scroll_settle(page, opts["anchor_wait"])
            v = await page.evaluate(JS_IN_VIEW, disc_sel)
            ok = v["inView"] and v["scrollY"] > before
            add(f"앵커 {n}: 번호 → 디스클레이머", "PASS" if ok else "FAIL",
                expected=f"하단 디스클레이머 {n} 이 화면에 보임",
                actual=f"보임={v['inView']}, 위치 top={v['top']}, 스크롤 {before}→{v['scrollY']}",
                screenshot="" if ok else await shot(page, shots_dir, f"{tag}_n{n}_down"))
            arrow = page.locator(arrow_sel).first
            if not await arrow.count():
                add(f"앵커 {n}: 화살표 → 번호", "FAIL", "디스클레이머에 화살표(Go to feature)가 없음",
                    screenshot=await shot(page, shots_dir, f"{tag}_n{n}_noarrow"))
                continue
            try:
                if not ok:
                    await arrow.scroll_into_view_if_needed(timeout=5000)
                    await page.wait_for_timeout(500)
                before2 = await page.evaluate("() => Math.round(window.scrollY)")
                await arrow.click(timeout=5000)
            except Exception as e:
                add(f"앵커 {n}: 화살표 → 번호", "FAIL", f"화살표 클릭 실패: {type(e).__name__}",
                    screenshot=await shot(page, shots_dir, f"{tag}_n{n}_arrowclick"))
                continue
            await wait_scroll_settle(page, opts["anchor_wait"])
            v2 = await page.evaluate(JS_IN_VIEW, num_sel)
            ok2 = v2["inView"] and v2["scrollY"] < before2
            add(f"앵커 {n}: 화살표 → 번호", "PASS" if ok2 else "FAIL",
                expected=f"상단 번호 {n} 이 화면에 보임",
                actual=f"보임={v2['inView']}, 위치 top={v2['top']}, 스크롤 {before2}→{v2['scrollY']}",
                screenshot="" if ok2 else await shot(page, shots_dir, f"{tag}_n{n}_up"))

        async def snap(name):
            return await shot(page, shots_dir, f"{tag}_{name}")

        sections = []
        if opts["diff"]:
            sections.append(("차이점", lambda: check_differences(page, add, snap, opts["shots_all"])))
        if opts["allspecs"]:
            sections.append(("All Specs", lambda: check_all_specs(page, add, snap)))
        if opts["dropdown"]:
            sections.append(("드롭다운", lambda: check_dropdown(page, add, snap, opts["dropdown_col"])))
        if opts["cta"]:
            cols = opts["cta_cols"] or sorted(int(c) for c in st["cols"] if str(c).isdigit()) or [1, 2, 3]
            sections.append(("CTA", lambda: check_cta(page, add, snap, cmp_url, cols, opts["wait"], opts["click_buy"])))
        for name, run in sections:
            try:
                await run()
            except Exception as e:
                add(f"{name}: 오류", "오류", f"{type(e).__name__}: {e}", screenshot=await snap(f"{name}_error"))
    except Exception as e:
        add("페이지 오류", "오류", f"{type(e).__name__}: {e}", screenshot=await shot(page, shots_dir, f"{tag}_error"))
    finally:
        await page.close()
    return rows


# ------------------------------------------------------------------ 리포트
def write_report(path, rows, args):
    from collections import Counter, OrderedDict
    wb = Workbook()
    ws = wb.active
    ws.title = "국가별 요약"
    page_names = [p[2] for p in select_pages(args)]
    by_site = OrderedDict()
    for r in rows:
        by_site.setdefault(r["site"], OrderedDict()).setdefault(r["page"], []).append(r)
    srows = []
    for site, pages in by_site.items():
        verdicts = []
        for pn in page_names:
            rs = pages.get(pn, [])
            if not rs:
                verdicts.append("-")
            elif all(x["result"] == "해당 없음" for x in rs):
                verdicts.append("해당 없음")
            elif any(x["result"] in ("FAIL", "오류") for x in rs):
                verdicts.append("FAIL")
            else:
                verdicts.append("PASS")
        fails = sum(1 for rs in pages.values() for x in rs if x["result"] in ("FAIL", "오류"))
        srows.append([site, *verdicts, fails,
                      "FAIL" if fails else ("해당 없음" if set(verdicts) <= {"해당 없음", "-"} else "PASS")])
    _sheet(ws, ["사이트", *page_names, "FAIL 수", "국가 판정"], srows, [9] + [13] * len(page_names) + [8, 10])
    for row in ws.iter_rows(min_row=2, min_col=2, max_col=len(page_names) + 3):
        for c in row:
            color = FILL.get(str(c.value))
            if color:
                c.fill = PatternFill("solid", fgColor=color)
    ws = wb.create_sheet("페이지별 요약")
    by = OrderedDict()
    for r in rows:
        by.setdefault((r["site"], r["group"], r["page"]), []).append(r)
    hdr = ["사이트", "그룹", "페이지", "진입", "Key Specs", "앵커 번호→디스클레이머", "화살표→번호",
           "차이점 보기", "All Specs", "드롭다운", "Buy", "Learn more", "FAIL 수", "판정"]
    out = []
    for (site, group, pg), rs in by.items():
        def res(prefix):
            sub = [x for x in rs if x["check"].startswith(prefix)]
            if not sub:
                return "-"
            c = Counter(x["result"] for x in sub)
            return "PASS" if set(c) == {"PASS"} else ", ".join(f"{k} {v}" for k, v in c.items())
        down = [x for x in rs if "번호 → 디스클레이머" in x["check"] and x["result"] != "확인불가"]
        up = [x for x in rs if "화살표 → 번호" in x["check"]]
        fails = sum(1 for x in rs if x["result"] in ("FAIL", "오류"))
        no_num = not down and any(x["result"] == "확인불가" for x in rs)
        skip = getattr(args, "no_anchor", False)
        out.append([site, group, pg, res("진입"), res("Key Specs"),
                    "생략" if skip else "번호 없음" if no_num else f"{sum(x['result'] == 'PASS' for x in down)}/{len(down)}",
                    "생략" if skip else "번호 없음" if no_num else f"{sum(x['result'] == 'PASS' for x in up)}/{len(up)}",
                    res("차이점"), res("All Specs"), res("드롭다운"), res("CTA: Buy"), res("CTA: Learn more"),
                    fails, "FAIL" if fails else ("해당 없음" if all(x["result"] == "해당 없음" for x in rs) else "PASS")])
    _sheet(ws, hdr, out, [8, 7, 15, 12, 12, 20, 14, 14, 14, 14, 14, 14, 8, 8], result_col=13)
    ws = wb.create_sheet("상세")
    _sheet(ws, ["사이트", "그룹", "페이지", "검증", "기대", "실제", "결과", "비고", "캡처"],
           [[r["site"], r["group"], r["page"], r["check"], r["expected"], r["actual"], r["result"], r["detail"],
             r["screenshot"]] for r in rows], [8, 7, 15, 26, 30, 50, 9, 45, 40], result_col=6)
    ws = wb.create_sheet("실행 정보")
    for r in [["실행 시각", dt.datetime.now().strftime("%Y-%m-%d %H:%M")], ["환경", f"live ({BASE_URL})"],
              ["사이트", f"{len(args.sites)}개: " + " ".join(args.sites)], ["페이지", ", ".join(p[2] for p in select_pages(args))],
              ["앵커 확인", "생략" if args.no_anchor else
               (f"번호 최대 {args.max_anchors}개" if args.max_anchors else "Key Specs 의 모든 번호")],
              ["차이점 보기", "생략" if args.no_diff else "켜기 → 같은 값 행 숨김·다른 값 행 노출·파란 글자 → 끄기"],
              ["All Specs", "생략" if args.no_allspecs else "카테고리 첫·마지막 버튼: 누르면 접힘 → 다시 누르면 펼침"],
              ["드롭다운", "생략" if args.no_dropdown else f"{args.dropdown_col}열에서 다른 모델 선택 → 모델명·코드·Key Specs 변경"],
              ["Buy / Learn more", "생략" if args.no_cta else
               f"{' '.join(map(str, args.cta_cols)) + '열' if args.cta_cols else '모든 열'} - Learn more 클릭 후 이동 확인, "
               + ("Buy 클릭 후 이동/레이어 확인" if args.click_buy else "Buy 는 클릭하지 않고 구매 링크 응답만 확인")],
              ["해당 없음", "그 나라에 PF 나 Compare 페이지가 없는 경우 (FAIL 아님)"]]:
        ws.append(r)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 120
    wb.save(path)


def _sheet(ws, header, rows, widths, result_col=None):
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F4E78")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in rows:
        ws.append(r)
        if result_col is not None:
            color = next((col for k, col in FILL.items() if str(r[result_col]).startswith(k)), None)
            if color:
                ws.cell(ws.max_row, result_col + 1).fill = PatternFill("solid", fgColor=color)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def select_pages(args):
    ps = PAGES
    if args.groups:
        ps = [p for p in ps if p[0].lower() in {g.lower() for g in args.groups}]
    if args.pages:
        ps = [p for p in ps if p[1] in set(args.pages)]
    return ps


# ------------------------------------------------------------------ main
async def run_site(browser, args, site, pages, opts, run_dir):
    """사이트 1개: 새 브라우저 컨텍스트에서 모든 페이지 검증 → sites/<site>.json 저장."""
    ctx = await browser.new_context(viewport={"width": args.width, "height": args.height})
    rows = []
    shots_dir = run_dir / "screenshots"
    try:
        for spec in pages:
            r = await check_page(ctx, BASE_URL, site, spec, opts, shots_dir)
            rows += r
            f = sum(1 for x in r if x["result"] in ("FAIL", "오류"))
            na = bool(r) and all(x["result"] == "해당 없음" for x in r)
            print(f"  [{site}] {spec[2]}: " + ("해당 없음" if na else f"검증 {len(r)}건, FAIL/오류 {f}건"))
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
    pages = select_pages(args)
    opts = {"wait": args.wait, "anchor_wait": args.anchor_wait, "max_anchors": args.max_anchors,
            "shots_all": args.shots_all, "anchor": not args.no_anchor, "diff": not args.no_diff,
            "allspecs": not args.no_allspecs, "dropdown": not args.no_dropdown, "dropdown_col": args.dropdown_col,
            "cta": not args.no_cta, "cta_cols": args.cta_cols, "click_buy": args.click_buy}
    todo = [s for s in args.sites if not (run_dir / "sites" / f"{s}.json").exists()]
    print(f"사이트 {len(args.sites)}개 (이번에 실행 {len(todo)}개, 동시 {args.parallel}개), 페이지 {len(pages)}개 → {run_dir}")
    sem = asyncio.Semaphore(max(1, args.parallel))
    async with async_playwright() as p:
        b = await p.chromium.launch(headless=args.headless, slow_mo=TIMEOUTS["slow_mo"])

        async def one(site):
            async with sem:
                t0 = dt.datetime.now()
                print(f"[{site}] 시작")
                try:
                    await run_site(b, args, site, pages, opts, run_dir)
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
    out = run_dir / f"ComparePage_LIVE_{run_dir.name.split('_', 2)[2]}.xlsx"
    write_report(out, rows, args)
    print("결과 파일:", out)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--sites", nargs="*", default=[], help="사이트코드 (생략하면 Live_기준 시트의 전체 사이트)")
    ap.add_argument("--exclude-sites", nargs="*", default=[], help="제외할 사이트코드")
    ap.add_argument("--mapping-list", default=MAPPING_LIST_FILE, help="사이트 목록을 읽을 매핑 리스트 파일")
    ap.add_argument("--groups", nargs="*", default=[], help="MX / PC / TV / Audio / DA 중 선택")
    ap.add_argument("--pages", nargs="*", default=[], help="페이지 키: " + ", ".join(p[1] for p in PAGES))
    ap.add_argument("--parallel", type=int, default=2, help="동시에 돌릴 사이트 수 (기본 2)")
    ap.add_argument("--resume", nargs="?", const="latest", default=None, help="이전 실행 이어서 (폴더 지정 가능)")
    ap.add_argument("--max-anchors", type=int, default=3,
                    help="페이지당 확인할 디스클레이머 번호 수 (기본 3 = 앞에서부터 3개, 0 = 전부)")
    ap.add_argument("--wait", type=int, default=TIMEOUTS["landing_wait"], help="Compare 페이지 로딩 대기(초)")
    ap.add_argument("--anchor-wait", type=int, default=TIMEOUTS["anchor_wait"], help="번호/화살표 클릭 후 대기(ms)")
    ap.add_argument("--headless", action="store_true", help="브라우저 창 없이 실행")
    ap.add_argument("--shots-all", action="store_true", help="PASS 인 Key Specs 화면도 캡처")
    ap.add_argument("--no-anchor", action="store_true", help="번호/화살표 앵커 검증 생략")
    ap.add_argument("--no-diff", action="store_true", help="Show me the differences 검증 생략")
    ap.add_argument("--no-allspecs", action="store_true", help="All Specs 접기/펼치기 검증 생략")
    ap.add_argument("--no-dropdown", action="store_true", help="제품명 드롭다운 검증 생략")
    ap.add_argument("--dropdown-col", type=int, default=1, help="드롭다운으로 모델을 바꿀 제품 열 (기본 1)")
    ap.add_argument("--no-cta", action="store_true", help="Buy / Learn more 검증 생략")
    ap.add_argument("--cta-cols", nargs="*", type=int, default=[], help="CTA 를 확인할 제품 열 (기본: 모든 열)")
    ap.add_argument("--click-buy", action="store_true", help="Buy 를 실제로 클릭 (라이브에서는 장바구니에 담김)")
    ap.add_argument("--width", type=int, default=1600)
    ap.add_argument("--height", type=int, default=1000)
    ap.add_argument("--out-dir", default=HERE / "result")
    asyncio.run(amain(ap.parse_args()))


if __name__ == "__main__":
    main()
