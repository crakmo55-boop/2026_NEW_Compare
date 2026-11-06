# -*- coding: utf-8 -*-
"""[라이브] Compare 팝업 Key Specs 노출 확인 - 라이브날 89개국(91개 사이트코드) 전체 확인용.

원천 데이터 비교는 하지 않는다. 라이브(www) 데이터만 보고 아래 두 가지를 확인한다.
1) Key Specs 노출: 'Live_기준' 시트의 DISPLAY_STATUS=Y 제품마다 Compare 팝업에 Key Specs 가 나오는지
     - 안 나오면 PF 카드 API 로 제품 등록 여부를 확인해 'Key Spec 미노출' / '라이브 미등록' 으로 구분
2) 비교 팝업 로직: 같은 PF 묶음(사이트 + IA 앞 4자리)에서 Key Specs 가 있는 기준 제품 A 를 잡고
     - 2개 비교 A+X : X 에 Key Specs 가 있으면 노출, 없으면 Key Specs 섹션 전체(A 포함) 미노출이어야 PASS
     - 3개 비교 A+B+X (X 는 Key Specs 없는 제품) : 전체 미노출이어야 PASS
                A+B+C (모두 Key Specs 있음)       : 노출 + 세 제품 값 모두 있어야 PASS

실행 전 사이트마다 Compare API 응답을 먼저 확인하고, 404 면 그 사이트는 'API 미응답' 으로 표시하고 넘어간다.
사이트 단위로 결과를 저장하므로 중간에 끊겨도 --resume 으로 남은 사이트부터 이어서 돌릴 수 있다.

사용 예
  python run_popup_keyspec_live.py --discover                  # 라이브 PF 에서 팝업 API 주소 확인 (기본 uk)
  python run_popup_keyspec_live.py --sites uk                  # 한 나라만
  python run_popup_keyspec_live.py --sites uk de fr --category "Mobile > Smartphones"
  python run_popup_keyspec_live.py                             # 91개 사이트 전체
  python run_popup_keyspec_live.py --resume                    # 마지막 실행을 이어서
  python run_popup_keyspec_live.py --api-base <주소>           # discover 로 찾은 주소로 실행
결과: compare/result/run_<날짜시간>/KeySpec_LIVE_<날짜시간>.xlsx  (사이트별 중간 결과는 같은 폴더 sites/)
"""
import argparse
import datetime as dt
import json
import sys
from collections import Counter, OrderedDict
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from config import (
    BASE_URL,
    FILL_COLORS,
    HERE,
    LIVE_SEARCH_API,
    MAPPING_LIST_FILE,
)
from ks_live_api import ApiNotFound, LiveApi

MAPPING_LIST = MAPPING_LIST_FILE
FILL = FILL_COLORS


# ------------------------------------------------------------------ 대상 목록
def load_targets(path=MAPPING_LIST_FILE, sites=None, category=""):
    """'Live_기준' 시트의 DISPLAY_STATUS=Y 제품 → {사이트: [제품, ...]} (시트 순서 유지)."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = next(w for w in wb.worksheets if w.title.upper().startswith("LIVE"))
    rows = ws.iter_rows(values_only=True)
    hdr = [str(h or "").strip() for h in next(rows)]
    i = {h: n for n, h in enumerate(hdr)}
    out = OrderedDict()
    want = {x.lower() for x in sites} if sites else None
    for n, r in enumerate(rows, start=2):
        site = str(r[i["SITE_CD"]] or "").strip()
        if str(r[i["DISPLAY_STATUS"]] or "").strip().upper() != "Y":
            continue
        if want and site.lower() not in want:
            continue
        cat = str(r[i["CATEGORY_NAME"]] or "").strip()
        if category and not cat.startswith(category):
            continue
        out.setdefault(site, []).append({
            "row": n, "site": site, "ia": str(r[i["IA_CD"]] or "").strip(), "category": cat,
            "fmy_id": str(r[i["FMY_ID"]] or "").strip(), "fmy_mkt": str(r[i["FMY_MKT_NM"]] or "").strip(),
            "model": str(r[i["REP_MODEL"]] or "").strip(),
        })
    wb.close()
    return out


# ------------------------------------------------------------------ 사이트 1개 검증
def check_set(api, site, members, ks_of):
    """비교 묶음(2~3개) 1건 검증. ks_of: 모델 → 단독 Key Specs 여부."""
    models = [m["model"] for m in members]
    expect = all(ks_of[m] for m in models)
    row = {"site": site, "size": len(models), "models": " + ".join(models),
           "names": " + ".join(m["fmy_mkt"] for m in members),
           "no_ks": ", ".join(m for m in models if not ks_of[m]),
           "expected": "노출" if expect else "전체 미노출", "actual": "", "count": 0, "values": "", "result": "", "note": ""}
    try:
        r = api.compare(site, models)
        row.update(actual="노출" if r["shown"] else "전체 미노출", count=r["count"],
                   values=", ".join(str(r["values"][m]) for m in models))
        if expect != r["shown"]:
            row["result"] = "FAIL"
            row["note"] = ("Key Specs 없는 제품이 섞였는데 팝업에 Key Specs 가 나옴" if r["shown"]
                           else "모두 Key Specs 가 있는데 팝업에 Key Specs 가 안 나옴")
        elif r["shown"] and any(r["values"][m] == 0 for m in models):
            row["result"] = "FAIL"
            row["note"] = "Key Specs 는 나오지만 값이 빈 제품이 있음: " + ", ".join(m for m in models if r["values"][m] == 0)
        else:
            row["result"] = "PASS"
    except Exception as e:
        row.update(result="오류", note=f"{type(e).__name__}: {e}")
    return row


def run_site(api, site, items, opts):
    """반환: {'site', 'status', 'products': [...], 'sets': [...]}"""
    res = {"site": site, "status": "", "products": [], "sets": []}

    # 사전 확인: 이 사이트에 새 Compare API 가 있는지
    probe = [t for t in items if t["model"]][:3]
    try:
        for t in probe:
            api.compare(site, [t["model"]])
            break
        res["status"] = "확인"
    except ApiNotFound as e:
        res["status"] = "API 미응답(404)"
        res["products"] = [{**t, "result": "API 미응답", "ks_count": "", "note": str(e)} for t in items]
        return res
    except Exception as e:
        res["status"] = f"API 오류: {type(e).__name__}"

    # 1) 제품별 Key Specs
    def one(t):
        p = {**t, "result": "", "ks_count": 0, "note": ""}
        if not t["model"]:
            p["result"] = "REP_MODEL 없음"
            return p
        try:
            r = api.compare(site, [t["model"]])
            p["ks_count"] = r["values"][t["model"]]
            if p["ks_count"] > 0:
                p["result"] = "Key Spec 노출"
            else:
                p["result"] = "Key Spec 미노출" if api.card(site, t["model"]) else "라이브 미등록"
        except Exception as e:
            p.update(result="오류", note=f"{type(e).__name__}: {e}")
        return p

    with ThreadPoolExecutor(opts["workers"]) as ex:
        products = list(ex.map(one, items))
    res["products"] = products
    if opts["no_sets"]:
        return res

    # 2) 비교 팝업 로직 (PF 묶음 = IA 앞 4자리)
    ks_of = {p["model"]: p["result"] == "Key Spec 노출" for p in products if p["model"]}
    groups = OrderedDict()
    for p in products:
        if p["result"] in ("Key Spec 노출", "Key Spec 미노출"):
            groups.setdefault(p["ia"][:4], []).append(p)
    jobs = []
    for g, ps in groups.items():
        have = [p for p in ps if ks_of[p["model"]]]
        lack = [p for p in ps if not ks_of[p["model"]]]
        forced = [p for p in have if p["model"] in opts["pair_base"]]
        base = (forced or have or [None])[0]
        if base is None:
            res["sets"].append({"site": site, "size": 2, "models": "", "names": f"PF {g}0000", "no_ks": "",
                                "expected": "", "actual": "", "count": 0, "values": "", "result": "확인불가",
                                "note": "이 PF 묶음에 Key Specs 있는 제품이 없어 기준을 못 잡음", "group": g})
            continue
        for p in ps:
            if p is not base:
                jobs.append((g, [base, p]))
        if not opts["no_triple"]:
            others = [p for p in have if p is not base]
            if others:
                b = others[0]
                for x in lack:
                    jobs.append((g, [base, b, x]))
                if len(others) >= 2:
                    jobs.append((g, [base, b, others[1]]))

    def run_job(job):
        g, members = job
        return {**check_set(api, site, members, ks_of), "group": g}

    with ThreadPoolExecutor(opts["workers"]) as ex:
        res["sets"] += list(ex.map(run_job, jobs))
    return res


# ------------------------------------------------------------------ 리포트
HDR_FILL = PatternFill("solid", fgColor="1F4E78")


def write_sheet(ws, header, rows, widths, result_col=None):
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in rows:
        ws.append(r)
        if result_col is not None:
            v = str(r[result_col])
            color = next((col for k, col in FILL.items() if v.startswith(k)), None)
            if color:
                ws.cell(ws.max_row, result_col + 1).fill = PatternFill("solid", fgColor=color)
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_report(path, results, args):
    wb = Workbook()
    ws = wb.active
    ws.title = "국가별 요약"
    rows = []
    for r in results:
        pc = Counter(p["result"] for p in r["products"])
        sc = Counter(s["result"] for s in r["sets"])
        rows.append([r["site"], r["status"], len(r["products"]), pc["Key Spec 노출"], pc["Key Spec 미노출"],
                     pc["라이브 미등록"], pc["REP_MODEL 없음"], pc["API 미응답"], pc["오류"],
                     len(r["sets"]), sc["PASS"], sc["FAIL"], sc["확인불가"], sc["오류"],
                     "FAIL" if (sc["FAIL"] or pc["오류"] or sc["오류"]) else
                     ("API 미응답" if r["status"].startswith("API 미응답") else "PASS")])
    write_sheet(ws, ["사이트", "API 상태", "대상 제품", "Key Spec 노출", "Key Spec 미노출", "라이브 미등록",
                     "REP_MODEL 없음", "API 미응답", "오류(제품)", "비교 건수", "비교 PASS", "비교 FAIL",
                     "비교 확인불가", "비교 오류", "국가 판정(비교 로직)"],
                rows, [10, 16, 9, 10, 10, 10, 10, 10, 9, 9, 9, 9, 10, 9, 14], result_col=14)

    prod_hdr = ["사이트", "엑셀 행", "CATEGORY_NAME", "IA_CD", "FMY_ID", "FMY_MKT_NM", "REP_MODEL",
                "결과", "Key Specs 항목 수", "비고"]
    prod_w = [9, 8, 42, 10, 9, 45, 18, 15, 9, 40]

    def prow(p):
        return [p["site"], p["row"], p["category"], p["ia"], p["fmy_id"], p["fmy_mkt"], p["model"],
                p["result"], p["ks_count"], p["note"]]

    ws = wb.create_sheet("Key Spec 미노출")
    write_sheet(ws, prod_hdr, [prow(p) for r in results for p in r["products"]
                               if p["result"] in ("Key Spec 미노출", "오류")], prod_w, result_col=7)

    set_hdr = ["사이트", "PF 묶음(IA)", "비교 수", "비교 모델", "비교 제품명", "Key Specs 없는 모델",
               "기대(팝업 Key Specs)", "실제(팝업 Key Specs)", "결과", "팝업 항목 수", "모델별 값 있는 항목 수", "비고"]
    set_w = [9, 10, 7, 45, 60, 20, 14, 14, 9, 9, 14, 45]

    def srow(s):
        return [s["site"], (s.get("group") or "") + "0000", f"{s['size']}개", s["models"], s["names"], s["no_ks"],
                s["expected"], s["actual"], s["result"], s["count"], s["values"], s["note"]]

    ws = wb.create_sheet("비교 FAIL")
    write_sheet(ws, set_hdr, [srow(s) for r in results for s in r["sets"] if s["result"] != "PASS"],
                set_w, result_col=8)
    ws = wb.create_sheet("비교 전체(2·3개)")
    write_sheet(ws, set_hdr, [srow(s) for r in results for s in r["sets"]], set_w, result_col=8)
    ws = wb.create_sheet("제품 전체")
    write_sheet(ws, prod_hdr, [prow(p) for r in results for p in r["products"]], prod_w, result_col=7)

    ws = wb.create_sheet("실행 정보")
    pc = Counter(p["result"] for r in results for p in r["products"])
    sc = Counter(s["result"] for s in r["sets"])
    info = [
        ["실행 시각", dt.datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["환경", f"라이브 ({BASE_URL}) - 원천 데이터 비교 없음"],
        ["대상 목록", f"{args.mapping_list} > Live_기준 시트, DISPLAY_STATUS=Y"
                     + (f", 사이트 {' '.join(args.sites)}" if args.sites else ", 전체 사이트")
                     + (f", 카테고리 '{args.category}'" if args.category else "")],
        ["Compare API", f"{args.api_base}/master/compare"],
        ["판정 기준", "제품: 단독 Compare 시 Key Specs 값이 있으면 '노출'. 비교: 묶음 중 하나라도 Key Specs 가 없으면 "
                     "팝업 Key Specs 전체 미노출, 모두 있으면 노출(값 모두 있음)이어야 PASS"],
        ["사이트 수", len(results)],
        ["제품 수", sum(len(r["products"]) for r in results)],
    ] + [[f"제품: {k}", v] for k, v in pc.most_common()] + [[f"비교: {k}", v] for k, v in sc.most_common()]
    for r in info:
        ws.append(r)
    ws.column_dimensions["A"].width = 20
    ws.column_dimensions["B"].width = 140
    wb.save(path)


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--mapping-list", default=MAPPING_LIST_FILE, help="사이트 목록을 읽을 매핑 리스트 파일")
    ap.add_argument("--sites", nargs="*", default=[], help="검증할 사이트코드 (예: uk de ae_ar). 생략하면 전체")
    ap.add_argument("--exclude-sites", nargs="*", default=[], help="제외할 사이트코드")
    ap.add_argument("--category", default="", help="CATEGORY_NAME 앞부분 필터 (예: 'Mobile > Smartphones')")
    ap.add_argument("--api-base", default=LIVE_SEARCH_API, help=f"라이브 검색 API 주소 (기본 {LIVE_SEARCH_API})")
    ap.add_argument("--workers", type=int, default=8, help="동시 호출 수 (기본 8)")
    ap.add_argument("--no-sets", action="store_true", help="비교 팝업 로직(2·3개 비교) 생략, Key Specs 노출만")
    ap.add_argument("--no-triple", action="store_true", help="3개 비교 생략 (2개 비교만)")
    ap.add_argument("--pair-base", nargs="*", default=[], help="비교 기준으로 쓸 REP_MODEL")
    ap.add_argument("--resume", nargs="?", const="latest", default=None,
                    help="이전 실행 이어서 (폴더 지정 가능, 생략하면 마지막 실행)")
    ap.add_argument("--discover", action="store_true", help="라이브 PF 에서 Compare 팝업 API 주소만 확인하고 종료")
    ap.add_argument("--discover-site", default="uk", help="--discover 에 쓸 사이트 (기본 uk)")
    ap.add_argument("--out-dir", default=HERE / "result")
    args = ap.parse_args()

    if args.discover:
        from ks_discover import discover
        discover(args.discover_site)
        return

    out_root = Path(args.out_dir)
    if args.resume:
        runs = sorted(out_root.glob("run_*")) if args.resume == "latest" else [Path(args.resume)]
        if not runs or not runs[-1].exists():
            sys.exit("이어서 실행할 이전 결과 폴더가 없습니다.")
        run_dir = runs[-1]
    else:
        run_dir = out_root / f"run_{dt.datetime.now():%Y%m%d_%H%M%S}"
    sites_dir = run_dir / "sites"
    sites_dir.mkdir(parents=True, exist_ok=True)

    targets = load_targets(args.mapping_list, args.sites, args.category)
    for s in args.exclude_sites:
        targets.pop(s, None)
    print(f"대상: 사이트 {len(targets)}개, 제품 {sum(len(v) for v in targets.values())}개  →  {run_dir}")

    api = LiveApi(args.api_base)
    opts = {"workers": args.workers, "no_sets": args.no_sets, "no_triple": args.no_triple,
            "pair_base": set(args.pair_base)}
    for n, (site, items) in enumerate(targets.items(), 1):
        f = sites_dir / f"{site}.json"
        if f.exists():
            print(f"  [{n}/{len(targets)}] {site}: 이전 결과 사용")
            continue
        t0 = dt.datetime.now()
        res = run_site(api, site, items, opts)
        f.write_text(json.dumps(res, ensure_ascii=False), encoding="utf-8")
        pc = Counter(p["result"] for p in res["products"])
        sc = Counter(s["result"] for s in res["sets"])
        print(f"  [{n}/{len(targets)}] {site}: {res['status']} | 노출 {pc['Key Spec 노출']} / 미노출 {pc['Key Spec 미노출']}"
              f" / 미등록 {pc['라이브 미등록']} | 비교 PASS {sc['PASS']} FAIL {sc['FAIL']}"
              f" | {(dt.datetime.now() - t0).seconds}초")

    results = [json.loads((sites_dir / f"{s}.json").read_text(encoding="utf-8"))
               for s in targets if (sites_dir / f"{s}.json").exists()]
    out = run_dir / f"KeySpec_LIVE_{run_dir.name[4:]}.xlsx"
    write_report(out, results, args)
    print("제품:", dict(Counter(p["result"] for r in results for p in r["products"])))
    print("비교:", dict(Counter(s["result"] for r in results for s in r["sets"])))
    print("결과 파일:", out)


if __name__ == "__main__":
    main()
