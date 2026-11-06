# -*- coding: utf-8 -*-
"""라이브(www.samsung.com) Compare 팝업 데이터 조회.

PF 에서 Compare 체크 → Compare 버튼을 누르면 팝업이 아래 API 를 호출하고,
응답의 resultData.keyFeatures 로 팝업의 'Key Specs' 섹션을 그린다. (QA pre-qa3 에서 화면/네트워크로 확인)
    {base}/master/compare?siteCode=<사이트>&pageType=PF&modelCodes=<모델1>,<모델2>[,<모델3>]
비교 모델 중 하나라도 Key Specs 가 없으면 keyFeatures 가 비어 Key Specs 섹션 전체가 안 나온다.

  QA   : https://searchapi-preqa.samsung.com/preqa3/v6/qa/b2c/product
  라이브: https://searchapi.samsung.com/v6/front/b2c/product   (기본값, --api-base 로 변경)
2026-10-22 확인: 라이브 card/detail 은 응답, master/compare 는 404 (새 Compare 미배포,
기존 팝업은 .../product/spec/compare 호출). 라이브날 주소는 --discover 로 확인.
"""
import json
import time
import urllib.error
import urllib.parse
import urllib.request

LIVE_SEARCH = "https://searchapi.samsung.com/v6/front/b2c/product"


class ApiNotFound(Exception):
    """Compare API 가 404 - 해당 사이트에 새 Compare 미배포 또는 주소 다름."""


class LiveApi:
    def __init__(self, api_base=LIVE_SEARCH, retries=3, timeout=40):
        self.base = api_base.rstrip("/")
        self.retries = retries
        self.timeout = timeout

    def _get(self, url):
        err = None
        for attempt in range(self.retries):
            try:
                with urllib.request.urlopen(url, timeout=self.timeout) as r:
                    return json.load(r)
            except urllib.error.HTTPError as e:
                if e.code == 404:
                    raise ApiNotFound(f"404 {url.split('?')[0]}")
                err = e
            except Exception as e:  # 네트워크 오류는 잠깐 쉬고 재시도
                err = e
            time.sleep(2 * (attempt + 1))
        raise err

    def compare(self, site, models):
        """모델 1~3개를 같이 Compare 팝업에 띄웠을 때의 Key Specs 상태.

        반환: {'shown': 팝업에 Key Specs 섹션이 나오는지, 'count': 항목 수,
               'values': {모델: 값이 있는 항목 수}}
        """
        url = f"{self.base}/master/compare?" + urllib.parse.urlencode({
            "siteCode": site, "pageType": "PF", "modelCodes": ",".join(models)})
        rd = self._get(url).get("response", {}).get("resultData") or {}
        kf = rd.get("keyFeatures") or []
        values = {m: 0 for m in models}

        def walk(items):
            for k in items or []:
                yield k
                yield from walk(k.get("children"))

        for k in walk(kf):
            for v in k.get("values") or []:
                if v.get("modelCode") in values and (_s(v.get("value")) or _s(v.get("icon"))):
                    values[v["modelCode"]] += 1
        return {"shown": bool(kf), "count": len(kf), "values": values}

    def card(self, site, model):
        """PF 카드 API 로 제품 등록 여부/제품군명 확인. 없으면 None."""
        url = f"{self.base}/card/detail/newhybris?" + urllib.parse.urlencode({
            "siteCode": site, "onlyRequestSkuYN": "N", "commonCodeYN": "N", "saleSkuYN": "N", "modelList": model})
        try:
            pl = (self._get(url).get("response", {}).get("resultData") or {}).get("productList") or []
        except ApiNotFound:
            return None
        return pl[0] if pl else None


def _s(v):
    return "" if v is None else str(v).strip()
