# -*- coding: utf-8 -*-
"""[설정 중앙화] Compare 라이브 자동화 공통 환경 설정 및 상수 정의 모듈."""
import os
from pathlib import Path

# 기준 디렉토리
HERE = Path(__file__).resolve().parent

# 1. 라이브 환경 도메인
BASE_URL = os.getenv("COMPARE_BASE_URL", "https://www.samsung.com")

# 2. 매핑 파일 경로 탐색
# 환경변수 -> 상위 'mapping table' 디렉토리 -> 로컬 fallback 순
def resolve_mapping_file():
    env_path = os.getenv("COMPARE_MAPPING_FILE")
    if env_path and Path(env_path).exists():
        return Path(env_path)
    
    # 기본 경로 1: 상위 3단계의 mapping table 폴더
    target_dir = HERE.parent.parent.parent / "mapping table"
    default_name = target_dir / "[VD,MX,DA] key_spec_family_mapping_list_251016.xlsx"
    if default_name.exists():
        return default_name
    
    # 기본 경로 2: mapping table 폴더 내 임의의 .xlsx 파일
    if target_dir.exists():
        excels = list(target_dir.glob("*.xlsx"))
        if excels:
            return excels[0]
            
    # 기본 경로 3: 로컬 data 폴더 fallback
    local_excel = HERE / "data" / "mapping_list.xlsx"
    return local_excel

MAPPING_LIST_FILE = resolve_mapping_file()

# 3. 백엔드 라이브 검색 API 엔드포인트
LIVE_SEARCH_API = os.getenv("COMPARE_LIVE_API", "https://searchapi.samsung.com/v6/front/b2c/product")

# 4. 통합 카테고리 마스터 테이블
# (그룹, 페이지 키, 화면 라벨, PF 경로, Compare 랜딩 경로)
CATEGORY_SPECS = [
    ("MX", "smartphones", "Smartphones", "smartphones/all-smartphones/", "smartphones/compare/"),
    ("MX", "tablets", "Tablets", "tablets/all-tablets/", "tablets/compare/"),
    ("MX", "watches", "Watches", "watches/all-watches/", "watches/compare/"),
    ("MX", "buds", "Galaxy Buds", "audio-sound/all-audio-sound/", "audio-sound/compare/"),
    ("PC", "galaxy-book", "Galaxy Book", "computers/all-computers/", "galaxy-book/galaxy-book-series/compare/"),
    ("PC", "monitors", "Monitors", "monitors/all-monitors/", "monitors/compare/"),
    ("TV", "tvs", "TVs", "tvs/all-tvs/", "tvs/compare/"),
    ("Audio", "soundbars", "Audio Devices", "audio-devices/all-audio-devices/", "audio-devices/compare/"),
    ("DA", "refrigerators", "Refrigerators", "refrigerators/all-refrigerators/", "refrigerators/compare/"),
    ("DA", "laundry", "Washers & Dryers", "washers-and-dryers/all-washers-and-dryers/", "washers-and-dryers/compare/"),
    ("DA", "dishwashers", "Dishwashers", "dishwashers/all-dishwashers/", "dishwashers/compare/"),
    ("DA", "cooking", "Cooking Appliances", "cooking-appliances/ovens/", "cooking-appliances/compare/"),
    ("DA", "microwaves", "Microwave Ovens", "microwave-ovens/all-microwave-ovens/", "microwave-ovens/compare/"),
    ("DA", "vacuums", "Vacuum Cleaners", "vacuum-cleaners/all-vacuum-cleaners/", "vacuum-cleaners/compare/"),
]

# 5. 공통 대기 시간 및 타임아웃 기본값
TIMEOUTS = {
    "page_load": 90000,         # goto 대기 (ms)
    "landing_wait": 10,         # Compare 랜딩 페이지 로딩 대기 (초)
    "pf_wait": 6,               # PF 로딩 대기 (초)
    "popup_wait": 3,            # 팝업 로딩 대기 (초)
    "anchor_wait": 1500,        # 번호/화살표 스크롤 대기 (ms)
    "chip_wait": 3000,          # 칩 클릭 후 변경 대기 (ms)
    "slow_mo": 0,               # 브라우저 동작 지연 (ms)
}

# 6. 리포트 서식 색상 팔레트
FILL_COLORS = {
    "PASS": "E2EFDA",
    "FAIL": "F8CBAD",
    "오류": "F8CBAD",
    "확인불가": "FFF2CC",
    "확인필요": "FFF2CC",
    "해당 없음": "D9D9D9",
    "Key Spec 노출": "E2EFDA",
    "Key Spec 미노출": "F8CBAD",
    "라이브 미등록": "D9D9D9",
    "API 미응답": "FFF2CC",
}

