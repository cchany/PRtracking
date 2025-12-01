import calendar
from datetime import datetime
from io import BytesIO
import re
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

# ===================== 분류 체계(범위 + 시장) =====================
GEOS = ["전세계", "한국", "중국", "유럽", "미국", "일본", "인도"]
MARKETS = [
    "스마트폰", "폴더블 스마트폰", "스마트폰 AP",
    "AI", "XR", "스마트워치",
    "보안",
    "TV", "OLED", "LCD TV", "디스플레이",
    "로봇청소기", "로봇",
    "반도체",
    "전기차",
]

# 허용 카테고리(화이트리스트): "범위 시장" 전 조합
ALLOWED_CATEGORIES = {f"{g} {m} 시장" for g in GEOS for m in MARKETS}
ALLOWED_CATEGORIES.add("미분류")

def _to_whitelist(cat: str) -> str:
    """화이트리스트 밖이면 '미분류'로 보정"""
    return cat if cat in ALLOWED_CATEGORIES else "미분류"

# ---------------- 지리/도메인 패턴 ----------------
# <범위> — 일반 키워드 + ‘지역명 … 시장/market’ 패턴까지 인식
GEO_PATTERNS = {
    "전세계": r"(전\s*세계.{0,20}시장|글로벌.{0,20}시장|전\s*세계|전세계|세계|글로벌)",
    "한국":   r"(한국.{0,20}시장|대한민국.{0,20}시장|국내.{0,20}시장|한국|대한민국|\b국내\b)",
    "중국":   r"(중국.{0,20}시장|중국)",
    "유럽":   r"(유럽.{0,20}시장|유럽)",
    "미국":   r"(미국.{0,20}시장|미국)",
    "일본":   r"(일본.{0,20}시장|일본)",
    "인도":   r"(인도.{0,20}시장|인도)",
}

# <시장> 패턴 (업데이트 포함)
DOMAIN_PATTERNS = {
    "폴더블 스마트폰": r"(폴더블\s*스마트폰|폴더블|클램셸|클램쉘|플립|플립폰|flip\b|fold\b|razr|레이저)",
    "스마트폰 AP": r"(\bAP\b|모바일\s*AP|\bSoC\b|chipset|칩셋|AP\s*원가|AP\s*비용|AP\s*공정)",
    # 🔧 스마트폰: '시장' 근접 단서로 제한(사이드바/연관글 노이즈 방지)
    "스마트폰": (
        r"((스마트폰|smart\s*phone|삼성폰|애플폰|mobile\s*phone|휴대폰).{0,15}시장|"
        r"시장.{0,15}(스마트폰|smart\s*phone|삼성폰|애플폰|mobile\s*phone|휴대폰))"
    ),

    "AI": r"(\bAI\b|인공지능|생성형\s*AI|Generative\s*AI|ChatGPT|Copilot|Gemini|LLM)",
    "XR": r"(\bXR\b|\bAR\b|\bVR\b|\bMR\b|헤드셋|스마트\s*안경|스마트안경)",
    "스마트워치": r"(스마트\s*워치|smart\s*watch|웨어러블)",

    "보안": r"(보안|사이버\s*보안|사이버\s*위협|사이버위협|위협|cyber\s*security|security)",

    "TV": r"(?:(?:\bTV\b|티비|television)(?:\s*시장)?)",
    "OLED": r"(?:OLED\s*TV\s*시장|OLED\s*시장|올레드\s*시장|OLED\s*TV)",
    "LCD TV": r"(?:LCD\s*TV\s*시장|LCD\s*시장)",
    "디스플레이": r"(?:(?:디스플레이|PC)(?!.{0,15}(?:TV|티비|OLED|올레드|LCD))(?:\s*시장)?|디스플레이 시장|PC)",

    "로봇청소기": r"(로봇\s*청소기|청소\s*로봇|robot\s*vacuum|vacuum\s*robot|로보락|Ecovacs|Dreame)",
    "로봇": r"(로봇\b|로봇공학|서비스\s*로봇|산업용\s*로봇|제조용\s*로봇|로봇산업)",

    "반도체": r"(반도체|파운드리|foundry|칩\b|chips\b|chip\b|메모리|memory|HBM|\bDRAM\b|D-?RAM|\bNAND\b|D램|디램|에이치비엠|"
             r"하이닉스|엔비디아|NVIDIA|AMD|인텔|Intel|TSMC|마이크론|Micron|wafer|fab|패키징)",
    "전기차": r"(전기차\b|전기차\s*시장|electric\s*vehicle|\bEV\b|\bBEV\b|\bPHEV\b)",
}

# 도메인 우선순위 (세부 → 일반)
DOMAIN_PRIORITY = [
    "폴더블 스마트폰", "스마트폰 AP",
    "OLED", "LCD TV", "TV",
    "XR", "스마트워치",
    "보안",
    "로봇청소기", "로봇",
    "반도체",
    "전기차",
    "디스플레이",
    "AI",
    "스마트폰",
]

# ---------------- 명시적 "<범위><시장> 시장" 최우선 탐지 ----------------
def _compile_explicit_patterns():
    geo_tokens = {
        "전세계": r"(전\s*세계.{0,20}시장|글로벌.{0,20}시장|전\s*세계|전세계|세계|글로벌)",
        "한국":   r"(한국.{0,20}시장|대한민국.{0,20}시장|국내.{0,20}시장|한국|대한민국|\b국내\b)",
        "중국":   r"(중국.{0,20}시장|중국)",
        "유럽":   r"(유럽.{0,20}시장|유럽)",
        "미국":   r"(미국.{0,20}시장|미국)",
        "일본":   r"(일본.{0,20}시장|일본)",
        "인도":   r"(인도.{0,20}시장|인도)",
    }
    market_tokens = {
        "스마트폰": r"(스마트폰|smart\s*phone|삼성폰|애플폰|휴대폰|mobile\s*phone)",
        "폴더블 스마트폰": r"(폴더블\s*스마트폰|폴더블|플립|폴드|flip\b|fold\b|클램셸|클램쉘|razr|레이저)",
        "스마트폰 AP": r"(\bAP\b|모바일\s*AP|\bSoC\b|chipset|칩셋)",
        "AI": r"(\bAI\b|인공지능|생성형\s*AI|Generative\s*AI|ChatGPT|Copilot|Gemini|LLM)",
        "XR": r"(\bXR\b|\bAR\b|\bVR\b|\bMR\b|헤드셋|스마트\s*안경|스마트안경)",
        "스마트워치": r"(스마트\s*워치|smart\s*watch|wearable)",
        "보안": r"(보안|사이버\s*보안|사이버\s*위협|cyber\s*security|security)",
        "TV": r"(?:\bTV\b|티비)",
        "OLED": r"(?:OLED\s*TV|OLED|올레드)",
        "LCD TV": r"(?:LCD-?TV|LCD)",
        "디스플레이": r"(디스플레이|PC)",
        "로봇청소기": r"(로봇\s*청소기|청소\s*로봇|robot\s*vacuum)",
        "로봇": r"(로봇\b|로봇공학|서비스\s*로봇|산업용\s*로봇|제조용\s*로봇)",
        "반도체": r"(반도체|파운드리|foundry|칩\b|하이닉스|메모리|HBM|DRAM|NAND|D램|디램|TSMC|Intel|인텔|Micron|마이크론)",
        "전기차": r"(전기차|electric\s*vehicle|\bEV\b|\bBEV\b|\bPHEV\b)",
    }
    patterns = []
    for g, gtok in geo_tokens.items():
        for m, mtok in market_tokens.items():
            # <범위> ... <시장> ... 시장
            p1 = rf"({gtok}).{{0,30}}({mtok}).{{0,10}}시장"
            # <시장> ... 시장 ... <범위>
            p2 = rf"({mtok}).{{0,10}}시장.{0,30}({gtok})"
            patterns.append((g, m, re.compile(p1, re.I)))
            patterns.append((g, m, re.compile(p2, re.I)))
    return patterns

EXPLICIT_PATTERNS = _compile_explicit_patterns()

# ---------------- 다중 국가 → '전세계' 강제 규칙 ----------------
_GEO_TOKEN_SPECIFIC = [
    r"(한국|대한민국|\b국내\b)",
    r"(중국)",
    r"(유럽)",
    r"(미국)",
    r"(일본)",
    r"(인도)",
]
_GEO_TOKEN_GLOBAL = r"(전\s*세계|전세계|글로벌|global|worldwide)"
GEO_RE_SPECIFICS = [re.compile(p, re.I) for p in _GEO_TOKEN_SPECIFIC]
GEO_RE_GLOBAL = re.compile(_GEO_TOKEN_GLOBAL, re.I)

def _multi_geo_triggers_world(text: str) -> bool:
    t = text or ""
    specific_hits = 0
    for rx in GEO_RE_SPECIFICS:
        if rx.search(t):
            specific_hits += 1
        if specific_hits >= 2:
            return True
    if specific_hits >= 1 and GEO_RE_GLOBAL.search(t):
        return True
    return False

def _find_explicit_geo_market(text: str):
    """
    본문에 '<범위><시장> 시장' 명시가 있으면 해당 조합 즉시 반환.
    단, 다중 국가 규칙이 트리거되면 범위는 '전세계'로 강제.
    """
    if _multi_geo_triggers_world(text):
        for g, m, rx in EXPLICIT_PATTERNS:
            if rx.search(text or ""):
                return "전세계", m
        return "전세계", None
    t = text or ""
    for g, m, rx in EXPLICIT_PATTERNS:
        if rx.search(t):
            return g, m
    return None, None

# ---------------- 공통 유틸 ----------------
def _copy_range_values(src_ws, dst_ws, src_range: str, dst_top_left: str):
    min_col, min_row, max_col, max_row = range_boundaries(src_range)
    col_letters = "".join([c for c in dst_top_left if c.isalpha()])
    row_digits = "".join([c for c in dst_top_left if c.isdigit()])
    dst_row0 = int(row_digits)
    dst_col0 = 0
    for i, ch in enumerate(reversed(col_letters.upper())):
        dst_col0 += (ord(ch) - 64) * (26 ** i)
    rows = max_row - min_row + 1
    cols = max_col - min_col + 1
    for r in range(rows):
        for c in range(cols):
            val = src_ws.cell(row=min_row + r, column=min_col + c).value
            dst_ws.cell(row=dst_row0 + r, column=dst_col0 + c, value=val)

def _rename_if_exists(wb, candidates, new_name):
    for name in candidates:
        if name in wb.sheetnames:
            wb[name].title = new_name
            return True
    last = candidates[-1] if candidates else None
    if last and last.startswith("re:"):
        rx = re.compile(last[3:], re.I)
        for s in wb.sheetnames:
            if rx.match(s):
                wb[s].title = new_name
                return True
    return False

def _fill_auto_numbers(ws, start_row: int = 5, col: int = 1, max_rows: int = 800):
    count = 0
    for i in range(start_row, max_rows + 1):
        val = ws.cell(row=i, column=2).value
        if val is None or str(val).strip() == "":
            break
        count += 1
        ws.cell(row=i, column=col, value=count)

def _update_countif_formulas(ws, month, base_sheet="CP"):
    for row in range(7, 501):
        ws[f"K{row}"] = f'=COUNTIF({base_sheet}_{month}!G:G,L{row})'

# ---------------- 크롤링(인터넷) ----------------
def _fetch_article_text(url: str) -> str:
    """
    기사 본문만 최대한 깨끗하게 추출.
    - 헤더/푸터/네비/사이드바/스크립트 제거 (노이즈 차단)
    - 여러 후보 컨테이너를 탐색
    """
    if not url or not isinstance(url, str) or not url.startswith(("http://", "https://")):
        return ""
    try:
        res = requests.get(url.strip(), timeout=5, headers={"User-Agent": "Mozilla/5.0"})
        if res.status_code != 200:
            return ""
        soup = BeautifulSoup(res.text, "html.parser")

        # 노이즈가 큰 영역 제거
        for sel in ["header", "nav", "footer", "aside", "script", "style", ".sidebar", ".breadcrumbs", ".breadcrumb", ".related", ".recommend", ".ad", ".ads"]:
            for n in soup.select(sel):
                n.decompose()

        selectors = [
            "article", ".article", "#articleBody", "#articeBody", "#news_body",
            ".news_body", ".article_body", ".article-body", ".content", "#content",
            ".post-content", ".entry-content", ".post_body", ".post-body"
        ]
        nodes = []
        for sel in selectors:
            nodes = soup.select(sel)
            if nodes:
                break
        if nodes:
            text = " ".join(n.get_text(separator=" ", strip=True) for n in nodes)
        else:
            text = soup.get_text(separator=" ", strip=True)

        return (text or "")[:4000]
    except Exception:
        return ""

# ---------------- 카테고리 분류 ----------------
def _regex_search(pattern, text):
    return re.search(pattern, text, flags=re.I) is not None

def _pick_geo(text):
    # 다중 국가 규칙 우선 적용
    if _multi_geo_triggers_world(text):
        return "전세계"
    # 명시적/일반 지리 패턴
    for geo, patt in GEO_PATTERNS.items():
        if _regex_search(patt, text):
            return geo
    return None

def _pick_domain(text):
    for key in DOMAIN_PRIORITY:
        patt = DOMAIN_PATTERNS[key]
        if _regex_search(patt, text):
            return key
    return None

def _compose_category(geo_label, domain_label):
    geo = geo_label if geo_label in GEOS else "전세계"
    dom = domain_label if domain_label in MARKETS else "스마트폰"
    return f"{geo} {dom} 시장"

def _classify_category_for_row(text_concat: str, source_hint: str):
    t = (text_concat or "")

    # 1) 명시 "<범위><시장> 시장" 최우선
    eg, em = _find_explicit_geo_market(t)
    if em:
        # 디스플레이 일반이면서 같은 문장에 TV/OLED/LCD가 있으면 세부로 재결정
        if em == "디스플레이" and _regex_search(r"(TV|티비|OLED|올레드|LCD)", t):
            dm = _pick_domain(t) or "디스플레이"
            return _to_whitelist(_compose_category(eg if eg else "전세계", dm))
        return _to_whitelist(_compose_category(eg if eg else "전세계", em))

    # 2) 일반 규칙: 도메인 → 지리 (다중국가 규칙은 _pick_geo 내부에서 처리)
    geo = _pick_geo(t)
    domain = _pick_domain(t)
    if domain:
        return _to_whitelist(_compose_category(geo if geo else "전세계", domain))

    # 3) 소스 보정 (디스플레이 전문 소스)
    if source_hint in ("OmdiaTV", "DSCC"):
        return _to_whitelist(_compose_category(geo if geo else "전세계", "디스플레이"))

    # 4) 최종 기본값: 어떤 규칙도 매칭되지 않으면 '미분류'
    return "미분류"

# ---------------- 카테고리 입력 루틴 ----------------
def _fill_categories(ws, source_hint: str, start_row: int = 5, max_rows: int = 800):
    for r in range(start_row, max_rows + 1):
        bval = ws.cell(row=r, column=2).value
        e_text = ws.cell(row=r, column=5).value
        f_url  = ws.cell(row=r, column=6).value
        if not (bval or e_text or f_url):
            break

        text_source = (str(e_text).strip() if e_text else "")

        # 🔧 요약문이 너무 짧으면(20→100자)만 URL 본문을 보강해서 사용
        if len(text_source) < 100 and f_url:
            fetched = _fetch_article_text(str(f_url))
            if fetched:
                text_source = fetched

        cat = _classify_category_for_row(text_source, source_hint)
        ws.cell(row=r, column=7, value=cat)

# ---------------- 메인 처리 ----------------
def process_monthly_copy(raw_bytes: bytes, monthly_bytes: bytes, month: int) -> bytes:
    raw_wb = load_workbook(BytesIO(raw_bytes), data_only=True)
    mon_wb = load_workbook(BytesIO(monthly_bytes))
    m = int(month)

    # 시트명 변경
    _rename_if_exists(mon_wb, ["CP_9", "cp_9", "CP-9", "re:^CP[_ -]?9$"], f"CP_{m}")
    _rename_if_exists(mon_wb, ["CP_9_work", "CP_9_Work", "re:^CP[_ -]?9[_ -]?work$"], f"CP_{m}_work")
    _rename_if_exists(mon_wb, ["IDC_9", "re:^IDC[_ -]?9$"], f"IDC_{m}")
    _rename_if_exists(mon_wb, ["IDC_9_work", "re:^IDC[_ -]?9[_ -]?work$"], f"IDC_{m}_work")
    _rename_if_exists(mon_wb, ["OmdiaTV_9", "Omdia TV_9", "re:^Omdia\s?TV[_ -]?9$"], f"OmdiaTV_{m}")
    _rename_if_exists(mon_wb, ["OmdiaTV_9_work", "Omdia TV_9_work", "re:^Omdia\s?TV[_ -]?9[_ -]?work$"], f"OmdiaTV_{m}_work")
    _rename_if_exists(mon_wb, ["DSCC_9", "re:^DSCC[_ -]?9$"], f"DSCC_{m}")
    _rename_if_exists(mon_wb, ["DSCC_9_work", "re:^DSCC[_ -]?9[_ -]?work$"], f"DSCC_{m}_work")
    _rename_if_exists(mon_wb, [f"9월 총평", "re:^9\s*월\s*총평$"], f"{m}월 총평")

    # RAW → 월 시트 값 복사
    _copy_range_values(raw_wb["CPR"],      mon_wb[f"CP_{m}"],      "B5:B800", "B5")
    _copy_range_values(raw_wb["CPR"],      mon_wb[f"CP_{m}"],      "D5:G800", "C5")
    _copy_range_values(raw_wb["IDC"],      mon_wb[f"IDC_{m}"],     "B5:B800", "B5")
    _copy_range_values(raw_wb["IDC"],      mon_wb[f"IDC_{m}"],     "D5:G800", "C5")
    _copy_range_values(raw_wb["Omdia TV"], mon_wb[f"OmdiaTV_{m}"], "B5:B800", "B5")
    _copy_range_values(raw_wb["Omdia TV"], mon_wb[f"OmdiaTV_{m}"], "D5:G800", "C5")
    _copy_range_values(raw_wb["DSCC"],     mon_wb[f"DSCC_{m}"],    "B5:B800", "B5")
    _copy_range_values(raw_wb["DSCC"],     mon_wb[f"DSCC_{m}"],    "D5:G800", "C5")

    # 번호 매기기
    for name in [f"CP_{m}", f"IDC_{m}", f"OmdiaTV_{m}", f"DSCC_{m}"]:
        _fill_auto_numbers(mon_wb[name])

    # A2 수식 업데이트 (IDC/OmdiaTV/DSCC)
    for name in [f"IDC_{m}", f"OmdiaTV_{m}", f"DSCC_{m}"]:
        mon_wb[name]["A2"] = f"=CP_{m}!A2"
        
    # ✅ CP_{m} 시트 A2 날짜 자동 업데이트
    current_year = datetime.now().year
    last_day = calendar.monthrange(current_year, m)[1]
    start_date = f"{current_year}/{m:02d}/01"
    end_date = f"{current_year}/{m:02d}/{last_day:02d}"
    mon_wb[f"CP_{m}"]["A2"] = f"[기간] {start_date}~ {end_date}"

    # 카테고리 자동 분류 (G열)
    _fill_categories(mon_wb[f"CP_{m}"],      "CP")
    _fill_categories(mon_wb[f"IDC_{m}"],     "IDC")
    _fill_categories(mon_wb[f"OmdiaTV_{m}"], "OmdiaTV")
    _fill_categories(mon_wb[f"DSCC_{m}"],    "DSCC")

    out = BytesIO()
    mon_wb.save(out)
    out.seek(0)
    return out.getvalue()
