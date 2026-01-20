from io import BytesIO
import re
from openpyxl import load_workbook
from datetime import date, datetime

MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
              "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]


def _find_month_summary_sheet(wb):
    """
    워크북에서 '{m}월 총평' 시트를 찾아 (시트객체, m) 반환.
    없으면 (None, None).
    """
    for name in wb.sheetnames:
        m = re.match(r"^(\d{1,2})\s*월\s*총평$", str(name).strip())
        if m:
            month = int(m.group(1))
            return wb[name], month
    return None, None


def _get_month_col_in_by_tier(month: int, year: int) -> int:
    """
    by Tier 시트의 월 컬럼 계산.
    기준:
      - 2025년 1월 시작열 = O(15)
      - 2026년 1월 시작열 = AA(27)
      - 2027년 1월 시작열 = AM(39)
    즉, 연도마다 12칸씩 증가하는 구조.
    """
    if not (1 <= month <= 12):
        raise ValueError(f"유효하지 않은 월입니다: {month}")
    if year < 2025:
        raise ValueError(f"지원하지 않는 연도입니다: {year}")

    start_col_2025 = 15  # O
    start_col = start_col_2025 + (year - 2025) * 12
    return start_col + (month - 1)


def _num(v):
    if v is None or v == "":
        return 0
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", ""))
    except Exception:
        return 0


# -----------------------------
# by Coverage 업데이트용 유틸
# -----------------------------
def _make_month_label(year: int, month: int) -> str:
    # 예: 2025, 12 -> "Dec-25"
    if not (1 <= month <= 12):
        raise ValueError("month must be 1~12")
    yy = str(year)[-2:]
    return f"{MONTH_ABBR[month-1]}-{yy}"


def _cell_to_month_label(v) -> str | None:
    """
    셀 값이
    - datetime/date -> 'Dec-25'
    - 'Dec-25' 문자열 -> 'Dec-25'
    그 외 -> None
    """
    if isinstance(v, (datetime, date)):
        return f"{MONTH_ABBR[v.month - 1]}-{str(v.year)[-2:]}"
    if isinstance(v, str):
        return v.strip()
    return None


def _find_row_by_label(ws, label: str, label_col: int = 1):
    for r in range(1, ws.max_row + 1):
        raw = ws.cell(row=r, column=label_col).value
        cell_label = _cell_to_month_label(raw)
        if cell_label == label:
            return r
    return None


def _contains_any(text: str, keywords) -> bool:
    t = (text or "").strip().lower()
    return any(k.lower() in t for k in keywords)


def _calc_coverage_from_work_sheet(checked_wb, *, sheet_name: str, tv_keywords: list[str]) -> list[float]:
    """
    지정한 *_work 시트에서 M7:N50을 읽어 6개 카테고리 합계를 반환:
    [Smartphone, AI, TV/Display, Semiconductor, Auto, IoT]
    """
    if sheet_name not in checked_wb.sheetnames:
        raise ValueError(f"검수완료파일에 '{sheet_name}' 시트가 없습니다.")

    ws = checked_wb[sheet_name]

    sums = {
        "smartphone": 0,
        "ai": 0,
        "tv_display": 0,
        "semi": 0,
        "auto": 0,
        "iot": 0,
    }

    for r in range(7, 51):  # 7~50
        m_val = _num(ws[f"M{r}"].value)
        n_txt = ws[f"N{r}"].value
        s = (str(n_txt) if n_txt is not None else "").strip()

        # 스마트폰
        if "스마트폰" in s:
            sums["smartphone"] += m_val

        # AI
        if "ai" in s.lower():
            sums["ai"] += m_val

        # TV/Display
        if _contains_any(s, tv_keywords):
            sums["tv_display"] += m_val

        # 반도체
        if "반도체" in s:
            sums["semi"] += m_val

        # 전기차
        if "전기차" in s:
            sums["auto"] += m_val

        # IoT
        if "iot" in s.lower():
            sums["iot"] += m_val

    return [
        sums["smartphone"],
        sums["ai"],
        sums["tv_display"],
        sums["semi"],
        sums["auto"],
        sums["iot"],
    ]


def _calc_omdia_tv_from_cp_work(checked_wb, *, month: int, keywords: list[str]) -> float:
    """
    CP_{m}_work 시트에서 M7:N50 중, N열 텍스트에 keywords가 포함된 행들의 M열 합계를 반환.
    (Omdia TV 단일 값)
    """
    sheet_name = f"CP_{month}_work"
    if sheet_name not in checked_wb.sheetnames:
        raise ValueError(f"검수완료파일에 '{sheet_name}' 시트가 없습니다.")

    ws = checked_wb[sheet_name]
    total = 0

    for r in range(7, 51):
        m_val = _num(ws[f"M{r}"].value)
        n_txt = ws[f"N{r}"].value
        s = (str(n_txt) if n_txt is not None else "").strip()

        if _contains_any(s, keywords):
            total += m_val

    return total


def _write_coverage_block_to_master(
    master_wb,
    *,
    year: int,
    month: int,
    start_col: int,   # B=2, H=8 ...
    values: list,
    sheet_name: str = "by Coverage",
):
    """
    Master의 'by Coverage' 시트에서 해당 월 라벨 행을 찾아
    start_col부터 values를 순서대로 입력.
    """
    if sheet_name not in master_wb.sheetnames:
        raise ValueError(f"Master 파일에 '{sheet_name}' 시트가 없습니다.")

    ws = master_wb[sheet_name]
    label = _make_month_label(year, month)  # 예: Dec-27
    row = _find_row_by_label(ws, label, label_col=1)

    if row is None:
        raise ValueError(f"Master by Coverage에서 '{label}' 행을 찾을 수 없습니다.")

    for i, v in enumerate(values):
        ws.cell(row=row, column=start_col + i, value=v)


def process_master_update(checked_bytes: bytes, master_bytes: bytes, *, year: int, month: int) -> bytes:
    """
    Step3:
    1) 검수완료파일 '{m}월 총평'에서 by Tier 업데이트
    2) by Coverage 업데이트
       - Counterpoint: CP_{m}_work -> B~G (6개)
       - IDC        : IDC_{m}_work -> H~M (6개)
       - Omdia TV   : CP_{m}_work -> N (1개)
    """

    # 1) 검수완료파일 로드 + {m}월 총평 찾기
    checked_wb = load_workbook(BytesIO(checked_bytes), data_only=True)
    summary_ws, found_month = _find_month_summary_sheet(checked_wb)
    if summary_ws is None or found_month is None:
        raise ValueError("검수완료파일에서 '{m}월 총평' 시트를 찾을 수 없습니다.")

    # ✅ 안전장치: period로 받은 month와 파일 내 총평 시트 month가 다르면 중단
    if found_month != month:
        raise ValueError(f"기간의 월({month})과 검수파일 총평 시트의 월({found_month})이 일치하지 않습니다.")

    # 2) D5:G8 데이터 읽기
    summary = {}
    for r in range(5, 9):  # 5~8
        summary[r] = {
            "D": summary_ws[f"D{r}"].value,
            "E": summary_ws[f"E{r}"].value,
            "F": summary_ws[f"F{r}"].value,
            "G": summary_ws[f"G{r}"].value,
        }

    # 3) 마스터 파일 로드
    master_wb = load_workbook(BytesIO(master_bytes), data_only=False)

    # 4) by Tier 업데이트
    if "by Tier" not in master_wb.sheetnames:
        raise ValueError("Master 파일에 'by Tier' 시트가 없습니다.")
    tier_ws = master_wb["by Tier"]

    month_col = _get_month_col_in_by_tier(month, year=year)

    mapping = [
        (3, 5, "E"),
        (4, 5, "F-E"),
        (5, 5, "G"),
        (6, 5, "D"),

        (8, 6, "E"),
        (9, 6, "F-E"),
        (10, 6, "G"),
        (11, 6, "D"),

        (13, 7, "E"),
        (14, 7, "F-E"),
        (15, 7, "G"),
        (16, 7, "D"),

        (18, 8, "E"),
        (19, 8, "F-E"),
        (20, 8, "G"),
        (21, 8, "D"),
    ]

    for target_row, src_row, kind in mapping:
        data = summary[src_row]
        if kind == "E":
            val = data["E"]
        elif kind == "G":
            val = data["G"]
        elif kind == "D":
            val = data["D"]
        elif kind == "F-E":
            val = _num(data["F"]) - _num(data["E"])
        else:
            val = None

        tier_ws.cell(row=target_row, column=month_col, value=val)

    # -----------------------------
    # 5) by Coverage 업데이트
    # -----------------------------
    # Counterpoint (B~G)
    cp_tv_keywords = ["tv", "디스플레이", "lcd", "led", "모니터", "oled", "xr"]
    cp_values6 = _calc_coverage_from_work_sheet(
        checked_wb,
        sheet_name=f"CP_{month}_work",
        tv_keywords=cp_tv_keywords,
    )
    _write_coverage_block_to_master(
        master_wb,
        year=year,
        month=month,
        start_col=2,   # B
        values=cp_values6,
    )

    # IDC (H~M)
    idc_tv_keywords = ["tv", "디스플레이", "lcd", "led", "모니터", "oled", "xr"]
    idc_values6 = _calc_coverage_from_work_sheet(
        checked_wb,
        sheet_name=f"IDC_{month}_work",
        tv_keywords=idc_tv_keywords,
    )
    _write_coverage_block_to_master(
        master_wb,
        year=year,
        month=month,
        start_col=8,   # H
        values=idc_values6,  # 6개면 H~M
    )

    # Omdia TV (N 단일 값) - CP에서 TV/디스플레이/OLED/LCD만
    omdia_keywords = ["tv", "디스플레이", "oled", "lcd"]
    omdia_tv_val = _calc_omdia_tv_from_cp_work(
        checked_wb,
        month=month,
        keywords=omdia_keywords,
    )
    _write_coverage_block_to_master(
        master_wb,
        year=year,
        month=month,
        start_col=14,  # N
        values=[omdia_tv_val],
    )

    # 6) 저장 후 반환
    out = BytesIO()
    master_wb.save(out)
    out.seek(0)
    return out.getvalue()
