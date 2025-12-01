from io import BytesIO
import re
from openpyxl import load_workbook


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


def _get_month_col_in_by_tier(month: int, year: int = 2025) -> int:
    if not (1 <= month <= 12):
        raise ValueError(f"유효하지 않은 월입니다: {month}")

    if year == 2025:
        return 15 + (month - 1)  # O=15
    elif year == 2026:
        return 27 + (month - 1)  # AA=27
    else:
        # 다른 연도 규칙이 생기면 여기서 추가
        raise ValueError(f"지원하지 않는 연도입니다: {year}")



def process_master_update(checked_bytes: bytes, master_bytes: bytes) -> bytes:
    """
    Step3: 검수완료파일 + 마스터 파일 → 마스터 by Tier 시트 업데이트

    1. 검수완료파일의 '{m}월 총평' 시트에서 D5:G8 가져오기
    2. 마스터 파일의 'by Tier' 시트에서
       - R~Z 열이 1~12월 → {m}월에 해당하는 열(col)을 찾고
       - 아래 규칙대로 값 채우기:

       (예: {m}=10 → X열)
       X3  = E5
       X4  = F5 - E5
       X5  = G5
       X6  = D5

       X8  = E6
       X9  = F6 - E6
       X10 = G6
       X11 = D6

       X13 = E7
       X14 = F7 - E7
       X15 = G7
       X16 = D7

       X18 = E8
       X19 = F8 - E8
       X20 = G8
       X21 = D8
    """

    # 1) 검수완료파일 로드 + {m}월 총평 찾기
    checked_wb = load_workbook(BytesIO(checked_bytes), data_only=True)
    summary_ws, month = _find_month_summary_sheet(checked_wb)
    if summary_ws is None or month is None:
        raise ValueError("검수완료파일에서 '{m}월 총평' 시트를 찾을 수 없습니다.")

    # 2) D5:G8 데이터 읽기
    # summary[row] = {"D":..., "E":..., "F":..., "G":...}
    summary = {}
    for r in range(5, 9):  # 5,6,7,8
        summary[r] = {
            "D": summary_ws[f"D{r}"].value,
            "E": summary_ws[f"E{r}"].value,
            "F": summary_ws[f"F{r}"].value,
            "G": summary_ws[f"G{r}"].value,
        }

    # 3) 마스터 파일 로드 + by Tier 시트 찾기
    master_wb = load_workbook(BytesIO(master_bytes), data_only=False)
    if "by Tier" not in master_wb.sheetnames:
        raise ValueError("Master 파일에 'by Tier' 시트가 없습니다.")
    tier_ws = master_wb["by Tier"]

    # 4) {m}월에 해당하는 열 (R~Z)
    month_col = _get_month_col_in_by_tier(month)
    # 2026년도에는 아래 코드 활성화
    # month_col = _get_month_col_in_by_tier(month, year=2026)

    # 안전한 숫자 계산용 헬퍼
    def _num(v):
        if v is None or v == "":
            return 0
        if isinstance(v, (int, float)):
            return v
        try:
            return float(str(v).replace(",", ""))
        except Exception:
            return 0

    # 5) 매핑 규칙 적용
    # (target_row, source_row, 'E'/'F-E'/'G'/'D')
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

    # 6) 변경된 마스터 파일 반환
    out = BytesIO()
    master_wb.save(out)
    out.seek(0)
    return out.getvalue()
