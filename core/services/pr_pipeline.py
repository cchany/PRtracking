import io
import pandas as pd
import numpy as np
from datetime import datetime

# ---------- 규칙/티어 로딩 ----------
def load_category_rules(fileobj):
    df = pd.read_excel(fileobj, sheet_name=0)
    df.columns = [c.strip() for c in df.columns]
    for c in ["category","keyword","scope"]:
        if c not in df.columns:
            raise ValueError("category_rules.xlsx에는 category, keyword, scope 컬럼이 필요합니다.")
    scoped = {}
    for _, row in df.iterrows():
        scope = str(row.get("scope","ALL")).strip() or "ALL"
        kws = [k.strip() for k in str(row.get("keyword","")).split(",") if str(k).strip()]
        scoped.setdefault(scope, []).append((row["category"], kws))
    return scoped

def classify_category(text, source, rules):
    hay = (text or "").lower()
    for scope in [source, "ALL"]:
        for cat, kws in rules.get(scope, []):
            if not kws: continue
            for kw in kws:
                if kw.lower() in hay:
                    return cat
    for scope in [source, "ALL"]:
        for cat, kws in rules.get(scope, []):
            if cat == "기타" and not kws:
                return cat
    return "기타"

def apply_tiers(df, tier_table_df):
    m = df.merge(tier_table_df, how="left", on="outlet")
    m["tier1"] = m["tier1"].fillna(0).astype(int)
    m["tier2"] = m["tier2"].fillna(0).astype(int)
    m["tier"] = np.where(m["tier1"]==1, "Tier1",
                np.where(m["tier2"]==1, "Tier2", "Unknown"))
    return m

def top_k_plus_others(series_counts, k=8):
    s = series_counts.sort_values(ascending=False)
    if len(s) <= k: return s
    top = s.iloc[:k]
    others = s.iloc[k:].sum()
    return pd.concat([top, pd.Series({"기타": others})])

# ---------- 월 파일 생성 ----------
def build_monthly_workbook(raw_df, year:int, month:int, cat_rules, tier_table_df) -> io.BytesIO:
    df = raw_df.copy()
    needed = ["date","outlet","title","body","url","source"]
    for c in needed:
        if c not in df.columns:
            raise ValueError(f"raw에 '{c}' 컬럼이 필요합니다.")
    df["date"] = pd.to_datetime(df["date"])
    df["category"] = df.apply(
        lambda r: classify_category(f"{r['title']} {r['body']}", str(r["source"]), cat_rules), axis=1
    )
    df = apply_tiers(df, tier_table_df)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as w:
        # CP_m
        cp = df[df["source"]=="CP"].copy()
        cp.to_excel(w, sheet_name=f"CP_{month}", index=False)

        # *_Work
        for src in ["CP","IDC","OmdiaTV","DSCC"]:
            sub = df[df["source"]==src].copy()
            sheet = f"{src}_{month}_Work"
            if not sub.empty:
                work = pd.DataFrame({
                    "C": sub["date"].dt.strftime("%Y-%m-%d").values,
                    "D": sub["outlet"].values,
                    "E": sub["title"].values,
                    "F": sub["body"].values,
                    "G": (sub["tier"]=="Tier1").astype(int).values,
                    "H": (sub["tier"]=="Tier2").astype(int).values,
                    "L": sub["category"].values,
                    "Z": sub["url"].values,
                })
            else:
                work = pd.DataFrame(columns=list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")[:30])
            work.to_excel(w, sheet_name=sheet, index=False)
            ws = w.sheets[sheet]
            ws.write(1, 5, "=COUNTA(E:E)")       # F2: 총게재수
            ws.write(2, 3, "수기 총게재수 확인")  # D3: 수기 검증칸

        # 총평
        df["month"] = df["date"].dt.to_period("M").dt.to_timestamp()
        outlet_counts = df.groupby(["outlet","tier"]).size().reset_index(name="count")
        summary = outlet_counts.copy()
        summary.insert(0, "month", f"{year}-{month:02d}")
        summary.to_excel(w, sheet_name="총평", index=False)

        # Top8+기타 + 차트
        for src in ["CP","트렌드포스","IDC","OmdiaTV","DSCC"]:
            sub = df[df["source"]==src]
            if sub.empty: 
                continue
            counts = sub.groupby("category").size()
            top8 = top_k_plus_others(counts, k=8).reset_index()
            top8.columns = ["category","count"]
            sht = f"{src}_Top8"
            top8.to_excel(w, sheet_name=sht, index=False)
            ws = w.sheets[sht]
            chart = w.book.add_chart({"type":"column"})
            chart.add_series({
                "name": f"{src} 카테고리",
                "categories": [sht, 1, 0, len(top8), 0],
                "values":     [sht, 1, 1, len(top8), 1],
            })
            chart.set_title({"name": f"{src} 상위 카테고리(Top8+기타)"})
            chart.set_legend({"position":"bottom"})
            ws.insert_chart("E2", chart)

        # Unknown outlets
        unknown = df[df["tier"]=="Unknown"]["outlet"].value_counts().reset_index()
        unknown.columns = ["outlet","count"]
        unknown.to_excel(w, sheet_name="Unknown_Outlets", index=False)

    out.seek(0)
    return out

# ---------- 마스터 파일 갱신 ----------
def _append_df_to_excel_bytes(existing_bytes, df, sheet_name:str) -> io.BytesIO:
    if existing_bytes is None:
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.to_excel(w, sheet_name=sheet_name, index=False)
        buf.seek(0)
        return buf
    buf = io.BytesIO(existing_bytes.getvalue())
    with pd.ExcelWriter(buf, engine="openpyxl", mode="a", if_sheet_exists="overlay") as w:
        book = w.book
        if sheet_name in book.sheetnames:
            exist = pd.read_excel(buf, sheet_name=sheet_name)
            startrow = (len(exist)+1) if not exist.empty else 0
            df.to_excel(w, sheet_name=sheet_name, index=False, header=exist.empty, startrow=startrow)
        else:
            df.to_excel(w, sheet_name=sheet_name, index=False)
    buf.seek(0)
    return buf

def update_master_from_monthly(monthly_bytes: io.BytesIO, master_bytes: io.BytesIO|None) -> io.BytesIO:
    xls = pd.ExcelFile(monthly_bytes)
    # month 추출
    cp_sheet = [s for s in xls.sheet_names if s.startswith("CP_")]
    month_str = "unknown"
    if cp_sheet:
        cp = pd.read_excel(monthly_bytes, sheet_name=cp_sheet[0])
        if "date" in cp.columns and not cp.empty:
            m = pd.to_datetime(cp["date"]).dt.to_period("M").dt.to_timestamp().iloc[0]
            month_str = m.strftime("%Y-%m")

    # by Tier
    frames = []
    for s in xls.sheet_names:
        if s.endswith("_Work"):
            dfw = pd.read_excel(monthly_bytes, sheet_name=s)
            if "D" in dfw.columns:
                dfw["outlet"] = dfw["D"]
                dfw["tier"] = np.where(dfw.get("G",0)==1, "Tier1",
                                       np.where(dfw.get("H",0)==1, "Tier2", "Unknown"))
                frames.append(dfw[["outlet","tier"]])
    if frames:
        bt = pd.concat(frames, ignore_index=True)
        bt = bt.groupby(["outlet","tier"]).size().reset_index(name="count")
        bt.insert(0, "month", month_str)
    else:
        bt = pd.DataFrame(columns=["month","outlet","tier","count"])

    # by Coverage
    frames = []
    for s in xls.sheet_names:
        if s.endswith("_Work"):
            dfw = pd.read_excel(monthly_bytes, sheet_name=s)
            source = s.split("_")[0]
            if "L" in dfw.columns:
                tmp = dfw[["L"]].copy()
                tmp.columns = ["category"]
                tmp["source"] = source
                frames.append(tmp)
    if frames:
        bc = pd.concat(frames, ignore_index=True)
        bc = bc.groupby(["source","category"]).size().reset_index(name="count")
        bc.insert(0, "month", month_str)
    else:
        bc = pd.DataFrame(columns=["month","source","category","count"])

    # 신규 PR
    frames = []
    for s in xls.sheet_names:
        if s.endswith("_Work"):
            dfw = pd.read_excel(monthly_bytes, sheet_name=s)
            source = s.split("_")[0]
            if set(["C","D","E","Z","L"]).issubset(dfw.columns):
                tmp = pd.DataFrame({
                    "date": pd.to_datetime(dfw["C"], errors="coerce"),
                    "outlet": dfw["D"],
                    "title": dfw["E"],
                    "url": dfw["Z"],
                    "source": source,
                    "category": dfw["L"],
                })
                frames.append(tmp)
    if frames:
        new_pr = pd.concat(frames, ignore_index=True)
    else:
        new_pr = pd.DataFrame(columns=["date","outlet","title","url","source","category"])

    # append 3 sheets
    master_buf = master_bytes
    master_buf = _append_df_to_excel_bytes(master_buf, bt, "by Tier")
    master_buf = _append_df_to_excel_bytes(master_buf, bc, "by Coverage")
    master_buf = _append_df_to_excel_bytes(master_buf, new_pr, "신규 PR")
    return master_buf
