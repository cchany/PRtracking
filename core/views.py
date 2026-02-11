from __future__ import annotations

import re
import uuid
from collections import Counter, defaultdict
from datetime import datetime, timezone

from django.conf import settings
from django.core.cache import cache
from django.http import HttpResponse, HttpResponseBadRequest, JsonResponse
from django.shortcuts import render
from django.urls import reverse

# PR tracking
from .forms import AnalyzeForm
from .services.xl_copy_simple import process_monthly_copy
from .services.xl_step2_tracking import process_tracking_from_work
from .services.xl_step3_master import process_master_update

# 네이버 뉴스
from .services.naver_news_client import NaverNewsClient
from .services.news_classifier import classify
from .services.news_excel_exporter import build_news_workbook


MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
}


def parse_period(period: str) -> tuple[int, int]:
    """
    "Dec-25" -> (2025, 12)
    """
    if not period:
        raise ValueError("기간(period)이 비어 있습니다.")
    p = period.strip()
    m = re.match(
        r"^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-(\d{2})$",
        p,
        re.IGNORECASE,
    )
    if not m:
        raise ValueError("기간 형식이 올바르지 않습니다. 예: Dec-25")

    mon = MONTH_MAP[m.group(1).lower()]
    yy = int(m.group(2))
    year = 2000 + yy
    return year, mon


def home(request):
    form = AnalyzeForm()

    if request.method == "POST":
        step = request.POST.get("step")

        # ---------- Step 1: 월 PR 분석 자동반영 ----------
        if step == "1":
            form = AnalyzeForm(request.POST, request.FILES)
            if not form.is_valid():
                return render(request, "core/home.html", {"form": form})

            month = form.cleaned_data["month"]
            raw = request.FILES["raw_file"].read()
            monthly = request.FILES["monthly_file"].read()

            try:
                result_bytes = process_monthly_copy(raw, monthly, month)
            except Exception as e:
                return HttpResponseBadRequest(f"처리 중 오류: {e}")

            fn = f"월 PR 분석_자동반영_{month}월.xlsx"
            resp = HttpResponse(
                result_bytes,
                content_type=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )
            resp["Content-Disposition"] = f'attachment; filename="{fn}"'
            return resp

        # ---------- Step 2: 검수본 업로드 → *_work 통합 ----------
        elif step == "2":
            checked_file = request.FILES.get("checked_file")
            if not checked_file:
                return HttpResponseBadRequest("검수 완료 파일이 업로드되지 않았습니다.")

            try:
                result_bytes = process_tracking_from_work(checked_file.read())
            except Exception as e:
                return HttpResponseBadRequest(f"Step2 처리 중 오류: {e}")

            fn = "월 PR Tracking_완성본.xlsx"
            resp = HttpResponse(
                result_bytes,
                content_type=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )
            resp["Content-Disposition"] = f'attachment; filename="{fn}"'
            return resp

        # ---------- Step 3: 마스터 반영 ----------
        elif step == "3":
            checked_file = request.FILES.get("checked_file_step3")
            master_file = request.FILES.get("master_file")
            period = request.POST.get("period")

            if not checked_file or not master_file:
                return HttpResponseBadRequest(
                    "검수 완료 파일과 마스터 파일을 모두 업로드해야 합니다."
                )

            try:
                year, month = parse_period(period)
                result_bytes = process_master_update(
                    checked_file.read(),
                    master_file.read(),
                    year=year,
                    month=month,
                )
            except Exception as e:
                return HttpResponseBadRequest(f"Step3 처리 중 오류: {e}")

            fn = f"PR_Master_업데이트_{period}.xlsx"
            resp = HttpResponse(
                result_bytes,
                content_type=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )
            resp["Content-Disposition"] = f'attachment; filename="{fn}"'
            return resp

        else:
            return HttpResponseBadRequest("잘못된 step 값입니다.")

    return render(request, "core/home.html", {"form": form})


def news_collect(request):
    """
    GET  -> 수집 페이지 렌더
    POST -> 네이버 OpenAPI로 수집 → 기간 필터 → 분류 → (통계+다운로드URL) JSON 반환

    프론트는 JSON을 받아:
    1) 페이지 하단에 통계 그래프 렌더
    2) download_url로 엑셀 다운로드 트리거
    """
    if request.method == "GET":
        return render(request, "core/news_collect.html")

    # -------------------------
    # POST
    # -------------------------
    companies_raw = (request.POST.get("companies") or "").strip()
    if not companies_raw:
        return HttpResponse("기업명을 입력해 주세요.", status=400)

    # 줄바꿈/콤마 모두 허용
    companies = []
    for line in companies_raw.replace(",", "\n").splitlines():
        name = line.strip()
        if name:
            companies.append(name)

    if not companies:
        return HttpResponse("기업명을 입력해 주세요.", status=400)

    # ✅ 기간(yyyy-mm-dd) 파싱
    start_date_raw = (request.POST.get("start_date") or "").strip()
    end_date_raw = (request.POST.get("end_date") or "").strip()
    if not start_date_raw or not end_date_raw:
        return HttpResponse("시작일과 종료일을 입력해 주세요.", status=400)

    try:
        start_dt = datetime.strptime(start_date_raw, "%Y-%m-%d").replace(tzinfo=timezone.utc)
        end_dt = datetime.strptime(end_date_raw, "%Y-%m-%d").replace(
            hour=23, minute=59, second=59, tzinfo=timezone.utc
        )
    except ValueError:
        return HttpResponse("날짜 형식이 올바르지 않습니다. (예: 2026-02-03)", status=400)

    if start_dt > end_dt:
        return HttpResponse("시작일은 종료일보다 늦을 수 없습니다.", status=400)

    # 안전장치
    max_per_company = int(request.POST.get("max_per_company") or 200)
    max_per_company = max(10, min(max_per_company, 500))

    client_id = getattr(settings, "NAVER_CLIENT_ID", "")
    client_secret = getattr(settings, "NAVER_CLIENT_SECRET", "")

    try:
        api = NaverNewsClient(
            client_id=client_id,
            client_secret=client_secret,
            timeout_sec=12,
            max_retries=3,
            min_interval_sec=0.12,
        )
    except Exception as e:
        return HttpResponse(f"네이버 API 설정 오류: {e}", status=500)

    # ✅ 통계 집계용
    company_total = Counter()                 # 회사별 기사 수
    company_category = defaultdict(Counter)   # 회사별 카테고리 카운트

    rows = []
    for company in companies:
        items = api.search_news(
            company=company,
            query=company,
            max_items=max_per_company,
            display=100,
            sort="date",
        )

        for it in items:
            if not it.pub_date:
                continue

            pub_dt = it.pub_date
            if pub_dt.tzinfo is None:
                pub_dt = pub_dt.replace(tzinfo=timezone.utc)

            # 기간 필터
            if not (start_dt <= pub_dt <= end_dt):
                continue

            c = classify(it.title, it.description)

            # ✅ 통계 누적
            company_total[it.company] += 1
            company_category[it.company][c.category] += 1

            rows.append({
                "company": it.company,
                "pub_date": pub_dt,
                "press": it.press,
                "category": c.category,
                "score": c.score,
                "matched_keywords": c.matched_keywords,
                "title": it.title,
                "description": it.description,
                "originallink": it.originallink,
                "naver_link": it.link,
                "uid": it.uid,
            })

    # pub_date 최신순 정렬 (aware/naive 혼합 방지)
    rows.sort(
        key=lambda x: x.get("pub_date") or datetime.min.replace(tzinfo=timezone.utc),
        reverse=True
    )

    # 엑셀 생성
    xlsx_bytes = build_news_workbook(rows)

    # ✅ 엑셀 임시 저장 + 다운로드 URL 발급
    job_id = uuid.uuid4().hex[:12]
    cache.set(f"news_xlsx:{job_id}", xlsx_bytes, timeout=60 * 30)  # 30분 보관

    download_path = reverse("news_download", kwargs={"job_id": job_id})
    download_url = request.build_absolute_uri(download_path)

    stats = {
        "company_total": dict(company_total),
        "company_category": {k: dict(v) for k, v in company_category.items()},
        "meta": {
            "start_date": start_date_raw,
            "end_date": end_date_raw,
            "companies": companies,
            "total_rows": len(rows),
        }
    }

    return JsonResponse({
        "job_id": job_id,
        "download_url": download_url,
        "stats": stats,
    })


def news_download(request, job_id: str):
    """
    job_id로 캐시에 저장된 엑셀을 다운로드
    """
    blob = cache.get(f"news_xlsx:{job_id}")
    if not blob:
        return HttpResponse("파일이 만료되었거나 존재하지 않습니다. 다시 수집해 주세요.", status=404)

    ts = datetime.now().strftime("%Y%m%d_%H%M")
    filename = f"naver_news_{ts}_{job_id}.xlsx"

    resp = HttpResponse(
        blob,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    resp["Content-Disposition"] = f"attachment; filename*=UTF-8''{filename}"
    return resp
