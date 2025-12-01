from django.shortcuts import render, redirect
from django.http import HttpResponse, HttpResponseBadRequest
from .forms import AnalyzeForm
from .services.xl_copy_simple import process_monthly_copy
from .services.xl_step2_tracking import process_tracking_from_work
from .services.xl_step3_master import process_master_update

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
            resp["Content-Disposition"] = f'attachment; filename=\"{fn}\"'
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
            resp["Content-Disposition"] = f'attachment; filename=\"{fn}\"'
            return resp

        # ---------- Step 3: 마스터 반영 ----------
        elif step == "3":
            checked_file = request.FILES.get("checked_file_step3")
            master_file = request.FILES.get("master_file")

            if not checked_file or not master_file:
                return HttpResponseBadRequest("검수 완료 파일과 마스터 파일을 모두 업로드해야 합니다.")

            try:
                result_bytes = process_master_update(
                    checked_file.read(),
                    master_file.read(),
                )
            except Exception as e:
                return HttpResponseBadRequest(f"Step3 처리 중 오류: {e}")

            fn = "PR_Master_업데이트.xlsx"
            resp = HttpResponse(
                result_bytes,
                content_type=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
            )
            resp["Content-Disposition"] = f'attachment; filename=\"{fn}\"'
            return resp

        else:
            return HttpResponseBadRequest("잘못된 step 값입니다.")

    # GET 요청이면 화면만 렌더링
    return render(request, "core/home.html", {"form": form})
