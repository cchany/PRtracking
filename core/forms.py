from django import forms

class AnalyzeForm(forms.Form):
    month = forms.IntegerField(min_value=1, max_value=12, label="월(1-12)")
    raw_file = forms.FileField(label="raw data 엑셀 (예: 시장조사업체_YYYYMMDD_YYYYMMDD.xlsx)")
    monthly_file = forms.FileField(label="월 PR 분석 엑셀 (예: 월 PR 분석.xlsx)")
