# 🕹️ PR Tracking & News Classification Hub
Django 기반 사내 자동화 도구로, **PR Tracking 자동화 페이지**와 **뉴스 기사 자동 분류 페이지** 두 가지 핵심 기능을 제공합니다.
엑셀 업로드 기반의 리포팅 자동화와 키워드 기반 기사 분류를 한 곳에서 처리합니다.

---

## 📖 프로젝트 개요

본 프로젝트는 월간 PR 리포팅 및 뉴스 기사 분류 작업의 반복 업무를 줄이기 위해 제작되었습니다.
업로드된 데이터에서 자동으로 결과 파일을 생성하고, 필요한 산출물을 빠르게 내려받을 수 있도록 구성되어 있습니다.

- **PR Tracking 자동화 페이지**: 월별 PR 분석/검수/마스터 반영까지 3단계 워크플로 제공
- **뉴스 기사 자동 분류 페이지**: 기사 텍스트와 소스 기반 자동 분류

---

## 🚀 주요 기능

### 🧩 PR Tracking 자동화 (Step 1~3)
- **Step 1: 월 PR 분석 자동반영**
  - 원본(raw) 데이터 + 월 PR 분석 파일 업로드
  - 월 기준으로 자동 반영된 엑셀 파일 생성
- **Step 2: 검수본 업로드 → 통합본 생성**
  - 검수 완료 파일 업로드 시 `_work` 통합 결과 생성
- **Step 3: 마스터 반영**
  - 검수본 + 마스터 파일 업로드 후 기간 입력
  - 월별 데이터 반영된 마스터 파일 생성

### 📰 뉴스 기사 자동 분류
- 기사 텍스트/소스 기반 카테고리 자동 분류
- 키워드 규칙 및 소스별 스코프에 따른 분류 지원
- 분류 결과를 엑셀로 산출

---

## 🧠 기술 스택

| 분야 | 기술 |
| --- | --- |
| Backend | Python / Django |
| Data | pandas / numpy |
| Excel I/O | openpyxl / xlsxwriter |
| Database | SQLite (기본) |

---

## ⚙️ 빠른 시작

### 1) 가상환경 생성 및 패키지 설치

```bash
python -m venv .venv
source .venv/bin/activate
pip install django pandas numpy openpyxl xlsxwriter
```

### 2) 서버 실행

```bash
python manage.py runserver
```

브라우저에서 `http://127.0.0.1:8000/`에 접속하면 메인 화면을 확인할 수 있습니다.

---

## 📌 사용 방법

### PR Tracking 자동화

1. **Step 1**: 월(1-12) 입력 후 raw/월 PR 분석 엑셀 업로드 → 결과 다운로드
2. **Step 2**: 검수 완료 파일 업로드 → `월 PR Tracking_완성본.xlsx` 다운로드
3. **Step 3**: 검수본 + 마스터 파일 업로드, 기간 입력 (예: `Dec-25`) → 마스터 업데이트 파일 다운로드

### 뉴스 기사 자동 분류

1. 분류 대상 기사 파일 업로드
2. 분류 규칙에 따라 자동 카테고리 부여
3. 결과 엑셀 다운로드

---

## 🧱 프로젝트 구조

```
/workspace/PRtracking
├── PRtracker/          # Django 프로젝트 설정
├── core/               # 핵심 앱 (폼, 뷰, 서비스 로직)
├── db.sqlite3          # 기본 SQLite DB
└── manage.py
```

---

## ⚠️ 참고 사항

- 기본 DB는 SQLite(`db.sqlite3`)를 사용합니다.
- 실제 운영 환경에서는 `DEBUG`, `SECRET_KEY`, `ALLOWED_HOSTS` 설정을 반드시 조정하세요.
