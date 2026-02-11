# core/services/news_classifier.py
from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Dict, List, Tuple


@dataclass
class Classified:
    category: str  # 스마트폰/반도체(메모리)/TV/로봇/기타
    score: int
    matched_keywords: List[str]


def _compile_keywords() -> Dict[str, List[str]]:
    # 키워드 룰은 “설명 가능”이 핵심: matched_keywords로 근거를 남김
    return {
        "반도체(메모리)": [
            "반도체", "파운드리", "foundry", "공정", "노드", "nm", "euv","dram", "nand", "hbm",
            "웨이퍼", "메모리", "낸드", "advanced packaging", "디램", "서버메모리",
            "칩", "chip", "tsmc", "삼성 파운드리", "삼성전자 파운드리", "삼성전자", "SK 하이닉스", "Sk Hynix"
        ],
        "스마트폰": [
            "스마트폰", "휴대폰", "갤럭시", "아이폰", "iphone", "galaxy",
            "폴더블", "foldable"
        ],
        "TV": [
            "tv", "티비", "television", "oled", "qled", "lcd", "미니led", "mini led",
            "패널", "panel", "세트", "리테일", "디스플레이"
        ],
        "로봇": [
            "로봇", "휴머노이드", "humanoid", "협동로봇", "cobot",
            "amr", "agv", "자율주행로봇", "로보틱스", "robot vacuum"
        ],
    }


KW = _compile_keywords()


def _normalize(text: str) -> str:
    t = (text or "").strip().lower()
    t = re.sub(r"\s+", " ", t)
    return t


def classify(title: str, description: str) -> Classified:
    t = _normalize(f"{title} {description}")

    # 점수 집계
    hits: Dict[str, List[str]] = {k: [] for k in ["스마트폰", "반도체(메모리)", "TV", "로봇"]}
    scores: Dict[str, int] = {k: 0 for k in hits.keys()}

    for cat, words in KW.items():
        for w in words:
            if w.lower() in t:
                hits[cat].append(w)
                scores[cat] += 1

    # 우선순위:
    # 메모리는 반도체보다 먼저(요구사항: 별도 분류)
    # TV/로봇/스마트폰/메모리/반도체 중 최고 스코어 선택하되, 동점이면 우선순위로 결정
    priority = ["반도체(메모리)", "TV", "로봇", "스마트폰"]

    best_cat = None
    best_score = 0
    for cat in priority:
        sc = scores.get(cat, 0)
        if sc > best_score:
            best_cat = cat
            best_score = sc
        elif sc == best_score and sc > 0:
            # 동점이면 우선순위 쪽(cat이 먼저) 유지
            pass

    if not best_cat or best_score == 0:
        return Classified(category="기타", score=0, matched_keywords=[])

    # 표시용 카테고리명을 요구사항 순서로 통일
    name_map = {
        "스마트폰": "스마트폰",
        "반도체(메모리)": "반도체(메모리)",
        "TV": "TV",
        "로봇": "로봇",
    }
    mk = hits.get(best_cat, [])
    # 중복 제거(표시용)
    mk_unique = []
    seen = set()
    for x in mk:
        xl = x.lower()
        if xl in seen:
            continue
        seen.add(xl)
        mk_unique.append(x)

    return Classified(category=name_map[best_cat], score=best_score, matched_keywords=mk_unique)
