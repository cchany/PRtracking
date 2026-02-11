# core/services/naver_news_client.py
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime
from typing import Any, Dict, List, Optional
import html
import re
import time

import requests


NAVER_NEWS_ENDPOINT = "https://openapi.naver.com/v1/search/news.json"


def _strip_html(s: str) -> str:
    if not s:
        return ""
    # Naver title/description can include <b> tags
    s = re.sub(r"<[^>]+>", "", s)
    s = html.unescape(s)
    return s.strip()


def _parse_pubdate(pub_date_raw: str) -> Optional[datetime]:
    """
    Naver returns RFC822 style, e.g. "Mon, 03 Feb 2026 08:41:00 +0900"
    """
    if not pub_date_raw:
        return None
    try:
        dt = parsedate_to_datetime(pub_date_raw)
        # Ensure timezone-aware
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt
    except Exception:
        return None


@dataclass
class NewsItem:
    company: str
    title: str
    description: str
    press: str
    pub_date: datetime
    originallink: str
    link: str  # naver link
    uid: str   # dedup key


class NaverNewsClient:
    def __init__(
        self,
        *,
        client_id: str,
        client_secret: str,
        timeout_sec: int = 10,
        max_retries: int = 3,
        min_interval_sec: float = 0.12,
    ):
        if not client_id or not client_secret:
            raise ValueError("NAVER_CLIENT_ID / NAVER_CLIENT_SECRET is required.")

        self.client_id = client_id
        self.client_secret = client_secret
        self.timeout_sec = timeout_sec
        self.max_retries = max_retries
        self.min_interval_sec = min_interval_sec
        self._last_call_ts = 0.0

    def _sleep_if_needed(self):
        gap = time.time() - self._last_call_ts
        if gap < self.min_interval_sec:
            time.sleep(self.min_interval_sec - gap)

    def search_news(
        self,
        *,
        company: str,
        query: str,
        days: int = 30,
        max_items: int = 200,
        display: int = 100,
        sort: str = "date",
    ) -> List[NewsItem]:
        """
        Fetch up to max_items within last `days` days.
        Naver pagination: start 1..1000, display <= 100
        """
        days = int(days)
        max_items = int(max_items)
        display = min(int(display), 100)
        if max_items <= 0:
            return []

        cutoff = datetime.now(timezone.utc) - timedelta(days=days)

        out: List[NewsItem] = []
        seen_uid: set[str] = set()

        start = 1
        while start <= 1000 and len(out) < max_items:
            self._sleep_if_needed()
            self._last_call_ts = time.time()

            params = {
                "query": query,
                "display": display,
                "start": start,
                "sort": sort,
            }

            data = self._request_json(params=params)

            items = data.get("items") or []
            if not items:
                break

            stop_because_old = False

            for it in items:
                title = _strip_html(str(it.get("title", "")))
                desc = _strip_html(str(it.get("description", "")))
                originallink = str(it.get("originallink", "")).strip()
                link = str(it.get("link", "")).strip()
                pub_raw = str(it.get("pubDate", "")).strip()
                pub_dt = _parse_pubdate(pub_raw)

                if pub_dt is None:
                    continue

                # Normalize to UTC for cutoff compare
                pub_utc = pub_dt.astimezone(timezone.utc)

                if pub_utc < cutoff:
                    # since sort=date (desc), we can stop early
                    stop_because_old = True
                    break

                # Dedup key (prefer originallink if present)
                base_url = originallink or link
                base_url = base_url.strip()

                # uid: URL + date + title (for safety)
                uid = f"{base_url}|{pub_utc.date().isoformat()}|{title}".lower()

                if uid in seen_uid:
                    continue
                seen_uid.add(uid)

                out.append(
                    NewsItem(
                        company=company,
                        title=title,
                        description=desc,
                        press=str(it.get("publisher", "") or "").strip(),  # 'publisher' sometimes absent
                        pub_date=pub_dt,
                        originallink=originallink,
                        link=link,
                        uid=uid,
                    )
                )

                if len(out) >= max_items:
                    break

            if stop_because_old:
                break

            start += display

        return out

    def _request_json(self, *, params: Dict[str, Any]) -> Dict[str, Any]:
        headers = {
            "X-Naver-Client-Id": self.client_id,
            "X-Naver-Client-Secret": self.client_secret,
            "Accept": "application/json",
        }

        last_err: Optional[Exception] = None
        for attempt in range(1, self.max_retries + 1):
            try:
                res = requests.get(
                    NAVER_NEWS_ENDPOINT,
                    headers=headers,
                    params=params,
                    timeout=self.timeout_sec,
                )
                if res.status_code == 200:
                    return res.json()
                # Some error payloads are json too, but not always
                raise RuntimeError(f"Naver API HTTP {res.status_code}: {res.text[:300]}")
            except Exception as e:
                last_err = e
                # backoff
                time.sleep(0.25 * attempt)

        raise RuntimeError(f"Naver API request failed: {last_err}")
