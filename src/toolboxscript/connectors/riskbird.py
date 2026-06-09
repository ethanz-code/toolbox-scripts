from __future__ import annotations

import json
import re
import time
import urllib.error
import urllib.parse
import urllib.request

from ..common.retry import backoff
from ..common.text import normalize_code, normalize_text, name_variants


BAD_STATUS_TERMS = ("注销", "吊销", "撤销", "迁出", "停业", "清算", "非正常")
RISK_COUNT_KEYS = {
    "dataCaseAbnormityCount": "经营异常",
    "dataCaseSeriousBreakLawCount": "严重违法",
    "dishonestyCount": "失信被执行人",
    "dishonestyHistoryCount": "历史失信被执行人",
    "zxzCount": "被执行人",
    "lsCount": "历史被执行人",
    "zbCount": "终本案件",
    "zblsCount": "历史终本案件",
    "xzgxfxzCount": "限制高消费",
    "xzgxflsCount": "历史限制高消费",
    "taxAbnormalCount": "非正常户",
    "taxAbnormalHistoryCount": "历史非正常户",
    "taxBreakLawCount": "税收违法",
    "taxBreakLawHistoryCount": "历史税收违法",
    "punishmentXzcfCount": "行政处罚",
    "cancelEasyCount": "简易注销",
    "companyCancellationFillingCount": "注销备案",
}


class RiskbirdClient:
    def __init__(self, cookie: str, logger=None):
        self.cookie = cookie
        self.logger = logger
        self.base = "https://www.riskbird.com"

    def _headers(self) -> dict[str, str]:
        return {
            "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 Chrome/125 Safari/537.36",
            "Cookie": self.cookie,
            "app-device": "WEB",
            "Referer": "https://www.riskbird.com/",
            "Origin": "https://www.riskbird.com",
            "Accept": "application/json",
            "Content-Type": "application/json",
        }

    def _request_json(self, method: str, url: str, *, params=None, body=None, retries: int = 5):
        last_error = None
        for attempt in range(1, retries + 1):
            full_url = url
            if params:
                query = urllib.parse.urlencode(params)
                full_url = f"{url}{'&' if '?' in url else '?'}{query}"
            data = None if body is None else json.dumps(body, ensure_ascii=False).encode("utf-8")
            req = urllib.request.Request(full_url, data=data, headers=self._headers(), method=method)
            started = time.perf_counter()
            try:
                with urllib.request.urlopen(req, timeout=40) as resp:
                    text = resp.read().decode("utf-8", "ignore")
                    elapsed_ms = int((time.perf_counter() - started) * 1000)
                    payload = json.loads(text) if text else None
                    return {
                        "ok": True,
                        "status": resp.status,
                        "url": full_url,
                        "elapsed_ms": elapsed_ms,
                        "text": text,
                        "data": payload,
                    }
            except urllib.error.HTTPError as exc:
                text = exc.read().decode("utf-8", "ignore") if hasattr(exc, "read") else ""
                elapsed_ms = int((time.perf_counter() - started) * 1000)
                last_error = exc
                if exc.code == 401:
                    refresh = self.refresh_token()
                    if refresh.get("ok"):
                        continue
                if exc.code == 429:
                    backoff(attempt, base=1.8, jitter=0.4)
                    continue
                return {
                    "ok": False,
                    "status": exc.code,
                    "url": full_url,
                    "elapsed_ms": elapsed_ms,
                    "text": text,
                    "error": str(exc),
                }
            except Exception as exc:
                last_error = exc
                backoff(attempt, base=1.4, jitter=0.4)
        return {
            "ok": False,
            "status": None,
            "url": url,
            "elapsed_ms": None,
            "text": "",
            "error": str(last_error) if last_error else "unknown error",
        }

    def refresh_token(self):
        result = self._request_json(
            "POST",
            f"{self.base}/riskbird-api/user/refreshToken",
            body={},
            retries=1,
        )
        if result.get("ok"):
            token = (((result.get("data") or {}).get("data") or {}).get("token")) or ""
            if token:
                self.cookie = re.sub(r"token=[^;]+", f"token={token}", self.cookie)
                if self.logger:
                    self.logger.emit("auth.refresh", ok=True)
        return result

    def company_search(self, key: str):
        return self._request_json(
            "POST",
            f"{self.base}/riskbird-api/api/v1/companies/search",
            body={
                "queryType": "1",
                "searchKey": key,
                "pageNo": 1,
                "range": 10,
                "selectConditionData": json.dumps({"status": "", "sort_field": "", "keyword_fields": ""}, ensure_ascii=False),
            },
        )

    def fast_search(self, key: str):
        return self._request_json(
            "GET",
            f"{self.base}/riskbird-api/api/v1/companies/fast-search",
            params={"limit": 100, "queryType": "1", "key": key},
        )

    def ent_query(self, ent_name: str, ent_id: str):
        return self._request_json(
            "GET",
            f"{self.base}/api/ent/query",
            params={"entName": ent_name, "entId": ent_id},
        )


def _extract_list(payload):
    if isinstance(payload, dict):
        for key in ("list", "records", "data"):
            value = payload.get(key)
            if isinstance(value, list):
                return value
        if isinstance(payload.get("data"), dict):
            return _extract_list(payload["data"])
    if isinstance(payload, list):
        return payload
    return []


def _candidate_name(item):
    return normalize_text(item.get("entName") or item.get("ENTNAME") or item.get("name") or item.get("perName") or "")


def _candidate_code(item):
    for key in ("uniscid", "taxId", "creditCode", "code"):
        value = item.get(key)
        if value:
            return normalize_code(value)
    return ""


def candidate_score(item, wanted_name: str, wanted_code: str) -> int:
    score = 0
    name = _candidate_name(item)
    code = _candidate_code(item)
    wanted_name_norm = normalize_text(wanted_name)
    wanted_code_norm = normalize_code(wanted_code)
    if wanted_code_norm and code and code == wanted_code_norm:
        score += 100
    if wanted_name and item.get("entName") == wanted_name:
        score += 80
    if wanted_name_norm and name == wanted_name_norm:
        score += 70
    elif wanted_name_norm and (wanted_name_norm in name or name in wanted_name_norm):
        score += 45
    if item.get("entStatus"):
        score += 3
    if wanted_code_norm and code and code != wanted_code_norm:
        score -= 15
    return score


def choose_best_candidate(candidates: list[dict], wanted_name: str, wanted_code: str):
    scored = [(candidate_score(item, wanted_name, wanted_code), item) for item in candidates if isinstance(item, dict)]
    if not scored:
        return None, 0
    scored.sort(key=lambda x: x[0], reverse=True)
    return scored[0][1], scored[0][0]


def search_candidates(client: RiskbirdClient, name: str, code: str):
    strategies = [
        ("company_search_code", code),
        ("company_search_name", name),
    ]
    variants = name_variants(name)
    for variant in variants:
        strategies.append(("fast_search_name", variant))
    seen = []
    strategy_payloads = []
    for strategy, key in strategies:
        if not key:
            continue
        if strategy == "fast_search_name":
            result = client.fast_search(key)
        else:
            result = client.company_search(key)
        payload = result.get("data") if result.get("ok") else None
        items = _extract_list(payload)
        deduped = []
        for item in items:
            if not isinstance(item, dict):
                continue
            fingerprint = (
                item.get("entid"),
                item.get("entName") or item.get("ENTNAME"),
                item.get("uniscid") or item.get("taxId"),
            )
            if fingerprint in seen:
                continue
            seen.append(fingerprint)
            deduped.append(item)
        strategy_payloads.append(
            {
                "strategy": strategy,
                "key": key,
                "response": result,
                "items": deduped,
            }
        )
    all_candidates = [item for payload in strategy_payloads for item in payload["items"]]
    return strategy_payloads, all_candidates


def analyze_detail(detail_response, wanted_code: str):
    if not detail_response or not detail_response.get("ok"):
        return [], "", "", {}
    data = detail_response.get("data") or {}
    basic_result = data.get("basicResult") or {}
    detail_result = data.get("detailResult") or {}
    api = basic_result.get("apiData") or {}
    list_data = api.get("list") if isinstance(api.get("list"), dict) else {}
    basic = list_data.get("jbxxInfo") if isinstance(list_data.get("jbxxInfo"), dict) else {}
    counts = {}
    counts.update(api.get("data") or {})
    counts.update(api.get("count") or {})
    counts.update((detail_result.get("apiData") or {}).get("count") or {})

    status = basic.get("entStatus") or basic_result.get("entStatus") or ""
    found_code = basic.get("uniscid") or basic.get("taxId") or ""
    reasons = []
    if any(term in status for term in BAD_STATUS_TERMS):
        reasons.append(f"经营状态：{status}")
    wanted_code_norm = normalize_code(wanted_code)
    found_code_norm = normalize_code(found_code)
    if wanted_code_norm and found_code_norm and wanted_code_norm != found_code_norm:
        reasons.append(f"信用代码不一致：表格为{wanted_code}，风鸟返回{found_code}")
    if basic.get("candate"):
        reasons.append(f"注销日期：{basic.get('candate')}")
    if basic.get("revdate"):
        reasons.append(f"吊销日期：{basic.get('revdate')}")
    for key, label in RISK_COUNT_KEYS.items():
        value = counts.get(key) or 0
        try:
            value = int(value)
        except Exception:
            value = 0
        if value > 0:
            reasons.append(f"{label}：{value}")

    return reasons, status, found_code, basic

