from __future__ import annotations

import re
import unicodedata


def normalize_text(text: str | None) -> str:
    value = unicodedata.normalize("NFKC", text or "")
    value = re.sub(r"\s+", "", value)
    return value


def normalize_code(text: str | None) -> str:
    value = normalize_text(text)
    return value.upper()


def safe_filename(text: str | None) -> str:
    value = normalize_text(text)
    value = re.sub(r"[^0-9A-Za-z\u4e00-\u9fff]+", "_", value)
    value = re.sub(r"_+", "_", value).strip("_")
    return value or "unknown"


def name_variants(name: str | None) -> list[str]:
    raw = normalize_text(name)
    if not raw:
        return []
    variants = [raw]
    cleaned = re.sub(r"[（）()【】\[\]{}·,，。!！?？;；:：/\\|\"'“”‘’`~\-_=+ ]+", "", raw)
    if cleaned and cleaned != raw:
        variants.append(cleaned)
    prefix_stripped = re.sub(r"^(北京|天津|上海|重庆|河北|山西|辽宁|吉林|黑龙江|江苏|浙江|安徽|福建|江西|山东|河南|湖北|湖南|广东|海南|四川|贵州|云南|陕西|甘肃|青海|台湾|内蒙古|广西|西藏|宁夏|新疆)", "", cleaned)
    if prefix_stripped and prefix_stripped not in variants:
        variants.append(prefix_stripped)
    return variants

