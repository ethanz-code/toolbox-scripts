from __future__ import annotations

import json
import os
import shutil
from pathlib import Path

from ..common.excel import load_input_rows, write_results_workbook, copy_file
from ..common.logging import JsonlLogger
from ..common.paths import new_run_dir
from ..common.retry import sleep_range
from ..common.text import safe_filename, normalize_code
from ..connectors.riskbird import RiskbirdClient, choose_best_candidate, search_candidates, analyze_detail


HEADERS = [
    "企业类型", "企业名称", "在线状态", "统一社会信用代码", "法定代表人", "注册资本", "所属省份",
    "风鸟命中方式", "风鸟候选数量", "风鸟企业名称", "风鸟统一社会信用代码", "信用代码匹配",
    "风鸟经营状态", "成立日期", "登记机关", "注册地址", "查询路径", "异常判断", "问题类型/依据", "风鸟链接",
]


def _read_cookie_from_args(args) -> str:
    if args.cookie_file:
        return Path(args.cookie_file).read_text(encoding="utf-8").strip()
    env = os.environ.get(args.cookie_env, "").strip()
    if env:
        return env
    raise RuntimeError(f"missing cookie: set {args.cookie_env} or pass --cookie-file")


def _summary_rows(results, run_dir: Path, input_path: str):
    counts = {}
    for item in results:
        counts[item["异常判断"]] = counts.get(item["异常判断"], 0) + 1
    rows = [
        ["项目", "数量"],
        ["总行数", len(results)],
        ["已确认命中", counts.get("已确认命中", 0)],
        ["有问题", counts.get("有问题", 0)],
        ["需复核", counts.get("需复核", 0)],
        ["未查到", counts.get("未查到", 0)],
        ["", ""],
        ["输入文件", input_path],
        ["运行目录", str(run_dir)],
    ]
    return rows


def _row_values(item: dict) -> list:
    return [item.get(h, "") for h in HEADERS]


def _build_issue_link(name: str, entid: str) -> str:
    if not entid:
        return ""
    from urllib.parse import quote
    return f"https://www.riskbird.com/ent/{quote(name)}.html?entid={entid}"


def run_riskbird_company_check(args):
    input_path = Path(args.input).expanduser().resolve()
    run_dir = new_run_dir(args.output_root, "riskbird-check")
    logger = JsonlLogger(run_dir / "log.jsonl")
    logger.emit("run.start", input=str(input_path), run_dir=str(run_dir))

    if args.save_input_copy:
        copy_file(input_path, run_dir / "input.xlsx")

    cookie = _read_cookie_from_args(args)
    headers, rows = load_input_rows(input_path)
    if args.limit and args.limit > 0:
        rows = rows[: args.limit]

    client = RiskbirdClient(cookie=cookie, logger=logger)
    results = []

    for idx, row in enumerate(rows, 1):
        name = (row.get("企业名称") or "").strip()
        code = (row.get("统一社会信用代码") or "").strip()
        row_id = safe_filename(code or name or str(idx))
        logger.emit("item.start", idx=idx, name=name, code=code)

        strategy_payloads, candidates = search_candidates(client, name, code)
        raw_search_path = run_dir / "raw" / "search" / f"{row_id}.json"
        raw_search_path.write_text(
            json.dumps(strategy_payloads, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        path_labels = []
        for payload in strategy_payloads:
            count = len(payload["items"])
            path_labels.append(f"{payload['strategy']}({count})")
        query_path = " | ".join(path_labels)

        if not candidates:
            result = {
                **row,
                "风鸟命中方式": "",
                "风鸟候选数量": 0,
                "风鸟企业名称": "",
                "风鸟统一社会信用代码": "",
                "信用代码匹配": "未返回",
                "风鸟经营状态": "",
                "成立日期": "",
                "登记机关": "",
                "注册地址": "",
                "查询路径": query_path,
                "异常判断": "未查到",
                "问题类型/依据": "多种检索方式都没有可用结果",
                "风鸟链接": "",
            }
            results.append(result)
            logger.emit("item.done", idx=idx, name=name, code=code, classification="未查到", hits=0)
            sleep_range(args.sleep_min, args.sleep_max)
            continue

        best, score = choose_best_candidate(candidates, name, code)
        entid = (best or {}).get("entid") or ""
        hit_name = (best or {}).get("entName") or (best or {}).get("ENTNAME") or ""
        hit_code = (best or {}).get("uniscid") or (best or {}).get("taxId") or ""
        hit_code_match = "一致" if normalize_code(hit_code) and normalize_code(hit_code) == normalize_code(code) else ("未返回" if not hit_code else "不一致")
        hit_mode = "exact" if hit_code_match == "一致" else ("fuzzy" if score >= 45 else "weak")

        detail_response = None
        detail_basic = {}
        detail_status = ""
        issues = []
        if entid:
            detail_response = client.ent_query(hit_name or name, entid)
            (run_dir / "raw" / "detail" / f"{row_id}.json").write_text(
                json.dumps(detail_response, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            issues, detail_status, detail_code, detail_basic = analyze_detail(detail_response, code)
            if detail_code and not hit_code:
                hit_code = detail_code
                hit_code_match = "一致" if normalize_code(detail_code) == normalize_code(code) else "不一致"
        else:
            (run_dir / "raw" / "detail" / f"{row_id}.json").write_text(
                json.dumps({"ok": False, "error": "missing entid from search candidate"}, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )

        if issues:
            classification = "有问题"
        else:
            classification = "已确认命中" if hit_code_match == "一致" and score >= 80 else "需复核"

        if detail_response is None or not detail_response.get("ok"):
            if classification == "已确认命中":
                classification = "需复核"
                issues = ["搜索命中但详情查询失败"]

        result = {
            **row,
            "风鸟命中方式": hit_mode,
            "风鸟候选数量": len(candidates),
            "风鸟企业名称": hit_name,
            "风鸟统一社会信用代码": hit_code,
            "信用代码匹配": hit_code_match,
            "风鸟经营状态": detail_status or (best or {}).get("entStatus", ""),
            "成立日期": detail_basic.get("esDate", "") if isinstance(detail_basic, dict) else "",
            "登记机关": detail_basic.get("regOrg", "") if isinstance(detail_basic, dict) else "",
            "注册地址": detail_basic.get("dom", "") if isinstance(detail_basic, dict) else "",
            "查询路径": query_path,
            "异常判断": classification,
            "问题类型/依据": "；".join(issues) if issues else "",
            "风鸟链接": _build_issue_link(hit_name or name, entid),
        }
        results.append(result)

        logger.emit(
            "item.done",
            idx=idx,
            name=name,
            code=code,
            classification=classification,
            candidates=len(candidates),
            score=score,
            issues=issues,
        )
        sleep_range(args.sleep_min, args.sleep_max)

    summary = {
        "total": len(results),
        "已确认命中": sum(1 for r in results if r["异常判断"] == "已确认命中"),
        "有问题": sum(1 for r in results if r["异常判断"] == "有问题"),
        "需复核": sum(1 for r in results if r["异常判断"] == "需复核"),
        "未查到": sum(1 for r in results if r["异常判断"] == "未查到"),
    }
    (run_dir / "summary.json").write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    summary_rows = _summary_rows(results, run_dir, str(input_path))
    all_rows = [HEADERS] + [_row_values(r) for r in results]
    problem_rows = [HEADERS] + [_row_values(r) for r in results if r["异常判断"] == "有问题"]
    review_rows = [HEADERS] + [_row_values(r) for r in results if r["异常判断"] == "需复核"]
    workbook_path = run_dir / "output" / "企业经营核验结果.xlsx"
    write_results_workbook(workbook_path, summary_rows, all_rows, problem_rows, review_rows)

    (run_dir / "output" / "results.json").write_text(json.dumps(results, ensure_ascii=False, indent=2), encoding="utf-8")

    logger.emit("run.finish", **summary, workbook=str(workbook_path))
    print(json.dumps({"run_dir": str(run_dir), "workbook": str(workbook_path), **summary}, ensure_ascii=False))

