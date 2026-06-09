from __future__ import annotations

import argparse
import sys

from .runners.riskbird_company_check import run_riskbird_company_check


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="toolboxscript")
    sub = parser.add_subparsers(dest="command", required=True)

    riskbird = sub.add_parser("riskbird-check", help="核验企业是否在风鸟上有异常")
    riskbird.add_argument("--input", required=True, help="输入 xlsx 文件路径")
    riskbird.add_argument("--output-root", default="runs", help="运行输出根目录")
    riskbird.add_argument("--limit", type=int, default=0, help="只处理前 N 行，用于测试")
    riskbird.add_argument("--sleep-min", type=float, default=0.8, help="单条企业最小间隔秒数")
    riskbird.add_argument("--sleep-max", type=float, default=2.2, help="单条企业最大间隔秒数")
    riskbird.add_argument("--cookie-env", default="RISKBIRD_COOKIE", help="风鸟 Cookie 环境变量名")
    riskbird.add_argument("--cookie-file", default="", help="可选，风鸟 Cookie 文件路径")
    riskbird.add_argument("--save-input-copy", action="store_true", help="把输入文件拷贝到运行目录")
    riskbird.set_defaults(func=run_riskbird_company_check)

    return parser


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)
    args.func(args)


if __name__ == "__main__":
    main(sys.argv[1:])

