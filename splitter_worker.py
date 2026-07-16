from __future__ import annotations

import base64
import sys

from splitter import SplitterError, process_workbooks


def encoded(value: str) -> str:
    return base64.b64encode(value.encode("utf-8")).decode("ascii")


def main() -> int:
    if len(sys.argv) != 4:
        print("status=error")
        print("message_b64=" + encoded("内部调用参数不完整。"))
        return 2
    try:
        def progress(percent: int, message: str) -> None:
            print(f"progress={percent}|{encoded(message)}", flush=True)

        stats = process_workbooks(sys.argv[1], sys.argv[2], sys.argv[3], progress=progress)
        print("status=ok")
        print(f"orders={stats.orders}")
        print(f"rows={stats.rows}")
        print(f"matched_rows={stats.matched_rows}")
        print(f"unmatched_rows={stats.unmatched_rows}")
        print(f"exceptional_orders={stats.exceptional_orders}")
        print("output_path_b64=" + encoded(stats.output_path))
        print("first_warning_b64=" + encoded(stats.warnings[0] if stats.warnings else ""))
        return 0
    except Exception as exc:
        message = str(exc) if isinstance(exc, SplitterError) else f"生成时发生错误：{exc}"
        print("status=error")
        print("message_b64=" + encoded(message))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
