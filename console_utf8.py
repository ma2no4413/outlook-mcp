"""コンソール出力を UTF-8 に固定する。

Windows のコンソール既定エンコーディングは環境の言語で決まり、英語環境では
cp1252 になる。このリポジトリの出力は日本語なので、そのままだと最初の print で
UnicodeEncodeError になり、処理が1行も進まないまま落ちる。
日本語版 Windows(cp932)では通ってしまうため、開発機では踏めない。

smoke_test.py と login.py で個別に対処していたが、あとから足した
validate_server_json.py で同じバグを再発させた。共有して一箇所に集約する。
日本語を print する実行可能スクリプトは、必ずこれを import すること。
"""
from __future__ import annotations

import sys


def force_utf8_output() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if reconfigure is not None:
            reconfigure(encoding="utf-8", errors="replace")


force_utf8_output()
