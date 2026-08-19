"""server.json が公式 MCP Registry の制約を満たすか確認する。

publish はタグを打った後にしか走らないため、そこで弾かれると
タグを消してやり直すことになる。実際 v0.2.6 で description の
100文字制限に引っかかった。CI で先に落とす。
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

import console_utf8  # noqa: F401  出力を UTF-8 に固定する(import した時点で効く)

ROOT = Path(__file__).resolve().parent
LIMITS = {"description": 100, "title": 100}


def main() -> int:
    errors: list[str] = []
    server = json.loads((ROOT / "server.json").read_text(encoding="utf-8"))

    for field, limit in LIMITS.items():
        value = server.get(field, "")
        if len(value) > limit:
            errors.append(f"{field} が {limit} 文字を超えています({len(value)} 文字)")

    for field in ("name", "description", "version", "packages"):
        if not server.get(field):
            errors.append(f"{field} がありません")

    name = server.get("name", "")
    if not re.fullmatch(r"io\.github\.[A-Za-z0-9-]+/[A-Za-z0-9._-]+", name):
        errors.append(f"name が io.github.<user>/<server> の形式ではありません: {name!r}")

    # 所有権はイメージのラベルと name の一致で確認される。食い違うと publish が弾かれる。
    dockerfile = (ROOT / "Dockerfile").read_text(encoding="utf-8")
    m = re.search(r'io\.modelcontextprotocol\.server\.name="([^"]+)"', dockerfile)
    if not m:
        errors.append("Dockerfile に io.modelcontextprotocol.server.name ラベルがありません")
    elif m.group(1) != name:
        errors.append(f"Dockerfile のラベル {m.group(1)!r} が server.json の name {name!r} と一致しません")

    # イメージのタグは server.json の version と揃っている必要がある。
    pkg = (server.get("packages") or [{}])[0]
    identifier = pkg.get("identifier", "")
    version = server.get("version", "")
    if identifier and not identifier.endswith(f":{version}"):
        errors.append(f"イメージのタグが version と一致しません: {identifier} / {version}")

    # サーバ本体が申告するバージョンとも揃える。
    src = (ROOT / "outlook_server.py").read_text(encoding="utf-8")
    m = re.search(r'^__version__ = "([^"]+)"', src, re.M)
    if m and m.group(1) != version:
        errors.append(f"outlook_server.__version__ {m.group(1)!r} が server.json の version {version!r} と違います")

    if errors:
        for e in errors:
            print(f"NG  {e}")
        return 1

    print(f"OK  server.json は妥当です(name={name} version={version})")
    print(f"OK  description {len(server['description'])}/100 文字")
    return 0


sys.exit(main())
