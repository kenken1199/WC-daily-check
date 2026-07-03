#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
測定値出力の txt ファイルを CSV に変換するスクリプト。

txt の各行は「No. 左の値 右の値」の形式（右の値は無い行もある）。
これを 1 値につき 1 行に展開し、以下の 3 列の CSV を出力する。

    測定値出力No. , 測定値(g) , 判定

・測定値出力No. … 左の値はその行のNo.、右の値はNo.+1
・測定値(g)      … 数値（末尾の * は除去）
・判定           … 元の値に * が付いていれば "NG"、無ければ "OK"

使い方:
    python convert_txt_to_csv.py 入力.txt [入力2.txt ...]
        → 各ファイルと同じ場所に「入力.csv」を作成する

    python convert_txt_to_csv.py 入力.txt -o 出力.csv
        → 出力先を指定する（入力が1つのときのみ）
"""

import sys
import csv
from pathlib import Path

# txt の文字コード（Windows 日本語環境の Shift-JIS）
INPUT_ENCODING = "shift_jis"
# CSV の文字コード（Excel で開けるよう BOM 付き UTF-8）
OUTPUT_ENCODING = "utf-8-sig"

HEADER = ["測定値出力No.", "測定値(g)", "判定"]


def convert(input_path: Path, output_path: Path) -> int:
    """1 ファイルを変換し、出力した行数（ヘッダーを除く）を返す。"""
    text = input_path.read_text(encoding=INPUT_ENCODING, errors="replace")

    rows = []
    for line in text.splitlines():
        parts = line.split()
        if not parts:
            continue

        # 先頭が整数の行だけをデータ行として扱う
        # （タイトル行・日付行はここで除外される）
        try:
            no = int(parts[0])
        except ValueError:
            continue

        values = parts[1:]  # 左の値・右の値（右は無いこともある）
        for offset, token in enumerate(values):
            judgment = "NG" if "*" in token else "OK"
            value = token.replace("*", "")
            rows.append([no + offset, value, judgment])

    with output_path.open("w", encoding=OUTPUT_ENCODING, newline="") as f:
        writer = csv.writer(f)
        writer.writerow(HEADER)
        writer.writerows(rows)

    return len(rows)


def main(argv):
    args = argv[1:]
    if not args:
        print(__doc__)
        return 1

    # -o による出力先指定に対応
    output_override = None
    if "-o" in args:
        i = args.index("-o")
        try:
            output_override = args[i + 1]
        except IndexError:
            print("エラー: -o の後に出力ファイル名を指定してください。")
            return 1
        del args[i:i + 2]

    if output_override and len(args) != 1:
        print("エラー: -o を使う場合、入力ファイルは 1 つだけ指定してください。")
        return 1

    for arg in args:
        input_path = Path(arg)
        if not input_path.is_file():
            print(f"見つかりません: {input_path}")
            continue

        if output_override:
            output_path = Path(output_override)
        else:
            output_path = input_path.with_suffix(".csv")

        count = convert(input_path, output_path)
        print(f"{input_path.name} → {output_path.name}（{count} 行）")

    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
