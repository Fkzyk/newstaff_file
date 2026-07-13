# -*- coding: utf-8 -*-
"""アップロードされた元ファイルから差し込み用テンプレートを作成するスクリプト。

リポジトリ管理者が一度だけ実行するもので、アプリの実行時には使わない。
元ファイル(個人情報入り)はリポジトリに含めず、個人情報を Jinja タグに
置き換えたテンプレートだけを templates/ に保存する。
"""
import re
import shutil
import sys
from pathlib import Path

import openpyxl
from docx import Document

REPO = Path(__file__).resolve().parent.parent
TEMPLATES = REPO / "templates"


def rewrite_paragraph(para, new_text):
    """段落全体のテキストを差し替える(先頭ランの書式を維持)。"""
    runs = para.runs
    if not runs:
        return False
    runs[0].text = new_text
    for r in runs[1:]:
        r.text = ""
    return True


# 段落全体のテキスト(空白正規化後)に対する置換ルール
# (正規表現, 置換後テキスト)
ANNAI_RULES = [
    (r"^2026年7月13日$", "{{ 発行日 }}"),
    (r"^伊藤 まゆ佳\s*様$", "{{ 氏名 }}　様"),
    (r"^●2026年8月3日（月）AM 8:10頃～8:30 集合$",
     "●{{ 入社式日 }}　{{ 集合時間 }} 集合"),
    (r"^場所：北九州本社\(福岡県北九州市小倉南区上葛原2-18-50\)$",
     "場所：{{ 入社式場所 }}"),
    (r"^住所：北九州市小倉北区神岳2丁目5-29$", "住所：{{ トレセン住所 }}"),
    (r"^\(場所\)：サンスカイホテル小倉 ?　\(住所\)：福岡県北九州市小倉北区幸町 2-1 ?$",
     "(場所)：{{ ホテル名 }}　(住所)：{{ ホテル住所 }}"),
    (r"^チェックイン：8月 ?2日\s+チェックアウト：8月7日$",
     "チェックイン：{{ チェックイン日 }}　　　　　チェックアウト：{{ チェックアウト日 }}"),
    (r"^・8月3日：入社式/入社手続き/オリエンテーション/トレーニングセンター研修$",
     "・{{ 入社日短 }}：入社式/入社手続き/オリエンテーション/トレーニングセンター研修"),
    (r"^・8月4日～7日：トレーニングセンター研修 ?※9:00開始$",
     "・{{ 研修期間 }}：トレーニングセンター研修　※9:00開始"),
    (r"^・8月8日～9日：公休 ?$", "・{{ 公休期間 }}：公休"),
    (r"^・8月10日：店舗配属\s*※勤務開始時間は決定次第ご連絡いたします。\s*$",
     "・{{ 配属日短 }}：店舗配属　{{ 配属備考 }}"),
]


def build_annai(src: Path):
    doc = Document(src)
    applied = set()
    for para in doc.paragraphs:
        text = para.text
        for i, (pat, repl) in enumerate(ANNAI_RULES):
            if re.match(pat, text):
                rewrite_paragraph(para, repl)
                applied.add(i)
                break
    missed = [ANNAI_RULES[i][0] for i in range(len(ANNAI_RULES)) if i not in applied]
    if missed:
        raise SystemExit(f"入社案内: 置換できなかったルールがあります: {missed}")
    out = TEMPLATES / "入社のご案内.docx"
    doc.save(out)
    print("wrote", out)


def build_keiyakusho(src: Path):
    wb = openpyxl.load_workbook(src)
    ws = wb["雇用契約書"]
    # 差し込み対象セルを空にしておく(実際の値はアプリが書き込む)
    ws["B3"] = ""      # 氏名 様
    ws["C7"] = ""      # 入社年月日
    ws["C8"] = ""      # 就業場所
    ws["W10"] = ""     # 役職
    ws["I16"] = ""     # 等級
    # 給与欄・備考欄は数式ではなく値をアプリ側で書き込む
    # (変換ソフトによって数式が再計算されない事故を防ぐため)
    for addr in ("I17", "T17", "AE17", "I18", "I19", "B33"):
        ws[addr] = ""
    out = TEMPLATES / "雇用契約書.xlsx"
    wb.save(out)
    print("wrote", out)


def main():
    up = Path(sys.argv[1])
    TEMPLATES.mkdir(exist_ok=True)
    build_annai(up / "58121f1d-20260713____________.docx")
    build_keiyakusho(up / "00a86777-___________.xlsx")
    # 社宅案内は個人情報を含まないためそのままコピー
    shutil.copy(up / "1f2810b9-__________.docx", TEMPLATES / "社宅利用申込のご案内.docx")
    shutil.copy(up / "34d0db03-__________2602.pdf", TEMPLATES / "社宅システム入力マニュアル.pdf")
    print("done")


if __name__ == "__main__":
    main()
