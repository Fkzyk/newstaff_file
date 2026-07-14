# -*- coding: utf-8 -*-
"""一人分の書類一式+メール下書きの生成、および修正後の再変換。"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

from . import defaults, documents, jirei
from .data import Person
from .emails import build_eml

# メールに添付する PDF の表示順(グロブパターン)
ATTACH_ORDER = [
    "入社のご案内.pdf",
    "雇用契約書.pdf",
    "*入社辞令*.pdf",
    "社宅利用申込のご案内.pdf",
    "社宅システム入力マニュアル.pdf",
]


def person_dir(base: Path, person: Person) -> Path:
    d = person.nyusha_date
    cohort = f"{d.year}年{d.month}月{d.day}日入社" if d else "入社日不明"
    return base / cohort / person.name.replace(" ", "").replace("　", "")


def _mail_context(person: Person, company: dict, attachments: list[Path],
                  shataku_text: str | None = None) -> dict:
    listing = "\n".join(f"・{p.stem}" for p in attachments)
    if shataku_text is None:
        shataku_text = defaults.MAIL_SHATAKU_PARAGRAPH_DEFAULT
    # 対象者には空行で挟んだ段落として入る。対象外は行ごと消える。
    paragraph = f"\n{shataku_text.strip()}\n" if person.shataku and shataku_text.strip() else ""
    return {
        "氏名": person.name,
        "入社日": defaults.fmt_md(person.nyusha_date),
        "入社日full": defaults.fmt_full_youbi(person.nyusha_date),
        "添付一覧": listing,
        "社宅段落": paragraph,
        "社宅文": paragraph,  # 旧テンプレート互換
        **company,
    }


def _collect_attachments(folder: Path) -> list[Path]:
    """フォルダ内のPDFを添付順に集める。"""
    found = []
    for pattern in ATTACH_ORDER:
        found.extend(sorted(folder.glob(pattern)))
    return found


def rebuild_mail(folder: Path, person: Person, company: dict,
                 subject_tpl: str, body_tpl: str,
                 shataku_text: str | None = None) -> Path:
    """フォルダ内のPDFを添付してメール下書きを(再)作成する。"""
    attachments = _collect_attachments(folder)
    ctx = _mail_context(person, company, attachments, shataku_text)
    subject = subject_tpl.format(**ctx)
    body = body_tpl.format(**ctx)
    (folder / "メール本文.txt").write_text(
        f"宛先: {person.email}\n件名: {subject}\n\n{body}", encoding="utf-8-sig")
    return build_eml(person.email, subject, body, attachments,
                     folder / "メール下書き.eml")


def generate_person(person: Person, base_dir: Path, cohort: dict, company: dict,
                    subject_tpl: str, body_tpl: str,
                    hakko_date: dt.date | None = None,
                    shataku_text: str | None = None,
                    progress=None) -> tuple[Path, list[str]]:
    """一人分の書類一式を生成する。戻り値は(出力フォルダ, 注意メッセージ)。"""
    def report(msg):
        if progress:
            progress(msg)

    notes: list[str] = []
    hakko_date = hakko_date or dt.date.today()
    folder = person_dir(base_dir, person)
    folder.mkdir(parents=True, exist_ok=True)

    report("入社のご案内を作成中…")
    annai = documents.render_annai(person, cohort, hakko_date, folder / "入社のご案内.docx")
    report("雇用契約書を作成中…")
    keiyaku = documents.render_keiyakusho(person, folder / "雇用契約書.xlsx")

    for f in (annai, keiyaku):
        report(f"{f.name} をPDFに変換中…")
        documents.convert_to_pdf(f)

    # 辞令は原本.docの等長置換で生成。PDF化はWord限定
    # (LibreOfficeでは飾り枠・テキストボックスが崩れるため)
    report("入社辞令を作成中…")
    jirei_doc = jirei.render_jirei_doc(person, folder, hakko_date)
    report(f"{jirei_doc.name} をPDFに変換中…")
    if documents.convert_to_pdf(jirei_doc, word_only=True) is None:
        notes.append(
            f"辞令のPDF化にはWordが必要です。「{jirei_doc.name}」をWordで開いて"
            "「PDFとして保存」し、修正反映ボタンでメールを作り直してください。")

    if person.shataku:
        report("社宅案内をコピー中…")
        shataku_annai, _manual = documents.copy_shataku_files(folder)
        report("社宅案内をPDFに変換中…")
        documents.convert_to_pdf(shataku_annai)

    report("メール下書きを作成中…")
    rebuild_mail(folder, person, company, subject_tpl, body_tpl, shataku_text)
    return folder, notes


def refresh_person(person: Person, base_dir: Path, company: dict,
                   subject_tpl: str, body_tpl: str,
                   shataku_text: str | None = None, progress=None) -> Path:
    """フォルダ内のWord/Excelを修正した後に呼ぶ。

    Word/Excel を PDF に変換し直し、メール下書きも作り直す。
    (Word/Excel 自体は上書きしないので、手修正した内容が反映される)
    """
    folder = person_dir(base_dir, person)
    if not folder.exists():
        raise FileNotFoundError(f"フォルダがありません: {folder}")
    for f in sorted(folder.glob("*.docx")) + sorted(folder.glob("*.xlsx")):
        if f.name.startswith("~$"):
            continue
        if progress:
            progress(f"{f.name} をPDFに変換中…")
        documents.convert_to_pdf(f)
    for f in sorted(folder.glob("*.doc")):  # 辞令(.doc)はWord限定で変換
        if f.name.startswith("~$"):
            continue
        if progress:
            progress(f"{f.name} をPDFに変換中…")
        documents.convert_to_pdf(f, word_only=True)
    # マニュアルPDFは変換対象外なのでそのまま残る
    if progress:
        progress("メール下書きを作り直し中…")
    rebuild_mail(folder, person, company, subject_tpl, body_tpl, shataku_text)
    return folder
