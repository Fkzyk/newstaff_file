# -*- coding: utf-8 -*-
"""一人分の書類一式+メール下書きの生成、および修正後の再変換。"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

from . import defaults, documents, jirei
from .data import Person
from .emails import build_eml

# メールに添付する PDF の表示順(グロブパターン)
# ファイル名は「YYYYMMDD_書類名(氏名).拡張子」ルール(CLAUDE.md)
ATTACH_ORDER = [
    "*入社のご案内*.pdf",
    "*雇用契約書*.pdf",
    "*入社辞令*.pdf",
    "*社宅利用申込のご案内*.pdf",
    "*社宅システム入力マニュアル*.pdf",
]


def person_dir(base: Path, person: Person) -> Path:
    """一人分の保存先: (保存先)/入社年月_氏名(例: 202608_伊藤まゆ佳)"""
    d = person.nyusha_date
    prefix = f"{d:%Y%m}" if d else "入社日不明"
    return base / f"{prefix}_{person.name.replace(' ', '').replace('　', '')}"


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

    # ファイル名ルール: 作成日付_ファイル名(氏名).拡張子
    sakusei = dt.date.today()

    def fname(title: str, ext: str) -> str:
        name_disp = person.name.replace(" ", "").replace("　", "")
        return f"{sakusei:%Y%m%d}_{title}({name_disp}).{ext}"

    report("入社のご案内を作成中…")
    annai = documents.render_annai(person, cohort, hakko_date,
                                   folder / fname("入社のご案内", "docx"))
    report("雇用契約書を作成中…")
    keiyaku = documents.render_keiyakusho(person, folder / fname("雇用契約書", "xlsx"))

    for f in (annai, keiyaku):
        report(f"{f.name} をPDFに変換中…")
        documents.convert_to_pdf(f)

    # 辞令は原本.docの等長置換で生成。PDF化はWord限定
    # (LibreOfficeでは飾り枠・テキストボックスが崩れるため)
    report("入社辞令を作成中…")
    try:
        jirei_doc = jirei.render_jirei_doc(person, folder, sakusei)
        report(f"{jirei_doc.name} をPDFに変換中…")
        if documents.convert_to_pdf(jirei_doc, word_only=True) is None:
            notes.append(
                f"辞令のPDF化にはWordが必要です。「{jirei_doc.name}」をWordで開いて"
                "「PDFとして保存」し、修正反映ボタンでメールを作り直してください。")
    except jirei.JireiError as e:
        notes.append(f"入社辞令だけ自動作成できませんでした(他の書類は作成済み): {e}")

    if person.shataku:
        report("社宅案内をコピー中…")
        shataku_annai, _manual = documents.copy_shataku_files(
            folder,
            annai_name=fname("社宅利用申込のご案内", "docx"),
            manual_name=fname("社宅システム入力マニュアル", "pdf"))
        report("社宅案内をPDFに変換中…")
        documents.convert_to_pdf(shataku_annai)

    report("メール下書きを作成中…")
    rebuild_mail(folder, person, company, subject_tpl, body_tpl, shataku_text)
    return folder, notes


def check_grades(people: list[Person]) -> tuple[bool, list[str]]:
    """等級・月給チェック: 全員の等級を給与テーブルと突合する。

    戻り値: (全員一致か, 一人ずつの結果メッセージ)
    """
    salary_table, _ = documents._load_salary_table()
    all_ok, lines = True, []
    for p in people:
        if not p.grade:
            all_ok = False
            lines.append(f"❌ {p.name}: 等級が空欄です")
        elif p.grade in salary_table:
            s = salary_table[p.grade]
            ok = s["月給"] == s["基本給"] + s["固定手当"]
            mark = "✅" if ok else "❌"
            if not ok:
                all_ok = False
            lines.append(
                f"{mark} {p.name}: {p.grade} → 月給 {s['月給']:,}円"
                f"(基本給 {s['基本給']:,} + 固定手当 {s['固定手当']:,})")
        else:
            all_ok = False
            lines.append(f"❌ {p.name}: 等級「{p.grade}」が給与テーブルにありません")
    return all_ok, lines


def generate_jobkan_mails(people: list[Person], base_dir: Path,
                          nyusha_date) -> tuple[Path, list[Path], list[str]]:
    """同じ入社日の新入社員全員をBCCに入れたジョブカン案内メール2通を作る。"""
    members = [p for p in people if p.nyusha_date == nyusha_date and p.email]
    if not members:
        raise ValueError("この入社日のメールアドレスが1件もありません")
    bcc = [p.email for p in members]
    d = nyusha_date
    folder = base_dir / f"{d:%Y%m}_ジョブカン案内メール({defaults.fmt_md(d)}入社)"
    folder.mkdir(parents=True, exist_ok=True)
    ctx = {"BCC人数": len(bcc), "入社日": defaults.fmt_md(d)}
    outs = []
    for i, (subj, body, name) in enumerate([
        (defaults.JOBKAN_MAIL1_SUBJECT, defaults.JOBKAN_MAIL1_BODY, "1_事前案内"),
        (defaults.JOBKAN_MAIL2_SUBJECT, defaults.JOBKAN_MAIL2_BODY, "2_登録日リマインド"),
    ], start=1):
        body = body.format(**ctx)
        eml = build_eml(", ".join(defaults.JOBKAN_TO), subj, body, [],
                        folder / f"ジョブカン{name}.eml", bcc=bcc)
        (folder / f"ジョブカン{name}.txt").write_text(
            f"宛先: {', '.join(defaults.JOBKAN_TO)}\nBCC: {', '.join(bcc)}\n"
            f"件名: {subj}\n\n{body}", encoding="utf-8-sig")
        outs.append(eml)
    return folder, outs, [p.name for p in members]


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
