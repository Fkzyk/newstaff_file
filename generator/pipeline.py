# -*- coding: utf-8 -*-
"""一人分の書類一式+メール下書きの生成、および修正後の再変換。"""
from __future__ import annotations

import datetime as dt
import re
from pathlib import Path

from . import defaults, documents, jirei
from .data import Person
from .emails import build_eml

# 入社案内メールに添付するPDF(グロブパターン・この順)
# 添付は「入社のご案内」+ 社宅対象者のみ社宅2点。
# 雇用契約書・入社辞令は生成はするがメールには添付しない(確定運用)。
ATTACH_ORDER = [
    "*入社のご案内*.pdf",
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
    # 添付ミス防止のため、実際に添付するファイル名をそのまま列挙する
    listing = "\n".join(f"・{p.name}" for p in attachments)
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


ARCHIVE_DIR = "旧版"


def _archive_old_files(folder: Path):
    """既存の書類ファイルを「旧版」サブフォルダへ退避する。

    上書きで手修正が消える事故と、新旧ファイルの混在(古い方を
    添付してしまう事故)を防ぐ。削除はしないので必要なら戻せる。
    """
    if not folder.exists():
        return
    dest = folder / ARCHIVE_DIR
    for f in sorted(folder.iterdir()):
        if f.is_file() and re.match(r"^\d{8}_", f.name):
            dest.mkdir(exist_ok=True)
            target = dest / f.name
            if target.exists():
                stamp = dt.datetime.now().strftime("%H%M%S")
                target = dest / f"{f.stem}_{stamp}{f.suffix}"
            f.rename(target)


def _office_files(folder: Path) -> list[Path]:
    files = []
    for pattern in ("*.docx", "*.xlsx", "*.doc"):
        files += [f for f in folder.glob(pattern) if not f.name.startswith("~$")]
    return sorted(set(files))


def find_stale_pdfs(folder: Path) -> list[str]:
    """Word/ExcelがPDFより新しい(=修正が反映されていない)ファイル名の一覧。"""
    stale = []
    for f in _office_files(folder):
        pdf = f.with_suffix(".pdf")
        if not pdf.exists() or pdf.stat().st_mtime + 1 < f.stat().st_mtime:
            stale.append(f.name)
    return stale


def ensure_pdfs_fresh(folder: Path, progress=None) -> list[str]:
    """PDFが古い・無いWord/Excelだけを変換し直す(辞令.docはWord限定)。"""
    refreshed = []
    for f in _office_files(folder):
        pdf = f.with_suffix(".pdf")
        if not pdf.exists() or pdf.stat().st_mtime + 1 < f.stat().st_mtime:
            if progress:
                progress(f"{f.name} をPDFに変換中…")
            try:
                if documents.convert_to_pdf(f, word_only=(f.suffix == ".doc")):
                    refreshed.append(f.name)
            except Exception:
                pass  # PDF変換できなくてもメール作成は続行(元ファイルは残る)
    return refreshed


def _collect_attachments(folder: Path) -> list[Path]:
    """フォルダ内のPDFを添付順に集める。

    同じ書類が複数日付分あっても、最新(ファイル名の日付が最大)だけを添付する。
    """
    found = []
    for pattern in ATTACH_ORDER:
        matches = [p for p in folder.glob(pattern) if p.is_file()]
        if matches:
            found.append(max(matches, key=lambda p: p.name))
    return found


def build_mail_content(folder: Path, person: Person, company: dict,
                       subject_tpl: str, body_tpl: str,
                       shataku_text: str | None = None
                       ) -> tuple[str, str, list[Path]]:
    """メールの(件名, 本文, 添付一覧)を組み立てる。PDFの鮮度も確認する。"""
    ensure_pdfs_fresh(folder)
    attachments = _collect_attachments(folder)
    ctx = _mail_context(person, company, attachments, shataku_text)
    return subject_tpl.format(**ctx), body_tpl.format(**ctx), attachments


def rebuild_mail(folder: Path, person: Person, company: dict,
                 subject_tpl: str, body_tpl: str,
                 shataku_text: str | None = None) -> Path:
    """フォルダ内のPDFを添付してメール下書きを(再)作成する。

    添付直前にPDFの鮮度を確認し、Word/Excelの方が新しければ
    自動で変換し直す(修正の反映漏れ防止)。
    """
    subject, body, attachments = build_mail_content(
        folder, person, company, subject_tpl, body_tpl, shataku_text)
    cc = defaults.ANNAI_CC
    (folder / "メール本文.txt").write_text(
        f"宛先: {person.email}\nCC: {', '.join(cc)}\n件名: {subject}\n\n{body}",
        encoding="utf-8-sig")
    return build_eml(person.email, subject, body, attachments,
                     folder / "メール下書き.eml", cc=cc)


def gmail_compose_url(to: str, subject: str, body: str,
                      cc: list[str] | None = None,
                      bcc: list[str] | None = None,
                      account: str | None = None) -> str:
    """Gmailの作成画面を開くURL(添付はGmailの仕様で自動では付かない)。"""
    from urllib.parse import urlencode
    params = {"view": "cm", "fs": "1", "to": to, "su": subject, "body": body}
    if cc:
        params["cc"] = ",".join(cc)
    if bcc:
        params["bcc"] = ",".join(bcc)
    if account:
        params["authuser"] = account
    return "https://mail.google.com/mail/?" + urlencode(params)


def gcal_event_url(title: str, day, details: str = "",
                   account: str | None = None) -> str:
    """Googleカレンダーの終日予定を1クリックで保存できる作成画面URL。"""
    from urllib.parse import urlencode
    end = day + dt.timedelta(days=1)
    params = {
        "action": "TEMPLATE", "text": title,
        "dates": f"{day:%Y%m%d}/{end:%Y%m%d}", "details": details,
    }
    if account:
        params["authuser"] = account
    return "https://calendar.google.com/calendar/render?" + urlencode(params)


def jobkan_calendar_events(people: list[Person]):
    """今後のジョブカン送信予定を(日付, タイトル, 説明)で返す(近い順)。"""
    today = dt.date.today()
    evs, seen = [], set()
    for p in people:
        if not p.nyusha_date or not p.email or p.nyusha_date in seen:
            continue
        seen.add(p.nyusha_date)
        ann = defaults.jobkan_announce_date(p.nyusha_date)
        auto = defaults.jobkan_auto_date(p.nyusha_date)
        md = defaults.fmt_md(p.nyusha_date)
        if ann >= today:
            evs.append((ann, f"【送信】ジョブカン事前案内メール（{md}入社）",
                        f"{md}入社の方へジョブカン事前案内メールを送る日。"
                        "入社書類作成アプリの③タブから送信。"))
        if auto >= today:
            evs.append((auto, f"【送信】ジョブカン登録日リマインドメール（{md}入社）",
                        f"{md}入社の方へ登録日リマインドメールを送る日。"))
    evs.sort()
    return evs


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
    # 既存の書類は上書きせず「旧版」へ退避(手修正の消失・新旧混在を防ぐ)
    _archive_old_files(folder)

    # ファイル名ルール: 作成日付_ファイル名(氏名).拡張子
    sakusei = dt.date.today()

    def fname(title: str, ext: str) -> str:
        name_disp = person.name.replace(" ", "").replace("　", "")
        return f"{sakusei:%Y%m%d}_{title}({name_disp}様).{ext}"

    # PDF変換は失敗しても止めない(元のWord/Excelは必ず残す)
    pdf_failed = []

    def to_pdf(f, word_only=False):
        report(f"{f.name} をPDFに変換中…")
        try:
            if documents.convert_to_pdf(f, word_only=word_only) is None:
                pdf_failed.append(f.name)
        except Exception:
            pdf_failed.append(f.name)

    report("入社のご案内を作成中…")
    annai = documents.render_annai(person, cohort, hakko_date,
                                   folder / fname("入社のご案内", "docx"))
    report("雇用契約書を作成中…")
    keiyaku = documents.render_keiyakusho(person, folder / fname("雇用契約書", "xlsx"))
    to_pdf(annai)
    to_pdf(keiyaku)

    # 辞令は原本.docの等長置換で生成。PDF化はWord限定
    # (LibreOfficeでは飾り枠・テキストボックスが崩れるため)
    report("入社辞令を作成中…")
    try:
        jirei_doc = jirei.render_jirei_doc(person, folder, sakusei)
        to_pdf(jirei_doc, word_only=True)
    except jirei.JireiError as e:
        notes.append(f"入社辞令だけ自動作成できませんでした(他の書類は作成済み): {e}")

    if person.shataku:
        report("社宅案内をコピー中…")
        shataku_annai, _manual = documents.copy_shataku_files(
            folder,
            annai_name=fname("社宅利用申込のご案内", "docx"),
            manual_name=fname("社宅システム入力マニュアル", "pdf"))
        to_pdf(shataku_annai)

    if pdf_failed:
        hint = documents.LAST_OFFICE_ERROR or ""
        notes.append(
            "PDFに変換できなかった書類があります(" + "、".join(pdf_failed) + ")。"
            "Word/Excelファイルは作成済みです。Microsoft Officeを一度終了してから"
            "アプリを再起動し、「作った書類を直したいとき」で作り直すとPDFが作られます。"
            + (f" [詳細: {hint}]" if hint else ""))

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


JOBKAN_MANUAL = documents.TEMPLATES / "jobkan_workflow.pdf"


def build_jobkan_reminders(people: list[Person], out_dir: Path) -> tuple[Path, list[str]]:
    """今後の入社月ごとに、ジョブカン案内メールの送信予定を
    カレンダー(.ics)に登録できるファイルを作る。当日に通知が出る。

    戻り値: (icsファイルパス, 予定の説明リスト)
    """
    today = dt.date.today()
    events, summary_lines = [], []
    seen = set()
    for p in people:
        if not p.nyusha_date or not p.email:
            continue
        key = p.nyusha_date
        if key in seen:
            continue
        seen.add(key)
        auto = defaults.jobkan_auto_date(p.nyusha_date)
        ann = defaults.jobkan_announce_date(p.nyusha_date)
        mstr = defaults.fmt_md(p.nyusha_date)
        if ann >= today:
            events.append((ann,
                f"ジョブカン事前案内メール送信({mstr}入社)",
                f"{mstr}入社の方へジョブカン事前案内メールを送る日です。"
                f"入社書類作成アプリで作成→送信してください。"))
            summary_lines.append(f"{ann:%Y/%m/%d}({defaults.WEEKDAYS[ann.weekday()]}) "
                                 f"事前案内メール送信 — {mstr}入社")
        if auto >= today:
            events.append((auto,
                f"ジョブカン登録日リマインドメール送信({mstr}入社)",
                f"{mstr}入社の方へ、ジョブカン登録メール送信当日のリマインドを送る日です。"))
            summary_lines.append(f"{auto:%Y/%m/%d}({defaults.WEEKDAYS[auto.weekday()]}) "
                                 f"登録日リマインドメール送信 — {mstr}入社")
    events.sort()
    out = out_dir / "ジョブカン送信リマインド.ics"
    out_dir.mkdir(parents=True, exist_ok=True)
    lines = ["BEGIN:VCALENDAR", "VERSION:2.0",
             "PRODID:-//nyusha-app//jobkan//JP", "CALSCALE:GREGORIAN"]
    for i, (d, summary, desc) in enumerate(events):
        end = d + dt.timedelta(days=1)
        lines += [
            "BEGIN:VEVENT",
            f"UID:jobkan-{d:%Y%m%d}-{i}@nyusha-app",
            f"DTSTART;VALUE=DATE:{d:%Y%m%d}",
            f"DTEND;VALUE=DATE:{end:%Y%m%d}",
            f"SUMMARY:{summary}",
            f"DESCRIPTION:{desc}",
            "BEGIN:VALARM", "TRIGGER:PT0S", "ACTION:DISPLAY",
            "DESCRIPTION:リマインド", "END:VALARM",
            "END:VEVENT",
        ]
    lines.append("END:VCALENDAR")
    out.write_text("\r\n".join(lines) + "\r\n", encoding="utf-8")
    return out, summary_lines


def jobkan_mail_contents(people: list[Person], nyusha_date
                         ) -> tuple[list[tuple[str, str, str]], list[str], list[str]]:
    """ジョブカン案内メール2通の(名前, 件名, 本文)とBCC・対象者名を返す。"""
    members = [p for p in people if p.nyusha_date == nyusha_date and p.email]
    if not members:
        raise ValueError("この入社日のメールアドレスが1件もありません")
    bcc = [p.email for p in members]
    ctx = {
        "BCC人数": len(bcc),
        "入社日": defaults.fmt_md(nyusha_date),
        "ジョブカン送信日": defaults.fmt_md(defaults.jobkan_auto_date(nyusha_date)),
        "事前案内送信日": defaults.fmt_md(defaults.jobkan_announce_date(nyusha_date)),
    }
    mails = [
        ("1_事前案内", defaults.JOBKAN_MAIL1_SUBJECT,
         defaults.JOBKAN_MAIL1_BODY.format(**ctx)),
        ("2_登録日リマインド", defaults.JOBKAN_MAIL2_SUBJECT,
         defaults.JOBKAN_MAIL2_BODY.format(**ctx)),
    ]
    return mails, bcc, [p.name for p in members]


def generate_jobkan_mails(people: list[Person], base_dir: Path,
                          nyusha_date) -> tuple[Path, list[Path], list[str]]:
    """同じ入社日の新入社員全員をBCCに入れたジョブカン案内メール2通を作る。"""
    mails, bcc, names = jobkan_mail_contents(people, nyusha_date)
    d = nyusha_date
    folder = base_dir / f"{d:%Y%m}_ジョブカン案内メール({defaults.fmt_md(d)}入社)"
    folder.mkdir(parents=True, exist_ok=True)
    # 添付PDFをフォルダにコピー(Gmailへのドラッグ用。差し替えたい場合はこのPDFを差し替え)
    attach = []
    if JOBKAN_MANUAL.exists():
        dest = folder / "入社時のワークフロー申請について.pdf"
        import shutil
        shutil.copy(JOBKAN_MANUAL, dest)
        attach = [dest]
    cc = defaults.JOBKAN_CC
    outs = []
    for name, subj, body in mails:
        eml = build_eml(", ".join(defaults.JOBKAN_TO), subj, body, attach,
                        folder / f"ジョブカン{name}.eml", bcc=bcc, cc=cc)
        (folder / f"ジョブカン{name}.txt").write_text(
            f"宛先(To): {', '.join(defaults.JOBKAN_TO)}\nCC: {', '.join(cc)}\n"
            f"BCC: {', '.join(bcc)}\n件名: {subj}\n\n{body}", encoding="utf-8-sig")
        outs.append(eml)
    return folder, outs, names


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
    ensure_pdfs_fresh(folder, progress=progress)
    if progress:
        progress("メール下書きを作り直し中…")
    rebuild_mail(folder, person, company, subject_tpl, body_tpl, shataku_text)
    return folder
