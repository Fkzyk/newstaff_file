# -*- coding: utf-8 -*-
"""各書類(Word/Excel)の生成とPDF変換。"""
from __future__ import annotations

import datetime as dt
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path

import openpyxl
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.shared import Mm, Pt
from docxtpl import DocxTemplate

from . import defaults
from .data import Person

REPO = Path(__file__).resolve().parent.parent
TEMPLATES = REPO / "templates"

ANNAI_TEMPLATE = TEMPLATES / "入社のご案内.docx"
KEIYAKU_TEMPLATE = TEMPLATES / "雇用契約書.xlsx"
SHATAKU_ANNAI = TEMPLATES / "社宅利用申込のご案内.docx"
SHATAKU_MANUAL = TEMPLATES / "社宅システム入力マニュアル.pdf"


# ---------------------------------------------------------------- 入社のご案内
def render_annai(person: Person, cohort: dict, hakko_date: dt.date, out: Path) -> Path:
    """入社のご案内.docx を生成する。"""
    d = person.nyusha_date
    if person.first_shift:
        haizoku_biko = f"※勤務開始時間は{person.first_shift}です。"
    else:
        haizoku_biko = "※勤務開始時間は決定次第ご連絡いたします。"
    checkout = cohort.get("チェックアウト日") or (
        defaults.fmt_md(person.tresen_end) if person.tresen_end else "")
    ctx = {
        "発行日": defaults.fmt_full(hakko_date),
        "氏名": person.name,
        "入社式日": defaults.fmt_full_youbi(d),
        "集合時間": cohort["集合時間"],
        "入社式場所": cohort["入社式場所"],
        "トレセン住所": cohort["トレセン住所"],
        "ホテル名": cohort["ホテル名"],
        "ホテル住所": cohort["ホテル住所"],
        "チェックイン日": cohort["チェックイン日"],
        "チェックアウト日": checkout,
        "入社日短": defaults.fmt_md(d),
        "研修期間": defaults.fmt_range(d + dt.timedelta(days=1), person.tresen_end)
        if person.tresen_end else "",
        "公休期間": defaults.fmt_range(person.kyukyu1, person.kyukyu2)
        if person.kyukyu1 and person.kyukyu2 else "",
        "配属日短": defaults.fmt_md(person.haizoku_date) if person.haizoku_date else "",
        "配属備考": haizoku_biko,
    }
    tpl = DocxTemplate(ANNAI_TEMPLATE)
    tpl.render(ctx)
    tpl.save(out)
    return out


# ---------------------------------------------------------------- 雇用契約書
def _load_salary_table() -> tuple[dict, dict]:
    """給与テーブルシートから 等級->給与 と 備考文言 を読み込む。"""
    wb = openpyxl.load_workbook(KEIYAKU_TEMPLATE, data_only=False)
    ws = wb["給与テーブル"]
    salary = {}
    for row in ws.iter_rows(min_row=3):
        grade = row[1].value  # B列
        if grade:
            salary[str(grade).strip()] = {
                "月給": row[2].value,   # C列
                "基本給": row[3].value,  # D列
                "固定手当": row[4].value,  # E列
            }
    notes = {addr: ws[addr].value for addr in ("L3", "L4", "L5", "L7", "L8", "L9")}
    return salary, notes


def _minashi_notes(grade: str, notes: dict) -> tuple[str, str]:
    """等級の頭文字(P/S/その他)に応じた固定手当・割増賃金の説明文。"""
    head = grade[:1]
    if head == "S":
        return notes["L4"], notes["L8"]
    if head == "P":
        return notes["L3"], notes["L7"]
    return notes["L5"], notes["L9"]


def render_keiyakusho(person: Person, out: Path) -> Path:
    """雇用契約書.xlsx を生成する。給与は等級から給与テーブルを引いて記入する。"""
    salary_table, notes = _load_salary_table()
    if person.grade not in salary_table:
        raise ValueError(f"等級「{person.grade}」が給与テーブルにありません")
    sal = salary_table[person.grade]
    minashi1, minashi2 = _minashi_notes(person.grade, notes)

    wb = openpyxl.load_workbook(KEIYAKU_TEMPLATE)
    ws = wb["雇用契約書"]
    date_fmt = 'yyyy"年"m"月"d"日"'
    ws["B3"] = f"{person.name}　様"
    ws["C7"] = dt.datetime.combine(person.nyusha_date, dt.time())
    ws["C7"].number_format = date_fmt
    ws["C8"] = (f"（雇入れ直後）{person.shozoku_name}店　　　　　"
                "（変更の範囲）会社の定める事業所")
    ws["W10"] = person.yakushoku
    ws["I16"] = person.grade
    ws["I17"] = sal["月給"]
    ws["T17"] = sal["基本給"]
    ws["AE17"] = sal["固定手当"]
    ws["I18"] = minashi1
    ws["I19"] = minashi2
    ws["B33"] = dt.datetime.combine(person.nyusha_date, dt.time())
    ws["B33"].number_format = date_fmt
    # 給与テーブル(全等級の給与一覧)は社外秘のため、本人に渡すファイルには残さない
    del wb["給与テーブル"]
    wb.save(out)
    return out


# ---------------------------------------------------------------- 入社辞令
def _add_ruby_run(paragraph, base: str, ruby: str, font: str, size_pt: int):
    """ふりがな(ルビ)付きのランを段落に追加する。"""
    r = paragraph.add_run()
    rt_sz = max(int(size_pt / 2) * 2, 10)  # ルビは本文の約半分
    ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    import xml.sax.saxutils as su
    xml = f"""<w:ruby xmlns:w="{ns}">
      <w:rubyPr>
        <w:rubyAlign w:val="center"/>
        <w:hps w:val="{rt_sz}"/>
        <w:hpsRaise w:val="{int(size_pt * 2)}"/>
        <w:hpsBaseText w:val="{size_pt * 2}"/>
        <w:lid w:val="ja-JP"/>
      </w:rubyPr>
      <w:rt>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>
            <w:sz w:val="{rt_sz}"/>
          </w:rPr>
          <w:t>{su.escape(ruby)}</w:t>
        </w:r>
      </w:rt>
      <w:rubyBase>
        <w:r>
          <w:rPr>
            <w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/>
            <w:sz w:val="{size_pt * 2}"/>
          </w:rPr>
          <w:t>{su.escape(base)}</w:t>
        </w:r>
      </w:rubyBase>
    </w:ruby>"""
    from docx.oxml import parse_xml
    r._r.append(parse_xml(xml))


def _set_font(run, font: str, size: int, bold=False):
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    run._element.rPr.rFonts.set(qn("w:eastAsia"), font)


def render_jirei(person: Person, company: dict, out: Path) -> Path:
    """入社辞令.docx を生成する(毛筆体の辞令レイアウト)。"""
    font = "HG正楷書体-PRO"
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Mm(210), Mm(297)
    sec.top_margin, sec.bottom_margin = Mm(30), Mm(25)
    sec.left_margin, sec.right_margin = Mm(25), Mm(25)

    def para(align=WD_ALIGN_PARAGRAPH.CENTER, before=0, after=8):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(before)
        p.paragraph_format.space_after = Pt(after)
        return p

    p = para(after=24)
    _set_font(p.add_run("入　社　辞　令"), font, 28, bold=True)

    p = para(after=20)
    if person.furigana:
        _add_ruby_run(p, person.name, person.furigana, font, 16)
        _set_font(p.add_run("　殿"), font, 16)
    else:
        _set_font(p.add_run(f"{person.name}　殿"), font, 16)

    reiwa = defaults.fmt_reiwa(person.nyusha_date)
    p = para(align=WD_ALIGN_PARAGRAPH.LEFT, after=12)
    _set_font(p.add_run(
        f"　貴殿を、{reiwa}付をもって、正社員として採用し、"
        f"{company['辞令配属先']}に配属します。"), font, 14)
    p = para(align=WD_ALIGN_PARAGRAPH.LEFT, after=36)
    _set_font(p.add_run("　今後の活躍と社業の発展に貢献されることを期待します。"), font, 14)

    p = para(after=28)
    _set_font(p.add_run(reiwa), font, 14)

    p = para(align=WD_ALIGN_PARAGRAPH.RIGHT, after=4)
    _set_font(p.add_run(f"{company['会社名']}　"), font, 16)
    p = para(align=WD_ALIGN_PARAGRAPH.RIGHT)
    _set_font(p.add_run(f"代表取締役社長　{company['社長名']}　"), font, 16)

    doc.save(out)
    return out


# ---------------------------------------------------------------- 社宅関係
def copy_shataku_files(out_dir: Path) -> list[Path]:
    """社宅案内(docx)とマニュアル(PDF)を出力フォルダにコピーする。"""
    annai = out_dir / "社宅利用申込のご案内.docx"
    manual = out_dir / "社宅システム入力マニュアル.pdf"
    shutil.copy(SHATAKU_ANNAI, annai)
    shutil.copy(SHATAKU_MANUAL, manual)
    return [annai, manual]


# ---------------------------------------------------------------- PDF変換
def _convert_with_ms_office(path: Path, out_pdf: Path) -> bool:
    """Windows + Microsoft Office がある場合はOfficeで変換する(再現度最優先)。"""
    if platform.system() != "Windows":
        return False
    try:
        import win32com.client  # type: ignore
    except ImportError:
        return False
    suffix = path.suffix.lower()
    try:
        if suffix in (".docx", ".doc"):
            app = win32com.client.DispatchEx("Word.Application")
            app.Visible = False
            try:
                doc = app.Documents.Open(str(path), ReadOnly=True)
                doc.ExportAsFixedFormat(str(out_pdf), 17)  # wdExportFormatPDF
                doc.Close(False)
            finally:
                app.Quit()
            return True
        if suffix in (".xlsx", ".xls"):
            app = win32com.client.DispatchEx("Excel.Application")
            app.Visible = False
            app.DisplayAlerts = False
            try:
                wb = app.Workbooks.Open(str(path), ReadOnly=True)
                wb.ExportAsFixedFormat(0, str(out_pdf))  # xlTypePDF
                wb.Close(False)
            finally:
                app.Quit()
            return True
    except Exception:
        return False
    return False


def _find_soffice() -> str | None:
    for name in ("soffice", "libreoffice"):
        path = shutil.which(name)
        if path:
            return path
    for cand in (
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    ):
        if Path(cand).exists():
            return cand
    return None


def convert_to_pdf(path: Path) -> Path:
    """WordやExcelをPDFに変換して同じフォルダに保存する。

    Microsoft Office(Windows) → LibreOffice の順に試す。
    """
    path = Path(path)
    out_pdf = path.with_suffix(".pdf")
    if _convert_with_ms_office(path, out_pdf):
        return out_pdf
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError(
            "PDF変換ソフトが見つかりません。Microsoft Office または "
            "LibreOffice (https://ja.libreoffice.org) をインストールしてください。")
    with tempfile.TemporaryDirectory() as tmp:
        res = subprocess.run(
            [soffice, "--headless", "--convert-to", "pdf",
             "--outdir", str(path.parent), str(path)],
            capture_output=True, text=True, timeout=180,
            env={**__import__("os").environ, "HOME": tmp},
        )
    if not out_pdf.exists():
        raise RuntimeError(f"PDF変換に失敗しました: {path.name}\n{res.stderr}")
    return out_pdf
