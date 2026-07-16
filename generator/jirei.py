# -*- coding: utf-8 -*-
r"""入社辞令の生成 — 原本.docの等長バイト置換方式。

【鉄則】辞令をゼロから作り直すのは絶対NG(金の飾り枠・書体・ルビが失われる)。
原本 templates/jirei_original.doc の氏名・ふりがな・日付だけを
バイト列レベルで置換し、レイアウト情報には一切触れない。

原本の本文(254文字)の構造:
  «F»EQ …(\s\up 31(ふりがな名字),名字)«E»␣«F»EQ …(\s\up 31(ふりがな名前),名前)«E»殿¶
  貴殿を、令和8年7月6日付をもって、…¶ … 　令和8年7月6日¶ …

- 本文全体の文字数(254)は絶対に変えない
- 氏名の長さ変化は EQフィールド内の hps32 の後のスペースで吸収
- それでも足りない場合はフィールド間のスペース1文字で吸収し、
  フィールド境界CP(Table streamのPLC内 5つ: 0, F1末, F2先頭, F2末, 254)を更新
- 日付は 8文字(7月6日型)は等長置換、9文字(10月5日型)は本文の読点1つと
  日付行の先頭全角空白で吸収。10文字(10月15日型)は未対応=手動作成
"""
from __future__ import annotations

import re
import struct
from pathlib import Path

from . import defaults

REPO = Path(__file__).resolve().parent.parent
ORIGINAL = REPO / "templates" / "jirei_original.doc"

FLD_BEGIN, FLD_END = "\x13", "\x15"
STORY_CPS = 254          # 本文ストーリーの総文字数(変えてはいけない)
ORIG_SEI, ORIG_MEI = "繁浪", "由香"
ORIG_FURI_SEI, ORIG_FURI_MEI = "しげなみ", "ゆか"
ORIG_DATE = "令和8年7月6日"


class JireiError(RuntimeError):
    pass


def _field_code(furi: str, base: str, pad: int) -> str:
    return (f'EQ \\* jc2 \\* "Font:HG正楷書体-PRO" \\* hps32{" " * pad} '
            f'\\o\\ad(\\s\\up 31({furi}),{base})')


def split_name(name: str, furigana: str) -> tuple[str, str, str, str]:
    """「伊藤 まゆ佳」「いとう まゆか」→ (伊藤, まゆ佳, いとう, まゆか)"""
    n = re.split(r"[ 　]+", name.strip())
    f = re.split(r"[ 　]+", furigana.strip())
    if len(n) != 2 or len(f) != 2:
        raise JireiError(
            f"氏名「{name}」/ふりがな「{furigana}」を姓・名に分割できません。"
            "シートの氏名とふりがなは姓と名の間にスペースを入れてください。")
    return n[0], n[1], f[0], f[1]


def build_jirei(sei: str, mei: str, furi_sei: str, furi_mei: str,
                reiwa_date: str, original: Path = ORIGINAL) -> bytes:
    raw = bytearray(original.read_bytes())

    # --- 本文ストーリーの位置を特定(最初のフィールド開始マーク) ---
    f1_marker = (FLD_BEGIN + "EQ").encode("utf-16-le")
    story_off = raw.find(f1_marker)
    if story_off < 0:
        raise JireiError("原本にEQフィールドが見つかりません")
    story = raw[story_off:story_off + STORY_CPS * 2].decode("utf-16-le")
    for expect in (ORIG_SEI, ORIG_MEI, ORIG_FURI_SEI, ORIG_FURI_MEI, ORIG_DATE):
        if expect not in story:
            raise JireiError(f"原本の本文に「{expect}」が見つかりません")
    if story.count(ORIG_DATE) != 2:
        raise JireiError("原本の日付の出現回数が想定(2回)と違います")

    # --- 旧フィールド範囲と旧CPを把握 ---
    f1_begin = 0                                   # ストーリーはフィールドで始まる
    f1_end = story.index(FLD_END)                  # F1の終了マークCP
    f2_begin = story.index(FLD_BEGIN, f1_end)
    f2_end = story.index(FLD_END, f2_begin)
    old_cps = [f1_begin, f1_end, f2_begin, f2_end, STORY_CPS]
    old_block_len = f2_end + 1 - f1_begin          # F1開始〜F2終了マークまで
    inter_space = story[f1_end + 1:f2_begin]       # フィールド間の文字(通常は半角スペース)

    # --- 新フィールドを等長で構成(スペースで調整) ---
    base1 = _field_code(furi_sei, sei, 0)
    base2 = _field_code(furi_mei, mei, 0)
    min_len = (1 + len(base1) + 1) + (1 + len(base2) + 1)  # マーク込み・間スペースなし
    budget = old_block_len - min_len               # 使える残り文字数
    if budget < 0:
        raise JireiError(
            f"氏名が長すぎて等長置換できません(超過{-budget}文字)。手動で作成してください。")
    # フィールド間スペースを優先して残し、余りをF1のhps32後に詰める
    keep_space = min(len(inter_space), budget) if inter_space else 0
    pad1 = budget - keep_space
    new_block = (FLD_BEGIN + _field_code(furi_sei, sei, pad1) + FLD_END
                 + " " * keep_space
                 + FLD_BEGIN + base2 + FLD_END)
    assert len(new_block) == old_block_len
    new_story = new_block + story[f2_end + 1:]

    # --- 新CP(フィールド境界が動いた場合に備えて算出) ---
    n_f1_end = new_story.index(FLD_END)
    n_f2_begin = new_story.index(FLD_BEGIN, n_f1_end)
    n_f2_end = new_story.index(FLD_END, n_f2_begin)
    new_cps = [0, n_f1_end, n_f2_begin, n_f2_end, STORY_CPS]

    # --- 日付置換(2箇所とも) ---
    if len(reiwa_date) == len(ORIG_DATE):
        new_story = new_story.replace(ORIG_DATE, reiwa_date)
    elif len(reiwa_date) == len(ORIG_DATE) + 1:
        # 本文: 直前の読点を1つ削って吸収 / 日付行: 先頭の全角空白を削って吸収
        if f"貴殿を、{ORIG_DATE}" not in new_story or f"　{ORIG_DATE}" not in new_story:
            raise JireiError("日付の吸収位置(読点・全角空白)が見つかりません")
        new_story = new_story.replace(f"貴殿を、{ORIG_DATE}", f"貴殿を{reiwa_date}", 1)
        new_story = new_story.replace(f"　{ORIG_DATE}", reiwa_date, 1)
    else:
        raise JireiError(
            f"日付「{reiwa_date}」({len(reiwa_date)}文字)は自動対応していません"
            "(10文字の日付は手動作成)。")
    if reiwa_date != ORIG_DATE and ORIG_DATE in new_story:
        raise JireiError("旧日付が置換しきれていません")
    if len(new_story) != STORY_CPS:
        raise JireiError(f"本文の文字数が変わってしまいました({len(new_story)}≠{STORY_CPS})")

    # --- 本文を書き戻し ---
    raw[story_off:story_off + STORY_CPS * 2] = new_story.encode("utf-16-le")

    # --- フィールドPLCのCPを更新(境界が動いた場合) ---
    if new_cps != old_cps:
        pattern = struct.pack("<5I", *old_cps)
        pos = raw.find(pattern)
        if pos < 0:
            raise JireiError("フィールドPLC(CP表)が見つかりません")
        if raw.find(pattern, pos + 1) >= 0:
            raise JireiError("フィールドPLCの候補が複数あり特定できません")
        raw[pos:pos + 20] = struct.pack("<5I", *new_cps)

    return bytes(raw)


def render_jirei_doc(person, out_dir: Path, sakusei_date) -> Path:
    """一人分の辞令.docを生成する。ファイル名: YYYYMMDD_入社辞令(氏名).doc"""
    sei, mei, furi_sei, furi_mei = split_name(person.name, person.furigana)
    reiwa = defaults.fmt_reiwa(person.nyusha_date)
    data = build_jirei(sei, mei, furi_sei, furi_mei, reiwa)
    name_disp = person.name.replace(" ", "").replace("　", "")
    out = out_dir / f"{sakusei_date:%Y%m%d}_入社辞令({name_disp}様).doc"
    out.write_bytes(data)
    return out
