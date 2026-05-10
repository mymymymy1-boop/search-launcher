# -*- coding: utf-8 -*-
"""
PHI at NYK ECSF Game 4 GamePreview Word Generator (2026-05-10)
※ Knicks vs 76ers Game 4 (5/10 3:30 PM ET, Xfinity Mobile Arena)
"""
import json, os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml, OxmlElement

DATA = json.load(open("/home/user/search-launcher/_data_nyk_phi_g4.json", "r", encoding="utf-8"))
OUT = "/home/user/search-launcher/PHI_at_NYK_GamePreview_R2G4_20260510.docx"

# --- Team colors (per word-format.md) ---
NYK = RGBColor(0xF5, 0x84, 0x26)   # NYK orange
PHI = RGBColor(0x00, 0x6B, 0xB6)   # PHI blue
BODY = RGBColor(0x33, 0x33, 0x33)
GREY = RGBColor(0x88, 0x88, 0x88)
BLACK = RGBColor(0, 0, 0)
FONT = "Noto Sans JP"


def set_run(run, text, size=13, bold=False, color=BODY, font=FONT):
    run.text = text
    run.font.name = font
    run.font.size = Pt(size)
    run.font.bold = bold
    if color is not None:
        run.font.color.rgb = color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts')
        rpr.append(rfonts)
    rfonts.set(qn('w:ascii'), font)
    rfonts.set(qn('w:hAnsi'), font)
    rfonts.set(qn('w:eastAsia'), font)


def add_p(doc, text, size=13, bold=False, color=BODY, align=None, space_after=2):
    p = doc.add_paragraph()
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.space_before = Pt(0)
    r = p.add_run()
    set_run(r, text, size, bold, color)
    return p


def add_page_break(doc):
    p = doc.add_paragraph()
    r = p.add_run()
    br = parse_xml(f'<w:br {nsdecls("w")} w:type="page"/>')
    r._element.append(br)


def set_cell_shading(cell, color):
    sh = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
    cell._tc.get_or_add_tcPr().append(sh)


def set_cell_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for edge in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{edge}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), 'BFBFBF')
        tcBorders.append(b)
    tcPr.append(tcBorders)


def write_cell(cell, text, size=11, bold=False, color=BODY, bg=None, align=None):
    cell.text = ""
    set_cell_borders(cell)
    if bg:
        set_cell_shading(cell, bg)
    p = cell.paragraphs[0]
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    if align is not None:
        p.alignment = align
    r = p.add_run()
    set_run(r, text, size, bold, color)
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER


def make_table(doc, headers, rows, widths_cm=None, header_bg="0F172A", header_color=RGBColor(0xFF, 0xFF, 0xFF)):
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = 'Table Grid'
    if widths_cm:
        for i, w in enumerate(widths_cm):
            for r in t.rows:
                r.cells[i].width = Cm(w)
    for j, h in enumerate(headers):
        write_cell(t.rows[0].cells[j], h, size=10, bold=True, color=header_color, bg=header_bg, align=WD_ALIGN_PARAGRAPH.CENTER)
    for i, row in enumerate(rows):
        bg = "F9FAFB" if i % 2 == 0 else None
        for j, val in enumerate(row):
            align = WD_ALIGN_PARAGRAPH.CENTER if j > 0 else WD_ALIGN_PARAGRAPH.LEFT
            write_cell(t.rows[1 + i].cells[j], str(val), size=10, color=BODY, bg=bg, align=align)
    return t


def main():
    doc = Document()
    s = doc.sections[0]
    s.top_margin = Cm(1.3); s.bottom_margin = Cm(1.3)
    s.left_margin = Cm(1.6); s.right_margin = Cm(1.6)
    s.page_width = Cm(21.6); s.page_height = Cm(27.9)

    sty = doc.styles['Normal']
    sty.font.name = FONT
    sty.font.size = Pt(13)
    rpr = sty.element.get_or_add_rPr()
    rfonts = rpr.find(qn('w:rFonts'))
    if rfonts is None:
        rfonts = OxmlElement('w:rFonts'); rpr.append(rfonts)
    rfonts.set(qn('w:ascii'), FONT); rfonts.set(qn('w:hAnsi'), FONT); rfonts.set(qn('w:eastAsia'), FONT)

    M = DATA["meta"]

    # ========= PAGE 1 =========
    add_p(doc, "NBA GAME PREVIEW", size=28, bold=True, color=BLACK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_p(doc, "プレーオフ第1ラウンド 第2回戦・第4戦", size=14, bold=True, color=BLACK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    add_p(doc, "フィラデルフィア・76ers", size=22, bold=True, color=PHI, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_p(doc, "@", size=14, bold=True, color=BLACK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_p(doc, "ニューヨーク・ニックス", size=22, bold=True, color=NYK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    add_p(doc, f"2026年5月10日（現地）／日本時間 5/11(月) 4:30 AM JST", size=12, color=BODY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_p(doc, f"{M['venue']} ／ {M['tv']}", size=12, color=BODY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_p(doc, f"シリーズ：ニックス3勝0敗リード（NYK勝てばカンファレンス・ファイナル進出／PHI負ければスイープ）", size=13, bold=True, color=BLACK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    # 試合の焦点
    add_p(doc, "■ 試合の焦点", size=18, bold=True, color=BLACK, space_after=6)
    for t in DATA["focus_points"]:
        add_p(doc, t, size=13, color=BODY, space_after=3)

    add_p(doc, "", size=8, space_after=4)

    # 直近10試合（5戦×2チーム）
    add_p(doc, "■ 直近の戦績（両チーム5戦）", size=16, bold=True, color=BLACK, space_after=4)
    add_p(doc, "ニックス（5連勝中）", size=12.5, bold=True, color=NYK, space_after=2)
    for line in DATA["recent_5_games"]["knicks"]:
        add_p(doc, "・" + line, size=12.5, color=BODY, space_after=1)
    add_p(doc, "76ers（3連敗中）", size=12.5, bold=True, color=PHI, space_after=2)
    for line in DATA["recent_5_games"]["sixers"]:
        add_p(doc, "・" + line, size=12.5, color=BODY, space_after=1)

    add_page_break(doc)

    # ========= PAGE 2: ストーリーライン6本 =========
    add_p(doc, "■ 注目ストーリーライン", size=18, bold=True, color=BLACK, space_after=8)
    stories = [
        ("1. 3-0からの逆転、NBA史上わずか1度",
         "プレーオフで3-0からの逆転は1947年以降、2025年のレイカーズvsデンバー戦が史上唯一とされる。76ersはホームでスイープ阻止に背水の陣。直近では昨年同じ東地区で見た歴史を、今度は自分達が起こす番だ。NBAではこのシチュエーションでHOMEチームが第4戦を勝つ確率は約34%（過去全156件・歴史的データ）。"),
        ("2. ブランソン圧巻、3戦平均31.3点",
         "シリーズ平均31.3点・8.0アシストでカンファレンス準決勝MVPクラスのパフォーマンス。第3戦は11-22FGで33点・9アシスト。マイク・ブラウンHCは『俺はライナス、ジェイレンは俺の毛布だ』と評す。第4戦も76ersのエッジコム・マクシーのガード勢の狙い目を上回る精度を維持できるか。"),
        ("3. エンビードの“復帰2戦目”、ヒップ痛Probable",
         "エンビードはGame 2を右足首・右臀部で欠場、Game 3は復帰し35分プレーで18点6リバ3ブロック（FG7-17）。Game 4はヒップ痛のProbableで出場濃厚だが、シリーズ平均16点に留まり、シーズン平均26.9点とは別人。1日空けの中1日でどこまで回復しているかが76ers延命の生命線。"),
        ("4. アヌノビー欠場リスクとブリッジス",
         "アヌノビーは右ハム筋挫傷でG3欠場、G4はQuestionable。G3でブリッジスが23点8-14FGとカバーしたが、G4も欠場ならブリッジスがウィングの守備負担を一身に受ける。G2終盤の負傷後の状態確認は試合前のメディカルスタッフ次第だ。"),
        ("5. ナーシュHCの精神的支柱、兄スティーブ氏への祈り",
         "ナーシュHCは4月29日に兄スティーブ氏（62歳・ノーザン・アイオワ大用具係26年）が急逝。葬儀のため5/5に離脱したが、Game 2前にチームに合流し『兄もコーチで居て欲しいと願っていた』と続行を選択。Game 3前の黙祷でNYKファンの一部が騒いだ件で批判もあった中、76ersは“兄に捧げる勝利”を狙う。"),
        ("6. タウンズの母の日メンタル、ハートのオールラウンド",
         "Game 4は母の日。タウンズは『仕事だよ。ただ、コートを離れて好きなことに集中できるから嬉しい』と冷静。Game 3は8点12リバ7アシストとアシスト・リバウンド両軸で試合を作った。ハートも12点11リバとブルーカラー・スタッツでいぶし銀の貢献を続けている。")
    ]
    for title, body in stories:
        add_p(doc, title, size=15, bold=True, color=BLACK, space_after=2)
        add_p(doc, body, size=13, color=BODY, space_after=8)

    add_page_break(doc)

    # ========= PAGE 3: 直近試合レビュー =========
    add_p(doc, "■ 直近試合レビュー", size=18, bold=True, color=BLACK, space_after=8)
    add_p(doc, "ニューヨーク・ニックス：直近3戦", size=14, bold=True, color=NYK, space_after=4)
    for r in DATA["game_reviews"]["knicks"]:
        add_p(doc, r["head"], size=12.5, bold=True, color=NYK, space_after=2)
        add_p(doc, r["body"], size=12.5, color=BODY, space_after=6)
    add_p(doc, "フィラデルフィア・76ers：直近3戦", size=14, bold=True, color=PHI, space_after=4)
    for r in DATA["game_reviews"]["sixers"]:
        add_p(doc, r["head"], size=12.5, bold=True, color=PHI, space_after=2)
        add_p(doc, r["body"], size=12.5, color=BODY, space_after=6)

    add_page_break(doc)

    # ========= PAGE 4: ケガ人 + チームスタッツ比較 =========
    add_p(doc, "■ ケガ人情報", size=18, bold=True, color=BLACK, space_after=6)

    add_p(doc, "ニューヨーク・ニックス", size=14, bold=True, color=NYK, space_after=2)
    inj_headers = ["選手", "Pos", "ケガ", "ステータス", "備考"]
    inj_rows_nyk = [[i["name"], i["pos"], i["injury"], i["status"], i["note"]] for i in DATA["knicks"]["injuries"]]
    make_table(doc, inj_headers, inj_rows_nyk, widths_cm=[3.5, 1.2, 3.5, 2.5, 7.0], header_bg="F58426")
    add_p(doc, "", size=4, space_after=2)

    add_p(doc, "フィラデルフィア・76ers", size=14, bold=True, color=PHI, space_after=2)
    inj_rows_phi = [[i["name"], i["pos"], i["injury"], i["status"], i["note"]] for i in DATA["sixers"]["injuries"]]
    make_table(doc, inj_headers, inj_rows_phi, widths_cm=[3.5, 1.2, 3.5, 2.5, 7.0], header_bg="006BB6")
    add_p(doc, "", size=6, space_after=4)

    # チームスタッツ比較
    add_p(doc, "■ チーム状況比較", size=16, bold=True, color=BLACK, space_after=4)
    stats_headers = ["項目", "ニューヨーク・ニックス", "フィラデルフィア・76ers", "優位"]
    stats_rows = [[s["item"], s["knicks"], s["sixers"], s["edge"]] for s in DATA["team_stats_compare"]]
    make_table(doc, stats_headers, stats_rows, widths_cm=[4.0, 5.5, 5.5, 1.5], header_bg="0F172A")
    add_p(doc, "出典：NBA.com / ESPN / CBS Sports / Basketball-Reference 2025-26", size=7.5, color=GREY, space_after=4)

    add_page_break(doc)

    # ========= PAGE 5: 注目選手平均 + 前回対戦 + キーマッチアップ =========
    add_p(doc, "■ 注目選手シーズン平均（2025-26）", size=18, bold=True, color=BLACK, space_after=4)
    add_p(doc, "ニューヨーク・ニックス", size=14, bold=True, color=NYK, space_after=2)
    avg_headers = ["選手", "GP", "MPG", "PPG", "RPG", "APG", "FG%", "3P%"]
    nyk_avg = [[p["name"], p["gp"], p["mpg"], p["ppg"], p["rpg"], p["apg"], f'{p["fg"]}%', f'{p["tp"]}%'] for p in DATA["knicks"]["key_players_for_avg"]]
    make_table(doc, avg_headers, nyk_avg, widths_cm=[4.5, 1.4, 1.6, 1.6, 1.6, 1.6, 2.0, 2.0], header_bg="F58426")
    add_p(doc, "", size=4, space_after=2)
    add_p(doc, "フィラデルフィア・76ers", size=14, bold=True, color=PHI, space_after=2)
    phi_avg = [[p["name"], p["gp"], p["mpg"], p["ppg"], p["rpg"], p["apg"], f'{p["fg"]}%', f'{p["tp"]}%'] for p in DATA["sixers"]["key_players_for_avg"]]
    make_table(doc, avg_headers, phi_avg, widths_cm=[4.5, 1.4, 1.6, 1.6, 1.6, 1.6, 2.0, 2.0], header_bg="006BB6")
    add_p(doc, "出典：ESPN / Basketball-Reference 2025-26レギュラーシーズン平均", size=7.5, color=GREY, space_after=6)

    # 前回対戦
    add_p(doc, "■ レギュラーシーズン直接対決（4試合 2-2）", size=16, bold=True, color=BLACK, space_after=4)
    for h in DATA["regular_season_h2h"]:
        winner_color = NYK if h["winner"] == "NYK" else PHI
        winner_ja = "ニックス" if h["winner"] == "NYK" else "76ers"
        venue_ja = "MSG" if h["at"] == "NYK" else "ザ・サウス・フィラデルフィア・スポーツ・コンプレックス"
        head = f"{h['date']}（@{venue_ja}）：{winner_ja}勝 {h['score']}"
        add_p(doc, head, size=12, bold=True, color=winner_color, space_after=2)
        add_p(doc, h["note"], size=12, color=BODY, space_after=6)

    add_p(doc, "", size=4, space_after=2)
    add_p(doc, "■ キーマッチアップ", size=16, bold=True, color=BLACK, space_after=4)
    for m in DATA["key_matchups"]:
        add_p(doc, m["title"], size=13, bold=True, color=BLACK, space_after=2)
        add_p(doc, m["body"], size=12, color=BODY, space_after=6)

    add_page_break(doc)

    # ========= PAGE 6-7: コメント =========
    add_p(doc, "■ 指揮官・選手コメント集", size=18, bold=True, color=BLACK, space_after=8)
    add_p(doc, "ニューヨーク・ニックス", size=14, bold=True, color=NYK, space_after=4)
    for c in DATA["comments"]["knicks"]:
        add_p(doc, f'【{c["speaker"]}｜{c["date"]} {c["context"]}】', size=11, bold=True, color=BLACK, space_after=2)
        add_p(doc, f'「{c["quote"]}」', size=13, color=BODY, space_after=1)
        add_p(doc, f'─ {c["src"]} ({c["date"]})', size=7.5, color=GREY, space_after=6)

    add_p(doc, "", size=4, space_after=2)
    add_p(doc, "フィラデルフィア・76ers", size=14, bold=True, color=PHI, space_after=4)
    for c in DATA["comments"]["sixers"]:
        add_p(doc, f'【{c["speaker"]}｜{c["date"]} {c["context"]}】', size=11, bold=True, color=BLACK, space_after=2)
        add_p(doc, f'「{c["quote"]}」', size=13, color=BODY, space_after=1)
        add_p(doc, f'─ {c["src"]} ({c["date"]})', size=7.5, color=GREY, space_after=6)

    add_page_break(doc)

    # ========= PAGE 8: NBA TOPICS =========
    add_p(doc, "■ NBA全体トピックス", size=18, bold=True, color=BLACK, space_after=8)
    for t in DATA["nba_topics"]:
        add_p(doc, t["title"], size=14, bold=True, color=BLACK, space_after=2)
        add_p(doc, t["body"], size=13, color=BODY, space_after=2)
        add_p(doc, f'（出典：{t["src"]}）', size=7.5, color=GREY, space_after=8)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print(f"SAVED: {OUT}")


if __name__ == "__main__":
    main()
