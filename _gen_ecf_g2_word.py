# -*- coding: utf-8 -*-
"""
CLE at NYK ECF Game 2 GamePreview Word Generator (2026-05-21)
Knicks vs Cavaliers Eastern Conference Finals Game 2
JST 2026-05-22 9:00 AM tipoff / MSG / ESPN
"""
import json, os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import nsdecls, qn
from docx.oxml import parse_xml, OxmlElement

DATA = json.load(open("/home/user/search-launcher/_data_nyk_cle_ecf_g2.json", "r", encoding="utf-8"))
OUT = "/home/user/search-launcher/CLE_at_NYK_GamePreview_ECF_G2_20260521.docx"

# Team colors
NYK = RGBColor(0xF5, 0x84, 0x26)
CLE = RGBColor(0x86, 0x00, 0x38)   # CLE wine
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
    add_p(doc, "2026 プレーオフ 東カンファレンス・ファイナル 第2戦", size=14, bold=True, color=BLACK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=8)
    add_p(doc, "クリーブランド・キャバリアーズ", size=22, bold=True, color=CLE, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_p(doc, "@", size=14, bold=True, color=BLACK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_p(doc, "ニューヨーク・ニックス", size=22, bold=True, color=NYK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=6)
    add_p(doc, "2026年5月21日（現地・木）／日本時間 5/22(金) 9:00 AM JST", size=12, color=BODY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=2)
    add_p(doc, f"{M['venue']} ／ {M['tv']}", size=12, color=BODY, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=4)
    add_p(doc, f"シリーズ：ニックス1勝0敗リード（G1は115-104 OT勝ち、22点差逆転）", size=13, bold=True, color=BLACK, align=WD_ALIGN_PARAGRAPH.CENTER, space_after=10)

    add_p(doc, "■ 試合の焦点", size=18, bold=True, color=BLACK, space_after=6)
    for t in DATA["focus_points"]:
        add_p(doc, t, size=13, color=BODY, space_after=3)

    add_p(doc, "", size=8, space_after=4)

    add_p(doc, "■ 直近の戦績（両チーム5戦）", size=16, bold=True, color=BLACK, space_after=4)
    add_p(doc, "ニックス（プレーオフ10戦9勝）", size=12.5, bold=True, color=NYK, space_after=2)
    for line in DATA["recent_5_games"]["knicks"]:
        add_p(doc, "・" + line, size=12.5, color=BODY, space_after=1)
    add_p(doc, "キャバリアーズ（連続Game 7突破後の長旅）", size=12.5, bold=True, color=CLE, space_after=2)
    for line in DATA["recent_5_games"]["cavs"]:
        add_p(doc, "・" + line, size=12.5, color=BODY, space_after=1)

    add_page_break(doc)

    # ========= PAGE 2: STORYLINES =========
    add_p(doc, "■ 注目ストーリーライン", size=18, bold=True, color=BLACK, space_after=8)
    stories = [
        ("1. 4Q 22点差逆転、99.9%勝率からの崩壊",
         "G1（5/19）はキャブズが3Q終了時に93-71で22点リード。残7:52時点のESPNアナリティクス勝率は99.9%だった。そこからニックスが44-11のランで一気に逆転、OT勝ち。フランチャイズ史上最大の4Q逆転勝ち、プレーバイプレー時代のNBA 4Q22点差逆転は1位タイ。ミッチェルは「We f---ing blew it（マジで台無しにした）」と認めた。G2の精神面の建て直しがキャブズの最大テーマ。"),
        ("2. ハーデン狙い継続か、アトキンソンHCの修正力",
         "マイク・ブラウンHCがG1試合後「ハーデンを攻めるのは秘密じゃなかった」と公言。ハーデンはG1でFG 5-16・3P 1-8・TO 6・+/- -6と崩壊。アトキンソンHCがG2でハーデンを隠す起用法（より早いスイッチ、シュルーダーとの併用、第4Q終盤の役割縮小）を打ち出すかが最大の戦術ポイント。"),
        ("3. ブランソン vs ミッチェル、トレード歴を持つ運命の対決",
         "2022年7月にニックスがブランソンを獲得した数週間後、ミッチェル獲得をクリーブランドに逆転され失った因縁。G1はブランソンが46分38点（4Q＋OTで17点）、ミッチェルは29点・6スティールも4Q＋OTでFG 1-6と失速。ミッチェルはウェストチェスター郡エルムスフォード（MSGから車で約45分）出身、地元での初CF。"),
        ("4. 連続Game 7のCLEと9日休みのNYK、コンディションの相剋",
         "キャブズは1回戦ラプターズ（4-3、Game 7勝利）、2回戦ピストンズ（4-3、0-2から4連勝）と過酷な連戦。22日間で14試合をこなしてECF入り。一方ニックスは76ersを4-0スイープ後9日間の休養。「疲労 vs リズム喪失」の競争。MブラウンHCも「両方経験している、どちらが有利か答えはない」と語る。"),
        ("5. マイク・ブラウン vs ケニー・アトキンソン、ウォリアーズ仲間対決",
         "両HCはスティーブ・カー時代のGSWアシスタント仲間。Brown：2008-09＆2022-23 COY（後者は満票受賞史上初）。アトキンソン：2024-25 COY満票受賞（前年）。G1ではブラウンの仕掛けが上回ったが、シリーズ通算でアトキンソンの応手が試される。"),
        ("6. アヌノビー復帰戦の精彩、G2でフル戻りなるか",
         "右ハム筋負傷からG1で復帰したOGアヌノビーは13点（FG 2-9）と精彩を欠いた。試合後「痛みより違和感が強い」と語る。G2でフィジカルとシュート感覚が戻れば、ブリッジスと並ぶウィングの2枚看板が完全復活し、ニックスのディフェンスは一段上のレベルに。両軍G2前の正式インジュリーレポートは『怪我人ゼロ』。")
    ]
    for title, body in stories:
        add_p(doc, title, size=15, bold=True, color=BLACK, space_after=2)
        add_p(doc, body, size=13, color=BODY, space_after=8)

    add_page_break(doc)

    # ========= PAGE 3 =========
    add_p(doc, "■ 直近試合レビュー", size=18, bold=True, color=BLACK, space_after=8)
    add_p(doc, "ニューヨーク・ニックス：直近3戦", size=14, bold=True, color=NYK, space_after=4)
    for r in DATA["game_reviews"]["knicks"]:
        add_p(doc, r["head"], size=12.5, bold=True, color=NYK, space_after=2)
        add_p(doc, r["body"], size=12.5, color=BODY, space_after=6)
    add_p(doc, "クリーブランド・キャバリアーズ：直近3戦", size=14, bold=True, color=CLE, space_after=4)
    for r in DATA["game_reviews"]["cavs"]:
        add_p(doc, r["head"], size=12.5, bold=True, color=CLE, space_after=2)
        add_p(doc, r["body"], size=12.5, color=BODY, space_after=6)

    add_page_break(doc)

    # ========= PAGE 4 =========
    add_p(doc, "■ ケガ人情報", size=18, bold=True, color=BLACK, space_after=6)

    add_p(doc, "ニューヨーク・ニックス", size=14, bold=True, color=NYK, space_after=2)
    inj_headers = ["選手", "Pos", "ケガ", "ステータス", "備考"]
    inj_rows_nyk = [[i["name"], i["pos"], i["injury"], i["status"], i["note"]] for i in DATA["knicks"]["injuries"]]
    make_table(doc, inj_headers, inj_rows_nyk, widths_cm=[3.5, 1.2, 3.5, 2.5, 7.0], header_bg="F58426")
    add_p(doc, "", size=4, space_after=2)

    add_p(doc, "クリーブランド・キャバリアーズ", size=14, bold=True, color=CLE, space_after=2)
    inj_rows_cle = [[i["name"], i["pos"], i["injury"], i["status"], i["note"]] for i in DATA["cavs"]["injuries"]]
    make_table(doc, inj_headers, inj_rows_cle, widths_cm=[3.5, 1.2, 3.5, 2.5, 7.0], header_bg="860038")
    add_p(doc, "", size=6, space_after=4)

    add_p(doc, "■ チーム状況比較", size=16, bold=True, color=BLACK, space_after=4)
    stats_headers = ["項目", "ニューヨーク・ニックス", "クリーブランド・キャバリアーズ", "優位"]
    stats_rows = [[s["item"], s["knicks"], s["cavs"], s["edge"]] for s in DATA["team_stats_compare"]]
    make_table(doc, stats_headers, stats_rows, widths_cm=[4.0, 5.5, 5.5, 1.5], header_bg="0F172A")
    add_p(doc, "出典：NBA.com / ESPN / CBS Sports / Basketball-Reference 2025-26", size=7.5, color=GREY, space_after=4)

    add_page_break(doc)

    # ========= PAGE 5 =========
    add_p(doc, "■ 注目選手シーズン平均（2025-26）", size=18, bold=True, color=BLACK, space_after=4)
    add_p(doc, "ニューヨーク・ニックス", size=14, bold=True, color=NYK, space_after=2)
    avg_headers = ["選手", "GP", "MPG", "PPG", "RPG", "APG", "FG%", "3P%"]
    nyk_avg = [[p["name"], p["gp"], p["mpg"], p["ppg"], p["rpg"], p["apg"], f'{p["fg"]}%', f'{p["tp"]}%'] for p in DATA["knicks"]["key_players_for_avg"]]
    make_table(doc, avg_headers, nyk_avg, widths_cm=[4.5, 1.4, 1.6, 1.6, 1.6, 1.6, 2.0, 2.0], header_bg="F58426")
    add_p(doc, "", size=4, space_after=2)
    add_p(doc, "クリーブランド・キャバリアーズ", size=14, bold=True, color=CLE, space_after=2)
    cle_avg = [[p["name"], p["gp"], p["mpg"], p["ppg"], p["rpg"], p["apg"], f'{p["fg"]}%', f'{p["tp"]}%'] for p in DATA["cavs"]["key_players_for_avg"]]
    make_table(doc, avg_headers, cle_avg, widths_cm=[4.5, 1.4, 1.6, 1.6, 1.6, 1.6, 2.0, 2.0], header_bg="860038")
    add_p(doc, "※ ハーデンはキャブズ加入後（2/4以降）のスタッツ。出典：ESPN / Basketball-Reference 2025-26", size=7.5, color=GREY, space_after=6)

    add_p(doc, "■ レギュラーシーズン直接対決（3試合 NYK 2-1）", size=16, bold=True, color=BLACK, space_after=4)
    for h in DATA["regular_season_h2h"]:
        winner_color = NYK if h["winner"] == "NYK" else CLE
        winner_ja = "ニックス" if h["winner"] == "NYK" else "キャバリアーズ"
        venue_ja = "MSG" if h["at"] == "NYK" else "Rocket Arena"
        head = f"{h['date']}（@{venue_ja}）：{winner_ja}勝 {h['score']}"
        add_p(doc, head, size=12, bold=True, color=winner_color, space_after=2)
        add_p(doc, h["note"], size=12, color=BODY, space_after=6)

    add_p(doc, "", size=4, space_after=2)
    add_p(doc, "■ キーマッチアップ", size=16, bold=True, color=BLACK, space_after=4)
    for m in DATA["key_matchups"]:
        add_p(doc, m["title"], size=13, bold=True, color=BLACK, space_after=2)
        add_p(doc, m["body"], size=12, color=BODY, space_after=6)

    add_page_break(doc)

    # ========= PAGE 6-7: COMMENTS =========
    add_p(doc, "■ 指揮官・選手コメント集", size=18, bold=True, color=BLACK, space_after=8)
    add_p(doc, "ニューヨーク・ニックス", size=14, bold=True, color=NYK, space_after=4)
    for c in DATA["comments"]["knicks"]:
        add_p(doc, f'【{c["speaker"]}｜{c["date"]} {c["context"]}】', size=11, bold=True, color=BLACK, space_after=2)
        add_p(doc, f'「{c["quote"]}」', size=13, color=BODY, space_after=1)
        add_p(doc, f'─ {c["src"]} ({c["date"]})', size=7.5, color=GREY, space_after=6)

    add_p(doc, "", size=4, space_after=2)
    add_p(doc, "クリーブランド・キャバリアーズ", size=14, bold=True, color=CLE, space_after=4)
    for c in DATA["comments"]["cavs"]:
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
