# -*- coding: utf-8 -*-
"""
CLE at NYK ECF Game 2 Excel stats sheet
2026-05-21 (JST 2026-05-22) / Knicks vs Cavaliers Eastern Conference Finals Game 2

Notes:
  * Cavs roster reflects May 2026 status:
    - Garland TRADED to Clippers (2/4 for Harden + 2nd)
    - Hunter TRADED to Sacramento (2/1, return: Schroder + Ellis)
    - Ty Jerome left to Memphis (offseason 2025)
    - Tristan Thompson retired
  * Column order: leftmost = newest game (G1 of ECF, 5/19)
    For Knicks: G1 ECF (5/19), R2 G4 vs PHI (5/10), R2 G3 vs PHI (5/8)
    For Cavs: G1 ECF (5/19), R2 G7 vs DET (5/17), R2 G6 vs DET (5/15)
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, datetime

OUT = "/home/user/search-launcher/nba_stats_CLE_NYK_ECF_G2_20260521.xlsx"
GAME_DATE = date(2026, 5, 21)

NYK_BG = "F58426"
CLE_BG = "860038"
SUB_BG = "374151"
GREY_FILL = "F3F4F6"

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def calc_age(bday_str, on_date):
    if not bday_str:
        return ""
    bd = datetime.strptime(bday_str, "%Y-%m-%d").date()
    yrs = on_date.year - bd.year
    if (on_date.month, on_date.day) < (bd.month, bd.day):
        yrs -= 1
    return yrs


WIDTHS = [5, 22, 5, 8, 8, 5, 7, 40] + [5, 5, 5, 5, 5, 5] * 3
COLS = 8 + 18


NYK_PLAYERS = [
    (11, "ジェイレン・ブランソン", "PG", 188, 86, "1996-08-31", 7, "ヴィラノバ大出身。25-26平均26.0点のスーパースター司令塔"),
    (25, "マイケル・ブリッジス", "SG", 198, 95, "1996-08-30", 7, "ヴィラノバ大出身。耐久性と守備力を兼備するウィング"),
    (8, "OG・アヌノビー", "SF", 201, 109, "1997-07-17", 8, "インディアナ大出身。ハム筋負傷からECF G1で復帰、3&D駒"),
    (3, "ジョシュ・ハート", "SF", 196, 96, "1995-03-06", 8, "ヴィラノバ大出身。リバウンドとハッスルが武器のグルーガイ"),
    (32, "カール=アンソニー・タウンズ", "C", 213, 112, "1995-11-15", 10, "ケンタッキー大出身。3点も打てる万能型ビッグマン"),
    (23, "ミッチェル・ロビンソン", "C", 213, 109, "1998-04-01", 7, "高卒組の大型C。リム守備とOREBが代名詞"),
    (2, "マイルズ・マクブライド", "PG", 188, 88, "2000-09-08", 4, "ウェストバージニア大出身。Deuceの愛称、対人守備の切り札"),
    (15, "ホセ・アルバラード", "PG", 183, 81, "1998-04-12", 4, "ジョージア工科大出身。2月にペリカンズから加入、闘犬PG"),
    (20, "ランドリー・シャメット", "SG", 196, 86, "1997-03-13", 7, "ウィチタ州立大出身。3P39%超のシューター"),
    (6, "ジョーダン・クラークソン", "SG", 196, 88, "1992-06-07", 11, "ミズーリ大出身。21年6thMan賞受賞のスコアラー"),
    (55, "アリエル・フクポルティ", "C", 213, 112, "2002-04-12", 1, "ドイツ出身の若手C。第3ユニットの控え"),
    (13, "タイラー・コレック", "PG", 188, 88, "2001-03-27", 1, "マーケット大出身2年目PG。Gリーグで2桁アシスト連発"),
    (4, "パコメ・ダディエ", "SF", 206, 91, "2005-07-27", 1, "フランス出身2年目。Gリーグと往復"),
    (5, "モハメド・ディアワラ", "PF", 206, 102, "2005-04-29", 0, "フランス出身の25年2巡指名ルーキー"),
    (10, "ジェレミー・ソーチャン", "PF", 203, 104, "2003-05-20", 3, "ベイラー大出身。2/13にスパーズから加入"),
]

# Cavs roster - corrected for May 2026
CLE_PLAYERS = [
    # スターター（G1 ECF）
    (45, "ドノバン・ミッチェル", "SG", 188, 97, "1996-09-07", 8, "ルイビル大出身。25-26平均27.9点、6度目のオールスター・エース"),
    (1, "ジェームズ・ハーデン", "PG", 196, 100, "1989-08-26", 16, "アリゾナ州立大出身、3度の得点王。2/4にLACから移籍"),
    (1, "マックス・ストルース", "SF", 198, 98, "1996-03-28", 6, "デポール大出身。スターター級3&Dウィング、3P 37.5%"),
    (4, "エヴァン・モブリー", "PF", 211, 97, "2001-06-18", 4, "USC出身。2024-25 DPOY、平均18.3点9.1リバ1.8ブロック"),
    (31, "ジャレット・アレン", "C", 206, 110, "1998-04-21", 8, "テキサス大出身。オールスターC、FG 68.7%の決定力"),
    # ベンチ
    (17, "デニス・シュルーダー", "PG", 185, 79, "1993-09-15", 12, "ドイツ出身。2/1にSACから加入、25年EuroBasket MVP"),
    (3, "キーオン・エリス", "SG", 193, 79, "2000-05-08", 4, "アラバマ大出身。2/1にSACから加入、ペリメーターD"),
    (15, "サム・メリル", "SG", 196, 92, "1996-05-20", 5, "ユタ州立大出身。3P 42.1%のシューター、G7で23点"),
    (32, "ディーン・ウェイド", "PF", 206, 100, "1996-11-20", 6, "カンザス州立大出身。ストレッチ4、G1で4Pプレー成功"),
    (2, "ロンゾ・ボール", "PG", 196, 86, "1997-10-27", 7, "UCLA出身。シーズン序盤に攻撃面で苦戦、ローテ外時期も"),
    (21, "ジェイロン・タイソン", "SF", 198, 98, "2002-12-02", 1, "カリフォルニア大出身。G1ではDNP-CD"),
    (40, "クレイグ・ポーター Jr.", "PG", 188, 84, "2000-05-09", 2, "ウィチタ州立大出身。3番手PG"),
    (31, "タイリース・プロクター", "PG", 193, 86, "2004-08-15", 0, "デューク大出身。25年ドラフト49位のルーキー"),
    (4, "サリオ・ニャン", "SF", 198, 89, "2004-04-15", 0, "セネガル/イタリア出身。25年ドラフト58位のルーキー"),
]


# BOX SCORES - columns: [G1 ECF 5/19, R2 G4 5/10, R2 G3 5/8]
NYK_BOX = {
    "G1_label": "ECF G1 5/19 ○115-104 OT vs CLE",
    "G2_label": "R2 G4 5/10 ○144-114 @PHI (sweep)",
    "G3_label": "R2 G3 5/8 ○108-94 @PHI",
    "data": {
        "ジェイレン・ブランソン":     [[46, 38, 1, 6, 6, 5], ["—", 28, "—", "—", "—", "—"], [38, 33, 3, 8, 9, 5]],
        "マイケル・ブリッジス":       [[42, 18, 2, 4, 1, 5], ["—", "—", "—", "—", "—", "—"], [37, 23, 2, 4, 1, 3]],
        "OG・アヌノビー":             [["—", 13, "—", "—", "—", "—"], "OUT", "OUT"],
        "ジョシュ・ハート":           [[31, 6, "—", "—", 3, 5], ["—", "—", "—", "—", "—", "—"], [40, 12, 0, 4, 3, 11]],
        "カール=アンソニー・タウンズ": [[40, 13, 1, 5, 5, 13], ["—", 16, "—", "—", "—", "—"], [26, 8, 0, 2, 7, 12]],
        "ミッチェル・ロビンソン":     [[14, "—", 0, 0, "—", 6], ["—", "—", 0, 0, "—", "—"], [19, 6, 0, 0, "—", 6]],
        "マイルズ・マクブライド":     [["—", 0, 0, 1, "—", "—"], ["—", 25, 7, "—", "—", "—"], [21, 3, 1, 5, 2, "—"]],
        "ホセ・アルバラード":         [["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"], ["—", 3, 1, "—", "—", "—"]],
        "ランドリー・シャメット":     [["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"], [26, 15, 2, 3, "—", 3]],
        "ジョーダン・クラークソン":   [["4:33", 3, 1, 1, "—", 1], ["—", "—", "—", "—", "—", "—"], ["—", 4, "—", "—", 3, 5]],
        "アリエル・フクポルティ":     ["DNP", "DNP", "DNP"],
        "タイラー・コレック":         ["DNP", "DNP", "DNP"],
        "パコメ・ダディエ":           ["DNP", "DNP", "DNP"],
        "モハメド・ディアワラ":       ["DNP", "DNP", "DNP"],
        "ジェレミー・ソーチャン":     ["DNP", "DNP", "DNP"],
    }
}

CLE_BOX = {
    "G1_label": "ECF G1 5/19 ●104-115 OT @NYK",
    "G2_label": "R2 G7 5/17 ○125-94 vs DET (突破)",
    "G3_label": "R2 G6 5/15 ●94-115 @DET",
    "data": {
        "ドノバン・ミッチェル":       [[41, 29, 4, 11, 3, 5],   ["—", 35, "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "ジェームズ・ハーデン":       [[42, 15, 1, 8, 3, 4],    ["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "マックス・ストルース":       [[25, 8, 2, 6, 1, 2],     ["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "エヴァン・モブリー":         [[29, 15, "—", "—", "—", 14], ["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "ジャレット・アレン":         [[35, 10, 0, 0, 3, 7],    ["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "デニス・シュルーダー":       [["—", 2, "—", "—", 5, "—"], ["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "キーオン・エリス":           [["5:45", 2, "—", "—", 0, 0], ["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "サム・メリル":               [["—", "—", "—", "—", "—", "—"], ["—", 23, "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "ディーン・ウェイド":         [["9+", 3, 1, 2, 0, 0],   ["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "ロンゾ・ボール":             [["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "ジェイロン・タイソン":       ["DNP",                   ["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "クレイグ・ポーター Jr.":     ["DNP",                   ["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "タイリース・プロクター":     ["DNP",                   ["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
        "サリオ・ニャン":             ["DNP",                   ["—", "—", "—", "—", "—", "—"], ["—", "—", "—", "—", "—", "—"]],
    }
}


def write_team_block(ws, start_row, team_name, hc, players, box, team_bg):
    HEADER_FONT = Font(name="MS Gothic", size=10, bold=True, color="FFFFFF")
    SUB_FONT = Font(name="MS Gothic", size=8, bold=True, color="FFFFFF")
    ROW_FONT = Font(name="MS Gothic", size=9, color="000000")
    GAME_FONT = Font(name="MS Gothic", size=8, color="000000")
    TANPYO_FONT = Font(name="MS Gothic", size=11, color="000000")
    DNP_FONT = Font(name="MS Gothic", size=8, italic=True, color="999999")
    BAND_FONT = Font(name="MS Gothic", size=14, bold=True, color="FFFFFF")

    HEADER_FILL = PatternFill("solid", fgColor=team_bg)
    SUB_FILL = PatternFill("solid", fgColor=SUB_BG)
    EVEN_FILL = PatternFill("solid", fgColor=GREY_FILL)

    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)

    band_row = start_row
    ws.cell(row=band_row, column=1, value=team_name).font = BAND_FONT
    ws.merge_cells(start_row=band_row, end_row=band_row, start_column=1, end_column=COLS)
    for c in range(1, COLS + 1):
        cell = ws.cell(row=band_row, column=c)
        cell.fill = PatternFill("solid", fgColor=team_bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[band_row].height = 22

    r1 = start_row + 1
    ws.cell(row=r1, column=1, value=f"HC {hc}").font = HEADER_FONT
    ws.merge_cells(start_row=r1, end_row=r1, start_column=1, end_column=3)
    for c in range(1, 4):
        cc = ws.cell(row=r1, column=c)
        cc.fill = HEADER_FILL; cc.alignment = CENTER; cc.border = BORDER
    headers_r1_dg = ["身長", "体重", "年齢", "NBA歴"]
    for i, h in enumerate(headers_r1_dg):
        cc = ws.cell(row=r1, column=4 + i, value=h)
        cc.font = HEADER_FONT; cc.fill = HEADER_FILL; cc.alignment = CENTER; cc.border = BORDER
    cc = ws.cell(row=r1, column=8, value="短評")
    cc.font = HEADER_FONT; cc.fill = HEADER_FILL; cc.alignment = CENTER; cc.border = BORDER
    for k, label in enumerate([box["G1_label"], box["G2_label"], box["G3_label"]]):
        c0 = 9 + k * 6
        ws.cell(row=r1, column=c0, value=label).font = HEADER_FONT
        ws.merge_cells(start_row=r1, end_row=r1, start_column=c0, end_column=c0 + 5)
        for cc_col in range(c0, c0 + 6):
            cc = ws.cell(row=r1, column=cc_col)
            cc.fill = HEADER_FILL; cc.alignment = CENTER; cc.border = BORDER
    ws.row_dimensions[r1].height = 18

    r2 = start_row + 2
    sub = ["No.", "選手名", "Pos", "", "", "", "", ""] + ["MIN", "PTS", "3PM", "3PA", "AST", "REB"] * 3
    for i, h in enumerate(sub):
        cc = ws.cell(row=r2, column=i + 1, value=h)
        cc.font = SUB_FONT; cc.fill = SUB_FILL; cc.alignment = CENTER; cc.border = BORDER
    ws.row_dimensions[r2].height = 18

    rr = start_row + 3
    for idx, p in enumerate(players):
        no, name, pos, h_cm, w_kg, bday, nba_yrs, tanpyo = p
        age = calc_age(bday, GAME_DATE)
        nba_label = f"{nba_yrs + 1}年目" if nba_yrs >= 0 else "1年目"
        row_fill = EVEN_FILL if idx % 2 == 0 else None

        def w(col, val, font=ROW_FONT, align=CENTER):
            cc = ws.cell(row=rr, column=col, value=val)
            cc.font = font; cc.alignment = align; cc.border = BORDER
            if row_fill:
                cc.fill = row_fill

        w(1, no)
        w(2, name, align=LEFT)
        w(3, pos)
        w(4, f"{h_cm}cm")
        w(5, f"{w_kg}kg")
        w(6, age)
        w(7, nba_label)
        w(8, tanpyo, font=TANPYO_FONT, align=LEFT)

        boxdata = box["data"].get(name, ["DNP", "DNP", "DNP"])
        for k in range(3):
            c0 = 9 + k * 6
            v = boxdata[k]
            if isinstance(v, list):
                for j, val in enumerate(v):
                    cc = ws.cell(row=rr, column=c0 + j, value=val)
                    cc.font = GAME_FONT; cc.alignment = CENTER; cc.border = BORDER
                    if row_fill:
                        cc.fill = row_fill
            else:
                ws.merge_cells(start_row=rr, end_row=rr, start_column=c0, end_column=c0 + 5)
                cc = ws.cell(row=rr, column=c0, value=v)
                cc.font = DNP_FONT; cc.alignment = CENTER; cc.border = BORDER
                if row_fill:
                    cc.fill = row_fill
                for jj in range(c0, c0 + 6):
                    ccc = ws.cell(row=rr, column=jj)
                    ccc.border = BORDER
                    if row_fill:
                        ccc.fill = row_fill

        ws.row_dimensions[rr].height = 28
        rr += 1

    return rr


def main():
    wb = Workbook()
    ws = wb.active
    ws.title = "NYK_CLE_ECF_G2_5_21"

    for i, w in enumerate(WIDTHS):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True

    ws.cell(row=1, column=1, value="2026 NBAプレーオフ 東カンファレンス・ファイナル 第2戦  クリーブランド・キャバリアーズ @ ニューヨーク・ニックス  2026/5/21(現地)・JST 5/22 9:00 タップオフ ／ シリーズ 1勝0敗（NYKリード／G1 OT勝ち22点差逆転）  ※ G列順は新しい試合が左").font = Font(name="MS Gothic", size=12, bold=True)
    ws.merge_cells(start_row=1, end_row=1, start_column=1, end_column=COLS)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    next_row = write_team_block(ws, 2, "ニューヨーク・ニックス（東3位 53勝29敗 / ECFシリーズ1-0リード）", "マイク・ブラウン", NYK_PLAYERS, NYK_BOX, NYK_BG)
    next_row += 1
    write_team_block(ws, next_row, "クリーブランド・キャバリアーズ（東4位 52勝30敗 / ECFシリーズ0-1）", "ケニー・アトキンソン", CLE_PLAYERS, CLE_BOX, CLE_BG)

    ws.freeze_panes = "I3"

    wb.save(OUT)
    print(f"SAVED: {OUT}")


if __name__ == "__main__":
    main()
