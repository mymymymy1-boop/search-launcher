# -*- coding: utf-8 -*-
"""
PHI at NYK ECSF Game 4 Excel stats sheet (1 tab, both teams stacked)
2026-05-10 / Knicks vs 76ers Game 4
"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, datetime

OUT = "/home/user/search-launcher/nba_stats_PHI_NYK_20260510.xlsx"
GAME_DATE = date(2026, 5, 10)

NYK_BG = "F58426"   # NYK orange
PHI_BG = "006BB6"   # PHI blue
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
COLS = 8 + 18  # 26


# ---------- KNICKS ROSTER + 短評 ----------
# (no, name_ja, pos, height_cm, weight_kg, birthday, nba_years, tanpyo)
NYK_PLAYERS = [
    # スターター（G4想定）
    (11, "ジェイレン・ブランソン", "PG", 188, 86, "1996-08-31", 7, "ヴィラノバ大出身。25-26平均26点のスーパースター司令塔"),
    (25, "マイケル・ブリッジス", "SG", 198, 95, "1996-08-30", 7, "ヴィラノバ大出身。耐久性と守備力を兼備するウィング"),
    (8, "OG・アヌノビー", "SF", 201, 109, "1997-07-17", 8, "【Q/右ハム】インディアナ大出身。3&D駒、G3欠場"),
    (3, "ジョシュ・ハート", "SF", 196, 96, "1995-03-06", 8, "ヴィラノバ大出身。リバウンドとハッスルが武器のグルーガイ"),
    (32, "カール=アンソニー・タウンズ", "C", 213, 112, "1995-11-15", 10, "ケンタッキー大出身。3点も打てる万能型ビッグマン"),
    # ベンチ
    (23, "ミッチェル・ロビンソン", "C", 213, 109, "1998-04-01", 7, "高卒組の大型C。リム守備とOREBが代名詞"),
    (2, "マイルズ・マクブライド", "PG", 188, 88, "2000-09-08", 4, "ウェストバージニア大出身。Deuceの愛称、対人守備の切り札"),
    (15, "ホセ・アルバラード", "PG", 183, 81, "1998-04-12", 4, "ジョージア工科大出身。2月にペリカンズから加入、闘犬PG"),
    (20, "ランドリー・シャメット", "SG", 196, 86, "1997-03-13", 7, "ウィチタ州立大出身。3P39%超のシューター"),
    (6, "ジョーダン・クラークソン", "SG", 196, 88, "1992-06-07", 11, "ミズーリ大出身。21年6thMan賞受賞のスコアラー"),
    (55, "アリエル・フクポルティ", "C", 213, 112, "2002-04-12", 1, "ドイツ出身の若手C。第3ユニットの控え"),
    (13, "タイラー・コレック", "PG", 188, 88, "2001-03-27", 1, "マーケット大出身2年目PG。プレーオフ出場機会限定的"),
    (4, "パコメ・ダディエ", "SF", 206, 91, "2005-07-27", 1, "フランス出身2年目。Gリーグと往復、Hawks戦でPO初出場"),
    (5, "モハメド・ディアワラ", "PF", 206, 102, "2005-04-29", 0, "フランス出身の25年2巡指名ルーキー"),
    (10, "ジェレミー・ソーチャン", "PF", 203, 104, "2003-05-20", 3, "ベイラー大出身。2/13に加入、フロントコート補強"),
]

# ---------- 76ERS ROSTER + 短評 ----------
PHI_PLAYERS = [
    # スターター（G4想定）
    (0, "タイリース・マクシー", "PG", 188, 90, "2000-11-04", 5, "ケンタッキー大出身、爆速ガード兼第2の柱"),
    (77, "VJ・エッジコム", "SG", 193, 86, "2005-07-30", 0, "ベイラー大出身、新人王最終候補のバハマ出身ルーキー"),
    (9, "ケリー・オウブレJr.", "SF", 203, 92, "1995-12-09", 10, "カンザス大出身、爆発力のあるアスレチック系3&D"),
    (8, "ポール・ジョージ", "PF", 203, 104, "1990-05-02", 15, "フレズノ州立大出身、コービー敬意で背番号8の9度オールスター"),
    (21, "ジョエル・エンビード", "C", 213, 127, "1994-03-16", 9, "【Probable/右臀部・足首】カンザス大、23年MVPの絶対エース"),
    # ベンチ
    (5, "クエンティン・グライムス", "SG", 196, 95, "2000-05-08", 4, "ヒューストン大出身、シックスマンの3&Dガード"),
    (1, "アンドレ・ドラモンド", "C", 208, 127, "1993-08-10", 13, "コネチカット大出身、ベテラン控えC、リバウンド職人"),
    (11, "ジャスティン・エドワーズ", "SF", 201, 92, "2003-12-15", 1, "ケンタッキー大出身、2年目のフォワードでローテ要員"),
    (14, "アデム・ボナ", "C", 208, 107, "2003-03-28", 1, "UCLA出身、運動能力高い若手控えセンター"),
    (25, "ドミニック・バーロウ", "PF", 206, 102, "2003-08-25", 3, "今季ツーウェイから本契約昇格、ハッスル系フォワード"),
    (12, "トレンドン・ワトフォード", "PF", 206, 105, "2000-11-09", 4, "LSU出身、夏に加入のスケールあるベンチPF"),
    (20, "キャメロン・ペイン", "PG", 188, 84, "1994-08-08", 10, "マレーシア州立大出身、ベテランPG、シーズン途中復帰"),
    (22, "ジョニ・ブルーム", "C", 208, 109, "2002-07-19", 0, "オーバーン大出身、2巡指名ルーキー、メニスカス手術明け"),
]


# ---------- BOX SCORES ----------
# G1 = 5/4 NYK 137-98 PHI @ NYK
# G2 = 5/6 NYK 108-102 PHI @ NYK (Embiid OUT)
# G3 = 5/8 NYK 108-94 PHI @ PHI (Embiid back, Anunoby OUT)
# Format per player: [MIN, PTS, 3PM, 3PA, AST, REB] or "DNP" or "OUT"

NYK_BOX = {
    "G1_label": "G1 5/4 ○137-98 vs PHI (MSG)",
    "G2_label": "G2 5/6 ○108-102 vs PHI (MSG)",
    "G3_label": "G3 5/8 ○108-94 @PHI",
    "data": {
        "ジェイレン・ブランソン":     [[31, 35, 3, 6, 3, 1],   [41, 26, "-", "-", 6, 1],   [38, 33, "-", "-", 9, 5]],
        "マイケル・ブリッジス":       [[27, 17, 3, 5, 5, 2],   [38, 18, "-", "-", "-", "-"], [37, 23, 2, 4, 1, 3]],
        "OG・アヌノビー":             [["-", 18, "-", "-", "-", "-"], [37, 24, "-", "-", "-", 5], "OUT"],
        "ジョシュ・ハート":           [[26, 8, 1, 2, 6, 8],    [44, 5, 1, 5, 6, 7],    ["-", 12, "-", "-", "-", 11]],
        "カール=アンソニー・タウンズ": [[20, 17, 3, 5, 6, 6],   [27, 20, 1, 1, 7, 10],  [26, 8, 0, 2, 7, 12]],
        "ミッチェル・ロビンソン":     [[12, 2, 0, 0, 1, 4],    "OUT",                  [19, "-", 0, 0, "-", "-"]],
        "マイルズ・マクブライド":     [["-", "-", "-", "-", "-", "-"], [21, 4, 1, 4, 1, 1], ["-", "-", "-", "-", "-", "-"]],
        "ホセ・アルバラード":         [["-", "-", "-", "-", "-", "-"], ["-", "-", "-", "-", "-", "-"], ["-", "-", "-", "-", "-", "-"]],
        "ランドリー・シャメット":     [["-", 15, "-", "-", "-", "-"], ["-", "-", "-", "-", "-", "-"], [26, 15, 2, 3, "-", 3]],
        "ジョーダン・クラークソン":   [["-", "-", "-", "-", "-", "-"], ["-", "-", "-", "-", "-", "-"], [8, 4, "-", "-", 3, 5]],
        "アリエル・フクポルティ":     [["-", "-", "-", "-", "-", "-"], [7, 2, 0, 0, "-", 3], ["-", "-", "-", "-", "-", "-"]],
        "タイラー・コレック":         ["DNP", "DNP", "DNP"],
        "パコメ・ダディエ":           ["DNP", "DNP", "DNP"],
        "モハメド・ディアワラ":       ["DNP", "DNP", "DNP"],
        "ジェレミー・ソーチャン":     ["DNP", "DNP", "DNP"],
    }
}

PHI_BOX = {
    "G1_label": "G1 5/4 ●98-137 @NYK",
    "G2_label": "G2 5/6 ●102-108 @NYK (Embiid OUT)",
    "G3_label": "G3 5/8 ●94-108 vs NYK (Embiid復帰)",
    "data": {
        "タイリース・マクシー":       [[27, 13, 0, 3, 2, 3],   [47, 26, "-", "-", 6, 6],   [44, 17, 1, 3, 7, 2]],
        "VJ・エッジコム":             [[28, 12, 2, 5, 2, 1],   ["-", 17, "-", "-", 3, 5],  ["-", 11, "-", "-", 4, 7]],
        "ケリー・オウブレJr.":         [["-", 12, "-", "-", 1, 5], ["-", 19, "-", "-", "-", "-"], [36, 22, 2, 5, 1, 8]],
        "ポール・ジョージ":           [[26, 17, 4, 6, 3, 3],   ["-", 19, 5, 13, 4, 6],   ["-", 15, "-", "-", "-", 5]],
        "ジョエル・エンビード":       [[25, 14, 0, 2, 1, 4],   "OUT",                   [35, 18, 0, 4, 5, 6]],
        "クエンティン・グライムス":   [["-", "-", "-", "-", "-", "-"], [24, 7, 1, 2, 3, 3], [22, 6, 2, "-", 2, 2]],
        "アンドレ・ドラモンド":       [["-", "-", "-", "-", "-", "-"], [15, 6, 0, 0, "-", 8], [15, "-", 0, 0, "-", "-"]],
        "ジャスティン・エドワーズ":   [["-", "-", "-", "-", "-", "-"], [4, 0, "-", "-", "-", "-"], ["-", 4, "-", "-", "-", "-"]],
        "アデム・ボナ":               [[4, "-", 0, 0, "-", "-"], [16, 2, 0, 0, 1, 7],   ["-", 1, "-", "-", "-", "-"]],
        "ドミニック・バーロウ":       ["DNP",                  ["-", "-", "-", "-", "-", "-"], ["-", 3, "-", "-", "-", "-"]],
        "トレンドン・ワトフォード":   ["DNP",                  ["-", "-", "-", "-", "-", "-"], ["-", 1, "-", "-", "-", "-"]],
        "キャメロン・ペイン":         ["DNP",                  ["-", "-", "-", "-", "-", "-"], ["-", 1, "-", "-", "-", "-"]],
        "ジョニ・ブルーム":           ["OUT",                  "OUT",                  "OUT"],
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

    # Team band
    band_row = start_row
    ws.cell(row=band_row, column=1, value=team_name).font = BAND_FONT
    ws.merge_cells(start_row=band_row, end_row=band_row, start_column=1, end_column=COLS)
    for c in range(1, COLS + 1):
        cell = ws.cell(row=band_row, column=c)
        cell.fill = PatternFill("solid", fgColor=team_bg)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[band_row].height = 22

    # Row 1: HC + per-game labels
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

    # Row 2: sub headers
    r2 = start_row + 2
    sub = ["No.", "選手名", "Pos", "", "", "", "", ""] + ["MIN", "PTS", "3PM", "3PA", "AST", "REB"] * 3
    for i, h in enumerate(sub):
        cc = ws.cell(row=r2, column=i + 1, value=h)
        cc.font = SUB_FONT; cc.fill = SUB_FILL; cc.alignment = CENTER; cc.border = BORDER
    ws.row_dimensions[r2].height = 18

    # Data rows
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

        # Game stats
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
                # DNP / OUT - merge 6 cells
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
    ws.title = "NYK_PHI_R2G4_5_10"

    for i, w in enumerate(WIDTHS):
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True

    # Title row
    ws.cell(row=1, column=1, value="2026 NBAプレーオフ 第2回戦・第4戦  フィラデルフィア・76ers @ ニューヨーク・ニックス  2026/5/10(現地)・JST 5/11 4:30 タップオフ ／ シリーズ 3勝1敗（NYKリード）").font = Font(name="MS Gothic", size=12, bold=True)
    ws.merge_cells(start_row=1, end_row=1, start_column=1, end_column=COLS)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 24

    # NYK block (home)
    next_row = write_team_block(ws, 2, "ニューヨーク・ニックス（東3位 53勝29敗 / シリーズ3-0リード）", "マイク・ブラウン", NYK_PLAYERS, NYK_BOX, NYK_BG)

    # Spacer
    next_row += 1

    # PHI block
    write_team_block(ws, next_row, "フィラデルフィア・76ers（東7位 45勝37敗 / シリーズ0-3）", "ニック・ナーシュ", PHI_PLAYERS, PHI_BOX, PHI_BG)

    # Freeze panes
    ws.freeze_panes = "I3"

    wb.save(OUT)
    print(f"SAVED: {OUT}")


if __name__ == "__main__":
    main()
