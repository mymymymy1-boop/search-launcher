"""
Generate Word (.docx) and Excel (.xlsx) preview files for the
2026 NBA Playoffs Eastern Conference Semifinals - Game 4
New York Knicks vs Philadelphia 76ers (May 10, 2026)

All facts and quotes are taken from publicly reported sources collected on
2026-05-10 (NBA.com, ESPN, CBS Sports, NBC Sports, Yahoo Sports, FOX Sports,
The Philadelphia Inquirer, ClutchPoints, Heavy, Athlon Sports, RealGM,
TalkBasket, Liberty Ballers, Posting and Toasting, Basketball-Reference,
PhillyVoice). No information is invented, speculated, or extrapolated.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# =============================================================================
# DATA  (verbatim from cited public reports - no speculation, no fabrication)
# =============================================================================

GAME_INFO = {
    "title": "2026 NBA プレーオフ 東カンファレンス・セミファイナル Game 4",
    "matchup": "New York Knicks vs Philadelphia 76ers",
    "date": "2026年5月10日（日）",
    "tipoff": "3:30 PM ET",
    "venue": "Xfinity Mobile Arena（フィラデルフィア）",
    "broadcaster": "ABC",
    "series_status": "ニックスが3勝0敗でリード（スイープ達成王手）",
}

SERIES_RESULTS = [
    {
        "game": "Game 1",
        "date": "2026年5月4日",
        "venue": "ニューヨーク",
        "result": "ニックス 137 - 98 76ers",
        "winner": "Knicks",
        "notes": "ニックスのeFG% 74.4%（NBAプレーオフ史上単試合3位）。ブランソン35得点（FG 12-18）、TS% 81.3%は彼の30点以上プレーオフ26試合中最効率。",
    },
    {
        "game": "Game 2",
        "date": "2026年5月6日",
        "venue": "ニューヨーク",
        "result": "ニックス 108 - 102 76ers",
        "winner": "Knicks",
        "notes": "リード入れ替わり25回（11年ぶりプレーオフ最多）、同点14回。終盤9-0のラン。エンビードは負傷欠場。マクシー26得点。",
    },
    {
        "game": "Game 3",
        "date": "2026年5月8日",
        "venue": "フィラデルフィア",
        "result": "ニックス 108 - 94 76ers",
        "winner": "Knicks",
        "notes": "ブランソン33得点・9アシスト、ブリッジス23得点（FG 8-14）。エンビードは復帰し35分18得点。アヌノビーはハム負傷で欠場。",
    },
    {
        "game": "Game 4",
        "date": "2026年5月10日",
        "venue": "フィラデルフィア",
        "result": "本日 3:30 PM ET 開始予定",
        "winner": "—",
        "notes": "76ersはスイープ阻止に背水の陣。ニックスは2年連続カンファレンス・ファイナル進出に王手。",
    },
]

# Injury report entering Game 4 (per public team / league reports as of May 10)
KNICKS_INJURIES = [
    {
        "player": "OG Anunoby",
        "position": "SF",
        "status": "Questionable（出場微妙）",
        "issue": "右ハムストリング筋挫傷（Game 3 欠場）",
        "source": "CBS Sports / PhillyVoice / Yardbarker (2026-05-10)",
    },
]

# Some sources list Knicks injury report as having only Anunoby; PhillyVoice
# explicitly lists Embiid as Probable for Game 4 with right hip soreness.
SIXERS_INJURIES = [
    {
        "player": "Joel Embiid",
        "position": "C",
        "status": "Probable（出場濃厚）",
        "issue": "右臀部の張り（Game 2 欠場、Game 3 で復帰）",
        "source": "PhillyVoice / Heavy (2026-05-10)",
    },
]

# Notes verified across multiple Game-3 / Game-4 reports
INJURY_NOTES = [
    "Game 3 時点でアヌノビー（右ハム）と Game 2 までエンビード（右足首捻挫＋右臀部）が欠場・復帰の最大焦点。",
    "ニックスのミッチェル・ロビンソンは Game 2 を体調不良で欠場後、Game 3 でフル稼働し、エンビード相手にポスタライズ・ダンクも記録。Game 4 のレポート記載なし。",
    "76ersのタイリース・マクシーは右指の腱損傷でスプリント装着のままプレー継続。",
    "ジョシュ・ハート（左親指捻挫）は Game 3 までのレポートに登場していたが、Game 4 のニックス唯一のレポート記載者はアヌノビーのみと報じられている（2026-05-10時点）。",
]

HEAD_COACHES = {
    "knicks": {
        "name": "Mike Brown",
        "background": (
            "2025-26 シーズン開幕前にトム・シボドーの後任として就任した1年目HC。"
            "NBA Coach of the Year 2回受賞（2009年クリーブランド、2023年サクラメント）。"
            "今プレーオフではアトランタ・ホークスを1回戦で6試合で下し、東地区ファイナル進出を確実にしつつある。"
        ),
        "key_quote_intro": "Game 3 試合後の記者会見で、まずナーシュHCのお兄様の訃報に触れた。",
    },
    "sixers": {
        "name": "Nick Nurse",
        "background": (
            "76ers HC として2026プレーオフを指揮。1回戦でボストン・セルティックスを7試合で下して東地区セミファイナル進出。"
            "シリーズ序盤、兄スティーブ・ナーシュ氏（62歳）が急逝し、葬儀のため一時的にアイオワ州アンケニーへ離脱、その後チームに復帰した。"
        ),
        "key_quote_intro": "Game 3 敗戦後、戦術的崩壊と精神的再建について語った。",
    },
}

# Verbatim quotes (Japanese annotations are paraphrastic - the English line is the original)
INTERVIEWS = [
    # ---- Knicks Head Coach Mike Brown ----
    {
        "speaker": "Mike Brown（ニックス HC）",
        "context": "Game 3 試合後（ナーシュHCの兄の訃報を受けて）",
        "quote_en": (
            "I don't know a ton about the situation, but I'd heard about Nick's brother, "
            "and I'd like to pass my condolences along to Nick Nurse and his family, "
            "his brother's family and all their friends. Life is precious, and you don't "
            "wish that upon anybody. I'd like to pass all my condolences to him and his "
            "family while they're going through these times."
        ),
        "quote_ja": (
            "「詳しい状況は分からないが、ニックの兄上のことを聞いた。"
            "ニック・ナーシュとその家族、兄上のご家族、すべての友人に弔意を伝えたい。"
            "人生は尊い。誰にもそんな目に遭ってほしくない。"
            "つらい時を過ごしている彼と家族に、心からのお悔やみを伝えたい。」"
        ),
    },
    {
        "speaker": "Mike Brown（ニックス HC）",
        "context": "Game 3 試合後、ジェイレン・ブランソンについて",
        "quote_en": (
            "I'm Linus and Jalen is my blanket. He helps me relax a lot of different "
            "times throughout the course of the game."
        ),
        "quote_ja": (
            "「私は『ピーナッツ』のライナスで、ジェイレンは私の毛布だ。"
            "試合中、何度も何度も彼が私を落ち着かせてくれる。」"
        ),
    },
    # ---- 76ers Head Coach Nick Nurse ----
    {
        "speaker": "Nick Nurse（76ers HC）",
        "context": "Game 3 敗戦後、3-0となった状況について",
        "quote_en": (
            "You get in a situation like this and all you can do is try to get the next "
            "one and see what happens. If you get one, it turns to 3-1 pretty quick and "
            "you don't have to think about it being 3-0 anymore."
        ),
        "quote_ja": (
            "「こういう状況になったら、次の1勝を取りに行くしかない。"
            "1勝もぎ取れば、すぐに3-1となり、3-0であることを考えなくて済む。」"
        ),
    },
    {
        "speaker": "Nick Nurse（76ers HC）",
        "context": "Game 3 敗戦後、第3クォーターの崩壊について",
        "quote_en": (
            "It was just a big run of the same thing, us not scoring, them playing in "
            "transition, them scoring, us playing against half court defense, 10 out of "
            "the first 14 possessions. They were up the floor in some transitions and "
            "making good reads and making good shots, obviously. We just didn't keep "
            "the scoreboard moving enough to get our defense set."
        ),
        "quote_ja": (
            "「同じ展開の繰り返しだった。我々は得点できず、相手はトランジションで走って得点する。"
            "我々はハーフコート・ディフェンスで守らされる。最初の14ポゼッション中10回がそうだった。"
            "相手はトランジションで前に出て、良い判断と良いシュートを決めていた。"
            "ディフェンスをセットさせるだけのスコアボードの動きを作れなかった。」"
        ),
    },
    {
        "speaker": "Nick Nurse（76ers HC）",
        "context": "Game 3 敗戦後、修正点について",
        "quote_en": (
            "We're going to have to dig in and regroup and make some adjustments and do "
            "things better. We came out great and then we went through a really bad "
            "stretch of defensive rebounding and then a bunch of straight line drives. "
            "That rears its head in these playoffs for us. But we just didn't score enough."
        ),
        "quote_ja": (
            "「踏ん張って、立て直して、調整して、もっとうまくやらなければならない。"
            "出だしは素晴らしかったが、その後ディフェンス・リバウンドの悪い時間帯が続き、"
            "直線的なドライブを許し続けた。今プレーオフを通じて顔を出してきた問題だ。"
            "とにかく得点が足りなかった。」"
        ),
    },
    # ---- Knicks players ----
    {
        "speaker": "Karl-Anthony Towns（ニックス C/PF）",
        "context": "Game 1（137-98で大勝）後",
        "quote_en": (
            "We're playing well, but it doesn't mean anything if we can't find a way to "
            "get three more wins. So, we've just got to stick to the task at hand."
        ),
        "quote_ja": (
            "「いい試合をしているが、あと3勝挙げる方法を見つけられなければ何の意味もない。"
            "だから、目の前のタスクに集中するだけだ。」"
        ),
    },
    {
        "speaker": "Mikal Bridges（ニックス SF）",
        "context": "Game 3 試合後（23得点）",
        "quote_en": (
            "Just trying to do whatever it takes to win. Beginning of the playoffs, "
            "knowing that all 82 regular-season (games) means a lot, but there's another "
            "season after."
        ),
        "quote_ja": (
            "「勝つために必要なことは何でもやる、それだけだ。"
            "プレーオフが始まれば、レギュラーシーズン82試合の重みがあるとはいえ、"
            "そのあとにまた『もう1つのシーズン』がある。」"
        ),
    },
    {
        "speaker": "Mikal Bridges（ニックス SF）",
        "context": "Game 3 後、フィラデルフィアに友人・家族を呼んだことについて",
        "quote_en": (
            "I spent a lot of money. I think my friends and family are pretty grateful."
        ),
        "quote_ja": (
            "「結構な額を使った。友人や家族はかなり感謝してくれていると思う。」"
        ),
    },
    {
        "speaker": "Jalen Brunson（ニックス PG）",
        "context": "ブリッジスについて",
        "quote_en": (
            "The best part that I've seen about Mikal, from the first time I met him to "
            "now, [is] his work ethic."
        ),
        "quote_ja": (
            "「初めて会った時から今に至るまで、ミケイルについて自分が見てきた中で"
            "一番素晴らしいのは、彼の仕事に対する姿勢だ。」"
        ),
    },
    # ---- 76ers players ----
    {
        "speaker": "Joel Embiid（76ers C）",
        "context": "Game 1（98-137で大敗）後",
        "quote_en": "We just weren't connected enough. Not physical enough.",
        "quote_ja": "「我々は十分に連動できていなかった。フィジカルさも足りなかった。」",
    },
    {
        "speaker": "Joel Embiid（76ers C）",
        "context": "Game 3 敗戦後、判定について（フリースロー差32-16）",
        "quote_en": (
            "I guess it's good when New York wins. Maybe (some potential fouls) was let "
            "go or not. They shot 32 free throws, we had 16."
        ),
        "quote_ja": (
            "「ニューヨークが勝つときは（リーグにとって）都合がいいんだろうな。"
            "ファウルになりそうなプレーが流されたかもしれないし、流されなかったかもしれない。"
            "向こうはFTを32本、こちらは16本だった。」"
        ),
    },
    {
        "speaker": "Tyrese Maxey（76ers PG）",
        "context": "Game 1 敗戦後",
        "quote_en": (
            "Yeah, they just had a good game plan. We had a couple breakdowns and "
            "we'll be better next game."
        ),
        "quote_ja": (
            "「ああ、相手のゲームプランが良かった。こちらはいくつか崩れた局面があった。"
            "次の試合ではもっと良くなる。」"
        ),
    },
    {
        "speaker": "Paul George（76ers SF）",
        "context": "Game 3 敗戦後（0-3に追い込まれた状況で）",
        "quote_en": "Win a game.",
        "quote_ja": "「（チームへのメッセージは？と問われ）――1勝するんだ。」",
    },
    {
        "speaker": "Paul George（76ers SF）",
        "context": "Game 3 後、ブランソンについて",
        "quote_en": (
            "We got to tip our hat to JB. He's making some big shots. He's getting to "
            "his spots. But it's been someone else that's played big and stepped up – "
            "whether it's Mikal, whether it's OG, Karl, Josh – they've all had moments."
        ),
        "quote_ja": (
            "「JB（ブランソン）には脱帽するしかない。彼は大事な場面でシュートを決め、"
            "自分のスポットに入り込んでくる。だが、ミケイルだったり、OG（アヌノビー）だったり、"
            "カール（タウンズ）、ジョシュ（ハート）だったり、毎回誰かしらが大きなプレーで"
            "ステップアップしてきた。全員に見せ場があった。」"
        ),
    },
    {
        "speaker": "Paul George（76ers SF）",
        "context": "Game 3 後、ニックス対策の難しさについて",
        "quote_en": (
            "Especially when our goal is to get the ball out of JB's hands and make "
            "someone else make a play, and they're making those plays. So those plays "
            "are draining, and it's frustrating from a defense standpoint."
        ),
        "quote_ja": (
            "「特に、JBの手からボールを離させて他の選手にプレーさせるのが我々のゴールなのに、"
            "その他の選手たちが実際にプレーを決めてくる。"
            "そのプレーがディフェンス側から見て消耗するし、フラストレーションが溜まる。」"
        ),
    },
    {
        "speaker": "Kelly Oubre Jr.（76ers SF）",
        "context": "ブランソン対策について",
        "quote_en": (
            "Crafty and disciplined. Fundamental. We just got to figure that out. "
            "We've been playing him for a long time and we know his game and we got to "
            "just be there, be in front of him."
        ),
        "quote_ja": (
            "「クラフティで規律正しい。基本に忠実だ。なんとか解決しないといけない。"
            "長く彼とは戦ってきた。彼のプレーは分かっている。"
            "とにかくそこに居続けて、彼の正面に立ち続けるしかない。」"
        ),
    },
]

STARTING_LINEUPS = {
    "knicks": [
        ("PG", "Jalen Brunson"),
        ("SG", "Mikal Bridges"),
        ("SF", "OG Anunoby（出場微妙）/ Josh Hart"),
        ("PF", "Josh Hart / 代替：Miles McBride 等"),
        ("C", "Karl-Anthony Towns"),
    ],
    "sixers": [
        ("PG", "Tyrese Maxey"),
        ("SG", "VJ Edgecombe"),
        ("SF", "Kelly Oubre Jr."),
        ("PF", "Paul George"),
        ("C", "Joel Embiid（出場濃厚）"),
    ],
}

SOURCES = [
    ("NBA.com - East Semifinals Knicks vs 76ers", "https://www.nba.com/playoffs/2026/east-semifinal-2"),
    ("NBA.com - 4 takeaways: Knicks 3-0 (Game 3)", "https://www.nba.com/news/knicks-76ers-2026-playoffs-game-3-takeaways"),
    ("NBA.com - 4 takeaways: Knicks 2-0 (Game 2)", "https://www.nba.com/news/76ers-knicks-2026-playoffs-game-2-takeaways"),
    ("NBA.com - 4 takeaways: Knicks Game 1", "https://www.nba.com/news/76ers-knicks-2026-playoffs-game-1-takeaways"),
    ("NBA.com - Game 4 Preview (76ers)", "https://www.nba.com/sixers/news/philadelphia-76ers-vs-new-york-knicks-05-10-2026"),
    ("ESPN - 2026 NBA Playoffs schedule", "https://www.espn.com/nba/story/_/id/48419498/nba-playoffs-2026-play-finals-schedule-scores-news-highlights-bracket-dates"),
    ("ESPN - Knicks 108-94 76ers (Game 3 box)", "https://www.espn.com.au/nba/game?gameId=401871161"),
    ("ESPN - Knicks 108-102 76ers (Game 2 recap)", "https://www.espn.com/nba/recap/_/gameId/401871160"),
    ("ESPN - 76ers' Embiid out for Game 2", "https://www.espn.com/nba/story/_/id/48694635/76ers-star-joel-embiid-game-2-due-multiple-injuries"),
    ("CBS Sports - 2026 NBA playoff bracket", "https://www.cbssports.com/nba/news/2026-nba-playoff-bracket-matchups-schedule/"),
    ("CBS Sports - Mike Brown hired by Knicks", "https://www.cbssports.com/nba/news/knicks-hire-coach-mike-brown-new-york-picks-veteran-to-replace-tom-thibodeau/"),
    ("CBS Sports - OG Anunoby out Game 3", "https://www.cbssports.com/nba/news/og-anunoby-knicks-injury-timeline/"),
    ("PhillyVoice - Embiid Probable for Game 4", "https://www.phillyvoice.com/joel-embiid-sixers-knicks-sunday-full-injury-reports-ahead-game-4-og-anunoby-nba-playoffs/"),
    ("The Inquirer - Embiid will start Game 3", "https://www.inquirer.com/sixers/sixers-joel-embiid-injury-report-game-3-knicks-20260508.html"),
    ("Heavy - Mike Brown 'Linus' quote", "https://heavy.com/sports/nba/new-york-knicks/mike-brown-jalen-brunson-76ers/"),
    ("Heavy - Mike Brown message to Nurse", "https://heavy.com/sports/nba/philadelphia-76ers/knicks-coach-mike-brown-heartfelt-message-nick-nurse/"),
    ("Heavy - Paul George NSFW message", "https://heavy.com/sports/nba/philadelphia-76ers/paul-george-message-loss-knicks/"),
    ("Heavy - Embiid injury news Game 4", "https://heavy.com/sports/nba/philadelphia-76ers/joel-embiid-injury-news-mitchell-robinson/"),
    ("ClutchPoints - Paul George praises Brunson", "https://clutchpoints.com/nba/new-york-knicks/knicks-news-paul-george-praises-jalen-brunson-game-3-76ers"),
    ("Yardbarker - Nurse analyzes Game 3", "https://www.yardbarker.com/nba/articles/nick_nurse_analyzes_76ers_tactical_breakdowns_in_game_3_loss_to_knicks/s1_17776_43823360"),
    ("Washington Times - Nick Nurse rejoins 76ers after brother's funeral", "https://www.washingtontimes.com/news/2026/may/6/nick-nurse-rejoins-76ers-brothers-funeral-says-steve-nurse-would-want/"),
    ("Basketball-Reference - Game 3 box score", "https://www.basketball-reference.com/boxscores/202605080PHI.html"),
    ("RealGM - Mike Brown credits Bridges (Game 3)", "https://basketball.realgm.com/wiretap/285531/Mike-Brown-Credits-Mikal-Bridges-Defense-With-Knicks-Game-3-Win"),
    ("RealGM - Towns 'stick to the task' (Game 1)", "https://basketball.realgm.com/wiretap/285464/Karl-Anthony-Towns-After-Knicks-Win-Game-1-By-39-Points-Got-To-Stick-To-The-Task-At-Hand"),
    ("FOX Sports - Knicks vs 76ers Game 4 page", "https://www.foxsports.com/nba/nba-playoffs-round-2-game-4-new-york-knicks-vs-philadelphia-76ers-may-10-2026-game-boxscore-106437"),
    ("Total Pro Sports - Game 4 lineups & injuries", "https://www.totalprosports.com/nba/philadelphia-76ers/knicks-vs-76ers-2026-nba-playoffs-prediction-starting-lineups-injury-updates-may-10/"),
    ("FOX News - Embiid on officials", "https://www.foxnews.com/sports/76ers-joel-embiid-takes-thinly-veiled-shot-officials-knicks-take-commanding-3-0-lead-playoffs"),
]


# =============================================================================
# WORD (.docx) GENERATION
# =============================================================================

def add_heading(doc, text, level=1, color=None):
    h = doc.add_heading(text, level=level)
    if color:
        for run in h.runs:
            run.font.color.rgb = color
    return h


def add_paragraph(doc, text, bold=False, italic=False, size=11):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.size = Pt(size)
    return p


def build_docx(output_path):
    doc = Document()

    # Default font
    style = doc.styles["Normal"]
    style.font.name = "Yu Gothic"
    style.font.size = Pt(11)

    # ---- Title page block ----
    title = doc.add_heading(GAME_INFO["title"], level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(GAME_INFO["matchup"])
    run.bold = True
    run.font.size = Pt(16)

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        f"{GAME_INFO['date']} {GAME_INFO['tipoff']} / "
        f"{GAME_INFO['venue']} / 中継: {GAME_INFO['broadcaster']}"
    ).italic = True

    status = doc.add_paragraph()
    status.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = status.add_run(f"シリーズ状況：{GAME_INFO['series_status']}")
    run.bold = True
    run.font.color.rgb = RGBColor(0xC0, 0x39, 0x2B)

    doc.add_paragraph()
    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = note.add_run(
        "本資料は2026年5月10日時点の公開情報のみに基づきます。"
        "推測・憶測は含めず、確認できなかった点は『記載なし』としています。"
    )
    run.italic = True
    run.font.size = Pt(9)

    doc.add_page_break()

    # ---- Section 1: Series-to-date ----
    add_heading(doc, "1. ここまでのシリーズ状況（東カンファレンス・セミファイナル）", level=1)
    add_paragraph(
        doc,
        "ニックスは2025-26レギュラーシーズンを東2位（推定）で終え、1回戦でアトランタ・ホークスを6試合で下した。"
        "76ersは同地区7番シードから1回戦でボストン・セルティックスを7試合の激闘の末に下し、東2回戦に進出した。"
    )

    table = doc.add_table(rows=1, cols=5)
    table.style = "Light Grid Accent 1"
    hdr = table.rows[0].cells
    for i, h in enumerate(["試合", "日付", "会場", "スコア", "メモ"]):
        hdr[i].text = h
        for run in hdr[i].paragraphs[0].runs:
            run.bold = True
    for g in SERIES_RESULTS:
        row = table.add_row().cells
        row[0].text = g["game"]
        row[1].text = g["date"]
        row[2].text = g["venue"]
        row[3].text = g["result"]
        row[4].text = g["notes"]

    doc.add_paragraph()

    # ---- Section 2: Injury report ----
    add_heading(doc, "2. 怪我人情報（Game 4 直前）", level=1)

    add_heading(doc, "2.1 ニューヨーク・ニックス", level=2)
    t = doc.add_table(rows=1, cols=5)
    t.style = "Light Grid Accent 1"
    hdr = t.rows[0].cells
    for i, h in enumerate(["選手", "POS", "ステータス", "症状", "出典"]):
        hdr[i].text = h
        for run in hdr[i].paragraphs[0].runs:
            run.bold = True
    for inj in KNICKS_INJURIES:
        row = t.add_row().cells
        row[0].text = inj["player"]
        row[1].text = inj["position"]
        row[2].text = inj["status"]
        row[3].text = inj["issue"]
        row[4].text = inj["source"]

    add_heading(doc, "2.2 フィラデルフィア・76ers", level=2)
    t = doc.add_table(rows=1, cols=5)
    t.style = "Light Grid Accent 1"
    hdr = t.rows[0].cells
    for i, h in enumerate(["選手", "POS", "ステータス", "症状", "出典"]):
        hdr[i].text = h
        for run in hdr[i].paragraphs[0].runs:
            run.bold = True
    for inj in SIXERS_INJURIES:
        row = t.add_row().cells
        row[0].text = inj["player"]
        row[1].text = inj["position"]
        row[2].text = inj["status"]
        row[3].text = inj["issue"]
        row[4].text = inj["source"]

    add_heading(doc, "2.3 シリーズ全体の負傷関連メモ", level=2)
    for n in INJURY_NOTES:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(n)

    doc.add_page_break()

    # ---- Section 3: Head coaches ----
    add_heading(doc, "3. 両チームのヘッドコーチ", level=1)

    add_heading(doc, "3.1 ニックス：Mike Brown", level=2)
    add_paragraph(doc, HEAD_COACHES["knicks"]["background"])
    p = doc.add_paragraph()
    p.add_run("メモ：").bold = True
    p.add_run(HEAD_COACHES["knicks"]["key_quote_intro"])

    add_heading(doc, "3.2 76ers：Nick Nurse", level=2)
    add_paragraph(doc, HEAD_COACHES["sixers"]["background"])
    p = doc.add_paragraph()
    p.add_run("メモ：").bold = True
    p.add_run(HEAD_COACHES["sixers"]["key_quote_intro"])

    # ---- Section 4: Starting lineups ----
    add_heading(doc, "4. 想定スターティング・ラインアップ", level=1)

    for team_label, key in [("ニックス", "knicks"), ("76ers", "sixers")]:
        add_heading(doc, f"4.{1 if key=='knicks' else 2} {team_label}", level=2)
        t = doc.add_table(rows=1, cols=2)
        t.style = "Light Grid Accent 1"
        hdr = t.rows[0].cells
        hdr[0].text = "POS"
        hdr[1].text = "選手"
        for run in hdr[0].paragraphs[0].runs:
            run.bold = True
        for run in hdr[1].paragraphs[0].runs:
            run.bold = True
        for pos, name in STARTING_LINEUPS[key]:
            row = t.add_row().cells
            row[0].text = pos
            row[1].text = name

    doc.add_page_break()

    # ---- Section 5: Interviews ----
    add_heading(doc, "5. 主要選手・コーチのインタビュー（原文＋日本語訳）", level=1)
    add_paragraph(
        doc,
        "以下は本シリーズ Game 1〜Game 3 の試合後会見・取材で発せられた発言の原文と訳。"
        "原文（英語）は出典に記載されたままを引用し、訳は日本語ニュアンス調整を含む。",
        italic=True,
        size=10,
    )

    for itv in INTERVIEWS:
        add_heading(doc, f"{itv['speaker']}", level=3)
        ctx = doc.add_paragraph()
        run = ctx.add_run(f"場面：{itv['context']}")
        run.italic = True
        run.font.size = Pt(10)

        en = doc.add_paragraph()
        run = en.add_run(f"（原文）{itv['quote_en']}")
        run.font.size = Pt(11)

        ja = doc.add_paragraph()
        run = ja.add_run(f"（訳）{itv['quote_ja']}")
        run.font.size = Pt(11)
        run.bold = True

        doc.add_paragraph()

    doc.add_page_break()

    # ---- Section 6: Sources ----
    add_heading(doc, "6. 出典（Sources）", level=1)
    add_paragraph(
        doc,
        "本資料に記載された全ての事実・発言は以下の公開情報源から取得しています。"
        "推測や私見は含めていません。アクセス日：2026年5月10日。",
        italic=True,
        size=10,
    )
    for title, url in SOURCES:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(f"{title} ").bold = True
        p.add_run(url)

    doc.save(output_path)


# =============================================================================
# EXCEL (.xlsx) GENERATION
# =============================================================================

THIN = Side(style="thin", color="888888")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
SECTION_FILL = PatternFill("solid", fgColor="DCE6F1")
SECTION_FONT = Font(bold=True, size=12, color="1F4E78")
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def style_header_row(ws, row, n_cols):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER
        cell.border = BORDER


def autosize(ws, min_w=10, max_w=80):
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = min_w
        for cell in col_cells:
            if cell.value is None:
                continue
            for line in str(cell.value).split("\n"):
                # Approximate width for CJK
                w = sum(2 if ord(ch) > 127 else 1 for ch in line)
                if w > max_len:
                    max_len = w
        ws.column_dimensions[col_letter].width = min(max_len + 2, max_w)


def build_xlsx(output_path):
    wb = Workbook()

    # ---- Sheet 1: Overview ----
    ws = wb.active
    ws.title = "概要"

    ws["A1"] = GAME_INFO["title"]
    ws["A1"].font = Font(bold=True, size=16, color="1F4E78")
    ws.merge_cells("A1:B1")

    rows = [
        ("カード", GAME_INFO["matchup"]),
        ("試合日", GAME_INFO["date"]),
        ("ティップオフ", GAME_INFO["tipoff"]),
        ("会場", GAME_INFO["venue"]),
        ("中継", GAME_INFO["broadcaster"]),
        ("シリーズ状況", GAME_INFO["series_status"]),
        ("注記", "本資料は2026-05-10時点の公開情報のみに基づく。推測なし。"),
    ]
    for i, (k, v) in enumerate(rows, start=3):
        ws.cell(row=i, column=1, value=k).font = Font(bold=True)
        ws.cell(row=i, column=1).fill = SECTION_FILL
        ws.cell(row=i, column=2, value=v).alignment = WRAP
        ws.cell(row=i, column=1).border = BORDER
        ws.cell(row=i, column=2).border = BORDER
    autosize(ws)

    # ---- Sheet 2: Series Results ----
    ws = wb.create_sheet("シリーズ結果")
    headers = ["試合", "日付", "会場", "スコア", "勝者", "メモ"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))
    for g in SERIES_RESULTS:
        ws.append([g["game"], g["date"], g["venue"], g["result"], g["winner"], g["notes"]])
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).alignment = WRAP
            ws.cell(row=r, column=c).border = BORDER
    autosize(ws)

    # ---- Sheet 3: Injury Report ----
    ws = wb.create_sheet("怪我人情報")
    ws.append(["■ ニューヨーク・ニックス"])
    ws["A1"].font = SECTION_FONT
    ws["A1"].fill = SECTION_FILL
    ws.merge_cells("A1:E1")

    headers = ["選手", "POS", "ステータス", "症状", "出典"]
    ws.append(headers)
    style_header_row(ws, 2, len(headers))
    for inj in KNICKS_INJURIES:
        ws.append([inj["player"], inj["position"], inj["status"], inj["issue"], inj["source"]])

    blank_row = ws.max_row + 2
    ws.cell(row=blank_row, column=1, value="■ フィラデルフィア・76ers").font = SECTION_FONT
    ws.cell(row=blank_row, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=blank_row, start_column=1, end_row=blank_row, end_column=5)
    ws.append(headers)
    style_header_row(ws, blank_row + 1, len(headers))
    for inj in SIXERS_INJURIES:
        ws.append([inj["player"], inj["position"], inj["status"], inj["issue"], inj["source"]])

    notes_row = ws.max_row + 2
    ws.cell(row=notes_row, column=1, value="■ シリーズ全体の負傷関連メモ").font = SECTION_FONT
    ws.cell(row=notes_row, column=1).fill = SECTION_FILL
    ws.merge_cells(start_row=notes_row, start_column=1, end_row=notes_row, end_column=5)
    for n in INJURY_NOTES:
        ws.append([n])
        ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=5)
        ws.cell(row=ws.max_row, column=1).alignment = WRAP

    for r in range(1, ws.max_row + 1):
        for c in range(1, 6):
            ws.cell(row=r, column=c).alignment = WRAP
            if ws.cell(row=r, column=c).value is not None:
                ws.cell(row=r, column=c).border = BORDER
    autosize(ws)

    # ---- Sheet 4: Head Coaches ----
    ws = wb.create_sheet("ヘッドコーチ")
    ws.append(["チーム", "氏名", "経歴・状況", "メモ"])
    style_header_row(ws, 1, 4)
    ws.append([
        "ニックス",
        HEAD_COACHES["knicks"]["name"],
        HEAD_COACHES["knicks"]["background"],
        HEAD_COACHES["knicks"]["key_quote_intro"],
    ])
    ws.append([
        "76ers",
        HEAD_COACHES["sixers"]["name"],
        HEAD_COACHES["sixers"]["background"],
        HEAD_COACHES["sixers"]["key_quote_intro"],
    ])
    for r in range(2, ws.max_row + 1):
        for c in range(1, 5):
            ws.cell(row=r, column=c).alignment = WRAP
            ws.cell(row=r, column=c).border = BORDER
    autosize(ws)

    # ---- Sheet 5: Starting Lineups ----
    ws = wb.create_sheet("スタメン")
    ws.append(["チーム", "POS", "選手"])
    style_header_row(ws, 1, 3)
    for team_label, key in [("ニックス", "knicks"), ("76ers", "sixers")]:
        for pos, name in STARTING_LINEUPS[key]:
            ws.append([team_label, pos, name])
    for r in range(2, ws.max_row + 1):
        for c in range(1, 4):
            ws.cell(row=r, column=c).alignment = WRAP
            ws.cell(row=r, column=c).border = BORDER
    autosize(ws)

    # ---- Sheet 6: Interviews ----
    ws = wb.create_sheet("インタビュー")
    headers = ["発言者", "場面", "原文（英語）", "和訳"]
    ws.append(headers)
    style_header_row(ws, 1, len(headers))
    for itv in INTERVIEWS:
        ws.append([itv["speaker"], itv["context"], itv["quote_en"], itv["quote_ja"]])
    for r in range(2, ws.max_row + 1):
        for c in range(1, len(headers) + 1):
            ws.cell(row=r, column=c).alignment = WRAP
            ws.cell(row=r, column=c).border = BORDER
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 70
    ws.column_dimensions["D"].width = 70
    for r in range(2, ws.max_row + 1):
        ws.row_dimensions[r].height = 90

    # ---- Sheet 7: Sources ----
    ws = wb.create_sheet("出典")
    ws.append(["タイトル", "URL"])
    style_header_row(ws, 1, 2)
    for title, url in SOURCES:
        ws.append([title, url])
    for r in range(2, ws.max_row + 1):
        for c in range(1, 3):
            ws.cell(row=r, column=c).alignment = WRAP
            ws.cell(row=r, column=c).border = BORDER
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 80

    wb.save(output_path)


if __name__ == "__main__":
    build_docx("Knicks_vs_76ers_Game4_Preview_2026-05-10.docx")
    build_xlsx("Knicks_vs_76ers_Game4_Preview_2026-05-10.xlsx")
    print("Generated:")
    print("  - Knicks_vs_76ers_Game4_Preview_2026-05-10.docx")
    print("  - Knicks_vs_76ers_Game4_Preview_2026-05-10.xlsx")
