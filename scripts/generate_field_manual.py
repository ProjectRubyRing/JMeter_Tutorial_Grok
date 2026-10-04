# -*- coding: utf-8 -*-
"""Apache JMeter 5.6.3 の現場手順書を Meiryo UI・モノトーンの Excel にする。"""

from __future__ import annotations

import math
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins
from openpyxl.worksheet.worksheet import Worksheet

OUT_PATH = Path(__file__).resolve().parents[1] / (
    "JMeter現場手順_取得からCLI実行とExcel見込み分析.xlsx"
)

# Monotone only.
C_BLACK = "1A1A1A"
C_DARK = "2B2B2B"
C_CHAR = "3D3D3D"
C_GRAY = "5C5C5C"
C_MID = "8A8A8A"
C_SILVER = "B8B8B8"
C_LINE = "C8C8C8"
C_PALE = "E8E8E8"
C_ZEBRA = "F4F4F4"
C_BG = "FAFAFA"
C_WHITE = "FFFFFF"
C_SOFT = "F7F7F7"
C_EMPH = "EDEDED"
C_FRAME = "E8E8E8"

FONT_NAME = "Meiryo UI"
ALLOWED_FILLS = {
    C_BLACK, C_DARK, C_CHAR, C_GRAY, C_MID, C_SILVER, C_LINE,
    C_PALE, C_ZEBRA, C_BG, C_WHITE, C_SOFT, C_EMPH, C_FRAME, "000000",
}
ALLOWED_FONT_COLORS = {C_BLACK, C_DARK, C_CHAR, C_GRAY, C_MID, C_WHITE, "000000"}

THIN = Border(
    left=Side(style="thin", color=C_LINE),
    right=Side(style="thin", color=C_LINE),
    top=Side(style="thin", color=C_LINE),
    bottom=Side(style="thin", color=C_LINE),
)
FRAME_BORDER = Border(
    left=Side(style="medium", color=C_GRAY),
    right=Side(style="medium", color=C_GRAY),
    top=Side(style="medium", color=C_GRAY),
    bottom=Side(style="medium", color=C_GRAY),
)
NONE_BORDER = Border()

DOC_LEFT = "JMeter 現場手順  取得 / シナリオ / CLI / Excel見込み分析"
DOC_RIGHT = "Meiryo UI  /  モノトーン"
DOC_FOOT_L = "対象: Apache JMeter 5.6.3  /  実行Java: 17 推奨  /  参照日 2026-10-04"


def F(size=10, bold=False, color=C_BLACK, italic=False, underline=None):
    return Font(
        name=FONT_NAME, size=size, bold=bold, color=color,
        italic=italic, underline=underline,
    )


def FILL(hex_color: str) -> PatternFill:
    return PatternFill("solid", fgColor=hex_color)


def AL(h="left", v="center", wrap=True, indent=0):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap, indent=indent)


def cjk_units(text: str) -> int:
    total = 0
    for ch in str(text):
        total += 2 if ord(ch) > 0x7F else 1
    return total


def row_height(text, width_units: float, font_pt=10, min_h=20, pad=8) -> int:
    usable = max(8.0, width_units * 0.92)
    lines = 0
    for para in str(text).split("\n"):
        units = max(1, cjk_units(para))
        lines += max(1, math.ceil(units / usable))
    return max(min_h, min(220, int(lines * (font_pt * 1.35) + pad)))


def set_workbook_defaults(wb: Workbook) -> None:
    try:
        wb._fonts[0].name = FONT_NAME
        wb._fonts[0].sz = 10
        wb._fonts[0].charset = 128
    except Exception:
        pass
    wb.properties.title = "JMeter現場手順 取得からCLI実行とExcel見込み分析"
    wb.properties.subject = "Apache JMeter 5.6.3 を Windows 11 で作成し Linux で CLI 実行し Excel で報告する"
    wb.properties.creator = "JMeter field manual"
    wb.properties.description = (
        "初心者向け現場手順。フォントは Meiryo UI。色はモノトーン。"
        "画像枠は後から図を挿入する。"
    )
    wb.properties.keywords = "JMeter,JMX,JTL,負荷試験,Meiryo UI"
    wb.properties.category = "手順書"


def style_sheet(ws: Worksheet, widths: dict[str, float], tab: str, freeze="A6") -> None:
    ws.sheet_view.showGridLines = False
    ws.sheet_view.view = "normal"
    ws.sheet_view.zoomScale = 100
    ws.sheet_view.showZeros = True
    ws.sheet_format.defaultRowHeight = 18
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.page_setup.horizontalCentered = True
    ws.page_setup.pageOrder = "overThenDown"
    ws.print_options.horizontalCentered = True
    ws.print_options.gridLines = False
    ws.page_margins = PageMargins(left=0.45, right=0.45, top=0.7, bottom=0.55, header=0.28, footer=0.28)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_properties.tabColor = tab
    ws.oddHeader.left.text = f'&"{FONT_NAME},Regular"{DOC_LEFT}'
    ws.oddHeader.right.text = f'&"{FONT_NAME},Regular"{DOC_RIGHT}'
    ws.oddFooter.left.text = f'&"{FONT_NAME},Regular"{DOC_FOOT_L}'
    ws.oddFooter.center.text = f'&"{FONT_NAME},Regular"&P / &N'
    ws.oddFooter.right.text = f'&"{FONT_NAME},Regular"図は IMG 番号の枠へ後から挿入'
    if freeze:
        ws.freeze_panes = freeze
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    ws.row_dimensions[1].height = 8
    ws.print_title_rows = "1:5"
    ws.sheet_format.defaultColWidth = 12


def merged_width(ws, last_col: int) -> float:
    total = 0.0
    for c in range(1, last_col + 1):
        total += ws.column_dimensions[get_column_letter(c)].width or 12
    return total


def paint(ws, row, start, end, fill, border=THIN):
    for c in range(start, end + 1):
        cell = ws.cell(row, c)
        cell.fill = fill
        cell.border = border


def merge(ws, rng, value, font, fill, align, border=THIN, height=None):
    ws.merge_cells(rng)
    from openpyxl.utils import range_boundaries

    min_col, min_row, max_col, max_row = range_boundaries(rng)
    cell = ws.cell(min_row, min_col, value)
    cell.font = font
    cell.fill = fill
    cell.alignment = align
    cell.border = border
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            tgt = ws.cell(r, c)
            tgt.border = border
            tgt.fill = fill
            if tgt is not cell:
                tgt.font = font
                tgt.alignment = align
    if height:
        ws.row_dimensions[min_row].height = height
    return cell


def put(ws, row, col, value, font=None, fill=None, align=None, border=THIN, height=None, num_fmt=None):
    cell = ws.cell(row, col, value)
    cell.font = font or F()
    if fill is not None:
        cell.fill = fill
    cell.alignment = align or AL()
    cell.border = border
    if num_fmt:
        cell.number_format = num_fmt
    if height:
        ws.row_dimensions[row].height = height
    return cell


def link_to_sheet(cell, sheet_name: str, size=10, bold=False):
    # 関係ファイルの Target に # を置くと Excel がブックを開けないことがある。
    # 数式の HYPERLINK ならシート内だけで完結する。
    label = cell.value if isinstance(cell.value, str) and cell.value else sheet_name
    label = label.replace('"', "")
    cell.value = f'=HYPERLINK("#\'{sheet_name}\'!A1","{label}")'
    cell.font = F(size, bold, C_CHAR, underline="single")


def banner(ws, last_col: int, kicker: str, title: str, lead: str) -> int:
    merge(ws, f"A1:{get_column_letter(last_col)}1", "", F(8), FILL(C_BLACK), AL(), NONE_BORDER, 8)
    merge(
        ws, f"A2:{get_column_letter(last_col)}2", kicker,
        F(9, False, C_GRAY), FILL(C_WHITE), AL("left", "center"), NONE_BORDER, 16,
    )
    merge(
        ws, f"A3:{get_column_letter(last_col)}3", title,
        F(16, True, C_WHITE), FILL(C_BLACK), AL("left", "center"), THIN, 30,
    )
    merge(
        ws, f"A4:{get_column_letter(last_col)}4", lead,
        F(9, False, C_CHAR), FILL(C_EMPH), AL("left", "center"), THIN, 32,
    )
    merge(
        ws, f"A5:{get_column_letter(last_col)}5",
        "読み方: 上から順に行う。グレーの枠（IMG-番号）は後からスクリーンショットを入れる場所。白黒の印刷でも意味が変わる箇所はない。",
        F(8, False, C_GRAY), FILL(C_WHITE), AL("left", "center"), NONE_BORDER, 16,
    )
    return 6


def section(ws, row, last_col, text) -> int:
    merge(
        ws, f"A{row}:{get_column_letter(last_col)}{row}", text,
        F(12, True, C_WHITE), FILL(C_DARK), AL("left", "center"), THIN, 24,
    )
    return row + 1


def note(ws, row, last_col, text, fill=C_EMPH, size=9, bold=False) -> int:
    width = merged_width(ws, last_col)
    h = row_height(text, width, size, 28, 8)
    merge(
        ws, f"A{row}:{get_column_letter(last_col)}{row}", text,
        F(size, bold, C_CHAR), FILL(fill), AL("left", "center"), THIN, h,
    )
    return row + 1


def headers(ws, row, titles, fill=C_BLACK) -> int:
    ws.row_dimensions[row].height = 28
    for i, title in enumerate(titles, start=1):
        put(ws, row, i, title, F(9, True, C_WHITE), FILL(fill), AL("center", "center"), THIN)
    return row + 1


def table(ws, row, rows, widths_for_height, font_size=9, zebra=True) -> int:
    """rows: list of list. Each inner list length matches headers already written."""
    for i, data in enumerate(rows):
        fill = FILL(C_WHITE if (i % 2 == 0 or not zebra) else C_ZEBRA)
        height = 22
        for c, value in enumerate(data, start=1):
            width = widths_for_height[c - 1]
            text = "" if value is None else str(value)
            height = max(height, row_height(text, width, font_size, 22, 8))
            font = F(font_size, False, C_BLACK)
            align = AL("left", "top")
            put(ws, row, c, value, font, fill, align, THIN)
        ws.row_dimensions[row].height = min(height, 160)
        row += 1
    return row


def finish(ws, last_col: int) -> None:
    end = get_column_letter(last_col)
    ws.print_area = f"A1:{end}{ws.max_row}"
    ws.sheet_view.zoomScale = 100
    if ws.freeze_panes is None:
        ws.freeze_panes = "A6"


class Images:
    def __init__(self):
        self.rows = []
        self.n = 0

    def add(self, ws, row, last_col, what, when, framing) -> int:
        self.n += 1
        img_id = f"IMG-{self.n:02d}"
        title = f"{img_id}    画像差し込み枠    {what}"
        merge(
            ws, f"A{row}:{get_column_letter(last_col)}{row}", title,
            F(10, True, C_WHITE), FILL(C_CHAR), AL("left", "center"), THIN, 22,
        )
        body = (
            f"このグレーの枠へ、後から図を入れる。\n"
            f"入れ方: 枠のセルを選択 → 挿入 → 画像 → このデバイス → 図を選ぶ → "
            f"図を右クリック → サイズとプロパティ → プロパティ → "
            f"「セルに合わせて移動やサイズ変更をする」 → 枠の中に収める。\n"
            f"撮るタイミング: {when}\n"
            f"画角: {framing}\n"
            f"文章は図の下に隠れてもよい。どの図かは 32_画像台帳 の {img_id} を見る。"
        )
        merge(
            ws, f"A{row + 1}:{get_column_letter(last_col)}{row + 1}", body,
            F(9, False, C_CHAR), FILL(C_FRAME), AL("left", "top"), FRAME_BORDER, 148,
        )
        self.rows.append((img_id, ws.title, what, when, framing))
        return row + 2


# ---------------------------------------------------------------------------
# Shared reference text
# ---------------------------------------------------------------------------

USER_PROPERTIES = """\
# --- 現場追記。既存行は消さず、ファイルの末尾に足す。---
# 結果を CSV にし、HTML ダッシュボードの日時解釈と揃える。
jmeter.save.saveservice.output_format=csv
jmeter.save.saveservice.print_field_names=true
jmeter.save.saveservice.timestamp_format=yyyy/MM/dd HH:mm:ss.SSS
jmeter.reportgenerator.date_format=yyyy/MM/dd HH:mm:ss.SSS
jmeter.save.saveservice.connect_time=true
jmeter.save.saveservice.sent_bytes=true
jmeter.save.saveservice.thread_counts=true
jmeter.save.saveservice.latency=true
jmeter.save.saveservice.assertion_results_failure_message=true
jmeter.save.saveservice.response_data=false
jmeter.save.saveservice.samplerData=false
jmeter.save.saveservice.responseHeaders=false
jmeter.save.saveservice.requestHeaders=false
jmeter.save.saveservice.url=true
summariser.name=summary
summariser.interval=30
summariser.out=true
summariser.log=true
jmeter.reportgenerator.apdex_satisfied_threshold=500
jmeter.reportgenerator.apdex_tolerated_threshold=1500
jmeter.reportgenerator.report_title=負荷試験
"""

SETENV_BAT = """\
rem bin\\setenv.bat
rem jmeter.bat が自動で読む。jmeter.bat 本体は編集しない。
set HEAP=-Xms1g -Xmx2g -XX:MaxMetaspaceSize=256m
set JMETER_LANGUAGE=-Duser.language=ja -Duser.region=JP
"""

SETENV_SH = """\
# bin/setenv.sh
# 改行は LF。所有者は JMeter を実行するユーザ。
export HEAP="-Xms1g -Xmx2g -XX:MaxMetaspaceSize=256m"
export TZ=Asia/Tokyo
"""

RUN_SH = """\
#!/bin/bash
# 1 回の試験につき 1 回実行する。結果ディレクトリは毎回新規。
set -euo pipefail
JMETER_HOME=/opt/jmeter/apache-jmeter-5.6.3
PLAN=/opt/jmeter/plans/order-search/order-search.jmx
PROP=/opt/jmeter/plans/order-search/load.properties
STAMP=$(date +%Y%m%d_%H%M%S)
OUT=/opt/jmeter/results/order-search_${STAMP}
mkdir -p "$OUT"
cat > "$OUT/run-conditions.txt" <<EOF
plan=$PLAN
threads=50
ramp=60
duration=600
stamp=$STAMP
hostname=$(hostname)
date=$(date '+%Y-%m-%d %H:%M:%S %Z')
EOF
# html-report は作らない。-o は「まだ無いディレクトリ」が条件。
"$JMETER_HOME/bin/jmeter" \\
  -n \\
  -t "$PLAN" \\
  -l "$OUT/result.jtl" \\
  -j "$OUT/jmeter.log" \\
  -q "$PROP" \\
  -Jthreads=50 \\
  -Jramp=60 \\
  -Jduration=600 \\
  -e \\
  -o "$OUT/html-report"
echo "結果ディレクトリ=$OUT"
"""

LOAD_PROPERTIES = """\
# plans/order-search/load.properties
# このシナリオだけの上書き。スレッド数は実行コマンドの -J で渡す。
jmeter.reportgenerator.report_title=order-search
search.host=app.example.internal
search.port=443
search.protocol=https
"""


W7 = {
    "A": 14, "B": 24, "C": 38, "D": 42, "E": 32, "F": 32, "G": 16,
}
W7_LIST = [14, 24, 38, 42, 32, 32, 16]
H7 = ["手順", "どこで", "操作", "入力する値", "うまくいったときの見え方", "つまずきやすい点", "画像"]


def new_sheet(wb, title, tab, widths=None):
    ws = wb.create_sheet(title)
    style_sheet(ws, widths or W7, tab)
    return ws


def use_active(wb, title, tab, widths=None):
    ws = wb.active
    ws.title = title
    style_sheet(ws, widths or W7, tab)
    return ws


# ---------------------------------------------------------------------------
# Sheets
# ---------------------------------------------------------------------------

def sheet_cover(wb, images: Images):
    ws = use_active(wb, "00_表紙", C_BLACK, {"A": 22, "B": 28, "C": 36, "D": 36, "E": 28, "F": 24, "G": 18})
    last = 7
    merge(ws, "A1:G1", "", F(8), FILL(C_BLACK), AL(), NONE_BORDER, 10)
    merge(ws, "A2:G2", "現場手順書", F(11, False, C_GRAY), FILL(C_WHITE), AL("left", "center"), NONE_BORDER, 16)
    merge(
        ws, "A3:G3", "Apache JMeter  取得から報告まで",
        F(20, True, C_WHITE), FILL(C_BLACK), AL("left", "center"), THIN, 36,
    )
    merge(
        ws, "A4:G4",
        "Windows 11 の画面でシナリオ（jmx）を作り、Linux サーバで CLI 実行し、"
        "jtl とログを端末へ戻して Excel で結果と見込みを報告する。",
        F(11, False, C_CHAR), FILL(C_EMPH), AL("left", "center"), THIN, 32,
    )
    r = 6
    r = section(ws, r, last, "この一冊でたどる仕事")
    rows = [
        ["1", "端末で準備", "Java 17 を入れ、JMeter 5.6.3 のバイナリを公式サイトから取り、展開する。", "05 → 08", "インストーラは無い。zip を展開するだけ。", "日本語パスと空白のあるフォルダは使わない。", ""],
        ["2", "画面でシナリオ", "スレッドグループと HTTP リクエストを組み立て、1 ユーザで成功を確認して jmx を保存する。", "09 → 16", "負荷はこの段階ではかけない。", "リスナーを有効なままサーバへ渡さない。", ""],
        ["3", "サーバで実行", "jmx と CSV を Linux へ置き、jmeter -n で実行し、コンソールの summary を見る。", "17 → 21", "GUI はサーバで開かない。", "同じ jtl へ追記しない。実行ごとに新フォルダ。", ""],
        ["4", "Excel で報告", "statistics.json とエラー行を読み、ひな形へ転記し、合格と見込みを文章で残す。", "22 → 27", "終了コード 0 だけでは合格にしない。", "ダッシュボードの数値と手計算の定義を混ぜない。", ""],
    ]
    r = headers(ws, r, ["順", "段階", "やること", "シート", "この段階のゴール", "守ること", "画像"])
    r = table(ws, r, rows, [22, 28, 36, 36, 28, 24, 18])
    r += 1
    r = section(ws, r, last, "文書の置き方")
    meta = [
        ["対象製品", "Apache JMeter 5.6.3（2026-10-04 時点の安定版配布）"],
        ["実行する Java", "Java 17（LTS）を推奨。5.6.3 の配布ページは Java 8 以上。次のメジャーは Java 17 以上が必須になる、と Changes にある。"],
        ["作る場所", "Windows 11 の端末。GUI はシナリオ作成と 1 ユーザの確認だけ。"],
        ["流す場所", "Linux の JMeter 実行サーバ。コマンドは CLI（非 GUI）。"],
        ["報告", "端末の Excel。フォント Meiryo UI。色はモノトーン。26_レポートひな形 を提出用に記入する。"],
        ["認証の例", "Cognito の /oauth2/token でクライアント認証し、API Gateway へはヘッダで API キーを渡す。34 から 39。12 の /api/login は画面操作の最小例。"],
        ["例示のパス", "C:\\Tools\\apache-jmeter-5.6.3 と /opt/jmeter/apache-jmeter-5.6.3 。ホスト名は example なので現場の名前に置き換える。"],
        ["許可", "負荷をかける相手は、試験してよいと合意した環境だけにする。本番への向き先を、この手順の例のまま使わない。"],
        ["印刷", "A3 横、幅を 1 ページに収める設定済み。A4 しか無いときは用紙だけ A4 に変えると縮小される。"],
    ]
    r = headers(ws, r, ["項目", "内容", "", "", "", "", ""])
    for i, (k, v) in enumerate(meta):
        fill = FILL(C_WHITE if i % 2 == 0 else C_ZEBRA)
        h = row_height(v, 28 + 36 + 36 + 28 + 24 + 18, 10, 24, 8)
        put(ws, r, 1, k, F(10, True), fill, AL("left", "top"), THIN, h)
        merge(ws, f"B{r}:G{r}", v, F(10), fill, AL("left", "top"), THIN, h)
        r += 1
    r += 1
    r = section(ws, r, last, "最初に開くシート")
    r = note(ws, r, last, "次のシート 01_使い方 を開く。全体を一度読んでから、05 の Java 準備へ進む。途中のシートだけを開いてコマンドを打たない。", C_SOFT, 10, True)
    r = images.add(
        ws, r, last,
        "端末、実行サーバ、試験対象の三者の配置図",
        "手順書を自分の現場に合わせたあとに撮る。無ければ手書きの構成を写真でもよい。",
        "左に Windows 11 端末、中央に Linux の JMeter 実行サーバ、右に試験対象。矢印は jmx を送る方向と、jtl を戻す方向の 2 本。",
    )
    finish(ws, last)
    return ws


def sheet_howto(wb, images: Images):
    ws = new_sheet(wb, "01_使い方", C_BLACK)
    last = 7
    r = banner(
        ws, last, "最初に読む",
        "この手順書の使い方",
        "JMeter を初めて開く人が、どのシートをどの順で読むかを固定する。",
    )
    r = section(ws, r, last, "30 分で到達する最初のゴール")
    r = note(
        ws, r, last,
        "最初の日のゴールは「自分の端末で JMeter が開き、HTTP リクエストを 1 件、スレッド 1、ループ 1 で成功させ、jmx を保存できる」こと。"
        "人数を増やした負荷、サーバへの転送、Excel の報告はその次の日でよい。負荷は GUI からはかけない。",
        C_SOFT, 10,
    )
    r = headers(ws, r, H7)
    steps = [
        ["1", "このシート", "下の読み順表で、今日やる範囲に印を付ける。", "初日は 05 から 14 まで。", "自分の担当範囲が言える。", "全部を初日にやろうとしない。", ""],
        ["2", "05 から 08", "Java と JMeter を入れ、日本語表示とヒープを設定する。", "setenv.bat と user.properties。全文は 31。", "jmeter.bat で窓が開く。", "Program Files に置かない。", "IMG-02 以降"],
        ["3", "09 から 13", "画面の名前を覚え、要素の追加場所をカタログで確認し、シナリオを組む。", "スレッド数は ${__P(threads,1)} 。", "ツリーが 12 の完成例に似ている。", "英語のままの項目があっても故障ではない。", ""],
        ["4", "14 から 16", "1 ユーザで実行し、結果ツリーで成功を見る。確認観点を通す。", "失敗したら人数を増やす前に直す。", "緑のチェックと応答コード 200。", "結果ツリーを有効なまま保存して負荷に使わない。", ""],
        ["5", "17 から 21", "jmx と CSV を Linux へ送り、run.sh で CLI 実行し、summary を監視する。", "スレッドは 1、次に少数、最後に目標。", "最終行の summary = を控える。", "SSH が切れても試験が死なないよう tmux を使う。", ""],
        ["6", "22 から 27", "結果一式を端末へ戻し、Excel へ読み、ひな形を記入する。", "数値の出典を statistics.json に揃える。", "総合判定と見込みの文章がある。", "エラーが 0 件でも応答時間が基準超過なら不合格。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = section(ws, r, last, "シートの読み順")
    order = [
        ["00_表紙", "何の文書か", "最初の 1 分"],
        ["01_使い方", "読む順", "最初の 5 分"],
        ["02_全体の流れ", "ファイルがどこを通るか", "作業前"],
        ["03_用語", "jmx、jtl、スレッド、ランプアップ", "分からない言葉が出たとき"],
        ["04_成果物と役割", "誰が何を残すか", "分担があるとき"],
        ["05_Javaの準備", "Java 17 と JAVA_HOME", "初日"],
        ["06_JMeterの取得", "公式サイト、zip、ハッシュ", "初日"],
        ["07_Windowsへ導入", "展開場所、起動", "初日"],
        ["08_初期設定", "setenv.bat、user.properties", "初日"],
        ["09_画面の見方", "メニューとツリー", "初日"],
        ["10_要素の追加", "右クリックの追加", "シナリオを作る直前"],
        ["11_機能カタログ", "機能一覧と負荷試験での扱い", "要素を足すたび"],
        ["12_シナリオ作成", "クリック手順", "シナリオを作る日"],
        ["13_データと相関", "CSV とトークン引き回し", "ログインがあるとき"],
        ["34_認証の全体像", "Cognito と API Gateway を足すときの地図", "トークンと API キーがあるとき。最初に 34"],
        ["35_Cognitoトークン", "トークン取得から CLI までのクリック順", "認証つきシナリオを作る日"],
        ["36_BASE64ヘッダ", "client_id と client_secret の Authorization", "35 の前処理を書くとき"],
        ["37_トークンの保持", "access_token を業務リクエストへ渡す", "JSON Extractor を書くとき"],
        ["38_正規表現抽出", "正規表現抽出の各欄と Regexp Tester", "JSON 以外、または抽出の練習"],
        ["39_APIキーヘッダ", "X-API-ID と、標準の X-API-Key", "業務リクエストのヘッダを書くとき"],
        ["14_GUIで実行", "1 ユーザの実行", "保存の前"],
        ["15_GUIの結果確認", "ツリー、統計、集計の見方", "GUI 実行のあと"],
        ["16_確認観点と注意", "合格の前に見るリスト", "実行の前と後"],
        ["17_JMXの保存と転送", "保存と WinSCP / scp", "サーバへ渡すとき"],
        ["18_Linuxサーバ準備", "同じバージョンの導入", "初回とバージョン変更時"],
        ["19_CLIの実行", "コマンドの意味と jmx の指定", "流すたび"],
        ["20_コンソール監視", "summary の読み方と止め方", "流れているあいだ"],
        ["21_JTLとログ", "出力先と中身", "終了直後"],
        ["22_結果を端末へ", "戻すファイル", "分析の前"],
        ["23_Excelへ読込", "JSON と CSV の読み方", "分析の日"],
        ["24_分析の手順", "指標の意味と切り方", "数値を書く前"],
        ["25_判定と見込み", "合格の書き方と外挿の限界", "所見を書くとき"],
        ["26_レポートひな形", "提出用の記入用紙", "提出の日"],
        ["27_計算例", "小さい JTL で数式を確認", "Excel の操作に慣れるとき"],
        ["28_トラブル対処", "症状から対処", "止まったとき"],
        ["29_チェックリスト", "前・中・後", "試験当日"],
        ["30_コマンド早見", "コピー用", "コマンドを打つとき"],
        ["31_設定の写し", "properties と run.sh の全文", "設定を書くとき"],
        ["32_画像台帳", "差し込み枠の一覧", "図を貼るとき"],
        ["33_改定履歴", "版と参照した公式ページ", "手順を直したとき"],
    ]
    r = headers(ws, r, ["シート", "役割", "いつ開く", "", "", "", ""])
    for i, (name, role, when) in enumerate(order):
        fill = FILL(C_WHITE if i % 2 == 0 else C_ZEBRA)
        put(ws, r, 1, name, F(9, True), fill, AL("left", "center"), THIN, 20)
        link_to_sheet(ws.cell(r, 1), name, 9, True)
        merge(ws, f"B{r}:D{r}", role, F(9), fill, AL("left", "center"), THIN)
        merge(ws, f"E{r}:G{r}", when, F(9), fill, AL("left", "center"), THIN)
        r += 1
    r += 1
    r = note(
        ws, r, last,
        "シートのタブは番号順。34 から 39 は末尾にある。Cognito でトークンを取り、API Gateway に API キーを載せるシナリオは、13 の次に 34 から 39 を読み、そのあと 14 へ戻る。12 の /api/login は、その認証を足す前の画面練習。",
        C_SOFT, 9,
    )
    r = section(ws, r, last, "画面の言葉について")
    r = note(
        ws, r, last,
        "手順のメニュー名は、日本語表示にしたときの名前を先に書き、英語名を括弧に入れている。"
        "JMeter は初期設定のままだと英語で開く（JMETER_LANGUAGE の既定が en）。08 の setenv.bat を入れてから起動すると日本語になる。"
        "一部の要素名（JSON Extractor、CSV Data Set Config、HTTP(S) Test Script Recorder など）は日本語表示でも英語のまま出ることがある。英語名でメニューを探す。",
    )
    r = section(ws, r, last, "図の入れ方（全シート共通）")
    r = note(
        ws, r, last,
        "各シートの「IMG-xx  画像差し込み枠」のグレーセルが貼る場所。"
        "1) グレーのセルをクリックする。 2) 挿入 → 画像 → このデバイス。 3) 図を右クリック → サイズとプロパティ → プロパティ。"
        "4) 「セルに合わせて移動やサイズ変更をする」を選ぶ。 5) 図の角をドラッグして枠の中に収める。"
        "枠を削除しない。図を差し替えるときは、古い図を選んで Delete してから新しい図を入れる。台帳の文章は残る。",
    )
    r = images.add(
        ws, r, last,
        "Excel の挿入メニューで画像を選んでいるところ",
        "この手順書へ最初の図を入れる練習のとき。",
        "挿入タブと「このデバイス」が見える。個人のファイル名は写さない。",
    )
    finish(ws, last)


def sheet_flow(wb, images: Images):
    ws = new_sheet(wb, "02_全体の流れ", C_DARK)
    last = 7
    r = banner(ws, last, "全体像", "ファイルと作業の流れ", "何を、どこで作り、どこへ渡し、何が戻るかを先に固定する。")
    r = section(ws, r, last, "登場するもの")
    r = headers(ws, r, ["名前", "実体", "作る場所", "残す場所", "中身", "負荷の数値に使うか", "画像"])
    rows = [
        ["シナリオ", "拡張子 .jmx の XML", "Windows の GUI で保存", "端末と Linux の plans", "リクエスト順、スレッド、アサーション", "直接は使わない。実行の設計図。", ""],
        ["データ", "UTF-8 の CSV", "Excel かテキストエディタ", "jmx と同じフォルダ", "ユーザ ID など、行ごとの入力", "使わない。個人情報を置かない。", ""],
        ["結果", "拡張子 .jtl の CSV", "Linux の -l が書く", "results/実行日時/", "1 リクエスト 1 行の実測", "詳細分析とエラー抽出に使う。", ""],
        ["実行ログ", "jmeter.log", "Linux の -j が書く", "jtl と同じフォルダ", "JMeter 自身の INFO と ERROR", "性能値ではない。異常の根拠に使う。", ""],
        ["条件メモ", "run-conditions.txt", "run.sh が書く", "jtl と同じフォルダ", "スレッド数、時間、ホスト名", "レポートの試験条件に転記する。", ""],
        ["HTML レポート", "html-report/index.html と statistics.json", "Linux の -e -o が書く", "jtl と同じフォルダ", "公式ダッシュボード", "報告用の代表値は statistics.json。", ""],
        ["報告", "26 のひな形を記入したブック", "Windows の Excel", "提出先の定め", "条件、結果、判定、見込み", "これが提出物。", ""],
    ]
    r = table(ws, r, rows, W7_LIST)
    r += 1
    r = section(ws, r, last, "ディレクトリ（例。現場の名前に変える）")
    r = note(
        ws, r, last,
        "端末\n"
        "C:\\Tools\\apache-jmeter-5.6.3\\          … プログラム。シナリオは置かない。\n"
        "C:\\LoadTest\\scenarios\\order-search\\   … order-search.jmx と users.csv と load.properties。\n"
        "C:\\LoadTest\\results\\                   … サーバから戻した実行フォルダを、実行ごとに置く。\n\n"
        "Linux\n"
        "/opt/jmeter/apache-jmeter-5.6.3/         … プログラム。端末と同じ 5.6.3。\n"
        "/opt/jmeter/plans/order-search/          … jmx、csv、load.properties、run.sh。\n"
        "/opt/jmeter/results/order-search_年月日時分秒/  … result.jtl、jmeter.log、run-conditions.txt、html-report/。",
        C_SOFT, 9,
    )
    r = section(ws, r, last, "1 回の試験で手が動く順")
    r = headers(ws, r, ["順", "誰の画面", "操作", "渡すもの", "確認", "失敗したら", "画像"])
    flow_steps = [
        ["1", "端末 GUI", "シナリオを直し、1 スレッドで成功を確認する。", "なし", "結果ツリーが成功", "14 に戻る。サーバへ送らない。", ""],
        ["2", "端末 GUI", "結果ツリーなどリスナーを無効にして jmx を保存する。", "jmx", "無効はグレー表示", "有効なリスナーのまま送らない。", ""],
        ["3", "端末 → Linux", "jmx、csv、properties、run.sh を plans へ置く。", "上記一式", "サーバ上で ls してサイズを見る", "CSV の文字コードを確認する。", ""],
        ["4", "Linux", "tmux を開き、スレッド 1 で run.sh を改変して流す。", "コマンド", "summary の Err が 0", "jmx を直す。負荷を上げない。", ""],
        ["5", "Linux", "合意した人数で本番の run.sh を流し、summary を見る。", "コンソール", "Active が目標に届く", "悪化が続くなら shutdown.sh。", ""],
        ["6", "Linux → 端末", "results のその回のフォルダを丸ごと戻す。", "jtl、log、html、条件", "index.html が開く", "index.html だけ戻さない。", ""],
        ["7", "端末 Excel", "statistics.json を読み、ひな形へ転記し、見込みを書く。", "提出ブック", "判定列が埋まっている", "定義の違う数値を混ぜない。", ""],
    ]
    r = table(ws, r, flow_steps, W7_LIST)
    r += 1
    r = note(ws, r, last, "認証が Cognito のクライアント認証のときは、送るファイルに secrets.properties と run-api.sh が増える。置き場所は /opt/jmeter/plans/api-gateway/ 。中身の作り方は 34 と 35。秘密の値は、この流れ図にも書かない。")
    r = images.add(
        ws, r, last,
        "端末の scenarios フォルダと、サーバの plans フォルダを並べた図",
        "転送の前後。左右に同じファイル名が並んでいる状態。",
        "jmx、csv、load.properties、run.sh の 4 つが両方にある。jtl はサーバ側の results にだけある。",
    )
    finish(ws, last)


def sheet_terms(wb, images: Images):
    ws = new_sheet(wb, "03_用語", C_DARK)
    last = 7
    r = banner(ws, last, "言葉", "はじめてでも止まらないための用語", "操作中に出てくる言葉を、現場での意味に限って書く。")
    r = headers(ws, r, ["用語", "短い意味", "この仕事での扱い", "よくある取り違え", "関連シート", "", ""])
    terms = [
        ["スレッド", "仮想ユーザ 1 人の動作の流れ。", "スレッド数＝同時に動かす人数、と報告では書く。厳密には同時接続数と一致しないことがある。", "スレッド数を「1 秒あたりのリクエスト数」と書かない。", "12, 19", "", ""],
        ["Ramp-Up 期間", "全員が起動し終わるまでの秒数。", "100 人を 100 秒なら、約 1 秒に 1 人ずつ増える。", "全員が最初の 1 秒で一斉に始まる、ではない。", "12, 25", "", ""],
        ["ループ回数", "1 人のスレッドがシナリオを何周するか。", "時間指定の試験では「無限」にし、スケジューラの持続時間で終わらせる。", "ループ 1 のまま長時間試験したつもりにならない。", "12", "", ""],
        ["持続時間", "スケジューラが試験を続ける秒数。", "報告の「試験時間」は原則この値。ランプアップを含む。", "持続時間をランプアップより短くすると、全員そろう前に終わる。", "12, 24", "", ""],
        ["サンプラー", "1 回のリクエスト。HTTP リクエストがそれ。", "jtl の 1 行は、原則サンプラー 1 回。", "トランザクションコントローラの親行と子行を両方数えると二重になる。", "11, 21", "", ""],
        ["トランザクション", "複数リクエストを 1 つの業務操作として束ねた名前。", "報告の行はトランザクション名に揃える。", "Generate parent sample のオンオフで行の意味が変わる。", "12, 24", "", ""],
        ["タイマ", "次のリクエストの前に待つ時間。思考時間の代わり。", "一様乱数タイマを業務の待ちに使う。", "タイマはスループットを下げる。入れたことをレポートに書く。", "11", "", ""],
        ["アサーション", "応答が期待どおりかの検査。失敗するとその行はエラーになる。", "応答コードの確認に使う。応答時間の合否は Excel 側を原則にする。", "遅いだけでアサーション失敗にすると、エラー率と遅さの両方が混ざる。", "16, 25", "", ""],
        ["リスナー", "結果を画面やファイルに出す部品。", "GUI の確認用。CLI の前に無効にする。ファイル出力は -l に任せる。", "結果ツリーを有効にした負荷試験はメモリを食って落ちる。", "14, 16", "", ""],
        ["jmx", "シナリオの保存形式。中身は XML。", "テキストで開けるが、手編集は最終手段。", "jmx にパスワードを書くと、ファイルを開いた人に見える。", "17", "", ""],
        ["jtl", "サンプル結果のファイル。この手順では中身は CSV。", "-l でパスを指定する。拡張子は jtl のままでよい。", "既存ファイルがあると追記される。実行ごとに新しいパス。", "19, 21", "", ""],
        ["jmeter.log", "JMeter プログラムのログ。サンプル結果ではない。", "-j で結果フォルダへ出す。", "jtl と log を同じものだと思わない。", "21", "", ""],
        ["プロパティ", "JMeter 全体の設定値。-J や properties ファイルで渡す。", "スレッド数は ${__P(threads,1)} で読む。", "${threads} と書くと -Jthreads を読まない。", "12, 19", "", ""],
        ["変数", "スレッドごとの値。CSV の列や抽出結果。", "${user} や ${token}。他のスレッドからは見えない。", "ユーザ定義変数は全スレッド共通の定数で、CSV の代わりにならない。", "13", "", ""],
        ["elapsed", "そのリクエストの所要時間（ミリ秒）。送信開始から応答完了まで。", "レポートの応答時間はこれを使う。", "Latency や Connect と混ぜない。", "21, 24", "", ""],
        ["Latency", "送信開始から、応答の最初のバイトまで。", "サーバが考え始めた時間の目安。", "elapsed の代わりに平均待ち時間と書かない。", "24", "", ""],
        ["Connect", "接続を確立するまでの時間（ミリ秒）。", "ここだけ大きいときはネットワークや接続数を疑う。", "Connect が 0 付近でも応答が遅いことはある。", "24", "", ""],
        ["スループット", "単位時間あたりに完了したサンプル数。", "報告では「件/秒」と、分母にした時間を書く。", "スレッド数そのものをスループットと呼ばない。", "24", "", ""],
        ["パーセンタイル", "速い順に並べたとき、その位置より遅いものが残る割合の境界。", "95% 値は「95% のリクエストはこの時間以内」。平均よりこちらを基準にする。", "平均が基準内でも 95% 値が基準超過なら、その取引は不合格にできる。", "24, 27", "", ""],
        ["エラー率", "失敗サンプル数 ÷ サンプル数。", "success が false の割合。終了コードとは別。", "HTTP 200 でもアサーション失敗ならエラー。", "20, 25", "", ""],
        ["APDEX", "満足・許容・失望の 3 段で体験を 0 から 1 に要約した指数。", "HTML レポートに出る。閾値は user.properties の 500ms / 1500ms が初期値。", "APDEX が高いことと、社内の合格基準を同じものにしない。基準は別途書く。", "21, 25", "", ""],
        ["CLI / 非 GUI", "画面を出さずに実行するモード。-n で指定。", "人数をかける試験は必ずこちら。", "サーバで jmeter とだけ打つと GUI を開こうとして失敗する。", "19", "", ""],
        ["見込み", "今回の実測の内側で言える余力と、次に試す負荷。", "曲がっていない範囲だけを書く。直線で 2 倍とは書かない。", "未測定の人数を、測定済みのように書かない。", "25", "", ""],
    ]
    r = table(ws, r, terms, [14, 24, 38, 42, 32, 32, 16])
    ws.auto_filter.ref = f"A6:G{r - 1}"
    r += 1
    r = images.add(
        ws, r, last,
        "用語を画面上で指したメモ（任意）",
        "教育で画面を見せながら説明するとき。無くても手順は進む。",
        "スレッドグループのスレッド数、Ramp-Up、ループ回数の 3 欄が同時に見える。",
    )
    finish(ws, last)


def sheet_roles(wb, images: Images):
    ws = new_sheet(wb, "04_成果物と役割", C_DARK)
    last = 7
    r = banner(ws, last, "分担", "誰が何を残すか", "一人でやるときも、成果物の名前と置き場所は同じにする。")
    r = headers(ws, r, ["成果物", "作る人", "置く場所", "ファイル名の例", "次の人が見ること", "残さないもの", ""])
    rows = [
        ["試験条件の合意", "依頼元と実施者", "チケットかメール", "条件は run-conditions に再掲", "対象 URL、人数、時間、合格基準", "口頭だけ", ""],
        ["jmx", "シナリオ作成者", "端末と Linux の plans", "order-search.jmx", "バージョン 5.6.3 で開くこと", "結果ツリーが有効な版", ""],
        ["CSV", "シナリオ作成者", "jmx の隣", "users.csv", "UTF-8、1 行目が列名", "本番パスワード、個人情報", ""],
        ["run.sh", "実施者", "plans/シナリオ名", "run.sh（LF）", "結果先が実行ごとに分かれること", "CRLF のままの sh", ""],
        ["コンソール控", "実施者", "結果フォルダかチケット", "summary の最終行を貼る", "Err と Active", "途中行だけを結果と称する", ""],
        ["jtl / log / html", "CLI が生成", "results/シナリオ_日時", "result.jtl ほか", "フォルダ一式が揃っている", "index.html 単体", ""],
        ["報告", "分析者", "提出先", "26 を記入した xlsx", "条件、数値、判定、見込み、図", "出典の無い平均値だけ", ""],
    ]
    r = table(ws, r, rows, W7_LIST)
    r += 1
    r = note(ws, r, last, "同じ JMeter のバージョンを、端末と Linux の両方に置く。片方が 5.6.3、片方が 5.4 だと、jmx の要素が読めずに CLI が起動直後に落ちることがある。両方で jmeter -v の結果を報告の試験条件へ残す。")
    finish(ws, last)


def sheet_java(wb, images: Images):
    ws = new_sheet(wb, "05_Javaの準備", C_CHAR)
    last = 7
    r = banner(ws, last, "Windows 11", "Java 17 を入れてから JMeter を取る", "JMeter は Java の上で動く。Java が無いと jmeter.bat は窓を開いてすぐ閉じる。")
    r = section(ws, r, last, "何を入れるか")
    r = note(
        ws, r, last,
        "配布ページ（download_jmeter.cgi）は JMeter 5.6.3 に Java 8 以上と書いている。"
        "Getting Started も Java 8 以上で、HTTPS の録音をするなら JDK（keytool があるもの）がよい、と書いている。"
        "Changes は「5.6 系の実行は Java 8 以上、Java 17 以上を推奨。次のメジャーは Java 17 以上が必須」と書いている。"
        "この手順は新規導入を Java 17 の JDK に揃える。既に Java 17 以上が入っていれば足さなくてよい。Java 8 しか無い端末で 5.6.3 を動かすこと自体は、配布条件の範囲内。",
    )
    r = headers(ws, r, H7)
    steps = [
        ["1", "PowerShell", "今の Java を見る。", "java -version", "openjdk version \"17 で始まる行、または 21。", "「用語 'java' は認識されません」なら未導入。", ""],
        ["2", "同じ窓", "JAVA_HOME を見る。", "echo $env:JAVA_HOME", "JDK のフォルダ（bin の一つ上）が表示される。", "bin まで含めたパスを JAVA_HOME にしない。", ""],
        ["3", "ブラウザ", "未導入なら Temurin 17 の JDK（Windows x64 の msi）を入れる。社内で指定の JDK があるならそれを使う。", "インストール先を控える。例: C:\\Program Files\\Eclipse Adoptium\\jdk-17...", "java -version が 17 になる。", "JMeter 用に Java の古い 8 を消して、他の業務アプリを壊さない。不明なら情報システムへ確認。", ""],
        ["4", "設定 → システム → バージョン情報 → システムの詳細設定 → 環境変数", "ユーザ環境変数 JAVA_HOME を JDK のホームにする。Path に %JAVA_HOME%\\bin を足す。", "JAVA_HOME=C:\\Program Files\\Eclipse Adoptium\\jdk-17.0.xx-hotspot のように、実在するフォルダ。", "新しい PowerShell で java -version が 17。", "設定したあとに開いたままの古い PowerShell は、古い環境変数のまま。窓を開き直す。", ""],
        ["5", "新しい PowerShell", "もう一度確認する。", "java -version\njavac -version", "どちらも 17。javac が無いのは JRE だけのことがある。録音しないなら JRE でも JMeter は起動できる。", "32 ビットの Java を 64 ビット Windows に混ぜるとメモリを積めない。x64 を使う。", "下の枠"],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(
        ws, r, last,
        "java -version と javac -version と echo $env:JAVA_HOME の 3 行",
        "環境変数を入れて、PowerShell を開き直した直後。",
        "バージョンが 17 であることと、JAVA_HOME が bin で終わっていないことが読める大きさ。",
    )
    r = note(ws, r, last, "会社の PC でインストーラが禁止されているときは、ここで止まって申請する。ポータブルな JDK の zip を使う運用が既にあるなら、その展開先を JAVA_HOME にする。JMeter 自体の zip とは別物。")
    finish(ws, last)


def sheet_download(wb, images: Images):
    ws = new_sheet(wb, "06_JMeterの取得", C_CHAR)
    last = 7
    r = banner(ws, last, "Windows 11", "公式サイトからバイナリだけを取る", "ソースの zip は実行に不要。ビルドもしない。")
    r = headers(ws, r, H7)
    steps = [
        ["1", "ブラウザ", "公式のダウンロードページを開く。検索広告の似たサイトは開かない。", "https://jmeter.apache.org/download_jmeter.cgi", "見出しに Apache JMeter 5.6.3 とある。", "バージョンが 5.6.3 でないときは、この手順の画面名と違うことがある。番号を報告に書く。", ""],
        ["2", "Binaries の行", "apache-jmeter-5.6.3.zip を保存する。Source の zip は取らない。", "保存先の例: C:\\LoadTest\\dist\\", "約 70MB 前後の zip。サイズは配布の更新で変わるので、ハッシュで判断する。", "tgz は Linux 用。Windows の展開に使えなくはないが、この手順は zip。", "IMG-03"],
        ["3", "同じページ", "sha512 のリンクを保存する。ページの説明どおり、ハッシュは本体サイト側のファイルを使う。", "apache-jmeter-5.6.3.zip.sha512", "テキストファイルが残る。", "ミラーページに書いてあるハッシュを、目視だけで信じない。", ""],
        ["4", "PowerShell", "ハッシュを計算して、sha512 ファイルの中の文字列と比べる。大文字小文字は無視してよい。", "certutil -hashfile C:\\LoadTest\\dist\\apache-jmeter-5.6.3.zip SHA512", "計算結果と sha512 ファイルのハッシュが一致。", "一致しない zip は展開しない。取り直す。", ""],
        ["5", "任意", "gpg が使えるなら、KEYS と .asc で署名も確認する。社内手順で必須なら必須に従う。", "gpg --verify apache-jmeter-5.6.3.zip.asc apache-jmeter-5.6.3.zip", "Good signature。", "SHA512 だけでも改ざん検知の最低限にはなる、と配布ページは説明している。署名の方が強い。", ""],
        ["6", "エクスプローラ", "ダウンロードした zip を右クリック → プロパティ。下に「許可する」があればオンにして OK。", "ブロック解除", "後で jar がブロックされて起動に失敗しにくい。", "中のファイルを一つずつ解除しなくてよいよう、展開前に zip を解除する。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(
        ws, r, last,
        "公式ダウンロードページの Binaries 行",
        "zip をクリックする直前。",
        "Apache JMeter 5.6.3、Binaries、zip、sha512 が 1 画面に入る。広告や別ツールのダウンロードボタンが混ざっていないこと。",
    )
    r = note(ws, r, last, "Linux サーバには、同じバージョンの apache-jmeter-5.6.3.tgz（または同じ zip）を別に取る。端末で展開したフォルダを丸ごとコピーしてもよい。その場合もバージョン番号のフォルダ名を変えない。18 でサーバ側を扱う。")
    finish(ws, last)


def sheet_install_win(wb, images: Images):
    ws = new_sheet(wb, "07_Windowsへ導入", C_CHAR)
    last = 7
    r = banner(ws, last, "Windows 11", "zip を空白の無いフォルダへ展開して起動する", "公式の Getting Started は、パスに空白があると特にクライアント／サーバモードで問題になると書いている。")
    r = headers(ws, r, H7)
    steps = [
        ["1", "エクスプローラ", "展開先のフォルダを作る。", "C:\\Tools", "C:\\Tools が見える。", "C:\\Program Files、デスクトップ、OneDrive、日本語のフォルダには置かない。", ""],
        ["2", "zip", "apache-jmeter-5.6.3.zip を C:\\Tools に展開する。", "右クリック → すべて展開、または Expand-Archive。", "C:\\Tools\\apache-jmeter-5.6.3\\bin がある。", "bin の中にさらに apache-jmeter-5.6.3 が入れ子になっていたら、一段間違えている。", "下の枠"],
        ["3", "bin フォルダ", "次の名前があることを目で見る。", "jmeter.bat / jmeterw.cmd / jmeter-server.bat / shutdown.cmd / stoptest.cmd / user.properties / jmeter.properties", "jmeter.bat がある。", "jmeter.properties をこの時点で編集しない。追記は user.properties。", ""],
        ["4", "bin", "jmeter.bat をダブルクリックする。", "黒い窓が残り、しばらくして JMeter の窓が開く。", "左にテスト計画、上にメニュー。", "黒い窓が一瞬で閉じるときは Java が無い。05 へ戻る。窓を閉じると JMeter も終わる。", "下の枠"],
        ["5", "開いた GUI", "バージョンを見る。", "メニュー ヘルプ（Help）→ About Apache JMeter", "Version 5.6.3 と出る。", "5.6.3 でないフォルダを後から PATH に混ぜない。", ""],
        ["6", "一度閉じる", "ファイル → 終了。黒い窓も閉じる。", "保存を聞かれたら、まだシナリオが無いので保存しない。", "プロセスが残っていない。", "jmeterw.cmd はコンソールが無く、起動失敗の原因が見えない。調査中は jmeter.bat。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(ws, r, last, "展開後の bin フォルダ", "jmeter.bat を押す直前。", "アドレスバーに C:\\Tools\\apache-jmeter-5.6.3\\bin と、jmeter.bat が見える。")
    r = images.add(ws, r, last, "起動直後の JMeter 全体", "初回起動でテスト計画だけがある状態。", "メニューバー、左のツリー、右のテスト計画の説明、右上の開始ボタン（緑の三角）が入る。")
    r = section(ws, r, last, "フォルダの意味（公式の構成）")
    r = headers(ws, r, ["フォルダ", "役割", "触るか", "置くもの", "置いてはいけないもの", "", ""])
    dirs = [
        ["bin", "起動スクリプトと設定", "setenv.bat を新規作成。user.properties の末尾に追記。", "setenv.bat", "シナリオの jmx。結果の jtl。", "", ""],
        ["lib", "補助の jar", "JDBC ドライバなどが必要なときだけ。", "ドライバの jar", "自作のサンプラー。", "", ""],
        ["lib/ext", "JMeter の部品とプラグイン", "Plugins Manager を使うときだけ jar を置く。", "プラグインの jar", "ただのライブラリ jar。公式は lib と lib/ext を分けている。", "", ""],
        ["docs ほか", "説明とライセンス", "読まなくてよい。消さない。", "なし", "業務ファイル", "", ""],
    ]
    r = table(ws, r, dirs, [14, 24, 38, 42, 32, 32, 16])
    r += 1
    r = note(ws, r, last, "環境変数 CLASSPATH を足しても JMeter は見ない。公式には、java -jar で起動するため CLASSPATH は無視されると書いてある。jar の追加は lib か lib/ext に置く。")
    finish(ws, last)


def sheet_config(wb, images: Images):
    ws = new_sheet(wb, "08_初期設定", C_CHAR)
    last = 7
    r = banner(ws, last, "Windows 11", "日本語表示、メモリ、結果 CSV を先に固定する", "jmeter.bat と jmeter.properties はアップグレードで置き換わる。自分の設定は setenv.bat と user.properties に書く。")
    r = headers(ws, r, H7)
    steps = [
        ["1", "bin", "メモ帳で新規ファイルを作り、31_設定の写し の setenv.bat を貼る。", "ファイル名 setenv.bat 。文字コード UTF-8 でよい。", "bin\\setenv.bat が jmeter.bat と並ぶ。", "setenv.bat.txt という名前になっていないか、拡張子を表示して確認する。", "下の枠"],
        ["2", "bin\\user.properties", "メモ帳で開く。既にあるコメントは消さない。末尾に 31 の user.properties 追記分を貼る。", "timestamp の書式と、reportgenerator.date_format を同じ文字列にする。", "2 行が対で入っている。", "jmeter.properties の同じキーを直接変えない。公式は user.properties へコピーして変える、と書いている。", "下の枠"],
        ["3", "保存", "user.properties を上書き保存してメモ帳を閉じる。", "UTF-8。", "JMeter を起動し直すと設定を読む。起動中の変更は読まない。", "片側だけ日付書式を変えると、HTML レポートの時刻がずれる。", ""],
        ["4", "起動", "jmeter.bat を起動する。", "メニューが日本語（ファイル、編集、検索、実行、オプション、ヘルプ）。", "ファイル メニューが読める。", "英語のままなら setenv.bat が読まれていない。ファイル名と場所を見る。", ""],
        ["5", "オプション → Look and Feel", "文字が小さいときは変更できる。必須ではない。", "既定のままでよい。", "ツリーの文字が読める。", "見た目の変更を手順の成否と混同しない。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(ws, r, last, "setenv.bat をメモ帳で開いた全体", "保存する直前。", "HEAP と JMETER_LANGUAGE の 2 つの set が見える。個人のユーザ名は写さない。")
    r = images.add(ws, r, last, "user.properties 末尾の追記", "保存する直前。", "timestamp_format と date_format の 2 行が同じ書式で並んでいるところを囲む。")
    r = section(ws, r, last, "この設定で決まること")
    r = headers(ws, r, ["設定", "値", "理由", "変えてよい場面", "変えると壊れること", "", ""])
    conf = [
        ["HEAP", "-Xms1g -Xmx2g", "既定は 1g/1g。GUI で結果ツリーを開くと 1g では足りなくなることがある。", "複数の大規模シナリオを同時に開くなら 4g まで。物理メモリの半分を超えない。", "大きすぎると PC 全体がスワップする。", "", ""],
        ["言語", "ja / JP", "メニューをこの手順書と揃える。", "英語のまま運用するチームは set 行を書かない。", "言語と手順書のメニュー名がずれるだけ。動作は同じ。", "", ""],
        ["output_format", "csv", "Excel が読める。HTML レポートも CSV を読む。", "xml はレスポンス本文の調査のときだけ、別実行で。", "xml の巨大ファイルを Excel で開こうとして失敗する。", "", ""],
        ["timestamp_format", "yyyy/MM/dd HH:mm:ss.SSS", "Excel で日時として読める。", "ms（エポック）のままにするなら date_format も ms の扱いに戻す。公式の既定は ms。", "date_format と不一致だとダッシュボードの時間軸がずれる。", "", ""],
        ["response_data", "false", "応答本文を jtl に残さない。個人情報とディスクを守る。", "障害解析で 1 回だけ on_error を検討する。ルールに従う。", "本文が入った jtl をメールで配ってしまう。", "", ""],
        ["summariser.interval", "30", "コンソールに 30 秒ごとの summary が出る。", "短い試験なら 10。長くしすぎると異常に気づくのが遅れる。", "0 に近いとログばかりで試験の邪魔になる。", "", ""],
        ["APDEX 閾値", "満足 500ms、許容 1500ms", "HTML レポートの APDEX 用。公式の既定値。", "取引ごとの契約があるなら reportgenerator.apdex_per_transaction で上書き。", "APDEX を合格基準そのものだと思って報告する。", "", ""],
    ]
    r = table(ws, r, conf, [14, 24, 38, 42, 32, 32, 16])
    r += 1
    r = note(ws, r, last, "Linux 側にも同じ user.properties の追記を入れる。ヒープ用は setenv.sh（31 に全文）。言語設定はサーバの CLI には不要。タイムゾーンは setenv.sh の TZ=Asia/Tokyo で、ログの時刻を端末と揃える。")
    finish(ws, last)


def sheet_screen(wb, images: Images):
    ws = new_sheet(wb, "09_画面の見方", C_CHAR)
    last = 7
    r = banner(ws, last, "GUI", "窓のどこを触るかを先に覚える", "負荷試験のボタンも同じ窓にある。押してよいボタンと、後で無効にするものを分ける。")
    r = headers(ws, r, ["場所", "名前", "何のため", "初心者が押すとき", "まだ押さないもの", "", ""])
    rows = [
        ["一番上", "メニューバー", "保存、実行、言語、HTML レポート生成。", "ファイル、実行、オプション。", "実行 → 開始 を、シナリオが空のまま押しても害は小さい。人数を書いたあとに無確認で押さない。", "", ""],
        ["メニューの下", "ツールバー", "新規、テンプレート、開く、保存、開始、停止。", "緑の三角は 14 で、スレッド 1 のときだけ。", "赤い停止を、開始前に連打しない。", "", ""],
        ["左", "テスト計画ツリー", "要素の親子。上が先に処理される、というより親子のスコープで効く。", "右クリックで追加。", "ルートのテスト計画を削除しようとしない。消せない。", "", ""],
        ["右", "要素の詳細", "選んだ要素の入力欄。", "名前を先に変える。", "コメント欄にパスワードを書かない。jmx に残る。", "", ""],
        ["右上付近", "開始（緑）", "書いたスレッド数どおりに動かす。", "14 の手順どおり。", "本試験の人数のまま押さない。", "", ""],
        ["右上付近", "停止（赤）とシャットダウン", "停止は強い中断。シャットダウンは今のサンプルを終えて止める。", "暴走したデバッグを止める。", "結果が欲しい試験を、いきなり赤い停止で切らない。可能ならシャットダウン。", "", ""],
        ["メニュー 実行", "全て消去（Ctrl+Shift+E）", "画面上のリスナーの表示を消す。jmx は消えない。", "デバッグのやり直しの前。", "jmx の保存と混同しない。", "", ""],
        ["メニュー ファイル", "テスト計画に名前をつけて保存", "jmx を書く。", "区切りのよいところで毎回。", "bin フォルダへ保存しない。", "", ""],
    ]
    r = table(ws, r, rows, [14, 24, 38, 42, 32, 32, 16])
    r += 1
    r = images.add(
        ws, r, last,
        "画面各部の名称を書き込んだ全体図",
        "起動直後。要素を足す前でも、足したあとでもよい。",
        "メニュー、ツールバーの緑の三角、左ツリー、右の詳細、の 4 つに引き出し線があると後から読んで分かる。",
    )
    r = section(ws, r, last, "ツリーの読み方")
    r = note(
        ws, r, last,
        "子要素は親の範囲に効く。テスト計画の直下に置いた HTTP リクエスト初期値は、配下の全部の HTTP リクエストに効く。"
        "あるリクエストだけヘッダを変えたいときは、そのリクエストの子に HTTP ヘッダマネージャを置く。"
        "同じ種類が親子の両方にあると、子が上書きする項目と、両方が足される項目がある。ヘッダはマージされる。初期値のホストは、リクエスト側が空のときだけ使われる。迷ったら、ホストは初期値だけに書き、各リクエストのサーバ名は空にする。",
    )
    r = section(ws, r, last, "実行順（1 回のリクエストの中）")
    r = note(
        ws, r, last,
        "公式の要素の実行順は次のとおり。1) 設定エレメント  2) 前処理  3) タイマ  4) サンプラー  5) 後処理  6) アサーション  7) リスナー。"
        "タイマはリクエストの前に待つ。ログインの後に待ちを入れたいなら、タイマをログインの子には置かず、次の業務リクエスト側のスコープに置く。"
        "アサーションが失敗すると、そのサンプルは success=false になり、エラー率に入る。",
    )
    finish(ws, last)


def sheet_add(wb, images: Images):
    ws = new_sheet(wb, "10_要素の追加", C_CHAR)
    last = 7
    r = banner(ws, last, "GUI", "要素は右クリックの「追加」から足す", "メニューの階層は、言語が違っても同じ位置にある。")
    r = headers(ws, r, H7)
    steps = [
        ["1", "左のツリー", "親にしたい要素を右クリックする。テスト計画から始める。", "追加（Add）を開く。", "分類の一覧が出る。", "詳細画面のボタンを探して追加しようとしない。追加はツリーの右クリック。", "下の枠"],
        ["2", "追加", "分類を選ぶ。", "スレッド（ユーザー）（Threads (Users)）。", "スレッドグループが見える。", "サンプラーをテスト計画の直下に足そうとすると、スレッドが無くて実行時に動かない。必ずスレッドグループの中。", ""],
        ["3", "分類の中", "名前をクリックする。", "スレッドグループ（Thread Group）。", "左のツリーにぶら下がる。右に入力欄が出る。", "足した直後の名前は英語や汎用名のまま。すぐ名前を変える。", ""],
        ["4", "名前", "右の画面の一番上の名前を変える。", "例: TG_注文検索", "ツリーの表示が変わる。", "スラッシュや絵文字は避ける。報告の表にそのまま出る。", ""],
        ["5", "無効と有効", "デバッグ用リスナーは、負荷の前に右クリック → 無効（Disable）。", "名前がグレーになり、アイコンに印が付く。", "jmx には残るが実行されない。", "削除と無効を間違えない。無効なら後で有効に戻せる。", ""],
        ["6", "順番", "同じ親の中で、上から処理されるコントローラがある。ドラッグで並べ替える。", "ログインを検索より上にする。", "見た目の順と実行順が一致する。", "タイマを意図しない位置へ落とすと、全部のリクエストの前に待ちが入る。", ""],
        ["7", "複製", "右クリック → コピー → 貼り付け、で同じリクエストを増やせる。", "名前を必ず変える。", "統計が名前ごとに分かれる。", "同じ名前が二つあると、レポートで 1 行に合算される。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(
        ws, r, last,
        "右クリックメニューの「追加」が開いているところ",
        "テスト計画を右クリックした直後。",
        "追加、の中に スレッド（ユーザー）、設定エレメント、リスナー が見える。",
    )
    r = section(ws, r, last, "追加メニューの分類（日本語）")
    r = headers(ws, r, ["メニュー", "英語", "中に入っているもの", "最初に使うもの", "後回し", "", ""])
    menus = [
        ["スレッド（ユーザー）", "Threads (Users)", "スレッドグループ、setUp、tearDown", "スレッドグループ", "Open Model Thread Group", "", ""],
        ["設定エレメント", "Config Element", "初期値、ヘッダ、クッキー、CSV", "HTTP リクエスト初期値、CSV", "DNS Cache、鍵ストア", "", ""],
        ["サンプラー", "Sampler", "実際のリクエスト", "HTTP リクエスト", "JDBC、Java リクエスト", "", ""],
        ["ロジックコントローラ", "Logic Controller", "まとまり、分岐、繰り返し", "トランザクション、一度だけ", "もしコントローラ、乱数", "", ""],
        ["前処理", "Pre Processors", "リクエスト前の加工", "最初は無し", "JSR223", "", ""],
        ["後処理", "Post Processors", "応答から値を抜く", "JSON Extractor、正規表現抽出", "JDBC 後処理", "", ""],
        ["アサーション", "Assertions", "成否の検査", "応答アサーション", "持続時間アサーション", "", ""],
        ["タイマ", "Timer", "待ち", "一様乱数タイマ", "定数スループットタイマ", "", ""],
        ["リスナー", "Listener", "結果の見え方", "結果をツリーで表示（後で無効）", "グラフ系全般", "", ""],
        ["Non-Testエレメント", "Non-Test Elements", "録音など、試験の木に属さないもの", "録音するときだけ Recorder", "プロキシと試験リクエスト用プロキシの取り違え", "", ""],
    ]
    r = table(ws, r, menus, [14, 24, 38, 42, 32, 32, 16])
    r += 1
    r = section(ws, r, last, "標準に無い機能を足すとき")
    r = note(
        ws, r, last,
        "Apache JMeter 本体に無い要素（Ultimate Thread Group など）は、サードパーティの Plugins Manager で入る。"
        "入れる手順は 11 の末尾。原則は、標準のスレッドグループとスケジューラで足りるあいだは入れない。"
        "入れた jar は、GUI の端末と Linux の lib/ext の両方に同じファイルを置く。片方にしか無いと、CLI は jmx を読んだ瞬間に失敗する。",
    )
    finish(ws, last)


def build() -> Path:
    wb = Workbook()
    set_workbook_defaults(wb)
    images = Images()
    sheet_cover(wb, images)
    sheet_howto(wb, images)
    sheet_flow(wb, images)
    sheet_terms(wb, images)
    sheet_roles(wb, images)
    sheet_java(wb, images)
    sheet_download(wb, images)
    sheet_install_win(wb, images)
    sheet_config(wb, images)
    sheet_screen(wb, images)
    sheet_add(wb, images)
    # 後半は別モジュール。循環を避けるため、関数の中で読み込む。
    from sheets_rest import build_rest
    build_rest(wb, images)
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    wb.save(OUT_PATH)
    return OUT_PATH


def _rgb(color) -> str | None:
    if color is None:
        return None
    theme = getattr(color, "theme", None)
    if theme is not None and getattr(color, "type", None) == "theme":
        return f"theme:{theme}"
    rgb = getattr(color, "rgb", None)
    if rgb is None:
        return None
    value = str(rgb).upper()
    if value in {"00000000", "NONE"}:
        return None
    if len(value) == 8:
        value = value[2:]
    return value


def verify(path: Path) -> None:
    from openpyxl import load_workbook

    wb = load_workbook(path)
    font_errors = fill_errors = font_color_errors = wrap_missing = 0
    checked = 0
    required = [
        "00_表紙", "01_使い方", "12_シナリオ作成", "19_CLIの実行",
        "21_JTLとログ", "23_Excelへ読込", "26_レポートひな形", "27_計算例",
        "31_設定の写し", "32_画像台帳",
        "34_認証の全体像", "35_Cognitoトークン", "36_BASE64ヘッダ",
        "37_トークンの保持", "38_正規表現抽出", "39_APIキーヘッダ",
    ]
    missing = [name for name in required if name not in wb.sheetnames]
    if missing:
        raise SystemExit(f"missing sheets: {missing}")
    img_ids = []
    for ws in wb.worksheets:
        if ws.freeze_panes is None:
            raise SystemExit(f"freeze missing: {ws.title}")
        if ws.page_setup.orientation != "landscape":
            raise SystemExit(f"orientation: {ws.title}")
        if ws.page_setup.fitToWidth != 1:
            raise SystemExit(f"fitToWidth: {ws.title}")
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                checked += 1
                text = str(cell.value)
                if text.startswith("IMG-") and "画像差し込み枠" in text:
                    img_ids.append(text.split()[0])
                name = cell.font.name if cell.font else None
                if name != FONT_NAME:
                    font_errors += 1
                    if font_errors <= 8:
                        print(f"FONT {name!r} {ws.title}!{cell.coordinate}")
                color = _rgb(cell.font.color) if cell.font and cell.font.color else None
                if color and not str(color).startswith("theme:") and color not in ALLOWED_FONT_COLORS:
                    font_color_errors += 1
                    if font_color_errors <= 8:
                        print(f"FONTCOLOR {color} {ws.title}!{cell.coordinate}")
                if cell.alignment is None or cell.alignment.wrap_text is not True:
                    wrap_missing += 1
                    if wrap_missing <= 8:
                        print(f"WRAP {ws.title}!{cell.coordinate}")
                fill = cell.fill
                if fill is not None and fill.fill_type not in (None, "none"):
                    rgb = _rgb(fill.fgColor)
                    if rgb and not str(rgb).startswith("theme:") and rgb not in ALLOWED_FILLS:
                        fill_errors += 1
                        if fill_errors <= 8:
                            print(f"FILL {rgb} {ws.title}!{cell.coordinate}")
    if len(img_ids) != len(set(img_ids)):
        raise SystemExit(f"duplicate image ids: {img_ids}")
    if len(img_ids) < 20:
        raise SystemExit(f"too few image slots: {len(img_ids)}")
    blob = "\n".join(
        str(c.value)
        for ws in wb.worksheets
        for row in ws.iter_rows()
        for c in row
        if isinstance(c.value, str)
    )
    for token in [
        "-n", "-t", "-l", "-j", "-e", "-o", "statistics.json", "${__P(threads,1)}", "shutdown.sh",
        "oauth2/token", "client_credentials", "X-API-ID", "X-API-Key",
        "basic_credentials", "${__P(access_token)}", "java.util.Base64",
        "TOKEN_NOT_FOUND", "secrets.properties", "$1$",
    ]:
        if token not in blob:
            raise SystemExit(f"missing token: {token}")
    print(
        f"sheets={len(wb.sheetnames)} cells={checked} images={len(img_ids)} "
        f"font_errors={font_errors} font_color_errors={font_color_errors} "
        f"fill_errors={fill_errors} wrap_missing={wrap_missing}"
    )
    if font_errors or font_color_errors or fill_errors or wrap_missing:
        raise SystemExit(1)


if __name__ == "__main__":
    out = build()
    print(out)
    verify(out)
