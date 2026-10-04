# -*- coding: utf-8 -*-
"""Cognito のトークン取得、BASE64、API キーヘッダ、正規表現抽出。"""

from __future__ import annotations

from openpyxl.utils import get_column_letter

from generate_field_manual import (
    AL, C_CHAR, C_DARK, C_EMPH, C_GRAY, C_SOFT, C_WHITE,
    FILL, F, H7, W7_LIST,
    banner, finish, headers, merge, new_sheet, note, put, section, table,
)

# JMeter 5.6.3 の関数一覧に __base64Encode は無い。
# java.util.Base64 は改行を入れない。Groovy の encodeBase64() は実装により
# 76 文字で改行することがあり、Authorization が壊れる。

GROOVY_BASIC = """\
// 名前: PP_basicを作る
// Language: groovy
// Cache compiled script if available: オン
// 値そのものはログに出さない。長さだけ出す。
String id = props.get("client_id")
String secret = props.get("client_secret")
if (id == null || secret == null || id.trim().isEmpty() || secret.trim().isEmpty()) {
    throw new IllegalStateException("client_id または client_secret が無い。jmeter.bat を -q secrets.properties 付きで起動し直す。")
}
String raw = id + ":" + secret
String b64 = java.util.Base64.getEncoder().encodeToString(raw.getBytes("UTF-8"))
vars.put("basic_credentials", b64)
log.info("basic_credentials length=" + b64.length())
"""

GROOVY_STORE = """\
// 名前: PP_トークンを全体へ
// Language: groovy
// この要素は JSON Extractor より下に置く。上にあると、まだ変数が無い。
String token = vars.get("access_token")
if (token == null || "TOKEN_NOT_FOUND".equals(token) || token.trim().isEmpty()) {
    throw new IllegalStateException("access_token が取れていない。応答と JSON Path を見る。")
}
props.put("access_token", token)
String expires = vars.get("expires_in")
log.info("access_token length=" + token.length() + " expires_in=" + expires)
"""

SECRETS_PROPERTIES = """\
# secrets.properties
# 秘密。git、メール、Excel、スクリーンショット、run-conditions.txt に入れない。
# 値は引用符で囲まない。引用符も値の一部になる。
# バックスラッシュが秘密に含まれるときだけ、\\\\ と二重にする。
client_id=置き換える
client_secret=置き換える
api_key=置き換える
"""

LOAD_API_PROPERTIES = """\
# plans/api-gateway/load.properties
# 秘密は書かない。secrets.properties とキーを重ねない。
cognito.host=example.auth.ap-northeast-1.amazoncognito.com
api.host=abc123def.execute-api.ap-northeast-1.amazonaws.com
api.stage=prod
scope=resource-server/read
"""

RUN_API_SH = """\
#!/bin/bash
# Cognito でトークンを 1 回取り、API Gateway を叩く。秘密の値は書かない。
set -euo pipefail
JMETER_HOME=/opt/jmeter/apache-jmeter-5.6.3
PLAN=/opt/jmeter/plans/api-gateway/api-gateway.jmx
PROP=/opt/jmeter/plans/api-gateway/load.properties
SECRET=/opt/jmeter/plans/api-gateway/secrets.properties
STAMP=$(date +%Y%m%d_%H%M%S)
OUT=/opt/jmeter/results/api-gateway_${STAMP}
mkdir -p "$OUT"
test -f "$PLAN"
test -f "$PROP"
test -f "$SECRET"
cat > "$OUT/run-conditions.txt" <<EOF
plan=$PLAN
threads=50
ramp=60
duration=600
stamp=$STAMP
hostname=$(hostname)
date=$(date '+%Y-%m-%d %H:%M:%S %Z')
secret_file=secrets.properties
api_header=X-API-ID
cognito=setUp で 1 回
EOF
"$JMETER_HOME/bin/jmeter" \\
  -n \\
  -t "$PLAN" \\
  -l "$OUT/result.jtl" \\
  -j "$OUT/jmeter.log" \\
  -q "$PROP" \\
  -q "$SECRET" \\
  -Jthreads=50 \\
  -Jramp=60 \\
  -Jduration=600 \\
  -e \\
  -o "$OUT/html-report"
echo "結果ディレクトリ=$OUT"
"""

TREE = """\
テスト計画 業務API
  setUp スレッドグループ SU_トークン          … 人数は 1 の直書き。-Jthreads では増やさない
    HTTPリクエスト HTTP_POST_cognito_token   … ホストは Cognito。API Gateway の初期値は継承しない
      JSR223 PreProcessor PP_basicを作る     … client_id:client_secret を Base64 にし、変数へ入れる
      HTTPヘッダマネージャ HM_トークン         … Content-Type と Authorization: Basic ${basic_credentials}
      JSON Extractor EX_access_token          … $.access_token と $.expires_in
      JSR223 PostProcessor PP_トークンを全体へ … 変数をプロパティ access_token にコピーする
      応答アサーション                          … 応答コードが 200 と等しい
  スレッドグループ TG_業務                     … 人数は ${__P(threads,1)}
    HTTPリクエスト初期値                        … API Gateway のホスト。このグループの中だけ
    トランザクションコントローラ TX_業務
      HTTPリクエスト HTTP_GET_orders
        HTTPヘッダマネージャ HM_API            … Bearer ${__P(access_token)} と X-API-ID
        応答アサーション                        … 応答コードが 200 と等しい
      一様乱数タイマ
  結果をツリーで表示                            … デバッグ中だけ有効。CLI の前に無効
"""


def _code(ws, row, last, text) -> int:
    # 1 行にまとめると行高の上限で全文が見えなくなる。16 行ずつ分ける。
    lines = text.rstrip("\n").split("\n")
    end = get_column_letter(last)
    for i in range(0, len(lines), 16):
        part = "\n".join(lines[i:i + 16])
        height = min(400, max(48, part.count("\n") * 15 + 24))
        merge(ws, f"A{row}:{end}{row}", part, F(9), FILL(C_WHITE), AL("left", "top"), height=height)
        row += 1
    return row


def _map(wb, images):
    ws = new_sheet(wb, "34_認証の全体像", C_DARK)
    last = 7
    r = banner(
        ws, last, "認証",
        "Cognito でトークンを取り、API Gateway へ渡す",
        "12 の /api/login は画面の練習。現場のクライアント認証は、この 34 から 39 の木に組み替える。",
    )
    r = section(ws, r, last, "知りたいことと、開くシート")
    r = headers(ws, r, ["知りたいこと", "開くシート", "このシートで分かること", "終了の目安", "", "", ""])
    rows = [
        ["JMeter の取得、インストール、初期設定", "05 から 08", "Java 17、zip、展開、setenv、user.properties。", "jmeter.bat で窓が開く。", "", "", ""],
        ["要素の追加と、各機能の意味", "10 と 11", "右クリックの追加。負荷試験で無効にするもの。", "カタログで名前が引ける。", "", "", ""],
        ["画面でシナリオを作る", "09、12、そして 35", "12 は練習。Cognito の完成手順は 35。", "左の木がこのシートの木と一致する。", "", "", ""],
        ["画面で 1 人だけ実行する", "14 と 35 の実行節", "緑の三角。プロパティは -q 付きで GUI を起動して渡す。", "トークン取得が緑、業務 API が緑。", "", "", ""],
        ["結果の見方、確認観点、注意", "15、16、20、21、35", "GUI の数字は提出値にしない。秘密は図に撮らない。", "観点の表を上から見て落ちが無い。", "", "", ""],
        ["Cognito のトークンエンドポイント", "35", "POST /oauth2/token。ホストは amazoncognito.com 側。", "応答 200 で access_token がある。", "", "", ""],
        ["client_id と client_secret の Base64", "36", "Authorization: Basic。java.util.Base64。", "公開例の文字列と PowerShell の結果が一致する。", "", "", ""],
        ["アクセストークンの保持", "37", "JSON Extractor のあと、プロパティへコピーする。", "業務側が ${__P(access_token)} を使う。", "", "", ""],
        ["正規表現抽出", "38", "各欄、Regexp Tester、トークン用の式。", "Match 1 の group 1 がトークン本体になる。", "", "", ""],
        ["API Gateway へ API キーをヘッダで渡す", "39", "合意が X-API-ID ならその名前。利用プランの標準は X-API-Key。", "業務リクエストにだけ付く。Cognito には付かない。", "", "", ""],
        ["実行サーバへの転送と配置", "17、18、35 の転送節", "jmx、load.properties、secrets.properties、run-api.sh。", "サーバで 4 ファイルのサイズが 0 でない。", "", "", ""],
        ["CLI、jmx、jtl、ログの指定", "19、21、30、31、35", "-n -t -l -j -q -q -e -o。html-report は先に作らない。", "結果ディレクトリのパスが表示される。", "", "", ""],
        ["端末で Excel に読む", "22 と 23", "フォルダごと戻す。正本は statistics.json。", "業務の label の行がシートに出る。", "", "", ""],
        ["Excel で結果と見込みを書く", "24、25、26、27、このシートの末尾", "トークン取得の 1 件を性能表に混ぜない。", "26 の総合判定と見込みの文章がある。", "", "", ""],
    ]
    r = table(ws, r, rows, [28, 24, 42, 36, 16, 14, 14])
    r += 1
    r = section(ws, r, last, "リクエストは 2 本。ホストが違う")
    r = note(
        ws, r, last,
        "1 本目は Cognito のトークンエンドポイント。マシン同士のクライアント認証（client_credentials）で、"
        "access_token を 1 回だけもらう。2 本目はそのトークンと API キーを付けて、API Gateway の業務 API を、合意した人数で叩く。"
        "ホストが違うので、HTTPリクエスト初期値をテスト計画の直下に置くと、トークン取得まで API Gateway へ飛んでしまう。"
        "初期値は TG_業務 の中に置く。トークン取得のリクエストには、プロトコル、サーバ、ポートを自分で書く。",
        C_SOFT, 10,
    )
    r = headers(ws, r, ["", "1 本目 トークン", "2 本目 業務 API", "置いてよいもの", "置いてはいけないもの", "", ""])
    pair = [
        ["ホストの例", "example.auth.ap-northeast-1.amazoncognito.com", "abc123def.execute-api.ap-northeast-1.amazonaws.com", "プロパティ cognito.host と api.host", "jmx への本番ホストの直書き。例の example のままの実行", "", ""],
        ["パス", "/oauth2/token", "/${__P(api.stage,prod)}/orders", "ステージ名はプロパティ", "トークン取得にステージのパスを付ける", "", ""],
        ["メソッドと本体", "POST。パラメータ表に grant_type と scope", "合意したメソッド。例は GET", "Content-Type はリクエストごとのヘッダ", "テスト計画直下の共通 Content-Type。JSON と form が混ざる", "", ""],
        ["認証", "Authorization: Basic と Base64", "Authorization: Bearer とアクセストークン", "Basic は setUp の子だけ", "業務側に Basic。トークン側に Bearer", "", ""],
        ["API キー", "付けない", "ヘッダ X-API-ID（合意が標準なら X-API-Key）", "値は ${__P(api_key)}", "キーの直書き。トークン取得への付与", "", ""],
        ["回数", "試験全体で 1 回（setUp、スレッド 1、ループ 1）", "スレッド数 × 持続時間のあいだ繰り返し", "setUp の人数は 1 の直書き", "setUp に ${__P(threads,1)}。50 人で Cognito を 50 回叩く", "", ""],
    ]
    r = table(ws, r, pair, [16, 36, 40, 32, 36, 14, 14])
    r += 1
    r = section(ws, r, last, "完成した木")
    r = _code(ws, r, last, TREE)
    r = images.add(
        ws, r, last,
        "Cognito と API Gateway を分けた左のツリー",
        "35 の手順が終わって保存する直前。",
        "setUp と TG_業務の親子が、折りたたまない状態で読める。ヘッダの値、トークン、API キーは写さない。",
    )
    r = images.add(
        ws, r, last,
        "端末、JMeter 実行サーバ、Cognito、API Gateway の配置",
        "自分の現場のホスト名に置き換えたあと。",
        "矢印は 3 本。jmx を実行サーバへ送る、実行サーバが Cognito と API Gateway を叩く、jtl を端末へ戻す。秘密の値は図に書かない。",
    )
    r = section(ws, r, last, "秘密の置き場所")
    r = headers(ws, r, ["置き場所", "入れてよいもの", "入れてはいけないもの", "理由", "", "", ""])
    secret_rows = [
        ["secrets.properties", "client_id、client_secret、api_key", "スレッド数、ホストの説明文", "実行のたびに -q で渡す。ファイル権限は 600。", "", "", ""],
        ["load.properties", "ホスト、ステージ、scope", "client_secret、api_key", "シナリオの説明用と秘密を分ける。キー名を secrets と重ねない。", "", "", ""],
        ["jmx", "${__P(client_id)} などの式", "秘密の実値、完成した Base64", "jmx は人の手に渡る。式なら実値はファイルに残らない。", "", "", ""],
        ["jmeter.log", "basic_credentials と access_token の文字数", "トークン文字列、API キー、client_secret", "結果フォルダごと端末へ戻すため、ログに秘密があると配布範囲が広がる。", "", "", ""],
        ["run-conditions.txt と 26", "ヘッダ名、setUp で 1 回、という事実", "ヘッダの値", "提出物を見る人は、キーそのものを必要としない。", "", "", ""],
        ["スクリーンショット", "欄の名前、式、応答コード 200", "Authorization の値、Response data のトークン", "図は 32 の枠へ後から入る。値が写った図は提出しない。", "", "", ""],
    ]
    r = table(ws, r, secret_rows, [24, 36, 36, 42, 14, 14, 14])
    r += 1
    r = section(ws, r, last, "Excel の報告で、トークン取得を性能に混ぜない")
    r = note(
        ws, r, last,
        "statistics.json には HTTP_POST_cognito_token と、業務の label（例: HTTP_GET_orders）が出る。"
        "トークン取得は setUp の 1 件なので、平均も 95% 値も性能の合格判定に使わない。"
        "26 の取引表に貼るのは業務の label だけ。トークン取得は、条件欄に次の 3 つを文章で書く。"
        "setUp で 1 回、応答コード 200、応答の expires_in が試験の持続秒より長い。"
        "401 はトークン側（期限、Bearer、scope）の候補、403 は API キー側（ヘッダ名、値、利用プラン、ステージ）の候補として、エラー内訳に件数を分ける。"
        "見込みに書いてよいのは、業務 API を人数で測った範囲だけ。Cognito の余力は、この手順では測っていない。",
        C_SOFT, 10,
    )
    finish(ws, last)


def _cognito(wb, images):
    ws = new_sheet(wb, "35_Cognitoトークン", C_CHAR)
    last = 7
    r = banner(
        ws, last, "手順",
        "トークン取得から業務リクエストまで、クリック順に作る",
        "保存先は C:\\LoadTest\\scenarios\\api-gateway\\api-gateway.jmx。12 の order-search.jmx は上書きしない。",
    )
    r = note(
        ws, r, last,
        "前提は 05 から 10 が終わっていること。公式のトークンエンドポイントは HTTPS の POST /oauth2/token だけ。"
        "クライアント認証の応答は JSON で、access_token、token_type（Bearer）、expires_in（秒）が入る。"
        "id_token と refresh_token は、この grant_type では返らない。Base64 の各欄は 36、変数とプロパティの違いは 37、"
        "正規表現で同じ値を抜く練習は 38、API キーのヘッダ名は 39。",
        C_SOFT, 9,
    )
    r = section(ws, r, last, "A. フォルダとプロパティ")
    r = headers(ws, r, H7)
    steps = [
        ["1", "エクスプローラ", "シナリオ用のフォルダを作る。", "C:\\LoadTest\\scenarios\\api-gateway", "フォルダがある。", "OneDrive やデスクトップには作らない。", ""],
        ["2", "メモ帳", "load.properties を UTF-8 で保存する。中身はこのシートの後半の写し。ホストと scope を現場の値にする。", "ファイル名 load.properties", "jmx を置くフォルダと同じ。", "secrets の 3 つのキーをこちらへ書かない。", ""],
        ["3", "メモ帳", "secrets.properties を別に作る。client_id、client_secret、api_key を現場の値にする。", "UTF-8（BOM 無し）。値に引用符を付けない。値は ASCII。", "3 行のキーがある。", "BOM 付きだと 1 行目のキー名が読めない。このファイルをメールで送らない。", ""],
        ["4", "PowerShell", "公開例で、Base64 の操作だけ先に確認する。実秘密はこのコマンドに貼らない。", "下の PowerShell。結果は 36 の期待値と一致する。", "ZGpjOTh1 で始まる 1 行。", "結果が違ったら、文字の前後に空白や改行が入っている。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = _code(
        ws, r, last,
        "$raw = \"djc98u3jiedmi283eu928:abcdef01234567890\"\n"
        "$bytes = [Text.Encoding]::UTF8.GetBytes($raw)\n"
        "[Convert]::ToBase64String($bytes)\n"
        "# 期待値（AWS の公開例）\n"
        "# ZGpjOTh1M2ppZWRtaTI4M2V1OTI4OmFiY2RlZjAxMjM0NTY3ODkw",
    )
    r = section(ws, r, last, "B. トークン取得（setUp）")
    r = headers(ws, r, H7)
    b = [
        ["5", "jmeter.bat", "プロパティを読ませてから GUI を開く。先に普通のダブルクリックで開くと、秘密が空のままになる。", "下の起動コマンド。", "JMeter の窓が開く。", "開いたあとに properties を保存し直しても、起動中の JMeter は読み直さない。閉じて、同じコマンドで開き直す。", ""],
        ["6", "ファイル", "名前を付けて保存。", "C:\\LoadTest\\scenarios\\api-gateway\\api-gateway.jmx", "jmx ができる。", "bin フォルダへ保存しない。", ""],
        ["7", "テスト計画を右クリック", "追加 → スレッド（ユーザー） → setUp スレッドグループ。メニューが英語なら setUp Thread Group。", "名前 SU_トークン。スレッド数 1。Ramp-Up 1。ループ回数 1。スケジューラはオフ。", "人数の欄が 1 という数字。", "${__P(threads,1)} をここに入れると、本番の人数の回数だけ Cognito を叩く。", "下の枠"],
        ["8", "SU_トークン", "エラー時の動作。", "停止（Stop Test）。", "トークンが取れなければ、業務スレッドが始まらない。", "継続のままだと、トークン無しの業務リクエストが大量に出る。", ""],
        ["9", "SU_トークンを右クリック", "追加 → サンプラー → HTTPリクエスト。", "名前 HTTP_POST_cognito_token。", "右に入力欄が出る。", "サーバ名を空にすると、別の初期値が効いたときに API Gateway へ飛ぶ。", ""],
        ["10", "HTTP リクエスト", "プロトコル、サーバ、ポート、メソッド、パス、Content encoding。", "https\n${__P(cognito.host)}\n443\nPOST\n/oauth2/token\nUTF-8", "パスが /oauth2/token。", "cognito-idp.リージョン.amazonaws.com のユーザープール URL は、この OAuth のトークンエンドポイントではない。ドメインは xxx.auth.リージョン.amazoncognito.com か、合意したカスタムドメイン。", "下の枠"],
        ["11", "パラメータ表", "Body Data は空のまま、Parameters に 3 行。", "grant_type = client_credentials\nclient_id = ${__P(client_id)}\nscope = ${__P(scope)}", "Use multipart/form-data はオフ。", "client_secret をこの表に入れない。秘密は Authorization ヘッダ側。scope を合意していないクライアントなら、scope の行だけ作らない。自分で %2F を書いた値を入れない。スラッシュは生の文字のまま。", ""],
        ["12", "リクエストを右クリック", "追加 → 前処理 → JSR223 PreProcessor。", "名前 PP_basicを作る。Language は groovy。", "スクリプト欄に 36 の全文が入っている。このシートの写しと同じ。", "Language が javascript のままだと、Java 17 では動かない環境がある。", "下の枠"],
        ["13", "リクエストを右クリック", "追加 → 設定エレメント → HTTPヘッダマネージャ。", "名前 HM_トークン。下の 2 行。", "このリクエストの子になっている。", "テスト計画の直下に置くと、業務リクエストの Content-Type と衝突する。子は親の同名ヘッダを置き換えるが、置き場所を分ける。", ""],
        ["14", "HM_トークン", "ヘッダを 2 つ入れる。", "Content-Type = application/x-www-form-urlencoded\nAuthorization = Basic ${basic_credentials}", "Basic のうしろに半角空白が 1 つある。", "Basic の前に空白。値の全体を引用符で囲む。Bearer をここに書く。", ""],
        ["15", "リクエストを右クリック", "追加 → 後処理 → JSON Extractor。37 の 4 欄。", "変数 access_token と expires_in。Match No. は 1;1。初期値 TOKEN_NOT_FOUND;EXPIRES_NOT_FOUND。", "Match No. が 1。", "Match No. を空にすると 0（ランダム）になる。トークンが毎回変わる抽取になる。", ""],
        ["16", "JSON Extractor の下", "追加 → 後処理 → JSR223 PostProcessor。JSON Extractor より下。", "名前 PP_トークンを全体へ。Language は groovy。37 の全文。", "ツリー上で JSON Extractor の次にある。", "上に置くと vars にまだ access_token が無く、毎回失敗する。", ""],
        ["17", "リクエストを右クリック", "追加 → アサーション → 応答アサーション。", "Field to Test は応答コード。一致は等しい。Patterns to Test は 200。", "200 以外は失敗。", "「含む」のまま 200 と書くと、別のコードにも一致し得る。", ""],
    ]
    r = table(ws, r, b, W7_LIST)
    r += 1
    r = images.add(
        ws, r, last,
        "setUp スレッドグループの人数欄",
        "手順 7 のあと。",
        "スレッド数 1、Ramp-Up 1、ループ 1、スケジューラがオフ。${__P(threads)} が無いことが読める。",
    )
    r = images.add(
        ws, r, last,
        "Cognito の HTTP リクエスト",
        "手順 11 のあと。",
        "メソッド POST、パス /oauth2/token、サーバ欄が ${__P(cognito.host)}。パラメータ 3 行。Body Data は空。client_secret の実値は写さない。",
    )
    r = section(ws, r, last, "C. 業務リクエスト（API Gateway）")
    r = headers(ws, r, H7)
    c = [
        ["18", "テスト計画を右クリック", "追加 → スレッド（ユーザー） → スレッドグループ。", "名前 TG_業務。人数 ${__P(threads,1)}、Ramp-Up ${__P(ramp,1)}、ループは無限、スケジューラの持続 ${__P(duration,60)}。エラー時は継続。", "12 の手順 5 から 9 と同じ式。", "ここに 50 と直書きすると、GUI の開始が 50 人になる。", ""],
        ["19", "TG_業務を右クリック", "追加 → 設定エレメント → HTTPリクエスト初期値。", "プロトコル https、サーバ ${__P(api.host)}、ポート 443。", "このグループの中だけに効く。", "テスト計画の直下に移すと、setUp の空欄までこのホストになる。setUp 側は空欄にしない手順にしてあるが、置き場所も守る。", ""],
        ["20", "TG_業務を右クリック", "追加 → ロジックコントローラ → トランザクションコントローラ。", "名前 TX_業務。Generate parent sample はオフ。", "子の HTTP が jtl に出る。", "親サンプルをオンにすると、件数が二重になりやすい。", ""],
        ["21", "TX_業務を右クリック", "追加 → サンプラー → HTTPリクエスト。", "名前 HTTP_GET_orders。メソッド GET。パス /${__P(api.stage,prod)}/orders。サーバ名は空。", "ホストは初期値から入る。", "パスの先頭のスラッシュを欠かすと、ステージが連結されない。", "下の枠"],
        ["22", "HTTP_GET_orders を右クリック", "追加 → 設定エレメント → HTTPヘッダマネージャ。", "名前 HM_API。39 の 2 行。", "このリクエストの子。", "SU_トークンのヘッダマネージャへ行を足すと、Cognito に API キーが付く。", ""],
        ["23", "HM_API", "ヘッダ。名前は合意メモのとおり。", "Authorization = Bearer ${__P(access_token)}\nX-API-ID = ${__P(api_key)}", "Bearer のうしろに半角空白。", "利用プランの API キー（API key source が HEADER）が契約なら、名前だけ X-API-Key にする。値の式は変えない。${access_token} と書くと、業務スレッドからは見えない。", ""],
        ["24", "HTTP_GET_orders を右クリック", "追加 → アサーション → 応答アサーション。", "応答コードが 200 と等しい。", "業務の失敗がエラー率に入る。", "本文の文言を「含む」で検査すると、文言変更で全件失敗する。", ""],
        ["25", "TX_業務を右クリック", "追加 → タイマ → 一様乱数タイマ。", "定数遅延 1000、ランダム遅延 2000。", "業務の前に 1 から 3 秒待つ。", "setUp の中に入れると、トークン取得の前に待つ。報告には思考時間ありと書く。", ""],
        ["26", "テスト計画を右クリック", "追加 → リスナー → 結果をツリーで表示。", "Filename は空。", "setUp も業務も、この 1 つのツリーに出る。", "Filename にパスを書くと、そこへ結果を書き始める。CLI の -l と二重になる。", ""],
        ["27", "保存", "Ctrl+S。", "api-gateway.jmx", "タイトルバーに * が無い。", "リスナーを有効なまま、このファイルを負荷用の正にするのは 16 で止める。先に 1 人の確認を行う。", ""],
    ]
    r = table(ws, r, c, W7_LIST)
    r += 1
    r = images.add(
        ws, r, last,
        "業務リクエストのパスと、子のヘッダマネージャ",
        "手順 23 のあと。実行前。",
        "パスに ${__P(api.stage,prod)} が見える。ヘッダ名 Authorization と X-API-ID、値は式のまま。実キーは写さない。",
    )
    r = section(ws, r, last, "D. 画面で 1 人だけ実行する")
    r = headers(ws, r, H7)
    d = [
        ["28", "起動の確認", "今開いている JMeter が、手順 5 の -q 付きで起動したものであること。", "違う開き方をしていたら、保存して閉じ、手順 5 で開き直して jmx を開く。", "secrets を変えたあとも、開き直している。", "ダブルクリック起動の窓で開始しても、プロパティは空。", ""],
        ["29", "人数", "TG_業務が式のままなら、GUI では 1 人、約 60 秒。", "${__P(threads,1)} と ${__P(duration,60)}。", "setUp は 1 のまま。", "-J は、この GUI 起動コマンドには付けない。", ""],
        ["30", "消去", "前回の表示を消す。", "実行 → 全て消去（Ctrl+Shift+E）。", "ツリーが空。jmx は残る。", "全て消去を、シナリオの削除だと思わない。", ""],
        ["31", "開始", "ツールバーの緑の三角。", "実行 → 開始。検証（Validate）では、この確認を終えない。", "右上の件数が増える。", "開始（タイマを無視）は、思考時間の確認では使わない。", "下の枠"],
        ["32", "トークン行", "結果ツリーで HTTP_POST_cognito_token を開く。", "緑。Sampler result の応答コード 200。", "Load time が 1 件分だけ見える。", "赤なら Response data の error の語を読む。invalid_client は 36。invalid_scope は scope。人数は増やさない。", ""],
        ["33", "送ったヘッダ", "同じサンプルの Request タブ。", "Authorization が Basic で始まり、${basic_credentials} という文字が残っていない。Content-Type が application/x-www-form-urlencoded。", "X-API-ID はこのリクエストに無い。", "値をチケットへ貼らない。図にも入れない。確認したらタブを戻す。", ""],
        ["34", "応答の形", "Response data をその場で見る。", "access_token、token_type が Bearer、expires_in が秒の数。", "expires_in をメモする。試験の持続秒は、この数より小さくする。", "画面をスクリーンショットしない。見終わったら別のタブへ移る。", ""],
        ["35", "業務行", "HTTP_GET_orders の Request タブ。", "URL のホストが API Gateway。Authorization が Bearer で始まり、${__P(access_token)} が残っていない。X-API-ID の値に ${__P(api_key)} が残っていない。", "応答コード 200、緑。", "401 は 37。403 は 39。ホストが Cognito のままなら、手順 10 と 19 の置き場所に戻る。", ""],
        ["36", "ログ", "結果フォルダではまだ無い。GUI の黒い窓、または bin の jmeter.log の末尾。", "basic_credentials length= と access_token length= と expires_in= がある。", "length が 0 ではない。TOKEN の文字数が 16 のままなら、初期値 TOKEN_NOT_FOUND を送っている。", "ログに eyJ で始まる文字列を出すスクリプトへ書き換えない。", ""],
        ["37", "停止と保存", "成功を見たらシャットダウンで止めてよい。直したら保存。", "Ctrl+S。", "jmx が更新される。", "ここでリスナーを無効にしてから、転送用の保存にする。16 の手順。", ""],
    ]
    r = table(ws, r, d, W7_LIST)
    r += 1
    r = images.add(
        ws, r, last,
        "結果ツリーでトークン取得が緑の画面",
        "手順 32。Sampler result タブ。",
        "サンプル名と応答コード 200。Request と Response data は開かない。トークンが写り込まない構図。",
    )
    r = section(ws, r, last, "E. 確認観点と、転送の前の注意")
    r = headers(ws, r, ["観点", "見方", "通った状態", "落ちている状態", "戻る手順", "", ""])
    checks = [
        ["トークンのホスト", "Request の URL", "amazoncognito.com か、合意したカスタムドメイン。", "execute-api のホストへ POST /oauth2/token している。", "10 と 19", "", ""],
        ["業務のホスト", "業務の Request の URL", "API Gateway のホスト。パスにステージがある。", "Cognito のホストへ業務パスを送っている。", "21", "", ""],
        ["Basic", "トークンの Request ヘッダ", "Basic のあとが置換済み。", "${basic_credentials} が残る。または Basic が無い。", "12 と 14、36", "", ""],
        ["Bearer", "業務の Request ヘッダ", "Bearer のあとが置換済み。", "${__P(access_token)} または ${access_token} が残る。", "16 と 23、37", "", ""],
        ["API キーの名前", "業務の Request ヘッダ", "合意した名前が 1 つある。", "名前が違う。またはトークン側にも付いている。", "39", "", ""],
        ["期限", "メモした expires_in と、本番の持続秒", "持続秒の方が小さい。例は 600 と 3600。", "持続の方が長い。後半が 401 になる。", "34 と run-api.sh の -Jduration", "", ""],
        ["リスナー", "ツリーの色", "転送前はグレー（無効）。Filename は空。", "有効なまま、または Filename が入っている。", "16", "", ""],
        ["秘密の混入", "jmx をメモ帳で開いて検索", "client_secret の実値も api_key の実値もヒットしない。", "実値が XML に見える。", "手順 14 と 23 を式に戻して保存し直す。", "", ""],
    ]
    r = table(ws, r, checks, [16, 24, 36, 40, 32, 16, 14])
    r += 1
    r = section(ws, r, last, "F. 実行サーバへ置くファイル")
    r = note(
        ws, r, last,
        "サーバのディレクトリは /opt/jmeter/plans/api-gateway/ 。JMeter 本体の導入は 18 のまま。バージョンは端末と同じ 5.6.3。"
        "置くファイルは 4 つ。api-gateway.jmx、load.properties、secrets.properties、run-api.sh。"
        "run-api.sh の改行は LF。secrets.properties を置いたら chmod 600 を、実行ユーザで付ける。"
        "結果の jtl は送らない。ワイルドカードで results を一緒に送らない。",
        C_SOFT, 9,
    )
    r = headers(ws, r, H7)
    f = [
        ["38", "端末", "リスナーを無効にして保存した 4 ファイルが、同じフォルダにある。", "dir C:\\LoadTest\\scenarios\\api-gateway", "jmx、load.properties、secrets.properties、run-api.sh。", "jmeter.log を同梱しない。", ""],
        ["39", "サーバ", "ディレクトリを作る。", "mkdir -p /opt/jmeter/plans/api-gateway", "loadtest が書き込める。", "root だけが書けるディレクトリにしない。", ""],
        ["40", "端末 PowerShell", "4 ファイルを送る。", "下の scp。", "プロンプトが戻る。", "secrets だけ送り忘れて、CLI が即失敗する。", "下の枠"],
        ["41", "サーバ", "名前、サイズ、改行、権限。", "cd /opt/jmeter/plans/api-gateway\nls -l\nfile run-api.sh\nchmod 600 secrets.properties", "jmx が 0 バイトではない。run-api.sh に CRLF が無い。secrets が 600。", "CRLF なら sed -i 's/\\r$//' run-api.sh のあと、もう一度 file。", ""],
    ]
    r = table(ws, r, f, W7_LIST)
    r += 1
    r = images.add(
        ws, r, last,
        "サーバの plans/api-gateway の ls -l",
        "chmod のあと。",
        "4 ファイルの名前とサイズ。secrets.properties の権限が 600。中身は開かない。",
    )
    r = section(ws, r, last, "G. CLI で jmx、jtl、ログを指定する")
    r = note(
        ws, r, last,
        "監視の 1 行の読み方と止め方は 20。ここでは、このシナリオのコマンドが何を指定しているかだけを固定する。"
        "run-api.sh が同じ指定を行う。手で打つときは、結果ディレクトリを先に作り、html-report は作らない。",
        C_SOFT, 9,
    )
    r = headers(ws, r, ["オプション", "このシナリオでの値", "指定しているもの", "間違うと", "", "", ""])
    opts = [
        ["-n", "-n", "CLI。画面を出さない。", "ディスプレイの無いサーバで GUI を開こうとして落ちる。", "", "", ""],
        ["-t", "/opt/jmeter/plans/api-gateway/api-gateway.jmx", "シナリオ。", "order-search.jmx を指定すると、別の試験を流す。", "", "", ""],
        ["-l", "/opt/jmeter/results/api-gateway_日時/result.jtl", "サンプル結果。1 行が 1 リクエスト。", "既存ファイルだと追記される。日時入りの新パスにする。", "", "", ""],
        ["-j", "同じフォルダ/jmeter.log", "JMeter 自身のログ。length の行はここ。", "付け忘れると、bin やカレントの jmeter.log を上書きする。", "", "", ""],
        ["-q", "load.properties と、もう一度 -q で secrets.properties", "ホストと、秘密。JMeter 5.6.3 は -q の重複を許す。後のファイルが同じキーを上書きする。", "パスが無いと警告だけ出して続行する。client_id が空になり invalid_client になる。jmeter.log の Can't find additional property file を見る。", "", "", ""],
        ["-J", "-Jthreads=50 -Jramp=60 -Jduration=600", "業務スレッドの式。setUp の 1 は変わらない。", "duration を expires_in 以上にすると、後半が 401 になる。", "", "", ""],
        ["-e -o", "同じフォルダ/html-report", "終了後の HTML と statistics.json。", "html-report を先に mkdir すると、空でないフォルダとして失敗する。", "", "", ""],
    ]
    r = table(ws, r, opts, [12, 42, 36, 42, 16, 14, 14])
    r += 1
    r = headers(ws, r, H7)
    g = [
        ["42", "tmux", "19 と同じ。セッションの中で流す。", "tmux new -s loadtest", "切断してもプロセスが残る。", "SSH の窓を閉じると、試験が終わる。", ""],
        ["43", "1 人", "run-api.sh をコピーし、threads=1、ramp=1、duration=60 にしたものを先に流す。", "bash run-api-1user.sh", "summary の Active が 1。Err が 0。", "最初から 50 で流さない。", ""],
        ["44", "本番", "合意した値の run-api.sh。", "bash /opt/jmeter/plans/api-gateway/run-api.sh", "最後に結果ディレクトリのパスが出る。", "そのパスをチケットに書く。次の実行と混ぜない。", "下の枠"],
        ["45", "回収と Excel", "22 のとおりフォルダごと端末へ戻す。23 で statistics.json を読む。", "26 に貼る行は業務の label だけ。トークン行の扱いは 34 の末尾。", "総合判定と見込みの文章がある。", "HTTP_POST_cognito_token の平均を、業務の性能として 26 に入れる。", ""],
    ]
    r = table(ws, r, g, W7_LIST)
    r += 1
    r = images.add(
        ws, r, last,
        "run-api.sh 実行後の summary",
        "1 人の実行で、summary が 1 行出たところ。",
        "Active と Err が読める。コマンドラインに秘密の値が展開されていない。",
    )
    r = section(ws, r, last, "写し（35 からコピーする）")
    r = note(ws, r, last, "GUI の起動。秘密の値はコマンドラインに直接書かない。-q のファイルに置く。", C_SOFT, 9)
    r = _code(
        ws, r, last,
        "cd C:\\Tools\\apache-jmeter-5.6.3\\bin\n"
        ".\\jmeter.bat -q C:\\LoadTest\\scenarios\\api-gateway\\load.properties -q C:\\LoadTest\\scenarios\\api-gateway\\secrets.properties",
    )
    r = note(ws, r, last, "load.properties", C_SOFT, 9)
    r = _code(ws, r, last, LOAD_API_PROPERTIES)
    r = note(ws, r, last, "secrets.properties の形。右の値は現場の値に置き換えてから保存する。", C_SOFT, 9)
    r = _code(ws, r, last, SECRETS_PROPERTIES)
    r = note(ws, r, last, "PP_basicを作る のスクリプト。", C_SOFT, 9)
    r = _code(ws, r, last, GROOVY_BASIC)
    r = note(ws, r, last, "PP_トークンを全体へ のスクリプト。", C_SOFT, 9)
    r = _code(ws, r, last, GROOVY_STORE)
    r = note(ws, r, last, "run-api.sh。改行は LF。threads、ramp、duration は、ファイルの中の -J と run-conditions.txt を同じ値にする。", C_SOFT, 9)
    r = _code(ws, r, last, RUN_API_SH)
    r = note(
        ws, r, last,
        "転送の例。\n"
        "scp C:\\LoadTest\\scenarios\\api-gateway\\api-gateway.jmx "
        "C:\\LoadTest\\scenarios\\api-gateway\\load.properties "
        "C:\\LoadTest\\scenarios\\api-gateway\\secrets.properties "
        "C:\\LoadTest\\scenarios\\api-gateway\\run-api.sh "
        "loadtest@jmeter-runner.example.internal:/opt/jmeter/plans/api-gateway/",
        C_WHITE, 9,
    )
    finish(ws, last)


def _base64(wb, images):
    ws = new_sheet(wb, "36_BASE64ヘッダ", C_CHAR)
    last = 7
    r = banner(
        ws, last, "ヘッダ",
        "client_id と client_secret を Base64 にして Authorization に入れる",
        "Cognito の client_secret_basic。文字列は client_id、コロン、client_secret。エンコード結果の前に Basic と半角空白を付ける。",
    )
    r = section(ws, r, last, "何をエンコードするか")
    r = note(
        ws, r, last,
        "エンコードする文字列は、client_id と client_secret を半角コロン 1 つでつなげたもの。コロンの前後に空白は入れない。"
        "エンコードするのはこの文字列だけで、Basic という語はエンコードしない。ヘッダの値は、Basic、半角空白、エンコード結果、の順。"
        "Cognito の公開例は、client_id が djc98u3jiedmi283eu928、client_secret が abcdef01234567890。"
        "この 2 つをコロンでつなげた文字列の Base64 が ZGpjOTh1M2ppZWRtaTI4M2V1OTI4OmFiY2RlZjAxMjM0NTY3ODkw。"
        "手順の確認はこの公開例で行い、実値の確認結果は残さない。",
        C_SOFT, 9,
    )
    r = headers(ws, r, ["段階", "文字", "例（公開例）", "入る場所", "誤り", "", ""])
    rows = [
        ["材料", "client_id:client_secret", "djc98u3jiedmi283eu928:abcdef01234567890", "スクリプトの中だけ。jmx には残さない。", "コロンが全角。前後に空白。", "", ""],
        ["エンコード結果", "標準の Base64。+ と / と末尾の = を使う。", "ZGpjOTh1M2ppZWRtaTI4M2V1OTI4OmFiY2RlZjAxMjM0NTY3ODkw", "変数 basic_credentials。", "URL 用の Base64（+ を - にする方式）。改行が入った結果。", "", ""],
        ["ヘッダ", "Authorization", "Basic のあと半角空白、そのあと変数。", "HM_トークンの値の欄。Basic ${basic_credentials}", "Bearer と書く。引用符で全体を囲む。", "", ""],
    ]
    r = table(ws, r, rows, [14, 28, 48, 36, 32, 14, 14])
    r += 1
    r = section(ws, r, last, "JMeter での作り方")
    r = note(
        ws, r, last,
        "JMeter 5.6.3 の関数一覧に、Base64 用の関数は無い。__groovy で 1 行に書く方法はあるが、引数のコンマを \\, と書く必要があり、秘密のコンマで壊れやすい。"
        "この手順は JSR223 PreProcessor に Groovy を書く。エンコードは java.util.Base64.getEncoder()。このエンコーダは改行を入れない。"
        "Groovy の byte[].encodeBase64() は使わない。長い秘密で 76 文字ごとに改行が入り、ヘッダが 2 行になって Cognito が拒否する。",
        C_SOFT, 9,
    )
    r = _code(ws, r, last, GROOVY_BASIC)
    r = headers(ws, r, ["行の意味", "コード", "結果", "失敗の見え方", "直し方", "", ""])
    lines = [
        ["秘密を読む", "props.get(\"client_id\")", "起動時の -q か -J の値。", "null でサンプラーが失敗し、ログに「client_id または client_secret が無い」。", "jmeter.bat を閉じて、-q 付きで開き直す。", "", ""],
        ["連結", "id + \":\" + secret", "コロン 1 つの文字列。画面には出さない。", "invalid_client。", "プロパティの値の前後に空白や引用符が無いか、ファイルを見る。", "", ""],
        ["エンコード", "java.util.Base64.getEncoder()", "改行無しの標準 Base64。", "ヘッダの途中に改行があり、400 や invalid_client。", "encodeBase64() に書き換えていないか見る。", "", ""],
        ["変数へ", "vars.put(\"basic_credentials\", b64)", "この setUp スレッドの中だけで見える変数。", "ヘッダに ${basic_credentials} が残る。", "前処理がこの HTTP の子であること。Language が groovy であること。", "", ""],
        ["ログ", "log.info で length だけ", "jmeter.log に文字数。", "ログに eyJ や秘密の原文がある。", "log.info の引数を、トークンや b64 そのものにしない。", "", ""],
    ]
    r = table(ws, r, lines, [14, 32, 32, 40, 36, 14, 14])
    r += 1
    r = images.add(
        ws, r, last,
        "JSR223 PreProcessor の Language とスクリプト",
        "保存前。公開例ではなく、現場のスクリプト全文が見える状態。",
        "Language が groovy。スクリプトに client_secret の実値が無い。log.info が length だけ。",
    )
    r = images.add(
        ws, r, last,
        "HM_トークンの 2 行",
        "ヘッダを入れた直後。実行前の画面。",
        "Content-Type が application/x-www-form-urlencoded。Authorization の値が Basic ${basic_credentials}。実値は無い。",
    )
    r = section(ws, r, last, "リクエスト本体との組み合わせ")
    r = note(
        ws, r, last,
        "ヘッダで Basic を送るとき、本体のパラメータは grant_type=client_credentials、client_id=${__P(client_id)}、必要なときだけ scope。"
        "client_secret は本体に書かない。公式は、ヘッダに秘密を載せた場合、本体の client_id は省略できる、とも書いている。"
        "この手順は公式のクライアント認証の例に合わせ、本体にも client_id を置く。ヘッダ側の client_id と、本体の client_id は、同じプロパティを使う。"
        "Content-Type は application/x-www-form-urlencoded。application/json のまま送ると、Cognito はパラメータを読めず、invalid_request か invalid_client になる。",
        C_SOFT, 9,
    )
    r = section(ws, r, last, "失敗したときの応答")
    r = headers(ws, r, ["応答の error", "HTTP", "まず疑うこと", "見る場所", "次の動作", "", ""])
    err = [
        ["invalid_client", "400 または 401", "Base64 の材料が違う。client_id と secret の取り違え。空白。改行。アプリクライアントに秘密が無い。", "Request の Authorization が Basic で始まること。値はログに残さず、その場で長さだけ見る。", "secrets.properties を直し、JMeter を起動し直す。負荷は続けない。", "", ""],
        ["invalid_request", "400", "grant_type が無い。Content-Type が form ではない。Body Data とパラメータを同時に使っている。", "パラメータ表と、Body Data が空であること。", "35 の手順 11 と 14。", "", ""],
        ["invalid_scope", "400", "scope がそのアプリクライアントに許可されていない。リソースサーバの識別子が違う。", "load.properties の scope。", "許可された scope に直す。%2F を二重にエンコードしていないか見る。", "", ""],
        ["unauthorized_client", "400", "そのアプリクライアントで client_credentials が有効ではない。", "ユーザープール側の設定。JMeter の式では直らない。", "設定の担当に確認する。JMeter の人数は増やさない。", "", ""],
    ]
    r = table(ws, r, err, [18, 16, 42, 40, 32, 14, 14])
    r += 1
    r = note(
        ws, r, last,
        "参照: Amazon Cognito の token endpoint。client_secret_basic の説明と、クライアント認証のリクエスト例。"
        " https://docs.aws.amazon.com/cognito/latest/developerguide/token-endpoint.html ",
        C_EMPH, 9,
    )
    finish(ws, last)


def _token(wb, images):
    ws = new_sheet(wb, "37_トークンの保持", C_CHAR)
    last = 7
    r = banner(
        ws, last, "保持",
        "access_token を、業務スレッドから見える場所へ置く",
        "JSON Extractor の変数は、そのスレッドの中だけ。setUp で取った値を、業務スレッドへ渡すにはプロパティにコピーする。",
    )
    r = section(ws, r, last, "変数とプロパティ")
    r = headers(ws, r, ["種類", "作り方", "参照の書き方", "見える範囲", "この手順での使い方", "", ""])
    rows = [
        ["変数", "JSON Extractor の vars、または vars.put", "${access_token}", "そのスレッドだけ。他のスレッドと setUp の外からは空。", "setUp の中で、抽出結果を受け取る。", "", ""],
        ["プロパティ", "props.put、または -J と -q", "${__P(access_token)}", "試験の中の全部のスレッド。", "業務リクエストの Bearer。", "", ""],
        ["取り違え", "業務側に ${access_token} と書く", "文字が置換されず、そのまま送られることがある。", "setUp の変数は業務側に無い。", "業務側は必ず ${__P(access_token)}。", "", ""],
    ]
    r = table(ws, r, rows, [14, 32, 28, 40, 36, 14, 14])
    r += 1
    r = section(ws, r, last, "JSON Extractor の欄")
    r = note(
        ws, r, last,
        "子として HTTP_POST_cognito_token に付ける。公式の名前は JSON Extractor / JSON PostProcessor。メニューが英語のままなら英語で探す。"
        "応答の形は { \"access_token\": \"...\", \"token_type\": \"Bearer\", \"expires_in\": 3600 }。"
        "パスは $.access_token と $.expires_in。セミコロンの数を、4 つの欄で揃える。",
        C_SOFT, 9,
    )
    r = headers(ws, r, ["欄（画面の英語）", "入れる値", "意味", "空や 0 にすると", "", "", ""])
    fields = [
        ["Apply to", "Main sample only", "このリクエストの本体だけを見る。", "サブサンプルを選ぶと、別の応答から抜こうとする。", "", "", ""],
        ["Names of created variables", "access_token;expires_in", "できる変数名。セミコロン区切り。", "名前が 1 つだと、2 つ目のパスの結果を捨てる。", "", "", ""],
        ["JSON Path expressions", "$.access_token;$.expires_in", "ルートのプロパティ。ドットの前に $。", "access_token とだけ書くと、パスとして解釈されない。", "", "", ""],
        ["Match No.", "1;1", "複数ヒットしたときの何件目か。1 は最初。", "空欄は 0。0 はランダム。トークンには使わない。", "", "", ""],
        ["Default Values", "TOKEN_NOT_FOUND;EXPIRES_NOT_FOUND", "ヒットしないときの値。", "空だと、失敗と未処理の区別が付かない。", "", "", ""],
        ["Compute concatenation var", "オフ", "全部をカンマでつなぐ変数は、今回は不要。", "オンでも動くが、見る変数が増える。", "", "", ""],
    ]
    r = table(ws, r, fields, [28, 36, 40, 40, 14, 14, 14])
    r += 1
    r = images.add(
        ws, r, last,
        "JSON Extractor の欄",
        "4 つの欄を入れたあと。",
        "変数名、JSON Path、Match No. が 1;1、Default に TOKEN_NOT_FOUND が見える。実トークンは写さない。",
    )
    r = section(ws, r, last, "プロパティへコピーする")
    r = note(
        ws, r, last,
        "JSR223 PostProcessor は、JSON Extractor より下に置く。後処理はツリーの上から順に走る。"
        "setUp スレッドグループは、通常のスレッドグループより前に終わる。コピーが終わってから TG_業務 が始まるので、"
        "業務の 1 件目の時点で ${__P(access_token)} は既にある。",
        C_SOFT, 9,
    )
    r = _code(ws, r, last, GROOVY_STORE)
    r = images.add(
        ws, r, last,
        "PostProcessor が JSON Extractor の下にあるツリー",
        "並びを直したあと。",
        "上から PreProcessor、ヘッダ、JSON Extractor、PostProcessor、応答アサーション。スクリプトの実トークンは写さない。",
    )
    r = section(ws, r, last, "業務側での使い方と、期限")
    r = note(
        ws, r, last,
        "HM_API の Authorization の値は Bearer ${__P(access_token)}。Bearer と値のあいだは半角空白 1 つ。"
        "expires_in は秒。Cognito の例では 3600。アプリクライアントの設定で変わるので、1 人実行の応答を正にする。"
        "run-api.sh の -Jduration は、この秒より小さくする。超えると、期限のあとの業務リクエストが 401 になる。"
        "超える試験は、この手順では組まない。途中で取り直すループは入れない。",
        C_SOFT, 9,
    )
    r = headers(ws, r, ["業務の応答", "意味の候補", "見るもの", "直し方", "", "", ""])
    bad = [
        ["200", "トークンは受理された。", "統計は 23 へ進む。", "ここをトークンの不具合として調べ続けない。", "", "", ""],
        ["401", "Bearer が無い、期限切れ、scope が API の認可と違う、トークンの種類が違う。", "Request に Bearer があるか。expires_in と経過秒。", "${__P(access_token)} になっているか。duration を短くする。scope を合意と揃える。", "", "", ""],
        ["403", "トークンは乗っていても、API キー側で拒否されていることが多い。", "ヘッダ名と 39。", "トークンの Path を疑う前に、ヘッダ名と利用プランを見る。", "", "", ""],
    ]
    r = table(ws, r, bad, [12, 40, 40, 42, 16, 14, 14])
    r += 1
    r = note(
        ws, r, last,
        "デバッグサンプラーで JMeter Properties を表示すると、client_secret と api_key と access_token が結果ツリーに出る。"
        "この手順ではデバッグサンプラーを足さない。確認は、業務リクエストの Request タブで、式が残っていないことだけを見る。",
        C_EMPH, 9,
    )
    finish(ws, last)


def _regex(wb, images):
    ws = new_sheet(wb, "38_正規表現抽出", C_CHAR)
    last = 7
    r = banner(
        ws, last, "抽出",
        "正規表現抽出の欄を、1 つずつ使う",
        "Cognito の本番の抽出は 37 の JSON Extractor。正規表現は、JSON 以外の応答と、抽出式の確認に使う。",
    )
    r = note(
        ws, r, last,
        "要素の名前は「正規表現抽出」（Regular Expression Extractor）。HTTP リクエストを右クリックし、追加 → 後処理。"
        "公式は、式をスラッシュで囲まない、と書いている。/\"access_token\"/ のようには書かない。"
        "JMeter のこの要素は Perl 5 系の正規表現。Java の文字列の例にある \\\\s を、そのまま欄へ二重には書かない。欄には \\s と 1 つ書く。",
        C_SOFT, 9,
    )
    r = section(ws, r, last, "欄")
    r = headers(ws, r, ["欄", "英語", "トークンの練習で入れる値", "意味", "間違えやすい値", "", ""])
    fields = [
        ["適用先", "Apply to", "Main sample only", "この応答の本体に式を当てる。", "Sub-samples only。埋め込みリソースを見にいく。", "", ""],
        ["確認するフィールド", "Field to check", "Body", "応答本文。", "Response Headers。本文のトークンはヒットしない。", "", ""],
        ["参照名", "Name of created variable", "access_token_re", "${access_token_re} になる。練習なので、JSON 側の access_token とは別名。", "本番と同じ access_token。下に置くと、JSON の結果を上書きする。", "", ""],
        ["正規表現", "Regular Expression", "\"access_token\"\\s*:\\s*\"([^\"]+)\"", "引用符の中を、グループ 1 として取る。", ".*access_token.* のように、次の項目まで飲み込む式。", "", ""],
        ["テンプレート", "Template", "$1$", "グループ 1 だけを変数に入れる。", "${1} や \\1。$0$ はキー名ごと全部になる。", "", ""],
        ["一致番号", "Match No.", "1", "最初の一致。", "0 はランダム。-1 は全件で、本体の変数は初期値のまま。", "", ""],
        ["初期値", "Default Value", "TOKEN_NOT_FOUND", "不一致のときの値。", "空欄。失敗に気づけない。", "", ""],
    ]
    r = table(ws, r, fields, [18, 28, 40, 36, 36, 14, 14])
    r += 1
    r = note(
        ws, r, last,
        "上の正規表現を、抽出の欄へタイプするときの文字は次のとおり。バックスラッシュは 1 つ。\n"
        "\"access_token\"\\s*:\\s*\"([^\"]+)\"\n"
        "テンプレートの欄へタイプする文字は $1$",
        C_WHITE, 9,
    )
    r = section(ws, r, last, "式の読み方")
    r = headers(ws, r, ["部分", "名前", "動作", "トークンの式での役割", "", "", ""])
    parts = [
        ["\"", "そのままの引用符", "その文字がある位置に一致する。", "JSON のキーと値を囲む引用符。", "", "", ""],
        ["\\s*", "空白が 0 回以上", "スペースやタブが無くても、あっても一致する。", "コロンの前後の空白差を吸収する。", "", "", ""],
        [":", "そのままのコロン", "コロン 1 つ。", "キーと値の区切り。", "", "", ""],
        ["([^\"]+)", "グループ。引用符以外が 1 文字以上", "次の引用符の手前で止まる。+ は 1 文字以上なので、空は採用しない。", "トークン本体。テンプレート $1$ が受け取る。", "", "", ""],
        ["$1$", "テンプレート", "グループ 1 を格納する、という指示。正規表現の一部ではない。", "変数の中身をトークンだけにする。", "", "", ""],
        ["$0$", "全体", "式に一致した全部。", "今回は使わない。使うと access_token というキー名も値に入る。", "", "", ""],
    ]
    r = table(ws, r, parts, [14, 28, 40, 42, 16, 14, 14])
    r += 1
    r = section(ws, r, last, "一致番号でできる変数")
    r = headers(ws, r, ["一致番号", "access_token_re の中身", "ほかにできる変数", "トークンに使うか", "", "", ""])
    match = [
        ["1", "テンプレートの結果。グループ 1。", "access_token_re_g0 が一致全体、_g1 がグループ 1。", "使う。", "", "", ""],
        ["0", "ヒットが複数あるとき、その場でどれか 1 つ。", "どれを取ったかは実行のたびに変わり得る。", "使わない。", "", "", ""],
        ["2", "2 番目のヒット。", "1 つしか無い応答では初期値。", "2 つ目のトークンが仕様のときだけ。", "", "", ""],
        ["-1", "初期値のまま。公式は、負の番号では参照名そのものを初期値にする、と書いている。", "access_token_re_1 に 1 件目、access_token_re_matchNr に件数。", "一覧を後で回すとき。Bearer には使わない。", "", "", ""],
    ]
    r = table(ws, r, match, [12, 36, 52, 32, 16, 14, 14])
    r += 1
    r = section(ws, r, last, "Regexp Tester で、式を先に確かめる")
    r = headers(ws, r, H7)
    steps = [
        ["1", "1 人で実行", "35 の手順で、トークン取得が緑になっている。", "結果ツリーに HTTP_POST_cognito_token がある。", "行をクリックできる。", "負荷の人数のまま、テスターのために再実行しない。", ""],
        ["2", "下の一覧", "結果ツリーの下の表示形式で、Regexp Tester を選ぶ。日本語だと正規表現テスターのことがある。", "正規表現の入力欄が出る。", "欄が見える。", "Response data のタブを開いたまま、画面を撮影しない。", ""],
        ["3", "式", "正規表現の欄に、上の式を貼る。スラッシュでは囲まない。", "\"access_token\"\\s*:\\s*\"([^\"]+)\"", "Test を押す。", "\\\\s と二重に貼ると、ヒット 0 になる。", "下の枠"],
        ["4", "結果", "Match count と、group 1 を見る。", "件数が 1。group 1 がトークン本体で、キー名を含まない。", "式を抽出要素へコピーしてよい。", "group 1 の文字列を、メールや図に残さない。確認したらテスターを閉じる。", ""],
        ["5", "要素へ", "練習用の正規表現抽出を、トークンリクエストの子として足す。", "参照名は access_token_re。テンプレート $1$。一致番号 1。初期値 TOKEN_NOT_FOUND。", "もう一度 1 人で実行し、業務の Bearer は依然として ${__P(access_token)} のまま。", "参照名を access_token にして JSON Extractor より下に置くと、プロパティへコピーする値が正規表現側になる。練習が終わったら、この要素を無効にする。", ""],
        ["6", "期限の練習", "別の式で expires_in を抜く。", "\"expires_in\"\\s*:\\s*([0-9]+)", "グループ 1 が秒の数字。", "引用符を数字の周りに書くと、JSON の数にはヒットしない。expires_in は引用符で囲まれていない。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(
        ws, r, last,
        "正規表現抽出の入力欄",
        "練習用の要素を作ったあと。実行結果は開かない。",
        "参照名 access_token_re、正規表現、テンプレート $1$、一致番号 1、初期値 TOKEN_NOT_FOUND が 1 画面に入る。",
    )
    r = images.add(
        ws, r, last,
        "Regexp Tester の欄",
        "Test を押す前が安全。押したあとに撮るなら、一致文字列が見えない範囲。",
        "ドロップダウンで Regexp Tester が選ばれている。応答本文と group の値がフレームに入っていない。",
    )
    r = section(ws, r, last, "JSON Extractor との使い分け")
    r = headers(ws, r, ["応答", "使う要素", "理由", "両方付けたとき", "", "", ""])
    choice = [
        ["Cognito のトークン JSON", "JSON Extractor", "パス $.access_token で足りる。空白の差を式で吸収しなくてよい。", "本番では正規表現側を無効にする。", "", "", ""],
        ["HTML、プレーンテキスト、区切りが崩れた 1 行", "正規表現抽出", "JSON Path は JSON 以外では空になる。", "初期値のままなら、式が本文と違う。", "", "", ""],
        ["始まりと終わりの文字列が固定", "境界抽出器", "正規表現より短い。11 のカタログ。", "境界の文字列が 2 回出ると、先の方を取る。", "", "", ""],
        ["成否だけ見たい", "応答アサーション", "値は残さない。エラー率にだけ効く。", "抽出とアサーションは別の要素。アサーションは変数を作らない。", "", "", ""],
    ]
    r = table(ws, r, choice, [28, 22, 42, 40, 16, 14, 14])
    r += 1
    r = note(
        ws, r, last,
        "参照: https://jmeter.apache.org/usermanual/component_reference.html の Regular Expression Extractor と JSON Extractor。"
        "式の例の読み方は https://jmeter.apache.org/usermanual/regular_expressions.html 。"
        "テンプレート $1$ はグループ 1、$0$ は一致全体。一致番号 0 はランダム、負の番号は全件。",
        C_EMPH, 9,
    )
    finish(ws, last)


def _apikey(wb, images):
    ws = new_sheet(wb, "39_APIキーヘッダ", C_CHAR)
    last = 7
    r = banner(
        ws, last, "ヘッダ",
        "API Gateway へ、API キーの値をヘッダで渡す",
        "値は ${__P(api_key)}。名前は合意メモの文字を、大文字小文字ごとそのまま入れる。",
    )
    r = section(ws, r, last, "名前を決めてから、欄を埋める")
    r = note(
        ws, r, last,
        "この手順書の記入例は、ヘッダ名を X-API-ID とする。値はその API キー。"
        "AWS の利用プランでメソッドに API キーを必須にし、API key source が HEADER のとき、公式の呼び方はヘッダ名 X-API-Key、値は API キー。"
        "その契約なら、名前の欄だけ X-API-Key と書く。値の式は同じ。"
        "x-apigw-api-id は別のヘッダで、中身は API の ID（abc123 のような識別子）。API キーの値ではない。"
        "プライベート API を、共有の VPC エンドポイント名で呼ぶときに追加で要ることがある。キーの代わりにはならない。",
        C_SOFT, 9,
    )
    r = headers(ws, r, ["渡したいもの", "ヘッダ名", "値の欄", "付くリクエスト", "無いときの見え方", "", ""])
    rows = [
        ["合意メモが X-API-ID", "X-API-ID", "${__P(api_key)}", "業務リクエストの子 HM_API だけ", "ゲートウェイの実装がこの名前を必須にしていれば 401 か 403。メッセージは実装による。", "", ""],
        ["利用プランの API キー", "X-API-Key", "${__P(api_key)}", "同上。名前だけ変える。", "公式の説明では 403 Forbidden。InvalidKeyParameter の系統。", "", ""],
        ["プライベート API の API ID", "x-apigw-api-id", "${__P(api.id)}。api_key とは別のプロパティ。", "その呼び方をする業務リクエスト", "名前解決はできても、別の API に届く、または 403。", "", ""],
        ["アクセストークン", "Authorization", "Bearer ${__P(access_token)}", "業務リクエスト", "401。39 のキーとは別の列として数える。", "", ""],
    ]
    r = table(ws, r, rows, [24, 22, 32, 32, 40, 14, 14])
    r += 1
    r = section(ws, r, last, "画面での入れ方")
    r = headers(ws, r, H7)
    steps = [
        ["1", "secrets.properties", "api_key の行に、発行された API キーの値を入れる。", "api_key=現場の値", "引用符無し。", "load.properties に同じキーを書かない。-q を 2 つ付けるとき、後のファイルが同じキーを上書きする。", ""],
        ["2", "HTTP_GET_orders を右クリック", "追加 → 設定エレメント → HTTPヘッダマネージャ。", "名前 HM_API。このリクエストの子。", "左の木で、業務リクエストの 1 段下。", "テスト計画や setUp に置かない。", "下の枠"],
        ["3", "1 行目", "トークン。", "名前 Authorization、値 Bearer ${__P(access_token)}", "Bearer の直後が半角空白。", "${access_token} ではない。37。", ""],
        ["4", "2 行目", "API キー。記入例。", "名前 X-API-ID、値 ${__P(api_key)}", "名前が合意メモと 1 文字も違わない。", "値の欄にキーを貼り付ける。jmx に残る。", ""],
        ["5", "標準の利用プランのとき", "手順 4 の名前だけを変える。", "名前 X-API-Key、値 ${__P(api_key)}", "run-conditions.txt の api_header の文字も、実際に送った名前へ変える。", "X-API-ID と X-API-Key を両方付けて、どちらが効いたか分からない状態で報告する。", ""],
        ["6", "1 人で実行", "35 の手順 35。", "業務の Request に、決めた名前がある。値の位置に ${__P(api_key)} が残っていない。", "応答 200。", "トークン取得の Request に、同じ名前が付いていないか見る。付いていたら HM の置き場所を戻す。", ""],
        ["7", "CLI", "secrets.properties を chmod 600 で置き、-q を 2 回付ける。", "35 の run-api.sh。", "1 人の CLI で Err が 0。", "api_key を -J のコマンドラインへ直書きして、シェル履歴に残す。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(
        ws, r, last,
        "HM_API のヘッダ 2 行",
        "実行前。",
        "名前 Authorization と X-API-ID（または合意した X-API-Key）。値は式。実キーとトークンは写さない。",
    )
    r = section(ws, r, last, "403 と 401 の分け方")
    r = headers(ws, r, ["応答", "先に見ること", "キー側で直すこと", "トークン側で直すこと", "", "", ""])
    split = [
        ["403 Forbidden", "送ったヘッダ名が、ゲートウェイの期待と一致しているか。", "X-API-Key が契約なのに X-API-ID を送っていないか。値の取り違え。利用プランにステージが無い。メソッドで API キー必須なのに、デプロイし直していない。", "403 の主因にしにくい。Bearer が完全に欠けると 401 であることが多い。両方欠けることもあるので、Request を 1 回見れば足りる。", "", "", ""],
        ["401 Unauthorized", "Authorization が Bearer で始まるか。", "キー名の間違いだけでは 401 にならない実装が多い。それでも Request にキーがあるかは見る。", "37。期限、scope、${__P(access_token)} の書き間違い。", "", "", ""],
        ["200 なのに中身が拒否", "応答アサーションがコードだけだと緑になる。", "本文のメッセージを、1 人の実行のときだけその場で読む。jtl には残さない設定のまま。", "本文を「含む」アサーションにしない。件数が多い試験のエラー率はコードと、必要な最小の検査に限る。", "", "", ""],
    ]
    r = table(ws, r, split, [18, 32, 42, 42, 14, 14, 14])
    r += 1
    r = note(
        ws, r, last,
        "参照: API キーを付けた呼び出しの公式の例は、ヘッダ X-API-Key。"
        " https://docs.aws.amazon.com/apigateway/latest/developerguide/api-gateway-api-key-call.html "
        "キーの出所が HEADER の説明は https://docs.aws.amazon.com/apigateway/latest/developerguide/api-gateway-api-key-source.html 。"
        "合意した名前が X-API-ID の現場は、公式の例の名前だけを合意の名前に変える。値の渡し方（プロパティ、業務リクエストの子、Cognito には付けない）は同じ。",
        C_EMPH, 9,
    )
    finish(ws, last)


def build_auth(wb, images) -> None:
    _map(wb, images)
    _cognito(wb, images)
    _base64(wb, images)
    _token(wb, images)
    _regex(wb, images)
    _apikey(wb, images)
