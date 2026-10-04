# -*- coding: utf-8 -*-
"""後半シート。generate_field_manual.build から呼ばれる。"""

from __future__ import annotations

from datetime import datetime, timedelta

from generate_field_manual import (
    AL, C_BLACK, C_CHAR, C_DARK, C_EMPH, C_GRAY, C_SOFT, C_WHITE, C_ZEBRA,
    FILL, F, H7, LOAD_PROPERTIES, RUN_SH, SETENV_BAT, SETENV_SH, USER_PROPERTIES,
    W7_LIST,
    banner, finish, headers, merge, new_sheet, note, put, section, table,
)


def _cat(wb, images):
    ws = new_sheet(wb, "11_機能カタログ", C_CHAR)
    last = 7
    r = banner(ws, last, "機能", "何を、どこに、負荷試験でどう扱うか", "追加はすべて、親を右クリック → 追加。英語のままの名前は英語で探す。")
    r = note(ws, r, last, "この表はフィルタできる。分類の列で「リスナー」だけに絞ると、負荷の前に無効にするものが一覧になる。", C_SOFT, 10)
    r = headers(ws, r, ["分類", "画面の名前", "英語名", "親と追加メニュー", "役割", "最初に入れる値", "負荷試験での扱い"])
    rows = [
        ["計画", "テスト計画", "Test Plan", "起動したときからある", "シナリオ全体の名前と、共通の変数を置く場所。", "名前を業務名にする。", "消さない。ユーザ定義変数を直下に置いてよい。"],
        ["スレッド", "スレッドグループ", "Thread Group", "テスト計画 → 追加 → スレッド（ユーザー）", "仮想ユーザの人数、起動の速さ、繰り返し、持続時間。", "人数 ${__P(threads,1)}、Ramp-Up ${__P(ramp,1)}、ループ無限、スケジューラの持続 ${__P(duration,60)}。", "GUI ではプロパティを渡さないので 1 人・1 秒・60 秒になる。CLI の -J で増やす。"],
        ["スレッド", "setUp スレッドグループ", "setUp Thread Group", "テスト計画 → 追加 → スレッド（ユーザー）", "本番スレッドの前に 1 回だけ走る準備。", "データの事前投入が必要なときだけ。", "毎回重い準備を入れると、試験時間の意味が変わる。"],
        ["スレッド", "tearDown スレッドグループ", "tearDown Thread Group", "同上", "本番のあとに 1 回だけ走る後始末。", "ログアウトやテストデータ削除。", "失敗しても結果 jtl の本体と混ぜて判断しない。"],
        ["コントローラ", "トランザクションコントローラ", "Transaction Controller", "スレッドグループ → 追加 → ロジックコントローラ", "複数の HTTP を 1 つの業務操作として束ねる。", "名前は TX_検索 のように業務名。Generate parent sample はオフ（公式ダッシュボードの既定の勧め）。", "報告の 1 行にしたい名前をここに付ける。"],
        ["コントローラ", "一度だけ実行されるコントローラ", "Once Only Controller", "スレッドグループ → 追加 → ロジックコントローラ", "そのスレッドの最初の周だけ実行する。ログイン向き。", "中にログインの HTTP リクエストを置く。", "ループのたびにログインしたいときは、この中に入れない。"],
        ["コントローラ", "ループコントローラ", "Loop Controller", "同上", "中だけを繰り返す。", "最初は使わない。スレッドグループのループで足りる。", "二重ループは回数が読みにくくなる。"],
        ["コントローラ", "シンプルコントローラ", "Simple Controller", "同上", "フォルダ。処理は変えず、ツリーを見やすくする。", "名前だけ付ける。", "名前は jtl に出ない。統計の名前にはならない。"],
        ["コントローラ", "もしコントローラ", "If Controller", "同上", "条件が真のときだけ子を実行する。", "最初のシナリオでは使わない。", "条件式の誤りで全件スキップしても、エラーにならないことがある。"],
        ["設定", "HTTPリクエスト初期値", "HTTP Request Defaults", "テスト計画またはスレッドグループ → 追加 → 設定エレメント", "プロトコル、ホスト、ポートを共通にする。", "https、ホスト名、443。各リクエスト側のサーバ名は空。", "画面の名前が「HTTP Request Defaults」のままのことがある。英語で探す。"],
        ["設定", "HTTPヘッダマネージャ", "HTTP Header Manager", "同上", "Content-Type や Authorization を付ける。", "Content-Type: application/json", "子に置くと、そのリクエストだけ追加のヘッダが乗る。"],
        ["設定", "HTTPクッキーマネージャ", "HTTP Cookie Manager", "スレッドグループ → 追加 → 設定エレメント", "Cookie をスレッドごとに保持する。", "追加するだけで、初期は空でよい。", "ログインセッションを Cookie で維持するサイトでは必須。"],
        ["設定", "HTTPキャッシュマネージャ", "HTTP Cache Manager", "同上", "ブラウザのキャッシュの真似。", "API の負荷試験では追加しない。", "キャッシュが乗ると、サーバに届かないリクエストが増え、負荷を過小評価する。"],
        ["設定", "CSV Data Set Config", "CSV Data Set Config", "スレッドグループ → 追加 → 設定エレメント", "CSV の 1 行を、スレッドの変数にする。", "ファイル名は jmx からの相対。共有は All threads。1 行目は列名。", "文字コードを UTF-8 にする。Excel の BOM 付きで先頭列が読めないことがある。"],
        ["設定", "ユーザー定義変数", "User Defined Variables", "テスト計画 → 追加 → 設定エレメント", "全スレッド共通の定数。", "ホスト名をここに置くこともある。ただし CLI で変えたい値はプロパティ ${__P} にする。", "スレッドごとに違うユーザ ID をここに書かない。CSV を使う。"],
        ["設定", "カウンタ", "Counter", "設定エレメント", "スレッドをまたいで増える番号。", "一意の注文番号が必要なとき。", "再実行で同じ番号から始まると、相手サーバが重複エラーにする。開始値の決め方をレポートに書く。"],
        ["サンプラー", "HTTPリクエスト", "HTTP Request", "コントローラまたはスレッドグループ → 追加 → サンプラー", "1 回の HTTP。", "メソッド、パス、必要なら Body。サーバ名は空（初期値に任せる）。", "名前を HTTP_GET_orders のように一意にする。同じ名前は 1 行に合算される。"],
        ["サンプラー", "デバッグサンプラー", "Debug Sampler", "追加 → サンプラー", "その時点の変数を結果ツリーに出す。", "トークンが取れているかの確認。", "負荷の前に無効にする。変数の中身が jtl に残り得る。"],
        ["後処理", "JSON Extractor", "JSON Extractor / JSON PostProcessor", "HTTP リクエスト → 追加 → 後処理", "JSON 応答から値を抜き、変数にする。", "変数名 token、JSON Path $.token、Match No. 1、Default TOKEN_NOT_FOUND。", "Default を空にすると、取りこぼしに気づけない。"],
        ["後処理", "正規表現抽出", "Regular Expression Extractor", "HTTP リクエスト → 追加 → 後処理", "テキストから抜く。HTML やプレーンテキスト向き。", "参照名、正規表現、テンプレート $1$、一致番号 1、初期値 NOT_FOUND。", "欲張った .* は遅く、誤って長い文字列を掴む。"],
        ["後処理", "境界抽出器", "Boundary Extractor", "同上", "開始文字列と終了文字列で抜く。正規表現より単純。", "左境界と右境界。", "境界文字列が応答に 2 回あると、意図しない方を取ることがある。"],
        ["アサーション", "応答アサーション", "Response Assertion", "リクエスト → 追加 → アサーション", "応答コードや本文が期待どおりか。", "応答コード、等しい、200。", "失敗はエラー率に入る。遅いこと自体はここでは見ない。"],
        ["アサーション", "アサーションの持続", "Duration Assertion", "同上", "所要時間がミリ秒の上限を超えたら失敗にする。", "最初は入れない。", "入れると、遅さと機能エラーが同じエラー率に混ざる。SLO は Excel で判定する。"],
        ["アサーション", "サイズアサーション", "Size Assertion", "同上", "応答のバイト数。", "空応答を失敗にしたいとき。", "圧縮の有無でサイズが変わり、誤検知する。"],
        ["アサーション", "JSON Assertion", "JSON Assertion / JSON Path Assertion", "追加 → アサーション", "JSON のパスが期待値か。", "メニューに無いバージョンでは、抽出＋応答アサーションで代用する。", "パスが無いだけの失敗と、HTTP の失敗をレポートで区別できるよう、メッセージを残す設定にする。"],
        ["タイマ", "定数タイマ", "Constant Timer", "追加 → タイマ", "毎回同じミリ秒だけ、次のリクエストの前に待つ。", "固定の待ちが仕様のとき。", "スコープを間違えると、ログインの前にも待ちが入る。"],
        ["タイマ", "一様乱数タイマ", "Uniform Random Timer", "追加 → タイマ", "待ちをばらす。人が画面を見る時間の代わり。", "定数遅延 1000、ランダム遅延 2000 なら、1〜3 秒の一様乱数。", "入れた待ちはスループットを下げる。レポートの条件に「思考時間あり」と書く。"],
        ["タイマ", "定数スループットタイマ", "Constant Throughput Timer", "追加 → タイマ", "目標の 1 分あたり件数に合わせて待つ。", "件数が契約で決まっているときだけ。", "目標がサーバの限界を超えると、タイマは待たなくなり、目標に届かない。届かなかった事実を書く。"],
        ["リスナー", "結果をツリーで表示", "View Results Tree", "追加 → リスナー", "1 件ずつの要求と応答を見る。デバッグの本体。", "デバッグ中だけ有効。", "CLI の前に無効。有効のままだとサンプルをメモリに貯めて落ちる。"],
        ["リスナー", "統計レポート", "Aggregate Report", "追加 → リスナー", "平均、中央値、90%、95%、99%、エラー率、スループット。", "GUI の少人数確認の目安。", "負荷の数値の正本にしない。正本は CLI の jtl と statistics.json。"],
        ["リスナー", "集計レポート", "Summary Report", "追加 → リスナー", "統計レポートより列が少ない要約。", "同上。", "同上。GUI の画面をスクリーンショットして報告にしない。"],
        ["リスナー", "結果を表で表示", "View Results in Table", "追加 → リスナー", "1 行 1 サンプルの表。", "失敗行を目で探すとき。", "負荷の前に無効。"],
        ["リスナー", "簡易データライタ", "Simple Data Writer", "追加 → リスナー", "リスナーからファイルへ書く。", "この手順では使わない。", "-l と二重に書くと混乱する。出力はコマンドの -l に一本化する。"],
        ["リスナー", "Backend Listener", "Backend Listener", "追加 → リスナー", "実行中に外部へメトリクスを送る。", "今回の基本手順では追加しない。", "送り先が無いとエラーログが増える。使うなら送信先を試験条件に書く。"],
        ["前処理", "JSR223 PreProcessor", "JSR223 PreProcessor", "追加 → 前処理", "Groovy でリクエスト前に加工する。", "最初は使わない。", "Language を javascript にしない。Java 11 以降 Nashorn は非推奨で、動かない環境がある。Groovy を使う。"],
        ["録音", "HTTP(S) Test Script Recorder", "HTTP(S) Test Script Recorder", "テスト計画 → 追加 → Non-Testエレメント", "ブラウザ操作を HTTP リクエストとして記録する。", "手動で API を書けるなら使わない。", "録音用プロキシと、JMeter が外へ出るときのプロキシは別物。16 を読む。"],
        ["テンプレート", "テンプレートから作成", "Templates", "ファイル → テンプレート（Templates）", "録音用や既製の骨組みを作る。", "中身を読めるようになってから。", "テンプレートのリスナーを無効にし忘れる。"],
    ]
    header_row = r - 1
    r = table(ws, r, rows, [12, 28, 36, 40, 36, 42, 36])
    ws.auto_filter.ref = f"A{header_row}:G{r - 1}"
    ws.freeze_panes = f"A{header_row + 1}"
    r += 1
    r = section(ws, r, last, "プラグインを足す手順（必要なときだけ）")
    r = headers(ws, r, H7)
    plug = [
        ["1", "方針", "標準の要素で書けるなら足さない。", "Ultimate Thread Group などが欲しくなったときだけ。", "カタログの標準要素でシナリオが組める。", "先にプラグインを入れて、サーバへ jar を送り忘れる。", ""],
        ["2", "ブラウザ", "jmeter-plugins.org の Plugins Manager の説明を開く。Apache 本体ではない。", "https://jmeter-plugins.org/wiki/PluginsManager/", "jar のダウンロードリンクがある。", "出所の違う同名 jar を置かない。", ""],
        ["3", "lib/ext", "jmeter-plugins-manager の jar を、端末の lib\\ext に置く。", "C:\\Tools\\apache-jmeter-5.6.3\\lib\\ext", "jar が ext の直下にある。", "lib 直下に置くと部品として認識されない。公式は部品を lib/ext、補助ライブラリを lib と分けている。", ""],
        ["4", "再起動", "JMeter を閉じ、jmeter.bat で開き直す。", "オプション → Plugins Manager", "窓が開く。", "起動したまま jar を置いてもメニューは増えない。", "下の枠"],
        ["5", "導入", "Available から必要なものだけ選び、Apply Changes and Restart。", "例は無理に入れない。", "再起動後、追加メニューに新しい要素が出る。", "Suggested を全部入れない。試験と無関係な要素が増える。", ""],
        ["6", "サーバ", "Plugins Manager が入った jar を、Linux の同じ相対位置へコピーする。lib と lib/ext の両方を見る。", "scp で jar だけ送る。", "サーバの lib/ext のファイル名が端末と一致。", "jmx だけ送って jar を送らないと、CLI は起動直後にプランを読めず終わる。", ""],
    ]
    r = table(ws, r, plug, W7_LIST)
    r += 1
    r = images.add(ws, r, last, "Plugins Manager を開いた画面", "jar を入れて再起動したあと。入れないとメニューが無いので、その場合はこの枠に「未使用」と記入して図は貼らなくてよい。", "オプションメニューと Plugins Manager の窓。導入したプラグイン名が読めること。")
    finish(ws, last)


def _scenario(wb, images):
    ws = new_sheet(wb, "12_シナリオ作成", C_CHAR)
    last = 7
    r = banner(ws, last, "GUI", "注文検索のシナリオを、クリック順に作る", "例のホストは app.example.internal。現場の試験対象に置き換える。スレッド数は最初からプロパティ参照にする。")
    r = note(ws, r, last, "完成形の木は次のとおり。上からこの順に足す。\nテスト計画 注文検索\n  HTTPリクエスト初期値\n  HTTPヘッダマネージャ（Content-Type）\n  スレッドグループ TG_注文検索\n    CSV Data Set Config\n    HTTPクッキーマネージャ\n    一度だけ実行されるコントローラ\n      HTTPリクエスト HTTP_POST_login\n        JSON Extractor（token）\n        応答アサーション（コード 200）\n    トランザクションコントローラ TX_検索\n      HTTPリクエスト HTTP_GET_orders\n        応答アサーション（コード 200）\n    一様乱数タイマ\n    結果をツリーで表示（デバッグ中だけ有効）", C_SOFT, 9)
    r = note(ws, r, last, "現場の認証が Cognito のクライアント認証で、API Gateway に API キーを載せるときは、この木のログインを使わない。34 の木を、35 のクリック順で api-gateway.jmx として別に作る。Base64 は 36、トークンの保持は 37、正規表現の練習は 38、ヘッダ名は 39。", C_EMPH, 9)
    r = headers(ws, r, H7)
    steps = [
        ["1", "テスト計画", "右の名前を変える。", "注文検索", "ツリーの一番上が 注文検索。", "ここにスレッド数は書かない。スレッド数はスレッドグループ。", ""],
        ["2", "テスト計画を右クリック", "追加 → 設定エレメント → HTTPリクエスト初期値（HTTP Request Defaults）。", "プロトコル https、サーバ app.example.internal、ポート 443。", "各リクエストにホストを書かなくてよくなる。", "ポートを空のままにすると 80 や 443 の既定に依存する。明示する。", "下の枠"],
        ["3", "テスト計画を右クリック", "追加 → 設定エレメント → HTTPヘッダマネージャ。", "名前 Content-Type、値 application/json", "JSON を送るリクエストの共通ヘッダになる。", "Authorization はまだ入れない。トークンはログイン後の変数。", ""],
        ["4", "テスト計画を右クリック", "追加 → スレッド（ユーザー） → スレッドグループ。", "名前 TG_注文検索。", "右に人数の欄が出る。", "サンプラーをこの外に置かない。", "下の枠"],
        ["5", "スレッドグループ", "エラー時の動作。", "継続（Continue）。", "1 件失敗しても、そのスレッドは次へ進む。", "停止（Stop Test）にすると、初回の失敗で全体が止まり、エラー率が小さくなる。デバッグ以外は継続。", ""],
        ["6", "スレッド数", "人数をプロパティにする。", "${__P(threads,1)}", "GUI では 1 人になる。", "欄に 100 と直書きすると、GUI の開始がそのまま 100 人になる。", ""],
        ["7", "Ramp-Up 期間（秒）", "全員が起動し終わる秒数。", "${__P(ramp,1)}", "1 人のときは 1 秒。", "100 人なのに ramp を 1 にすると、ほぼ一斉起動になる。試験条件と違う。", ""],
        ["8", "ループ回数", "無限にチェックを入れる。", "無限ループ（Infinite）。", "回数の欄がグレーになる。", "無限のままスケジューラをオフにすると、手動で止めるまで終わらない。次の手順でスケジューラをオンにする。", ""],
        ["9", "スケジューラ", "スケジューラをオンにし、持続時間だけ入れる。", "持続時間（秒） ${__P(duration,60)}、起動遅延は空または 0。", "GUI では約 60 秒で終わる。", "持続時間を Ramp-Up より短くすると、全員そろう前に終わる。", ""],
        ["10", "同じユーザを繰り返す", "ループの各周で Cookie を維持するかの指定。画面の「各反復で同じユーザ」に相当するチェック。", "ログインを一度だけにして検索を繰り返すならオン。", "2 周目も同じセッションで検索できる。", "オフだと周ごとに接続や Cookie が切れ、ログインが一度だけのコントローラと矛盾することがある。", ""],
        ["11", "スレッドグループを右クリック", "追加 → 設定エレメント → CSV Data Set Config。", "13 の手順で中身を入れる。ファイルが無い間は、ログインを固定値で先に通してよい。", "要素がツリーに出る。", "CSV のパスを絶対パスの C:\\Users\\自分 にすると、Linux で読めない。", ""],
        ["12", "スレッドグループを右クリック", "追加 → 設定エレメント → HTTPクッキーマネージャ。", "名前以外は空でよい。", "要素がある。", "クリアのチェックを不用意にオンにすると、周ごとに Cookie が消える。", ""],
        ["13", "スレッドグループを右クリック", "追加 → ロジックコントローラ → 一度だけ実行されるコントローラ。", "名前 ONCE_ログイン。", "そのスレッドの最初だけ子が動く。", "この中に検索リクエストを入れると、検索が 1 回しか出ない。", ""],
        ["14", "ONCE_ログインを右クリック", "追加 → サンプラー → HTTPリクエスト。", "名前 HTTP_POST_login。メソッド POST。パス /api/login。サーバ名は空。", "Body に JSON を書く場所がある。", "サーバ名も書くと、初期値とどちらが効くか迷う。空にする。", "下の枠"],
        ["15", "HTTP_POST_login", "Body Data に JSON を書く。パラメータ表と Body を同時に使わない。", "{\"user\":\"${user}\",\"password\":\"${password}\"}", "変数名が CSV の列名と一致。", "パラメータ表に同じ内容を入れると、JSON にならない。", ""],
        ["16", "HTTP_POST_loginを右クリック", "追加 → 後処理 → JSON Extractor。", "Names of created variables: token\nJSON Path expressions: $.access_token\nMatch No.: 1\nDefault Values: TOKEN_NOT_FOUND", "応答に access_token が無いと変数は TOKEN_NOT_FOUND。", "Default を空にすると、空トークンのまま次が 401 になり、原因が分かりにくい。", "下の枠"],
        ["17", "HTTP_POST_loginを右クリック", "追加 → アサーション → 応答アサーション。", "Field to Test: 応答コード。Pattern Matching: 等しい。Patterns: 200。", "200 以外は赤い失敗になる。", "「含む」のまま 200 と書くと、2000 にも一致し得る。等しいにする。", ""],
        ["18", "スレッドグループを右クリック", "追加 → ロジックコントローラ → トランザクションコントローラ。", "名前 TX_検索。Generate parent sample はオフ。Include duration of timer は、待ちを取引時間に入れたいときだけオン。", "子の HTTP が個別に jtl へ出る。", "親サンプルをオンにすると、親と子の両方が出て二重計上しやすい。報告方針を先に決める。この手順はオフ。", ""],
        ["19", "TX_検索を右クリック", "追加 → サンプラー → HTTPリクエスト。", "名前 HTTP_GET_orders。メソッド GET。パス /api/orders。", "パラメータ q = ${keyword}。", "パスに日本語を直書きすると、エンコードの差でサーバが 400 を返すことがある。", ""],
        ["20", "TX_検索または GET のヘッダ", "このリクエストの子に HTTPヘッダマネージャを足す。", "名前 Authorization、値 Bearer ${token}", "ログインで取ったトークンが付く。", "テスト計画直下のヘッダに Bearer を書くと、ログインリクエストにも付く。", ""],
        ["21", "HTTP_GET_ordersを右クリック", "追加 → アサーション → 応答アサーション。", "応答コードが 200 と等しい。", "検索 API の成功がエラー率に反映される。", "本文の一部を「含む」で検査すると、文言変更で全件失敗する。コードを先に見る。", ""],
        ["22", "TX_検索を右クリック", "追加 → タイマ → 一様乱数タイマ。", "定数遅延 1000、ランダム遅延 2000（ミリ秒）。", "検索の前に 1〜3 秒待つ。", "スレッドグループ直下に置くと、ログインの前にも待ちが入る。TX_検索の子にする。", ""],
        ["23", "スレッドグループを右クリック", "追加 → リスナー → 結果をツリーで表示。", "名前は既定のままでよい。ファイル名は空。", "14 で 1 件ずつ見られる。", "Filename にパスを書くと、GUI がそこに巨大ファイルを書く。空のまま。", ""],
        ["24", "ファイルメニュー", "テスト計画に名前をつけて保存。", "C:\\LoadTest\\scenarios\\order-search\\order-search.jmx", "jmx ができる。", "bin フォルダや OneDrive に保存しない。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(ws, r, last, "完成した左ツリー", "手順 24 の保存直前。", "このシート冒頭の木と同じ親子が、折りたたまない状態で読める。")
    r = images.add(ws, r, last, "スレッドグループの入力欄", "手順 9 のあと。", "スレッド数、Ramp-Up、無限ループ、スケジューラ、持続時間が 1 画面に入る。値は ${__P(...)} が見える。")
    r = images.add(ws, r, last, "HTTPリクエスト初期値", "手順 2 のあと。", "プロトコル、サーバ名、ポート。パスは空。")
    r = images.add(ws, r, last, "ログインの HTTP リクエスト", "手順 15 のあと。", "メソッド POST、パス、Body Data の JSON。パスワードの実値は写さない。${password} のままでよい。")
    r = section(ws, r, last, "録音で下書きを作る場合")
    r = note(ws, r, last, "ブラウザ操作から下書きを作るときは、テスト計画 → 追加 → Non-Testエレメント → HTTP(S) Test Script Recorder。ポートは 8888、対象コントローラを先に作って指定する。開始すると bin に ApacheJMeterTemporaryRootCA.crt ができる。Windows の信頼されたルートへ入れる必要がある。ブラウザのプロキシを 127.0.0.1:8888 にして操作し、終わったら停止して静的ファイルを消す。会社のプロキシとぶつかる。証明書のインストールが禁止なら録音は使わず、このシートの手入力で作る。録音は作成用であり、負荷実行の方法ではない。")
    finish(ws, last)


def _data(wb, images):
    ws = new_sheet(wb, "13_データと相関", C_CHAR)
    last = 7
    r = banner(ws, last, "データ", "CSV でユーザを渡し、応答からトークンを受け渡す", "プロパティ（-J）と変数（CSV や抽出）は別物。ここは変数。")
    r = section(ws, r, last, "users.csv を作る")
    r = headers(ws, r, H7)
    steps = [
        ["1", "Excel", "1 行目に列名、2 行目以降に試験用アカウントを書く。", "user, password, keyword\nuser01, Passw0rd-example, book", "列名に空白や全角括弧が無い。", "本番のパスワードや個人名を入れない。試験用の語にする。", ""],
        ["2", "Excel", "名前を付けて保存で CSV UTF-8 を選ぶ。", "C:\\LoadTest\\scenarios\\order-search\\users.csv", "jmx と同じフォルダ。", "「CSV（コンマ区切り）」だけだと Shift_JIS になり、Linux 上の JMeter が日本語を壊す。", ""],
        ["3", "メモ帳", "先頭の列名の前に見えない文字（BOM）が無いか確認する。疑わしいときは UTF-8（BOM 無し）で保存し直す。", "1 行目が user で始まる。", "JMeter の変数名が user になる。", "BOM が付くと変数名が user ではなくなり、${user} が置換されない。", ""],
        ["4", "CSV Data Set Config", "ファイル名、エンコーディング、共有、末尾の動作を入れる。", "Filename: users.csv\nFile encoding: UTF-8\nSharing mode: All threads\nRecycle on EOF: True\nStop thread on EOF: False", "相対パスは jmx のあるフォルダから解決される。", "Ignore first line を false のまま、1 行目をデータとして読んでしまう設定にしない。列名行は「先頭行を変数名として読む」動作に任せる。Allow quoted data は True。", "下の枠"],
        ["5", "使い回し", "行数より試行回数が多いときの方針を決める。", "使い回してよい試験: Recycle True。1 行を 1 人だけに使いたい試験: Recycle False、Stop thread on EOF True。", "方針をレポートの条件に 1 文書く。", "行が尽きたあとも空の変数でリクエストし続け、400 の山を作る。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(ws, r, last, "CSV Data Set Config の入力欄", "手順 4 のあと。", "Filename が users.csv で、絶対パスではない。Encoding が UTF-8。Sharing mode が見える。")
    r = section(ws, r, last, "トークンの受け渡し")
    r = note(ws, r, last, "ログインの JSON Extractor が token という変数を、そのスレッドだけに作る。次のリクエストのヘッダ Authorization の値を Bearer ${token} にする。他のスレッドのトークンは見えない。これが正しい。\n${__P(threads,1)} はプロパティで、CSV では作れない。逆に ${token} を -Jtoken= で渡しても、抽出結果では上書きされるので、トークンをコマンドラインで渡す設計にはしない。")
    r = headers(ws, r, ["確認", "どこを見る", "期待", "失敗の見え方", "直し方", "", ""])
    checks = [
        ["変数が置換された", "結果ツリー → リクエスト → Request Body", "user01 など CSV の値が入っている。${user} という文字が残っていない。", "本文に ${user} が残る。", "列名と変数名の不一致、BOM、CSV が jmx の隣に無い。", "", ""],
        ["トークンが取れた", "結果ツリー → HTTP_POST_login の次のリクエストのヘッダ", "Bearer の後に TOKEN_NOT_FOUND 以外の文字列。", "Bearer TOKEN_NOT_FOUND または Bearer ${token}。", "JSON Path が応答と違う。応答の実物をツリーで見て Path を直す。", "", ""],
        ["パスがサーバで解決される", "jmx を Linux の plans/order-search へ置いたあと、同じ階層に users.csv がある", "CLI が CSV を開く。", "jmeter.log に FileNotFound。", "フォルダ階層を端末と揃える。絶対パスにしない。", "", ""],
    ]
    r = table(ws, r, checks, [14, 24, 38, 42, 32, 32, 16])
    r += 1
    r = images.add(ws, r, last, "JSON Extractor の 4 欄", "ログインリクエストの子として設定したあと。", "変数名、JSON Path、Match No.、Default Values の 4 つが読める。実トークンは写さない。")
    finish(ws, last)


def _gui_run(wb, images):
    ws = new_sheet(wb, "14_GUIで実行", C_CHAR)
    last = 7
    r = banner(ws, last, "GUI", "1 人だけで流して、成功を目で見る", "公式の Getting Started は、GUI はシナリオ作成用、負荷試験は CLI、と書いている。ここは作成の確認。")
    r = headers(ws, r, H7)
    steps = [
        ["1", "保存", "実行の前に jmx を保存する。", "Ctrl+S。", "タイトルバーに * が無い。", "未保存のまま実行すると、落ちたときにシナリオが消える。", ""],
        ["2", "人数の確認", "スレッドグループを開き、人数がプロパティの既定 1 であることを確認する。", "${__P(threads,1)} のまま。-J は GUI では付けない。", "1 人、Ramp-Up 1 秒、持続 60 秒。", "欄を 50 に書き換えて緑の三角を押さない。", ""],
        ["3", "リスナー", "結果をツリーで表示 が有効（グレーではない）であることを確認する。他のグラフ用リスナーは無効。", "ツリーだけ有効。", "左のツリーでリスナーが黒字。", "統計レポートを開いたままにすると、少ない件数なら動くが、癖になる。デバッグはツリーに一本化する。", ""],
        ["4", "消去", "前の実行結果が残っていたら消す。", "実行 → 全て消去（Ctrl+Shift+E）。", "ツリーが空になる。jmx は残る。", "全て消去はシナリオの削除ではない。", ""],
        ["5", "開始", "ツールバーの緑の三角（開始）を押す。", "メニューなら 実行 → 開始。", "右上に緑の表示が出て、件数が増える。", "開始（タイマを無視）は待ち時間を飛ばす確認用。思考時間の確認では使わない。", "下の枠"],
        ["6", "停止の判断", "60 秒待つか、ログインと検索が 1 回ずつ成功したら止めてよい。", "シャットダウン（Shutdown）を使う。赤の停止は、応答が返らず固まったとき。", "件数の増加が止まる。", "停止したあとに、停止状態のまま保存して、次に開始できないことがある。もう一度開始できるか見る。", ""],
        ["7", "1 件目", "結果ツリーで HTTP_POST_login をクリックする。", "緑のチェック。応答コード 200。", "Sampler result に 200。", "赤い三角なら、応答コードと Response data を読む。人数は増やさない。", "下の枠"],
        ["8", "リクエスト本文", "Request タブを開く。", "CSV のユーザが入っている。", "${user} が残っていない。", "残っていたら 13 へ戻る。", ""],
        ["9", "次のリクエスト", "HTTP_GET_orders の Request ヘッダを見る。", "Authorization: Bearer のあとが TOKEN_NOT_FOUND ではない。", "検索も緑。", "401 ならトークンの Path か、ヘッダを付けた場所がログインより前。", ""],
        ["10", "応答時間の目安", "Sampler result の Load time（elapsed）を見る。", "この環境で何ミリ秒だったかをメモする。合格判定には使わない。", "1 件の値が分かる。", "1 件の値を「性能」として報告しない。", ""],
        ["11", "保存", "直したら保存する。", "Ctrl+S。", "jmx が更新される。", "結果ツリーを有効なまま、このファイルを負荷用の正とすると危ない。17 で無効化してから転送用に保存する。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(ws, r, last, "開始ボタンを押した直後の全体", "件数が 1 以上になり、まだ実行中か、終わった直後。", "緑の三角の位置と、右上の実行中表示、左ツリーの結果ツリーが見える。")
    r = images.add(ws, r, last, "結果ツリーでログインが緑の画面", "Sampler result タブ。", "サンプル名、応答コード 200、Load time が見える。トークンやパスワードの本文は写さない。Response data は畳む。")
    r = note(ws, r, last, "検証（Validate）ボタンは、スレッドグループを短く確認するためのもの。待ち時間の確認や、プロパティ既定どおりの 60 秒の確認には、開始を使う。Validate の結果だけで転送しない。")
    finish(ws, last)


def _gui_results(wb, images):
    ws = new_sheet(wb, "15_GUIの結果確認", C_CHAR)
    last = 7
    r = banner(ws, last, "GUI", "画面の結果を、何の判断に使ってよいか", "GUI の数字はデバッグ用。提出用の数字は CLI の statistics.json と jtl。")
    r = headers(ws, r, ["画面", "開く場所", "見る欄", "判断してよいこと", "判断してはいけないこと", "その後", ""])
    rows = [
        ["結果をツリーで表示", "リスナーをクリック", "緑か赤か。応答コード。Request。Response data。", "その 1 件が機能的に成功したか。変数が入ったか。", "平均応答時間、スループット、同時ユーザ耐性。", "原因を直してからもう一度 1 人で実行。", "下の枠"],
        ["Sampler result", "ツリーの中のタブ", "Thread Name、Load time、Latency、Connect、bytes。", "失敗が接続（Connect）なのか、処理（Latency 以降）なのかの当たり。", "パーセンタイル。", "メモだけ残す。", ""],
        ["統計レポート", "追加してある場合", "平均、中央値、90% Line、95% Line、99% Line、エラー%、スループット。", "数件のデバッグで、桁がおかしくないか。", "提出値。件数が多いと GUI が重く、値も落ちる。", "見たら無効に戻す。", "下の枠"],
        ["集計レポート", "追加してある場合", "平均、最小、最大、エラー率。パーセンタイル列は統計レポートより少ない。", "エラーが混ざっていないかの概観。", "95% 値。この画面には薄い。", "提出に使わない。", ""],
        ["HTML レポート生成", "ツール → Generate HTML report", "結果ファイル、user.properties、空の出力フォルダ。", "手元の小さい jtl でレポートの見方を練習すること。", "GUI で流した数件を、負荷試験の報告書にすること。", "本番の HTML は Linux の -e -o で作る。", ""],
    ]
    r = table(ws, r, rows, [16, 22, 36, 40, 36, 32, 14])
    r += 1
    r = images.add(ws, r, last, "結果ツリーの Request タブ", "変数が置換されたリクエスト。", "URL とヘッダ名が見える。Authorization の値は隠す。")
    r = images.add(ws, r, last, "統計レポートを開いた例（数件）", "1 人で数周したあと。負荷の証拠としては使わない、と分かった状態で撮る。", "Label、# Samples、Average、95% Line、Error %、Throughput の列見出しが見える。")
    r = section(ws, r, last, "統計レポートの列（提出値ではないが、言葉の対応）")
    r = headers(ws, r, ["列", "意味", "単位", "CLI 側の対応", "読み違い", "", ""])
    cols = [
        ["Label", "サンプラーまたはトランザクションの名前", "文字", "jtl の label、statistics.json のキー", "同じ名前は 1 行にまとまる。", "", ""],
        ["# Samples", "その名前の件数", "件", "sampleCount", "親と子を両方数えていないか。", "", ""],
        ["Average", "elapsed の平均。失敗も含む。", "ミリ秒", "meanResTime", "平均が小さいのに最大だけ巨大、は外れ値。", "", ""],
        ["Median", "中央値。", "ミリ秒", "medianResTime", "平均より中央値がずっと小さいときは、遅い側に尾がある。", "", ""],
        ["90/95/99% Line", "その割合のリクエストが収まる時間。", "ミリ秒", "pct1 / pct2 / pct3（既定 90/95/99）", "GUI とダッシュボードで計算方法が違い、件数が少ないと差が見える。公式もその差に触れている。", "", ""],
        ["Error %", "失敗の割合。", "%", "errorPct", "HTTP 200 でもアサーション失敗はエラー。", "", ""],
        ["Throughput", "その名前の件数 / 最初から最後までの時間。", "件/秒", "throughput", "スレッド数ではない。思考時間を入れると下がる。", "", ""],
        ["Received KB/sec", "受信バイトの速度。", "KB/秒", "receivedKBytesPerSec", "応答が小さい API では小さな値で正常。", "", ""],
    ]
    r = table(ws, r, cols, [16, 22, 36, 40, 36, 32, 14])
    finish(ws, last)


def _checks(wb, images):
    ws = new_sheet(wb, "16_確認観点と注意", C_DARK)
    last = 7
    r = banner(ws, last, "判断", "流す前と、流したあとに見るリスト", "ここを飛ばして人数だけ増やさない。")
    r = section(ws, r, last, "GUI で 1 人を流したあとの観点")
    r = headers(ws, r, ["観点", "見方", "合格の目安", "不合格の見え方", "戻るシート", "", ""])
    rows = [
        ["対象が試験用", "初期値のホスト名", "合意した試験環境。", "本番のホスト名が入っている。", "12", "", ""],
        ["人数が 1", "スレッドグループの式", "${__P(threads,1)} のまま。", "整数が直書きされている。", "12", "", ""],
        ["ログインが成功", "結果ツリー", "緑、応答コード 200。", "赤、401、500、接続拒否。", "14", "", ""],
        ["変数が入った", "Request タブ", "${user} が残っていない。", "ドル記号がそのまま送られている。", "13", "", ""],
        ["トークンが入った", "検索のヘッダ", "TOKEN_NOT_FOUND ではない。", "空、またはその文字列のまま。", "13", "", ""],
        ["検索が成功", "結果ツリー", "HTTP_GET_orders が緑。", "200 以外。本文が想定と違う。", "14", "", ""],
        ["待ちの位置", "ツリーの親子", "一様乱数タイマが TX_検索の中。", "ログインの前で毎回待っている。", "12", "", ""],
        ["秘密情報", "Request と jmx", "パスワードが例示用。ログに残す設定がオフ。", "本番パスワードが jmx に見える。", "08, 13", "", ""],
    ]
    r = table(ws, r, rows, [14, 24, 38, 42, 32, 32, 16])
    r += 1
    r = section(ws, r, last, "CLI の前に必ずやること")
    r = headers(ws, r, H7)
    before = [
        ["1", "リスナー", "結果ツリー、統計、集計、表、グラフをすべて無効にする。", "右クリック → 無効。", "名前がグレー。", "削除すると次のデバッグで作り直すことになる。無効で残す。", ""],
        ["2", "ファイル名", "リスナーの Filename が空。", "空欄。", "GUI が勝手に jtl を書かない。", "空でないと、転送先でもそのパスに書こうとして失敗する。", ""],
        ["3", "応答本文", "user.properties で response_data=false。", "08 と 31。", "jtl に本文が無い。", "障害の 1 回だけオンにする場合は、結果の配布範囲を決めてから。", ""],
        ["4", "保存", "無効化した状態で保存する。", "order-search.jmx", "転送するのはこのファイル。", "有効な状態の jmx を別名で残し、転送時に取り違える。", ""],
        ["5", "版", "端末で jmeter.bat -v が 5.6.3。サーバも 5.6.3。", "30 のコマンド。", "両方の文字列を試験条件に書く。", "違う版のまま送る。", ""],
    ]
    r = table(ws, r, before, W7_LIST)
    r += 1
    r = section(ws, r, last, "やってはいけないこと")
    r = note(ws, r, last, "GUI で目標人数を流さない。公式も「Don't run load test using GUI mode」と書いている。\n"
             "同じ result.jtl に 2 回目を追記しない。JMeter は存在する結果ファイルへ追記する。分析が混ざる。\n"
             "思考時間を 0 にした結果と、入れた結果を、同じ合格基準で比べない。条件を併記する。\n"
             "ランプアップ中の遅さを、定常状態の性能として書かない。書くなら「起動中を含む」と明記する。\n"
             "終了コード 0 を合格と書かない。サンプルの失敗ではプロセスが失敗扱いにならない。summary の Err と jtl を見る。\n"
             "応答時間の合否を Duration Assertion だけで取り、エラー率と混同しない。\n"
             "JMeter はブラウザではない。JavaScript を実行しない。画面の操作ではなく、合意した API を打つ。埋め込みリソースの取得をオンにすると、HTML からリンクされた静的ファイルも打つ。API 試験ではオフのまま。\n"
             "配布物の HTTP 実装は、試験のために証明書の検証を広く受け入れる。相手の証明書が不正でも通ることがある。接続できたことを、証明書が正しい証拠にしない。")
    r = section(ws, r, last, "CLI のあとの観点")
    r = headers(ws, r, ["観点", "どこを見る", "問題なし", "問題あり", "次の動作", "", ""])
    after = [
        ["起動した人数", "summary の Active と Started", "Started が -Jthreads と一致し、Active がそれに届いた。", "Active が増えない。", "ランプが長すぎる、スレッドがエラーで止まった、を jmeter.log で見る。", "", ""],
        ["エラー", "summary の Err", "基準以内。0 が理想だが、基準は 26 に書いた値。", "基準超過、または試験中に増え続ける。", "20 の止め方。エラー行を 23 で読む。", "", ""],
        ["応答の悪化", "30 秒ごとの Avg と Max", "定常に入ったあと、右肩上がりが続かない。", "Avg が区間ごとに上がり続ける。", "その回は不合格の材料。見込みで「頭打ち」と書く。", "", ""],
        ["終わり方", "最終の summary =", "Finished が Started に一致。", "Finished が少ないままプロセスが居る。", "shutdown.sh を試し、残れば 28。", "", ""],
        ["ログ", "jmeter.log の ERROR", "ERROR が無い、または説明できる。", "OutOfMemoryError、FileNotFound、Unknown host。", "28 の該当行。", "", ""],
        ["結果の形", "results フォルダ", "jtl、log、run-conditions.txt、html-report/index.html が揃う。", "jtl が 0 バイト。html が無い。", "21。", "", ""],
        ["二重計上", "statistics.json のキー", "報告する名前が 1 セットだけ。", "TX_検索 と HTTP_GET_orders を足して件数にした。", "どちらを報告行にするか 24 で決める。この手順は HTTP 名を明細、必要なら TX を別表。親サンプルはオフ。", "", ""],
    ]
    r = table(ws, r, after, [14, 24, 38, 42, 32, 32, 16])
    finish(ws, last)


def _transfer(wb, images):
    ws = new_sheet(wb, "17_JMXの保存と転送", C_CHAR)
    last = 7
    r = banner(ws, last, "転送", "jmx を保存し、Linux の plans へ置く", "送るのは設計図とデータ。結果ファイルは送らない。")
    r = headers(ws, r, H7)
    steps = [
        ["1", "GUI", "16 の無効化が終わっていることを確認して保存する。", "ファイル → テスト計画に名前をつけて保存。", "C:\\LoadTest\\scenarios\\order-search\\order-search.jmx", "別名保存で古い不良版を上書きしない。日付をフォルダ名に付けて世代を残してよい。", "下の枠"],
        ["2", "フォルダ", "同じフォルダに次が揃っているか見る。", "order-search.jmx、users.csv、load.properties、run.sh", "4 つある。run.sh の改行は LF。", "jmeter.log や前回の jtl を同梱しない。", ""],
        ["3", "run.sh", "31 の全文を run.sh として保存する。Windows 11 のメモ帳なら、保存時の改行コードを LF にする。", "UTF-8、LF。", "サーバで file run.sh が CRLF ではない。", "CRLF のままだと Linux が $'\\r' で失敗する。", ""],
        ["4", "サーバ", "置き場所を作る。SSH ユーザは JMeter を実行するユーザと同じにする。", "ssh loadtest@jmeter-runner.example.internal\nmkdir -p /opt/jmeter/plans/order-search", "ディレクトリができる。", "root で mkdir すると、あとから書けない。", ""],
        ["5", "送る", "PowerShell の scp か WinSCP。中身はテキストなので、転送モードはテキストでもバイナリでも jmx は通常壊れない。CSV は UTF-8 のバイトを維持する。", "scp C:\\LoadTest\\scenarios\\order-search\\* loadtest@jmeter-runner.example.internal:/opt/jmeter/plans/order-search/", "プロンプトが戻る。", "ワイルドカードで results まで送らない。scenarios フォルダの中だけを送る。", "下の枠"],
        ["6", "サーバで確認", "名前、サイズ、改行を見る。", "cd /opt/jmeter/plans/order-search\nls -l\nfile run.sh order-search.jmx users.csv", "jmx と csv のサイズが端末と同じ。run.sh に CRLF の表示が無い。", "0 バイトの jmx は転送失敗。送り直す。", ""],
        ["7", "WinSCP の場合", "左が端末、右がサーバ。左で 4 ファイルを選び、右の order-search へドラッグする。", "ホスト、ポート 22、ユーザ loadtest。", "右ペインに 4 ファイル。", "親フォルダごとドロップして階層が一段ずれる。jmx の相対パスが CSV を見失う。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(ws, r, last, "名前を付けて保存ダイアログ", "jmx を scenarios\\order-search に保存するとき。", "ファイル名 order-search.jmx と、フォルダパスが見える。")
    r = images.add(ws, r, last, "WinSCP または scp の完了", "転送直後。", "サーバ側に jmx、csv、load.properties、run.sh が並び、サイズが見える。結果ファイルは写っていない。")
    r = note(ws, r, last, "load.properties の中のホスト名も、試験対象に合わせてから送る。31 の例は search.host=app.example.internal。シナリオ側を ${__P(search.host,app.example.internal)} にして初期値に使ってもよい。その場合、HTTP リクエスト初期値のサーバ欄をその式にする。式を使わないなら、jmx 内のホストを現場の値に変えて保存し直す。どちらにしたかを試験条件に書く。")
    r = note(ws, r, last, "api-gateway のシナリオは、同じ送り方で 4 ファイルを /opt/jmeter/plans/api-gateway/ へ置く。secrets.properties は chmod 600。run-conditions.txt とスクリーンショットに値を出さない。コマンドは 35 の末尾。")
    finish(ws, last)


def _linux(wb, images):
    ws = new_sheet(wb, "18_Linuxサーバ準備", C_GRAY)
    last = 7
    r = banner(ws, last, "Linux", "実行サーバに同じ JMeter を置く", "初回と、バージョンを上げるときだけ。試験のたびに入れ直さない。")
    r = headers(ws, r, H7)
    steps = [
        ["1", "OS", "実行ユーザを決める。", "ユーザ名の例: loadtest。root では走らせない。", "id で loadtest と出る。", "root で走らせると、結果ファイルが他の人から扱えず、事故時の影響が大きい。", ""],
        ["2", "Java", "Java 17 があるか見る。無ければディストリビューションの手順で 17 の JDK を入れる。", "java -version", "17 以上。", "端末が 17、サーバが 8、でも 5.6.3 は動くことがある。新規は 17 に揃える。", ""],
        ["3", "配置", "tgz を /opt/jmeter に展開する。空白の無いパス。", "sudo mkdir -p /opt/jmeter\nsudo tar -xzf apache-jmeter-5.6.3.tgz -C /opt/jmeter\nsudo chown -R loadtest:loadtest /opt/jmeter", " /opt/jmeter/apache-jmeter-5.6.3/bin/jmeter がある。", "bin という名前のサブディレクトリを改名しない。公式が禁止している。", ""],
        ["4", "実行権", "スクリプトに実行権を付ける。", "chmod +x /opt/jmeter/apache-jmeter-5.6.3/bin/jmeter /opt/jmeter/apache-jmeter-5.6.3/bin/*.sh", "ls で jmeter が実行可能。", "jmeter.bat を Linux で動かそうとしない。", ""],
        ["5", "setenv.sh", "31 の setenv.sh を bin に置く。改行は LF。", "export HEAP と export TZ=Asia/Tokyo。", "ファイルがある。", "jmeter 本体のスクリプトを編集しない。setenv.sh は公式が読む場所。", ""],
        ["6", "user.properties", "31 の追記を、サーバの bin/user.properties の末尾にも入れる。端末と同じ。", "timestamp と date_format の対。", "両方ある。", "サーバだけ書式が ms のままだと、戻した CSV の日時が端末の練習と違う。", ""],
        ["7", "バージョン", "GUI を開かずにバージョンだけ見る。", "/opt/jmeter/apache-jmeter-5.6.3/bin/jmeter -v", "5.6.3 と出て、プロセスが終わる。", "引数なしで jmeter を起動すると GUI を開こうとして、ディスプレイの無いサーバでは失敗する。", "下の枠"],
        ["8", "ディレクトリ", "plans と results を作る。", "mkdir -p /opt/jmeter/plans /opt/jmeter/results", "loadtest が書き込める。", "results を /tmp にすると再起動で消えるサーバがある。", ""],
        ["9", "ファイル数", "同時接続が多い前に、オープンできるファイル数を見る。", "ulimit -n", "65535 など、スレッド数より十分大きい。小さいときは 28。", "既定の 1024 のまま数百スレッドで接続を開くと、Too many open files になる。", ""],
        ["10", "時刻", "端末と時刻帯が説明できること。", "date", "JST と読める。setenv.sh の TZ が Asia/Tokyo。", "サーバが UTC のままログを読むと、9 時間ずれた障害時刻と比べてしまう。run-conditions.txt に date の出力を残すので、それを報告に書く。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(ws, r, last, "java -version と jmeter -v", "準備が終わったサーバの端末画面。", "両方のバージョン文字列と、プロンプトが loadtest であることが読める。")
    r = note(ws, r, last, "複数台で同時に負荷をかける分散実行（jmeter-server と -r）は、この手順書の範囲外。1 台の CLI で目標人数に届き、かつ実行サーバ自身の CPU が先に飽和していないことを見てから検討する。分散にすると、jtl の集約と時刻の扱いが増える。")
    finish(ws, last)


def _cli(wb, images):
    ws = new_sheet(wb, "19_CLIの実行", C_GRAY)
    last = 7
    r = banner(ws, last, "CLI", "jmx・jtl・ログ・HTML の指定を固定して実行する", "公式の形は jmeter -n -t 計画 -l 結果 -e -o レポートディレクトリ。この手順は -j と -q と -J を足す。")
    r = section(ws, r, last, "オプションの意味")
    r = headers(ws, r, ["オプション", "例", "何を指定するか", "省略すると", "注意", "", ""])
    opts = [
        ["-n", "-n", "CLI モード。画面を出さない。", "GUI を開こうとする。", "負荷の実行では必ず付ける。", "", ""],
        ["-t", "-t /opt/jmeter/plans/order-search/order-search.jmx", "シナリオの jmx。", "実行できない。", "相対パスにするなら、今いるディレクトリに注意。絶対パスを推奨。", "", ""],
        ["-l", "-l /opt/jmeter/results/order-search_日時/result.jtl", "サンプル結果の出力先。", "結果ファイルが残らず、再分析できない。", "既存ファイルがあると追記される。毎回新しいパス。", "", ""],
        ["-j", "-j /opt/jmeter/results/order-search_日時/jmeter.log", "JMeter 自身のログ。", "起動したディレクトリに jmeter.log を作り、次回は上書きする。公式は -j で実行ごとに残す、と説明している。", "jtl とは別ファイル。", "", ""],
        ["-q", "-q /opt/jmeter/plans/order-search/load.properties", "そのシナリオだけのプロパティファイル。", "bin/user.properties だけが効く。", "同じキーは -J が -q より優先。", "", ""],
        ["-J", "-Jthreads=50", "プロパティ threads を 50 にする。${__P(threads,1)} が 50 になる。", "式の既定値。threads なら 1。", "-Dthreads=50 は Java のシステムプロパティで、__P からは読めない。", "", ""],
        ["-e", "-e", "終了後に HTML ダッシュボードを作る。", "jtl だけで終わり、HTML は無い。", "jtl の列がダッシュボードの要件を満たすこと。08 の設定で満たす。", "", ""],
        ["-o", "-o /opt/jmeter/results/.../html-report", "HTML の出力先ディレクトリ。", "-e だけでは bin 配下の既定フォルダに出ることがある。", "公式: フォルダは存在しないか、空であること。先に mkdir しない。", "", ""],
        ["-g", "-g result.jtl -o html-report", "試験を再実行せず、既存 jtl から HTML だけ作る。", "使わない。", "作り直しのとき。試験中には使わない。", "", ""],
        ["-f", "-f", "既存の結果ファイルとレポート先を削除してから走る。", "追記や「出力先が空でない」で失敗し得る。", "削除は取り消せない。原則は新フォルダ。-f は同じパスでやり直すときだけ。", "", ""],
        ["-L", "-Ljmeter.engine=DEBUG", "ログの詳細度。", "INFO 相当。", "DEBUG のまま負荷を流すとログが巨大になる。調査の 1 回だけ。", "", ""],
    ]
    r = table(ws, r, opts, [12, 36, 40, 36, 40, 18, 14])
    r += 1
    r = section(ws, r, last, "実行手順")
    r = headers(ws, r, H7)
    steps = [
        ["1", "SSH", "実行ユーザで入る。", "ssh loadtest@jmeter-runner.example.internal", "プロンプトが loadtest。", "別ユーザだと plans が読めない。", ""],
        ["2", "tmux", "セッションを作る。SSH が切れても試験を残すため。", "tmux new -s loadtest", "画面下に tmux の表示。", "tmux 無しで流すと、PC がスリープした瞬間に試験が SIGHUP で死ぬ。", ""],
        ["3", "切離しの練習", "デタッチと復帰を、試験の前に一度やっておく。", "デタッチは Ctrl+b のあとに d。戻るのは tmux attach -t loadtest", "コマンドが続き、プロンプトが戻る。", "Ctrl+c をデタッチと間違えて試験を殺す。", ""],
        ["4", "事前確認", "計画ファイルとユーザを見る。", "test -f /opt/jmeter/plans/order-search/order-search.jmx && echo PLAN_OK\nhead -n 2 /opt/jmeter/plans/order-search/users.csv", "PLAN_OK。1 行目が列名。", "PLAN_OK が出ないまま run.sh を叩かない。", ""],
        ["5", "小さい実行", "run.sh をコピーして、threads=1、ramp=1、duration=60 に変えたもので 1 回流す。", "cp run.sh run-1user.sh して 3 か所を直す。bash run-1user.sh", "summary が出て、結果ディレクトリができる。", "最初から 50 で流さない。", ""],
        ["6", "本番の実行", "合意した値の run.sh を流す。31 の例は 50 人、60 秒で全員起動、600 秒継続。", "bash /opt/jmeter/plans/order-search/run.sh", "すぐにログが流れ、summary + が 30 秒ごとに出る。", "html-report を先に mkdir していると -o が失敗する。run.sh は mkdir していない。", "下の枠"],
        ["7", "場所の記録", "スクリプトが最後に表示する結果ディレクトリを控える。", "/opt/jmeter/results/order-search_年月日時分秒", "そのパスをチケットに書く。", "次の実行のパスと混ぜる。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = note(ws, r, last, "1 行で書くときの形（run.sh と同じ意味）。結果ディレクトリは先に作り、html-report は作らない。\n"
             "/opt/jmeter/apache-jmeter-5.6.3/bin/jmeter -n -t /opt/jmeter/plans/order-search/order-search.jmx -l /opt/jmeter/results/order-search_日時/result.jtl -j /opt/jmeter/results/order-search_日時/jmeter.log -q /opt/jmeter/plans/order-search/load.properties -Jthreads=50 -Jramp=60 -Jduration=600 -e -o /opt/jmeter/results/order-search_日時/html-report",
             C_SOFT, 9)
    r = images.add(ws, r, last, "run.sh を実行した直後のコンソール", "summary が 1 行でも出たところ。", "コマンドの末尾と、summary の 1 行目が見える。パスワードは写さない。")
    r = section(ws, r, last, "プロパティがシナリオへ届いているかの確認")
    r = note(ws, r, last, "1 人の実行の jmeter.log にエラーが無く、summary の Active が 1 なら、threads の既定または -Jthreads=1 が届いている。\n"
             "50 の実行で Active が 1 のままなら、-J のキーが scenarios の式と違う。式が ${__P(threads,1)} なら、キーは threads。${__P(USERS,1)} なら -JUSERS=50。大文字小文字は区別される。\n"
             "GUI のユーザ定義変数に threads という変数を作っても、-Jthreads の代わりにはならない。__P はプロパティを読む。")
    finish(ws, last)


def _watch(wb, images):
    ws = new_sheet(wb, "20_コンソール監視", C_GRAY)
    last = 7
    r = banner(ws, last, "監視", "summary の 1 行を、左から読む", "30 秒ごとに 2 種類出る。+ はその 30 秒。= は開始からの累計。公式のリスナー説明にある summariser。")
    r = note(ws, r, last,
             "例（数値は説明用）:\n"
             "summary +   1500 in 00:00:30 =   50.0/s Avg:   120 Min:    40 Max:  2100 Err:     0 (0.00%) Active: 50 Started: 50 Finished: 0\n"
             "summary =   4500 in 00:01:30 =   50.0/s Avg:   118 Min:    35 Max:  2100 Err:     3 (0.07%) Active: 50 Started: 50 Finished: 0",
             C_SOFT, 9)
    r = headers(ws, r, ["欄", "例", "意味", "監視で見ること", "異常の目安", "", ""])
    rows = [
        ["summary +", "+", "直前の区間。既定 30 秒。", "今起きていること。", "この行だけを最終結果として貼らない。", "", ""],
        ["summary =", "=", "開始から今まで。", "試験全体の現在値。終わりのこの行を控える。", "途中の = を最終値と書く。", "", ""],
        ["件数 in 時間", "1500 in 00:00:30", "その範囲のサンプル数と経過。", "件数が 0 の区間が続かないか。", "0 が続くのは、タイマ、停止、到達不能。", "", ""],
        ["件/秒", "50.0/s", "その範囲のスループット。", "人数を増やしたのに件/秒が増えない、は飽和の候補。", "思考時間を入れた試験は、人数×操作の間隔で頭打ちになる。それが仕様なら異常ではない。", "", ""],
        ["Avg Min Max", "ミリ秒", "elapsed の平均、最小、最大。失敗も含む。", "Avg が区間ごとに上がり続けるか。", "Max だけ大きいのは外れ値。Avg と Max が一緒に上がるのは全体が遅い。", "", ""],
        ["Err", "3 (0.07%)", "失敗数と割合。", "0 のままか、基準を超えたか。", "増えていく途中なら、止める判断を 16 の観点で行う。", "", ""],
        ["Active", "50", "いま動いているスレッド。", "目標人数に届くか。", "Ramp-Up の途中は目標未満で正常。持続時間に入っても未満なら異常。", "", ""],
        ["Started", "50", "起動した累計。", "-Jthreads と一致するか。", "一致しないときは式がプロパティを読んでいない。", "", ""],
        ["Finished", "0", "終わったスレッド。", "試験中は 0 付近。終了時に Started と一致。", "試験の途中で Finished が増えるのは、エラーでスレッドを止める設定か、CSV の終端で止める設定。", "", ""],
    ]
    r = table(ws, r, rows, [14, 22, 36, 40, 42, 20, 14])
    r += 1
    r = images.add(ws, r, last, "summary の + と = が数行出ているコンソール", "試験の途中。", "+ と = の違いが分かるよう、両方を入れる。ウィンドウのタイトルやホスト名も入るとよい。")
    r = section(ws, r, last, "止め方")
    r = headers(ws, r, H7)
    stop = [
        ["1", "様子を見る", "Err が基準以内で、Active が目標どおりなら、持続時間まで手を出さない。", "tmux の画面を残す。", "30 秒ごとに行が増える。", "SSH の窓を閉じない。閉じるなら先にデタッチ。", ""],
        ["2", "穏やかに止める", "別の SSH で同じサーバへ入り、シャットダウンを送る。実行中のサンプルが終わってから止まる。", "/opt/jmeter/apache-jmeter-5.6.3/bin/shutdown.sh", "コンソールに停止のログが出て、最終の summary = が出る。", "jmeter.log が結果フォルダにあるので、bin の jmeter.log だけを見に行かない。", ""],
        ["3", "すぐに止める", "エラーが急増し、対象をこれ以上叩きたくないとき。", "/opt/jmeter/apache-jmeter-5.6.3/bin/stoptest.sh", "プロセスが終わる。", "最後の数秒のサンプルが jtl に無いことがある。run-conditions とログに「stoptest で中断」と書く。", ""],
        ["4", "Ctrl+c", "tmux の中で、上記が効かないときの最終手段。", "Ctrl+c を 1 回。", "プロンプトが戻る。", "連打すると後処理が飛ばない。押した事実を報告に書く。", ""],
        ["5", "終了の記録", "最後の summary = の行を、結果フォルダのメモかチケットへコピーする。", "Err、Avg、Active、Started、Finished。", "後からコンソールを閉じても残る。", "画面を閉じたあとに記憶で数値を書かない。", ""],
    ]
    r = table(ws, r, stop, W7_LIST)
    r += 1
    r = note(ws, r, last, "終了コードが 0 でも Err が 0 とは限らない。逆に、止め方によっては 0 以外で終わっても、それまでの jtl は分析できる。合否は終了コードではなく、サンプルの失敗と応答時間で決める。jmeterengine.stopfail.system.exit は、スレッドを止めきれなかったときの話で、アサーション失敗の話ではない。")
    finish(ws, last)


def _jtl(wb, images):
    ws = new_sheet(wb, "21_JTLとログ", C_GRAY)
    last = 7
    r = banner(ws, last, "出力", "結果フォルダの中身と、各列の意味", "出力先はコマンドが決める。シナリオのリスナーでは決めない。")
    r = section(ws, r, last, "1 回の実行でできるもの")
    r = headers(ws, r, ["パス（例）", "誰が書くか", "中身", "Excel で使うか", "欠けるときの原因", "", ""])
    rows = [
        ["/opt/jmeter/results/order-search_日時/result.jtl", "-l", "1 サンプル 1 行の CSV。拡張子は jtl。", "エラー行の抽出と、時間帯の切り分けに使う。代表値の正本にはしない。", "パスの親ディレクトリが無い。-l を付け忘れた。", "", ""],
        ["同じフォルダ/jmeter.log", "-j", "起動、エラー、summary のログ写し。", "ERROR 行を報告の付録にする。性能表にはしない。", "-j を付けず、bin やカレントに書いて探している。", "", ""],
        ["同じフォルダ/run-conditions.txt", "run.sh", "スレッド数、時間、ホスト、日付。", "26 の試験条件へ転記する。", "run.sh を使わず手入力し、メモを作り忘れた。", "", ""],
        ["同じフォルダ/html-report/index.html", "-e -o", "ダッシュボード。ブラウザで開く。", "グラフの図の元。フォルダごと無いと開いても白紙。", "-o のフォルダが先に存在して空でなかった。", "", ""],
        ["html-report/statistics.json", "同じ生成", "取引ごとの件数、平均、中央値、パーセンタイル、エラー率、スループット。", "報告の数値の正本。23 で Excel に読む。", "HTML 生成が失敗すると無い。jmeter.log の reportgenerator を見る。", "", ""],
    ]
    r = table(ws, r, rows, [36, 14, 36, 40, 36, 16, 14])
    r += 1
    r = images.add(ws, r, last, "結果ディレクトリの ls -l", "試験が終わった直後。", "result.jtl、jmeter.log、run-conditions.txt、html-report が 1 画面に入る。サイズが 0 でないこと。")
    r = images.add(ws, r, last, "html-report/index.html の Statistics 表", "ブラウザで index.html を開いたところ。", "Transactions、Samples、Error %、Average、90th、95th、99th、Throughput が見える。")
    r = section(ws, r, last, "result.jtl の列（この手順の user.properties）")
    r = headers(ws, r, ["列", "例", "意味", "単位", "分析での使い方", "注意", ""])
    cols = [
        ["timeStamp", "2026/10/04 15:00:01.234", "サンプルが終わった時刻。書式は user.properties。", "日時", "定常区間の切り出し。", "date_format と不一致だと HTML の軸がずれる。", ""],
        ["elapsed", "256", "所要時間。送信開始から応答を読み終わるまで。", "ミリ秒", "応答時間の本体。", "Latency と足さない。", ""],
        ["label", "HTTP_GET_orders", "要素の名前。", "文字", "報告の行のキー。", "同名は区別できない。", ""],
        ["responseCode", "200", "HTTP のコード。接続失敗時は Non HTTP response code で始まる長い文字。", "文字として扱う", "エラー内訳。", "Excel が 200 を数値にすると、文字コードの行と型が混ざる。23 で文字にする。", ""],
        ["responseMessage", "OK", "理由の短い文。", "文字", "エラー内訳。", "成功行はほぼ OK。", ""],
        ["threadName", "TG_注文検索 1-1", "スレッドグループ名と番号。", "文字", "特定ユーザの追跡。", "個人名ではない。", ""],
        ["success", "true", "アサーションを含めた成否。", "true/false", "エラー率の分子は false。", "HTTP 200 でも false がある。", ""],
        ["failureMessage", "(空またはアサーション文)", "失敗の理由。assertion_results_failure_message=true のとき。", "文字", "エラー分類。", "成功行は空。", ""],
        ["bytes", "1024", "受信バイト。", "バイト", "転送量。", "本文を保存していなくてもサイズは入る。", ""],
        ["sentBytes", "512", "送信バイト。", "バイト", "転送量。", "設定を false にすると列が欠ける。", ""],
        ["grpThreads", "50", "そのグループの動いているスレッド数。", "人", "その時点の負荷の目安。", "allThreads と、グループが 1 つなら同じ。", ""],
        ["allThreads", "50", "全グループの合計。", "人", "同上。", "setUp の分が混じる時間帯がある。", ""],
        ["Latency", "180", "最初の 1 バイトを受け取るまで。", "ミリ秒", "サーバ処理と接続の切り分け。", "elapsed より大きくはならない。", ""],
        ["Connect", "20", "接続確立まで。", "ミリ秒", "ここだけ大きいならネットワークや接続待ち。", "Keep-Alive で 0 に近い行が混ざる。", ""],
        ["URL", "https://...", "url=true のとき。", "文字", "打ち先の確認。", "クエリに秘密が付く設計なら、この列の配布に注意。", ""],
    ]
    r = table(ws, r, cols, [16, 28, 40, 18, 32, 36, 14])
    r += 1
    r = note(ws, r, last, "HTML ダッシュボードが求める列は、公式の Generating Dashboard にある。bytes、label、Latency、responseCode、responseMessage、success、thread_counts、threadName、elapsed、Connect、失敗メッセージ。時刻書式は時間を含める。この手順の 31 はその既定に、Excel で読める日時と date_format の対を足したもの。response_data は公式のダッシュボード要件ではないので false のままにする。")
    finish(ws, last)


def _back(wb, images):
    ws = new_sheet(wb, "22_結果を端末へ", C_CHAR)
    last = 7
    r = banner(ws, last, "回収", "その回の結果フォルダを、丸ごと端末へ戻す", "index.html だけを戻すと、グラフが欠ける。")
    r = headers(ws, r, H7)
    steps = [
        ["1", "Linux", "戻すパスを 1 つに決める。試験中のフォルダと、前の回を混ぜない。", "ls -d /opt/jmeter/results/order-search_*", "時刻入りのフォルダが並ぶ。", "最新と思ったフォルダが、失敗した試し実行であることがある。run-conditions.txt の threads を見てから決める。", ""],
        ["2", "Linux", "jtl の行数の当たりを見る。0 や 1 ではないこと。", "wc -l /opt/jmeter/results/order-search_日時/result.jtl", "ヘッダ 1 行 + サンプル行。", "数万行を cat で開かない。", ""],
        ["3", "端末", "受け皿を作る。", "mkdir C:\\LoadTest\\results は初回だけ。", "フォルダがある。", "scenarios フォルダへ結果を混ぜない。", ""],
        ["4", "端末 PowerShell", "フォルダごとコピーする。", "scp -r loadtest@jmeter-runner.example.internal:/opt/jmeter/results/order-search_20261004_150000 C:\\LoadTest\\results\\", "プロンプトが戻る。", "途中で切れたら、そのフォルダを削除して取り直す。中途半端な html-report は真っ白になる。", ""],
        ["5", "WinSCP", "右の results から、その回のフォルダを左へドラッグする。", "フォルダ単位。", "左に html-report まで見える。", "result.jtl だけドラッグして json を忘れる。", "下の枠"],
        ["6", "端末", "中身を確認する。", "result.jtl、jmeter.log、run-conditions.txt、html-report\\index.html、html-report\\statistics.json", "5 つがある。", "jtl をダブルクリックして Excel が直接開こうとする。数十万行で固まる。23 の手順で開く。", ""],
        ["7", "ブラウザ", "index.html を開いて、図が描かれているか見る。", "html-report\\index.html をダブルクリック。", "Statistics に行がある。", "file をメールに添付して単体で送ると、相手のブラウザではグラフが出ない。フォルダごと渡すか、必要な図だけをこの手順書の枠へ貼る。", "下の枠"],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(ws, r, last, "端末の results フォルダ", "scp または WinSCP のあと。", "1 回分のフォルダが開いていて、jtl、log、txt、html-report が見える。")
    r = images.add(ws, r, last, "ブラウザのダッシュボード先頭", "index.html を開いた直後。", "タイトルと、エラーの有無が分かる要約。個人情報のリクエスト URL が拡大されていれば隠す。")
    finish(ws, last)


def _excel_in(wb, images):
    ws = new_sheet(wb, "23_Excelへ読込", C_CHAR)
    last = 7
    r = banner(ws, last, "Excel", "正本は statistics.json。jtl はエラーと切り分けに使う", "jtl をダブルクリックしない。データ量で Excel が固まる。")
    r = section(ws, r, last, "A. statistics.json を読む（報告の数値）")
    r = headers(ws, r, H7)
    steps = [
        ["1", "Excel", "空のブックではなく、26 のひな形があるこのブックのコピーを作る。", "ファイル → 名前を付けて保存。試験名をファイル名にする。", "手順書の原本は上書きしない。", "原本の 26 に本番数値を入れて、次の試験と混ぜる。", ""],
        ["2", "データ", "JSON を取り込む。リボンの文言は版で少し違う。無いときはリボンの検索で「JSON」を探す。", "データ → データの取得 → ファイルから → JSON から。", "statistics.json を選べる。", "index.html を選ばない。", "下の枠"],
        ["3", "Power Query", "レコードの一覧になっているので、テーブルへ変換し、列を展開する。", "変換 → テーブルへ変換 → 展開。", "transaction、sampleCount、errorCount、errorPct、meanResTime、medianResTime、pct1ResTime、pct2ResTime、pct3ResTime、throughput の列が見える。", "Total 行と各 label が混在する。Total は合計であって、平均の平均ではない。報告では取引行を使い、Total は参考と書く。", ""],
        ["4", "型", "時間の列を小数、件数を整数、errorPct を数値にする。", "列の型の変更。", "pct2ResTime が 95 パーセンタイル（既定）。", "pct1 を 95% だと思い込む。既定は 90、95、99 の順。user.properties で変えていれば、変えた値を報告に書く。", ""],
        ["5", "読込", "閉じて読み込む。新しいシートに出してよい。", "閉じて読み込む。", "シートに表ができる。", "この表の数値を、手計算の別定義で上書きしない。", ""],
        ["6", "転記", "報告に使う行だけ、26 の記入欄へ値として貼る。", "値の貼り付け。数式の貼り付けではない。", "26 のサンプル数と平均が入る。", "Power Query のシートを消すと、更新できない。残す。", ""],
    ]
    r = table(ws, r, steps, W7_LIST)
    r += 1
    r = images.add(ws, r, last, "JSON の取り込みで statistics.json を選んだ画面", "ファイル選択ダイアログ。", "選んでいるファイル名が statistics.json と読める。")
    r = section(ws, r, last, "B. エラー行だけ jtl から読む")
    r = headers(ws, r, H7)
    b = [
        ["1", "データ", "テキストまたは CSV から、result.jtl を選ぶ。", "データ → テキストまたは CSV から。", "プレビューが出る。", "ここで「読み込む」をすぐ押して全行をシートに置かない。", "下の枠"],
        ["2", "プレビュー", "元のファイルを 65001 Unicode (UTF-8)、区切り記号をコンマにする。", "データの変換 を押して Power Query へ進む。", "1 行目が列名。", "Shift_JIS のまま開くと日本語の label が壊れる。この手順の jtl は ASCII 中心だが、UTF-8 で開く。", ""],
        ["3", "型", "responseCode をテキストにする。timeStamp を日時、elapsed を整数、success を true/false にする。", "列をクリック → データ型。", "200 が文字列 200。", "数値の 200 と、Non HTTP response code の文字列が同じ列でエラーになる。テキストにすれば両方残る。", ""],
        ["4", "絞る", "success が false の行だけにする。", "success 列のフィルタ → false。", "失敗行だけが残る。", "true を消すフィルタを間違えて、成功だけをエラー分析と称する。", ""],
        ["5", "件数を見る", "フィルタ後の行数が、Excel に置ける量か見る。数万行の失敗なら、responseCode でグループ化してから読み込む。", "ホーム → グループ化。列 responseCode と label。件数。", "内訳の表になる。", "百万行の失敗をシートに置いて、ファイルが開けなくなる。", ""],
        ["6", "読み込む", "閉じて読み込む。", "新しいシート。名前を エラー内訳 にする。", "26 から参照できる。", "全件 jtl のシートを報告ブックに残して、開くたびに数分待つ。", ""],
    ]
    r = table(ws, r, b, W7_LIST)
    r += 1
    r = images.add(ws, r, last, "CSV 取り込みのプレビュー", "区切り記号をコンマにしたあと。", "列名 timeStamp、elapsed、label、responseCode、success が見える。データの中身の個人情報は写さない。")
    r = section(ws, r, last, "C. 全行をシートに置いてよいとき")
    r = note(ws, r, last, "目安は、サンプル数がおおよそ 20 万行までで、PC のメモリに余裕があるとき。それを超える場合は A と B だけにする。全行を置く場合も、ダブルクリックではなく「テキストまたは CSV から」を使う。置いたあとは 27 の数式と同じ考え方でピボットテーブルを作る。ピボットの値フィールド: サンプル数は行数、平均は elapsed の平均、エラーは success が false の件数。パーセンタイルはピボット標準機能には無いので、代表値は statistics.json の pct2ResTime を使う。")
    r = note(ws, r, last, "Cognito の setUp で出る HTTP_POST_cognito_token は 1 件の準備。26 の取引表には業務の label だけを値で貼る。トークン取得は条件欄に、1 回、応答 200、expires_in が持続秒より長い、と書く。34 の末尾。")
    finish(ws, last)


def _analyze(wb, images):
    ws = new_sheet(wb, "24_分析の手順", C_CHAR)
    last = 7
    r = banner(ws, last, "分析", "数値の定義を揃えてから、文章を書く", "平均だけを並べて合格にしない。")
    r = headers(ws, r, ["順", "やること", "使うもの", "書き方", "落とし穴", "", ""])
    rows = [
        ["1", "条件を写す。", "run-conditions.txt と jmeter -v の記録。", "人数、Ramp-Up 秒、持続秒、思考時間の有無、JMeter と Java の版、対象ホスト。", "記憶で人数を書く。", "", ""],
        ["2", "報告する名前を決める。", "statistics.json のキー。", "HTTP_POST_login と HTTP_GET_orders を明細にする。親サンプルはオフなので TX_検索 は行として出ていない想定。", "コントローラ名とリクエスト名を足す。", "", ""],
        ["3", "件数とエラー率を書く。", "sampleCount、errorCount、errorPct。", "エラー率の隣に件数を書く。0.1% でも母数が 100 なら失敗は 0 件の丸めと区別が付かないことがあるので、件数を併記する。", "割合だけを書く。", "", ""],
        ["4", "応答時間を 4 つ書く。", "mean、median、pct2（95）、max。", "単位はミリ秒。平均と 95% の両方。", "平均だけが基準内なので合格、とする。", "", ""],
        ["5", "スループットを書く。", "throughput。単位は件/秒。", "「その label の全サンプルを、その label の最初から最後までの時間で割った値」と注釈を 1 行付ける。", "スレッド数をスループットと書く。", "", ""],
        ["6", "区間を言葉にする。", "ダッシュボードの Response Times Over Time と、必要なら jtl の時刻。", "Ramp-Up の 60 秒を含む値であることを書く。定常だけを採るときは、開始 60 秒を除いた再集計だと書く。", "黙って定常だけを採り、条件の持続時間と母数を一致させない。", "", ""],
        ["7", "エラーを分類する。", "エラー行の responseCode と failureMessage。", "接続失敗、HTTP 5xx、HTTP 4xx、アサーション、の 4 つに数を分ける。", "全部を「サーバが遅い」と書く。接続失敗は届いていない。", "", ""],
        ["8", "遅さの内訳を見る。", "Connect、Latency、elapsed。", "Connect が大きいなら経路。Latency が elapsed の大半なら相手の処理。elapsed だけが大きく Latency が小さいなら、応答のダウンロードに時間がかかっている。", "3 つを足して報告する。elapsed はすでに全体。", "", ""],
        ["9", "合否を行ごとに付ける。", "26 の基準セル。", "1 行でも基準超過なら、その取引は不合格。総合は最悪の行に合わせる。", "総合の平均が基準内なので全取引合格、とする。", "", ""],
        ["10", "図を 2 枚貼る。", "ダッシュボードの全体と、問題のあった取引の時間推移。", "26 の画像枠。図の下に、何の試験の何のグラフかを書く。", "装飾のグラフを増やして、基準との対応を書かない。", "", ""],
    ]
    r = table(ws, r, rows, [10, 28, 36, 42, 40, 20, 16])
    r += 1
    r = section(ws, r, last, "時間の読み分け")
    r = note(ws, r, last, "elapsed = 接続 + サーバが最初のバイトを返すまで + 残りのボディを読み終わるまで、をまとめたもの。内訳を足し算で再計算して elapsed と比べる必要はない。レポートでは elapsed を応答時間と呼び、内訳を見るときだけ Connect と Latency の語を使う。\n"
             "パーセンタイルは「速い方から数えて、その割合の位置」。95% 値が 800 ミリ秒なら、サンプルの 95% は 800 ミリ秒以内に終わり、5% はそれより遅い。平均が 200 でも 95% が 800 なら、遅い側がいる。合格基準を平均だけに置かない理由がこれ。\n"
             "ダッシュボードのパーセンタイルと、GUI の統計レポートのパーセンタイルは、件数が少ないと計算差が出る。公式の Generating Dashboard も、分布が広いと差が見える、と書いている。提出値は statistics.json に揃える。27 の手計算は、定義を理解するための小さい例で、提出値の置き換えではない。")
    r = images.add(ws, r, last, "Response Times Over Time のグラフ", "HTML レポートから、報告に使う取引だけを表示した状態。", "縦軸がミリ秒、横軸が時刻、凡例に label 名。ランプアップの区間が目で分かる。")
    r = note(ws, r, last, "認証つきの試験では、グラフに出す label も 26 に書く label も業務 API だけにする。HTTP_POST_cognito_token は 1 件なので、応答時間の推移にも合格判定にも入れない。401 と 403 は 37 と 39 で分けてから、エラー内訳の件数にする。")
    finish(ws, last)


def _judge(wb, images):
    ws = new_sheet(wb, "25_判定と見込み", C_DARK)
    last = 7
    r = banner(ws, last, "報告", "合格の書き方と、見込みを書いてよい範囲", "見込みは、測っていない人数を測ったように書く欄ではない。")
    r = section(ws, r, last, "判定のルール（26 の数式もこれ）")
    r = headers(ws, r, ["ルール", "内容", "例", "書いてはいけない例", "", "", ""])
    rules = [
        ["取引ごと", "サンプル数、エラー率、平均、95% 値を、その取引の基準と比べる。", "検索の 95% が 3000ms 以下、かつエラー率 1% 未満。", "全取引を混ぜた平均だけで検索が合格。", "", "", ""],
        ["総合", "取引が 1 つでも不合格なら総合不合格。全部合格なら総合合格。", "ログイン合格、検索不合格 → 総合不合格。", "重要でないから検索の不合格を総合では合格。", "", "", ""],
        ["母数", "件数を必ず書く。", "失敗 3 / 12000。", "エラー率 0%（件数を書かず、丸め）。", "", "", ""],
        ["区間", "ランプアップを含むかどうかを書く。", "持続 600 秒、うち起動 60 秒を含む。", "黙って後ろ 5 分だけを採って条件と違う母数にする。採るならその旨を書く。", "", "", ""],
        ["中断", "止めた試験は「中断」とし、中断時点の値だと書く。", "Err が 5% を超えたので stoptest。", "中断の数値を、最後まで流した試験と同列に比較する。", "", "", ""],
        ["環境", "1 台の実行サーバからの結果だと書く。", "実行サーバの CPU が先に飽和していないことを、可能なら別途確認した、と書く。確認していなければ「未確認」。", "JMeter 側の CPU を見ずに、相手サーバの限界だと断定する。", "", "", ""],
    ]
    r = table(ws, r, rules, [14, 40, 40, 42, 20, 16, 16])
    r += 1
    r = section(ws, r, last, "見込みに書いてよいこと")
    r = note(ws, r, last, "書いてよいのは、今回の実測の内側だけ。\n"
             "1) 目標人数のとき、95% 値が基準の何割だったか。基準 3000ms に対して 1500ms なら、応答時間の基準に対して半分を使った、と書く。\n"
             "2) エラーが 0 のまま、応答時間が区間ごとに上がり続けていなければ、「この人数では頭打ちの兆候は、このグラフの範囲では見えない」と書く。\n"
             "3) 次に試す人数は、今回の 1.5 倍など、一度に大きく飛ばさない値を提案する。提案であって、耐えることの断言ではない。\n"
             "4) 応答時間が人数に対して曲がり始めた、エラーが出始めた、件/秒が増えなくなった、のどれかがあれば、その人数を「今回見えた限界の兆候」と書く。それより上は未測定と書く。\n\n"
             "書いてはいけないのは、直線で引き延ばした結果。「50 人で 95% が 800ms だから 200 人でも 3200ms」は、待ち行列で遅れが曲がるので根拠にならない。測定していない 200 人を合格とも不合格とも書かない。")
    r = section(ws, r, last, "所見の文例")
    r = headers(ws, r, ["状況", "書き出しの例", "続けて書く事実", "見込みの文例", "", "", ""])
    samples = [
        ["基準内で頭打ち無し", "目標 50 人、持続 600 秒（起動 60 秒を含む）で、検索の 95% 値は基準 3000ms に対して 900ms、エラーは 0/件数。", "Response Times Over Time で、起動後に右肩上がりが続いていない。", "この条件の内側では余力がある。次は 75 人で同じ持続時間を測る。75 人より上は未測定。", "", "", ""],
        ["エラーは無いが 95% が超過", "エラーは 0 だが、検索の 95% 値が 4200ms で基準 3000ms を超えた。総合不合格。", "平均は 1000ms で、平均だけ見ると基準内に見える。", "平均を基準に替えない。遅い側の 5% を減らす対処のあと、同じ 50 人を再測定する。人数を増やす見込みは書かない。", "", "", ""],
        ["途中からエラーが増えた", "開始 4 分後から HTTP 500 が増え、累計エラー率が基準を超えたので shutdown した。", "それ以前の 95% 値は参考。試験としては中断。", "4 分時点の件数までは動作し、その後は不合格。連続 600 秒の耐性は未確認。", "", "", ""],
        ["件/秒が人数に比例しない", "20 人から 50 人へ増やしても、検索の件/秒が 30 から 32 へしか増えなかった。", "思考時間は 1〜3 秒で、人数の割に件/秒の上限が低い設計ではない。", "30 件/秒付近に飽和の兆候がある。次は 20 と 50 の間を刻むか、相手サーバの CPU を同時に見る。倍の人数は提案しない。", "", "", ""],
    ]
    r = table(ws, r, samples, [18, 42, 40, 42, 16, 14, 14])
    finish(ws, last)


def _report(wb, images):
    ws = new_sheet(wb, "26_レポートひな形", C_BLACK, {"A": 28, "B": 22, "C": 18, "D": 18, "E": 18, "F": 18, "G": 36})
    last = 7
    r = banner(ws, last, "提出用", "負荷試験の結果と見込み（記入用紙）", "グレーのセルが記入欄。白い数値の判定は数式。取引を 1 行でも不合格にすると総合が不合格になる。")
    r = note(ws, r, last, "このシートを試験ごとにコピーして記入する。手順書の原本には本番の数値を残さない。基準の例（エラー率 1%、平均 1000ms、95% 3000ms）は仮の値なので、合意した値に変える。変えると判定列が再計算される。", C_SOFT, 10)

    r = section(ws, r, last, "1. 文書")
    labels = [
        ("試験名", "order-search の例を消して記入"),
        ("実施日", "yyyy-mm-dd"),
        ("実施者", ""),
        ("確認者", ""),
        ("対象環境", "試験用 / ステージングなど。本番と書かないこと"),
        ("対象ホスト", "合意したホスト名"),
        ("JMeter", "5.6.3（jmeter -v の文字列）"),
        ("Java", "実行サーバの java -version"),
        ("実行サーバ", "ホスト名"),
        ("結果フォルダ", "order-search_年月日時分秒"),
    ]
    # two columns of pairs
    for i in range(0, len(labels), 2):
        left = labels[i]
        right = labels[i + 1] if i + 1 < len(labels) else ("", "")
        fill_l = FILL("E8E8E8")
        fill_in = FILL(C_EMPH)
        put(ws, r, 1, left[0], F(9, True), fill_l, AL("left", "center"), height=22)
        merge(ws, f"B{r}:C{r}", "", F(10), fill_in, AL("left", "center"))
        put(ws, r, 4, right[0], F(9, True), fill_l, AL("left", "center"))
        merge(ws, f"E{r}:G{r}", "", F(10), fill_in, AL("left", "center"))
        r += 1

    r += 1
    r = section(ws, r, last, "2. 目的と合格基準")
    put(ws, r, 1, "目的", F(9, True), FILL("E8E8E8"), AL("left", "top"), height=48)
    merge(ws, f"B{r}:G{r}", "例: 注文検索 API が、合意した同時ユーザ数で、エラー率と 95% 応答時間の基準を満たすかを確認する。", F(9), FILL(C_EMPH), AL("left", "top"))
    r += 1
    put(ws, r, 1, "エラー率の上限", F(9, True), FILL("E8E8E8"), AL("left", "center"), height=22)
    cell_err = put(ws, r, 2, 0.01, F(10, True), FILL(C_EMPH), AL("center", "center"), num_fmt="0.00%")
    err_ref = f"$B${r}"
    merge(ws, f"C{r}:G{r}", "例として 1% を入れてある。失敗数÷サンプル数。このセルを変えると下の判定が変わる。", F(9), FILL(C_WHITE), AL("left", "center"))
    r += 1
    put(ws, r, 1, "平均の上限（ミリ秒）", F(9, True), FILL("E8E8E8"), AL("left", "center"), height=22)
    put(ws, r, 2, 1000, F(10, True), FILL(C_EMPH), AL("center", "center"), num_fmt="#,##0")
    avg_ref = f"$B${r}"
    merge(ws, f"C{r}:G{r}", "elapsed の平均。失敗サンプルを含めた statistics.json の meanResTime を使う。", F(9), FILL(C_WHITE), AL("left", "center"))
    r += 1
    put(ws, r, 1, "95%値の上限（ミリ秒）", F(9, True), FILL("E8E8E8"), AL("left", "center"), height=22)
    put(ws, r, 2, 3000, F(10, True), FILL(C_EMPH), AL("center", "center"), num_fmt="#,##0")
    p95_ref = f"$B${r}"
    merge(ws, f"C{r}:G{r}", "statistics.json の pct2ResTime（既定で 95）。pct1 は 90 なので取り違えない。", F(9), FILL(C_WHITE), AL("left", "center"))
    r += 2

    r = section(ws, r, last, "3. 試験条件（run-conditions.txt から転記）")
    cond = ["同時ユーザ数（-Jthreads）", "Ramp-Up 秒", "持続秒", "思考時間", "親サンプル", "区間の注記"]
    cond_hint = ["50", "60", "600", "検索前に 1〜3 秒", "オフ", "起動 60 秒を含む"]
    for name, hint in zip(cond, cond_hint):
        put(ws, r, 1, name, F(9, True), FILL("E8E8E8"), AL("left", "center"), height=22)
        merge(ws, f"B{r}:D{r}", "", F(10), FILL(C_EMPH), AL("left", "center"))
        merge(ws, f"E{r}:G{r}", "記入例: " + hint, F(8, False, C_GRAY), FILL(C_WHITE), AL("left", "center"))
        r += 1
    r += 1

    r = section(ws, r, last, "4. 取引ごとの結果（statistics.json から値で貼る）")
    r = note(ws, r, last, "A〜F が記入。G が判定。サンプル数が空の行は判定しない。エラー率は C÷B、平均は D、95% は E を、上の基準と比べる。", C_WHITE, 9)
    header_at = r
    heads = ["トランザクション", "サンプル数", "エラー数", "平均ms", "95%ms", "件/秒", "判定"]
    for c, h in enumerate(heads, 1):
        put(ws, r, c, h, F(9, True, C_WHITE), FILL(C_BLACK), AL("center", "center"), height=22)
    r += 1
    first_tx = r
    for _ in range(8):
        put(ws, r, 1, "", F(9), FILL(C_EMPH), AL("left", "center"), height=20)
        for c, fmt in [(2, "#,##0"), (3, "#,##0"), (4, "#,##0.0"), (5, "#,##0.0"), (6, "0.00")]:
            put(ws, r, c, None, F(9), FILL(C_EMPH), AL("center", "center"), num_fmt=fmt)
        formula = (
            f'=IF(OR(A{r}="",B{r}="",B{r}=0),"",'
            f'IF(AND(C{r}/B{r}<={err_ref},D{r}<={avg_ref},E{r}<={p95_ref}),"合格","不合格"))'
        )
        put(ws, r, 7, formula, F(9, True), FILL(C_WHITE), AL("center", "center"))
        r += 1
    last_tx = r - 1
    overall_row = r
    put(ws, r, 1, "総合", F(10, True, C_WHITE), FILL(C_DARK), AL("left", "center"), height=24)
    for c in range(2, 7):
        put(ws, r, c, "", F(9), FILL(C_DARK), AL("center", "center"))
    overall = (
        f'=IF(COUNTIF(G{first_tx}:G{last_tx},"不合格")>0,"不合格",'
        f'IF(COUNTIF(G{first_tx}:G{last_tx},"合格")=0,"","合格"))'
    )
    put(ws, r, 7, overall, F(12, True, C_WHITE), FILL(C_DARK), AL("center", "center"))
    r += 2

    r = section(ws, r, last, "5. エラー内訳と所見")
    blocks = [
        ("エラー内訳", "接続失敗、HTTP 4xx、HTTP 5xx、アサーション失敗、の件数。0 なら 0 と書く。"),
        ("応答時間の所見", "平均と 95% の差、時間とともに悪化したか。グラフの番号を書く。"),
        ("中断の有無", "最後まで流した / shutdown / stoptest。時刻と、そのときの Err。"),
        ("総合の文章", "総合セルが合格か不合格かと、どの取引が理由かを 3 行以内で書く。"),
    ]
    for name, hint in blocks:
        put(ws, r, 1, name, F(9, True), FILL("E8E8E8"), AL("left", "top"), height=36)
        merge(ws, f"B{r}:G{r}", hint, F(9, False, C_GRAY), FILL(C_EMPH), AL("left", "top"))
        ws.row_dimensions[r].height = 48
        r += 1
    r += 1
    r = section(ws, r, last, "6. 見込み（測った範囲の内側だけ）")
    outlook = [
        ("今回の人数で、95% 値は基準の何割か", "例: 基準 3000ms に対して検索は 900ms で、約 0.3 倍を使用。"),
        ("頭打ち、またはエラーが出始めた人数", "見えなければ「この試験の範囲では見えない」と書く。"),
        ("次に測る人数", "今回の 1.5 倍など。断言ではなく次の試験の提案。"),
        ("書いていない範囲", "例: 75 人より上は未測定。直線で 200 人を推定していない。"),
    ]
    for name, hint in outlook:
        put(ws, r, 1, name, F(9, True), FILL("E8E8E8"), AL("left", "top"), height=36)
        merge(ws, f"B{r}:G{r}", "", F(10), FILL(C_EMPH), AL("left", "top"))
        ws.row_dimensions[r].height = 36
        # hint goes to comment via adjacent? Keep hint in the cell as gray placeholder people overwrite.
        ws.cell(r, 2).value = hint
        ws.cell(r, 2).font = F(9, False, C_GRAY)
        r += 1
    r += 1
    r = section(ws, r, last, "7. 図（ダッシュボードから貼る）")
    r = images.add(ws, r, last, "提出用: Statistics の表、またはそのスクリーンショット", "数値を転記した同じ実行の HTML レポート。", "試験名と日時が報告の結果フォルダと一致すると分かる範囲。")
    r = images.add(ws, r, last, "提出用: 応答時間の時間推移", "問題の有無が目で分かる取引を表示したグラフ。", "縦軸ミリ秒、横軸時刻、凡例。")
    r = images.add(ws, r, last, "提出用: エラーまたは件/秒のグラフ（無ければ枠に「該当なし」と書く）", "エラーが 0 で件/秒も安定なら、その旨を枠の文字として残し、図は省略してよい。", "省略するときは図を置かず、理由をこの枠の文に残す。")
    r = section(ws, r, last, "8. 承認")
    for name in ["作成日", "作成者", "確認日", "確認者"]:
        put(ws, r, 1, name, F(9, True), FILL("E8E8E8"), AL("left", "center"), height=22)
        merge(ws, f"B{r}:D{r}", "", F(10), FILL(C_EMPH), AL("left", "center"))
        merge(ws, f"E{r}:G{r}", "", F(9), FILL(C_WHITE), AL("left", "center"))
        r += 1
    # silence unused
    _ = (cell_err, header_at, overall_row, C_ZEBRA)
    finish(ws, last)


def _calc(wb, images):
    ws = new_sheet(wb, "27_計算例", C_CHAR, {
        "A": 22, "B": 12, "C": 20, "D": 14, "E": 16, "F": 18, "G": 12,
        "H": 12, "I": 12, "J": 12, "K": 12, "L": 16,
    })
    last = 12
    # style_sheet only set A-G if we passed 12 keys. We passed 12. Good.
    r = banner(ws, last, "手計算", "小さい結果で、平均・95%・件/秒の意味を確認する", "この表は説明用の 23 行。提出用の数値は statistics.json。数式は Excel 2021 または Microsoft 365 の FILTER を使う。")
    r = note(ws, r, last, "下の生データは架空。02_検索の 500 ミリ秒の行だけ success が false。平均は失敗を含む。右端の「成功のみ平均」は失敗を除いた値で、提出の meanResTime とは違う。違いを見るための列。", C_SOFT, 9)

    headers_raw = ["timeStamp", "elapsed", "label", "responseCode", "responseMessage", "threadName", "success", "bytes", "sentBytes", "Latency", "Connect", ""]
    # banner returned 6, note used one row. Put a section then headers.
    r = section(ws, r, last, "架空の jtl（23 行）")
    raw_header = r
    for c, h in enumerate(["timeStamp", "elapsed", "label", "responseCode", "responseMessage", "threadName", "success", "bytes", "sentBytes", "Latency", "Connect"], 1):
        put(ws, r, c, h, F(8, True, C_WHITE), FILL(C_BLACK), AL("center", "center"), height=20)
    put(ws, r, 12, "", F(8), FILL(C_BLACK), AL(), height=20)
    r += 1
    data_start = r

    login = [120, 140, 150, 160, 180, 190, 210, 400]
    search = [200, 220, 240, 260, 300, 350, 500, 800, 1200]
    detail = [80, 90, 100, 110, 130, 140]
    rows_spec = []
    t = datetime(2026, 10, 4, 15, 0, 0)
    for i, ms in enumerate(login):
        rows_spec.append((t + timedelta(seconds=3 * i), ms, "01_ログイン", "200", "OK", True, 40))
    t = datetime(2026, 10, 4, 15, 0, 30)
    for i, ms in enumerate(search):
        ok = ms != 500
        rows_spec.append((t + timedelta(seconds=4 * i), ms, "02_検索", "200" if ok else "500", "OK" if ok else "Server Error", ok, 80 if ok else 120))
    t = datetime(2026, 10, 4, 15, 2, 0)
    for i, ms in enumerate(detail):
        rows_spec.append((t + timedelta(seconds=2 * i), ms, "03_詳細", "200", "OK", True, 30))

    for i, (ts, ms, label, code, msg, ok, lat_delta) in enumerate(rows_spec):
        fill = FILL(C_WHITE if i % 2 == 0 else C_ZEBRA)
        put(ws, r, 1, ts, F(8), fill, AL("center", "center"), num_fmt="yyyy/mm/dd hh:mm:ss")
        put(ws, r, 2, ms, F(8), fill, AL("right", "center"), num_fmt="#,##0")
        put(ws, r, 3, label, F(8), fill, AL("left", "center"))
        put(ws, r, 4, code, F(8), fill, AL("center", "center"))
        put(ws, r, 5, msg, F(8), fill, AL("left", "center"))
        put(ws, r, 6, f"TG 1-{i+1}", F(8), fill, AL("left", "center"))
        put(ws, r, 7, ok, F(8), fill, AL("center", "center"))
        put(ws, r, 8, 800, F(8), fill, AL("right", "center"))
        put(ws, r, 9, 200, F(8), fill, AL("right", "center"))
        put(ws, r, 10, max(ms - 30, 1), F(8), fill, AL("right", "center"))
        put(ws, r, 11, 10, F(8), fill, AL("right", "center"))
        put(ws, r, 12, "", F(8), fill, AL())
        ws.row_dimensions[r].height = 18
        r += 1
    data_end = r - 1
    r += 1
    r = section(ws, r, last, "ラベルごとの計算（数式）")
    r = note(ws, r, last, "95% は PERCENTILE.INC。位置は (件数-1)×0.95。件数 9 の検索は、遅い方から数えて 800 と 1200 のあいだになり 1040。平均は失敗の 500ms を含むので 452.2。成功だけなら 446.3。件/秒は、そのラベルの最終時刻と最初時刻の差（秒）で件数を割る。", C_WHITE, 9)
    sum_header = r
    sum_heads = ["label", "件数", "エラー数", "エラー率", "平均ms", "中央値ms", "95%ms", "最小", "最大", "経過秒", "件/秒", "成功のみ平均"]
    for c, h in enumerate(sum_heads, 1):
        put(ws, r, c, h, F(8, True, C_WHITE), FILL(C_DARK), AL("center", "center"), height=28)
    r += 1
    first_sum = r
    for label in ["01_ログイン", "02_検索", "03_詳細"]:
        put(ws, r, 1, label, F(8, True), FILL(C_EMPH), AL("left", "center"), height=22)
        b = f'COUNTIFS($C${data_start}:$C${data_end},A{r})'
        put(ws, r, 2, f"={b}", F(8), FILL(C_WHITE), AL("right", "center"), num_fmt="#,##0")
        put(ws, r, 3, f'=COUNTIFS($C${data_start}:$C${data_end},A{r},$G${data_start}:$G${data_end},FALSE)', F(8), FILL(C_WHITE), AL("right", "center"), num_fmt="#,##0")
        put(ws, r, 4, f'=IF(B{r}=0,"",C{r}/B{r})', F(8), FILL(C_WHITE), AL("right", "center"), num_fmt="0.00%")
        put(ws, r, 5, f'=IF(B{r}=0,"",AVERAGEIFS($B${data_start}:$B${data_end},$C${data_start}:$C${data_end},A{r}))', F(8), FILL(C_WHITE), AL("right", "center"), num_fmt="#,##0.0")
        # Excel は、FILTER を _xlfn. なしで保存したブックを開けない。
        # MINIFS / MAXIFS / PERCENTILE.INC も接頭辞が無いと #NAME? になる。
        # 開いたあとの数式バーでは接頭辞は外れ、FILTER や MINIFS と表示される。
        put(ws, r, 6, f'=IF(B{r}=0,"",MEDIAN(_xlfn.FILTER($B${data_start}:$B${data_end},$C${data_start}:$C${data_end}=A{r})))', F(8), FILL(C_WHITE), AL("right", "center"), num_fmt="#,##0.0")
        put(ws, r, 7, f'=IF(B{r}=0,"",_xlfn.PERCENTILE.INC(_xlfn.FILTER($B${data_start}:$B${data_end},$C${data_start}:$C${data_end}=A{r}),0.95))', F(8), FILL(C_WHITE), AL("right", "center"), num_fmt="#,##0.0")
        put(ws, r, 8, f'=_xlfn.MINIFS($B${data_start}:$B${data_end},$C${data_start}:$C${data_end},A{r})', F(8), FILL(C_WHITE), AL("right", "center"), num_fmt="#,##0.0")
        put(ws, r, 9, f'=_xlfn.MAXIFS($B${data_start}:$B${data_end},$C${data_start}:$C${data_end},A{r})', F(8), FILL(C_WHITE), AL("right", "center"), num_fmt="#,##0.0")
        put(ws, r, 10, f'=(_xlfn.MAXIFS($A${data_start}:$A${data_end},$C${data_start}:$C${data_end},A{r})-_xlfn.MINIFS($A${data_start}:$A${data_end},$C${data_start}:$C${data_end},A{r}))*86400', F(8), FILL(C_WHITE), AL("right", "center"), num_fmt="#,##0.0")
        put(ws, r, 11, f'=IF(J{r}=0,"",B{r}/J{r})', F(8), FILL(C_WHITE), AL("right", "center"), num_fmt="0.000")
        put(ws, r, 12, f'=AVERAGEIFS($B${data_start}:$B${data_end},$C${data_start}:$C${data_end},A{r},$G${data_start}:$G${data_end},TRUE)', F(8), FILL(C_WHITE), AL("right", "center"), num_fmt="#,##0.0")
        r += 1
    r += 1
    r = note(ws, r, last, "検算の目安（この架空データ）: 01_ログインは件数 8、エラー 0、平均 193.8、中央値 170、95% 333.5、経過 21 秒、件/秒 0.381。02_検索は件数 9、エラー 1、エラー率 11.11%、平均 452.2、中央値 300、95% 1040、経過 32 秒、件/秒 0.281、成功のみ平均 446.3。03_詳細は件数 6、平均 108.3、中央値 105、95% 137.5、経過 10 秒、件/秒 0.600。Excel が再計算した値と小数 1 桁で揃えば、数式は意図どおり。", C_EMPH, 9)
    r = images.add(ws, r, last, "この計算例シートを Excel で開いて再計算した画面（任意）", "数式を理解したあとに撮る。提出物ではない。", "02_検索の行で、平均と 95% とエラー率が見える。")
    _ = (headers_raw, raw_header, sum_header, first_sum)
    finish(ws, last)


def _trouble(wb, images):
    ws = new_sheet(wb, "28_トラブル対処", C_DARK)
    last = 7
    r = banner(ws, last, "対処", "症状から、最初に見る場所を決める", "人数を増やす対処は、ここに無い。")
    r = headers(ws, r, ["症状", "最初に見る", "切り分け", "対処", "再発防止", "", ""])
    rows = [
        ["jmeter.bat の黒い窓がすぐ閉じる", "新しい PowerShell で java -version", "java が無い、または JAVA_HOME が古い窓のまま。", "05。窓を開き直す。", "setenv の前に Java を確認する。", "", ""],
        ["メニューが英語", "bin\\setenv.bat の有無とファイル名", "setenv.bat.txt になっている。", "拡張子を表示して改名。JMeter を再起動。", "08 のスクショを残す。", "", ""],
        ["日本語が CSV で化ける", "users.csv の文字コード", "Excel の ANSI CSV を使った。", "CSV UTF-8 で保存し直し、File encoding を UTF-8。", "13。", "", ""],
        ["${user} がそのまま送られる", "結果ツリーの Request", "列名不一致、BOM、ファイル未配置。", "1 行目を user にする。jmx と同じフォルダ。", "相対パスを守る。", "", ""],
        ["TOKEN_NOT_FOUND", "ログインの Response data", "JSON Path が実応答と違う。", "実応答に合わせて Path を直す。Default は残す。", "13 の枠の設定を維持。", "", ""],
        ["CLI が即終了し jtl が無い", "jmeter.log の末尾", "jmx の要素が未知、または -t のパス違い。", "プラグイン jar、パス、バージョンを揃える。", "16 の版の確認。", "", ""],
        ["FileNotFound users.csv", "jmeter.log と ls", "csv が jmx の隣に無い。", "plans/order-search に並べる。", "17 の 4 ファイル。", "", ""],
        ["$ '\\r' : command not found", "file run.sh", "CRLF。", "sed -i 's/\\r$//' run.sh またはメモ帳で LF 保存し直す。", "転送前に LF。", "", ""],
        ["-o のフォルダが空でない", "jmeter.log", "html-report を先に作った。", "その回は捨て、新しい STAMP で実行。同じパスで消してよいときだけ -f。", "run.sh の mkdir 対象に html-report を入れない。", "", ""],
        ["HTML の時刻がずれる", "timestamp_format と date_format", "片方だけ変えた。", "31 の 2 行を同じ書式にして、jtl から -g で作り直す。", "08。", "", ""],
        ["Active が 1 のまま", "シナリオの式と -J のキー", "__P の名前と -J が不一致。", "-Jthreads と ${__P(threads,1)} を一致させる。", "19 の確認節。", "", ""],
        ["Too many open files", "ulimit -n と jmeter.log", "ファイルディスクリプタ不足。", "実行ユーザの nofile を上げてからやり直す。上げ方は OS の手順に従う。", "18 の ulimit 確認。", "", ""],
        ["OutOfMemoryError", "jmeter.log", "リスナーが有効、ヒープが小さい、応答本文を保存している。", "リスナー無効、response_data=false、HEAP を 2g の範囲で見直す。人数を先に増やさない。", "16。", "", ""],
        ["Err が最初から 100%", "エラー行の responseCode", "ホスト名、証明書、パス、認証のどれかで全件失敗。", "1 人の GUI に戻る。負荷を続けない。", "14 を通ってから CLI。", "", ""],
        ["SSH を閉じたら試験が消えた", "ps に jmeter が無い", "tmux 無しで切断した。", "その回は中断として記録。取り直す。", "19 の tmux。", "", ""],
        ["Excel が jtl で固まる", "タスクマネージャ", "ダブルクリックで全行を開いた。", "Excel を強制終了してよい。23 の Power Query でエラー行だけにする。", "jtl を関連付けで開かない。", "", ""],
        ["statistics.json が無い", "html-report の有無と log", "HTML 生成に失敗した。", "log の原因を直して jmeter -g result.jtl -o 新しい空フォルダ。", "08 の列設定。", "", ""],
        ["判定が全部不合格になる", "26 の基準セルと単位", "秒で書かれた基準とミリ秒の実測を比べている。", "基準はミリ秒。95% の列に pct2 を貼っているか確認。", "26 の注記。", "", ""],
        ["総合が空のまま", "取引のサンプル数", "行が空。", "A と B を入れると判定が出る。", "値の貼り付け。", "", ""],
        ["invalid_client", "トークン取得の Response data の error", "Base64、client_id、client_secret、Content-Type。", "36。secrets を直し、JMeter を起動し直す。", "秘密をログや図に残さない。", "", ""],
        ["${basic_credentials} が残る", "トークン取得の Request ヘッダ", "前処理が子でない、Language が groovy でない、-q が無い。", "35 の手順 5 と 12。", "GUI は -q 付きで開く。", "", ""],
        ["Bearer の位置に式が残る", "業務の Request ヘッダ", "${access_token} と書いた。または PostProcessor が JSON Extractor より上。", "37。業務側は ${__P(access_token)}。", "setUp の人数は 1 のまま。", "", ""],
        ["トークン取得が API Gateway へ飛ぶ", "Request のホスト", "HTTPリクエスト初期値がテスト計画の直下にある。トークン側のサーバ名が空。", "35 の手順 10 と 19。初期値は TG_業務 の中。", "ホストを 2 つ書く。", "", ""],
        ["403 Forbidden", "業務の responseCode と送ったヘッダ名", "X-API-Key が契約なのに X-API-ID を送っている。値、利用プラン、ステージ。", "39。名前を合意に合わせる。", "トークンの Path を先に疑って、名前を見ない。", "", ""],
        ["試験の後半だけ 401", "expires_in と -Jduration", "持続秒がトークンの期限以上。", "duration を expires_in より小さくして取り直す。", "期限切れの行を、サーバの性能低下と書く。", "", ""],
        ["Can't find additional property file", "jmeter.log", "-q のパスが違う。JMeter は警告だけで続行する。", "パスを直し、1 人からやり直す。", "ファイルが無くても開始したので成功だと思わない。", "", ""],
    ]
    header = r
    r = table(ws, r, rows, [22, 24, 36, 42, 28, 16, 14])
    # header row is header-1? headers() was called and returns next, so header variable is the row AFTER headers if I set header = r after headers.
    # I set r = headers then header = r which is first data row. Fix filter to header-1.
    ws.auto_filter.ref = f"A{header - 1}:G{r - 1}"
    ws.freeze_panes = f"A{header}"
    finish(ws, last)


def _check(wb, images):
    ws = new_sheet(wb, "29_チェックリスト", C_DARK)
    last = 7
    r = banner(ws, last, "当日", "印を付けながら進める", "済の列に ○ を入れる。スキップした項目は空のままにせず、理由をメモする。")
    r = headers(ws, r, ["段階", "項目", "確認のしかた", "済", "メモ", "", ""])
    items = [
        ["前", "試験対象と人数と時間と合格基準が文章で残っている", "チケットの文言を 26 の条件へ写せる", "", "", "", ""],
        ["前", "端末とサーバの JMeter が 5.6.3", "jmeter -v と jmeter.bat -v", "", "", "", ""],
        ["前", "Java 17 が両方で説明できる", "java -version", "", "", "", ""],
        ["前", "user.properties の日時書式が対で入っている", "timestamp_format と date_format", "", "", "", ""],
        ["前", "スレッド数が ${__P(threads,1)}", "jmx を開かず、保存済みの画面か、次回 GUI で見る", "", "", "", ""],
        ["前", "1 人の GUI でログインと検索が緑", "14", "", "", "", ""],
        ["前", "リスナーが無効、Filename が空", "ツリーがグレー", "", "", "", ""],
        ["前", "CSV が UTF-8 で jmx の隣にある", "列名と BOM", "", "", "", ""],
        ["前", "Cognito のとき、secrets.properties が BOM 無しで、jmx に実値が無い", "35 の観点。メモ帳で jmx を検索", "", "", "", ""],
        ["前", "setUp の人数が 1。業務側が ${__P(access_token)}。ヘッダ名が合意どおり", "35 と 39", "", "", "", ""],
        ["前", "1 人の GUI でトークン取得と業務 API が緑。図に秘密が無い", "35 の手順 32 から 35", "", "", "", ""],
        ["前", "run.sh が LF", "file コマンド", "", "", "", ""],
        ["前", "結果先が新しいフォルダになる", "run.sh の STAMP", "", "", "", ""],
        ["中", "tmux の中で動いている", "デタッチしても ps に残る", "", "", "", ""],
        ["中", "最初は 1 人で CLI が成功している", "Active 1、Err 0", "", "", "", ""],
        ["中", "本番人数で Active が目標に届いた", "summary の Active", "", "", "", ""],
        ["中", "Err を 30 秒ごとに見ている", "+ の行", "", "", "", ""],
        ["中", "悪化やエラー増で止める基準を手元に置いている", "16 と 26 の基準", "", "", "", ""],
        ["後", "最終の summary = をコピーした", "チケットかメモ", "", "", "", ""],
        ["後", "jtl、log、run-conditions、html-report、statistics.json がある", "ls", "", "", "", ""],
        ["後", "フォルダごと端末へ戻した", "22", "", "", "", ""],
        ["後", "index.html がグラフ付きで開く", "ブラウザ", "", "", "", ""],
        ["後", "26 の取引行と総合判定を入れた", "数式が合格または不合格", "", "", "", ""],
        ["後", "見込みに、未測定の範囲を書いた", "25 の禁止例に触れていない", "", "", "", ""],
        ["後", "図の番号と試験フォルダ名が一致する", "画像枠", "", "", "", ""],
        ["後", "原本の手順書を上書きしていない", "提出ファイルが別名", "", "", "", ""],
    ]
    header = r
    for i, row in enumerate(items):
        fill = FILL(C_WHITE if i % 2 == 0 else C_ZEBRA)
        # 済 and メモ are input-like
        for c, val in enumerate(row, 1):
            this = FILL(C_EMPH) if c in (4, 5) else fill
            put(ws, r, c, val, F(9, c == 1), this, AL("left", "center"), height=24)
        r += 1
    ws.auto_filter.ref = f"A{header - 1}:G{r - 1}"
    ws.freeze_panes = f"A{header}"
    finish(ws, last)


def _commands(wb, images):
    ws = new_sheet(wb, "30_コマンド早見", C_GRAY)
    last = 7
    r = banner(ws, last, "早見", "コピーするコマンド", "ホスト名、人数、日時は現場の値に変える。")
    r = headers(ws, r, ["場面", "どこで", "コマンド", "成功の見え方", "注意", "", ""])
    rows = [
        ["Java 確認", "端末 PowerShell", "java -version", "17 で始まる", "古い窓は開き直す。", "", ""],
        ["JAVA_HOME", "端末 PowerShell", "echo $env:JAVA_HOME", "JDK のホーム。bin で終わらない", "ユーザ環境変数。", "", ""],
        ["zip のハッシュ", "端末 PowerShell", "certutil -hashfile C:\\LoadTest\\dist\\apache-jmeter-5.6.3.zip SHA512", "sha512 ファイルと一致", "不一致なら展開しない。", "", ""],
        ["ブロック解除", "端末 PowerShell", "Unblock-File -Path C:\\LoadTest\\dist\\apache-jmeter-5.6.3.zip", "エラーが出ない", "展開前に行う。", "", ""],
        ["バージョン", "端末 cmd", "C:\\Tools\\apache-jmeter-5.6.3\\bin\\jmeter.bat -v", "5.6.3 と出て戻る", "GUI は開かない。", "", ""],
        ["バージョン", "Linux", "/opt/jmeter/apache-jmeter-5.6.3/bin/jmeter -v", "5.6.3", "引数無しは GUI を試しに行く。", "", ""],
        ["ディレクトリ", "Linux", "mkdir -p /opt/jmeter/plans/order-search /opt/jmeter/results", "ls で見える", "実行ユーザで作る。", "", ""],
        ["転送", "端末 PowerShell", "scp C:\\LoadTest\\scenarios\\order-search\\* loadtest@jmeter-runner.example.internal:/opt/jmeter/plans/order-search/", "プロンプトが戻る", "results は送らない。", "", ""],
        ["改行", "Linux", "file /opt/jmeter/plans/order-search/run.sh", "CRLF ではない", "CRLF なら sed -i 's/\\r$//' run.sh", "", ""],
        ["tmux", "Linux", "tmux new -s loadtest", "セッションが開く", "離脱は Ctrl+b のあと d。復帰は tmux attach -t loadtest", "", ""],
        ["1 人で実行", "Linux", "bash /opt/jmeter/plans/order-search/run-1user.sh", "summary の Active が 1", "本番の run.sh と間違えない。", "", ""],
        ["本番実行", "Linux", "bash /opt/jmeter/plans/order-search/run.sh", "結果ディレクトリのパスが表示される", "html-report は事前に作らない。", "", ""],
        ["穏やかに停止", "Linux 別窓", "/opt/jmeter/apache-jmeter-5.6.3/bin/shutdown.sh", "最終 summary が出る", "実行中の 1 件は終わる。", "", ""],
        ["即時停止", "Linux 別窓", "/opt/jmeter/apache-jmeter-5.6.3/bin/stoptest.sh", "プロセスが終わる", "中断と記録する。", "", ""],
        ["行数", "Linux", "wc -l /opt/jmeter/results/order-search_日時/result.jtl", "2 行以上", "cat しない。", "", ""],
        ["HTML だけ作り直す", "Linux", "/opt/jmeter/apache-jmeter-5.6.3/bin/jmeter -g /opt/jmeter/results/order-search_日時/result.jtl -o /opt/jmeter/results/order-search_日時/html-report2", "index.html ができる", "出力先は空または未作成。", "", ""],
        ["回収", "端末 PowerShell", "scp -r loadtest@jmeter-runner.example.internal:/opt/jmeter/results/order-search_日時 C:\\LoadTest\\results\\", "html-report まで届く", "途中失敗は取り直す。", "", ""],
        ["GUI をプロパティ付きで開く", "端末 PowerShell", "C:\\Tools\\apache-jmeter-5.6.3\\bin\\jmeter.bat -q C:\\LoadTest\\scenarios\\api-gateway\\load.properties -q C:\\LoadTest\\scenarios\\api-gateway\\secrets.properties", "窓が開く", "秘密を -J でコマンドに直書きしない。", "", ""],
        ["公開例の Base64", "端末 PowerShell", "$raw = \"djc98u3jiedmi283eu928:abcdef01234567890\"; [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($raw))", "ZGpjOTh1M2ppZWRtaTI4M2V1OTI4OmFiY2RlZjAxMjM0NTY3ODkw", "実秘密はこのコマンドに貼らない。36。", "", ""],
        ["Cognito の CLI", "Linux", "bash /opt/jmeter/plans/api-gateway/run-api.sh", "結果ディレクトリのパス。summary の Active が -Jthreads と一致", "先に threads=1 のコピーで流す。35。", "", ""],
        ["秘密ファイルの権限", "Linux", "chmod 600 /opt/jmeter/plans/api-gateway/secrets.properties", "ls -l で 600", "中身を cat して画面に残さない。", "", ""],
        ["ファイル数", "Linux", "ulimit -n", "スレッド数より十分大きい", "小さいときは試験を始めない。", "", ""],
    ]
    header = r
    r = table(ws, r, rows, [16, 18, 55, 28, 36, 12, 12])
    ws.auto_filter.ref = f"A{header - 1}:G{r - 1}"
    ws.freeze_panes = f"A{header}"
    finish(ws, last)


def _settings(wb, images):
    ws = new_sheet(wb, "31_設定の写し", C_GRAY, {"A": 28, "B": 90, "C": 28, "D": 22, "E": 18, "F": 16, "G": 14})
    last = 7
    r = banner(ws, last, "写し", "設定ファイルの全文", "このシートからコピーする。08、17、18、19 はここを参照する。")
    blocks = [
        ("bin/setenv.bat（端末）", SETENV_BAT, "jmeter.bat と同じフォルダ。拡張子が .bat であること。"),
        ("bin/setenv.sh（Linux）", SETENV_SH, "改行 LF。chmod は jmeter 本体側で足りる。このファイルは source される。"),
        ("bin/user.properties の末尾（端末と Linux の両方）", USER_PROPERTIES, "既存のコメントは消さない。timestamp と date_format は同じ文字列のまま対で置く。"),
        ("plans/order-search/load.properties", LOAD_PROPERTIES, "ホスト名を現場の値に変える。スレッド数はここに書かず -J で渡す。"),
        ("plans/order-search/run.sh", RUN_SH, "改行 LF。threads、ramp、duration の 3 つは、ファイル内の -J と run-conditions.txt の両方を同じ値にする。"),
    ]
    from sheets_auth import (
        GROOVY_BASIC, GROOVY_STORE, LOAD_API_PROPERTIES, RUN_API_SH, SECRETS_PROPERTIES,
    )
    blocks.extend([
        ("plans/api-gateway/load.properties", LOAD_API_PROPERTIES, "Cognito と API Gateway のホスト。秘密は書かない。35 の手順で使う。"),
        ("plans/api-gateway/secrets.properties", SECRETS_PROPERTIES, "client_id、client_secret、api_key。UTF-8 の BOM 無し。値は引用符で囲まない。chmod 600。レポートに貼らない。"),
        ("PP_basicを作る（Groovy）", GROOVY_BASIC, "HTTP_POST_cognito_token の子。Language は groovy。java.util.Base64。値はログに出さない。"),
        ("PP_トークンを全体へ（Groovy）", GROOVY_STORE, "JSON Extractor より下。変数 access_token をプロパティへコピーする。"),
        ("plans/api-gateway/run-api.sh", RUN_API_SH, "改行 LF。-q は load.properties と secrets.properties の 2 回。-t が api-gateway.jmx。"),
    ])
    for title, body, caution in blocks:
        r = section(ws, r, last, title)
        r = note(ws, r, last, caution, C_SOFT, 9)
        text = body.rstrip() + "\n"
        lines = text.split("\n")
        # Excel の行高の上限は 409 ポイント。超えるとブックを開けないことがある。
        chunk = 24
        for i in range(0, len(lines), chunk):
            part = "\n".join(lines[i:i + chunk])
            h = min(400, max(72, part.count("\n") * 14 + 28))
            merge(ws, f"A{r}:G{r}", part, F(9), FILL(C_WHITE), AL("left", "top"), height=h)
            r += 1
        r += 1
    r = note(ws, r, last, "優先順位: 同じキーは、コマンドの -J が最も強い。次が -q の load.properties。次が user.properties。jmeter.properties は配布の既定で、直接編集しない。-D は Java のシステムプロパティで、${__P} は読まない。")
    finish(ws, last)


def _images(wb, images):
    ws = new_sheet(wb, "32_画像台帳", C_BLACK)
    last = 7
    r = banner(ws, last, "台帳", "差し込み枠の一覧", "図を入れる場所は各シートのグレー枠。この表は、何を撮るかを失わないための控。")
    r = headers(ws, r, ["番号", "シート", "撮るもの", "撮るタイミング", "画角", "入れたか", ""])
    header = r
    data = []
    for img_id, sheet, what, when, framing in images.rows:
        data.append([img_id, sheet, what, when, framing, "", ""])
    r = table(ws, r, data, [12, 24, 36, 36, 40, 14, 14])
    # 入れたか column should be input fill. Retouch last loop rows.
    # table already wrote them. Repaint column 6 as input.
    for row in range(header, r):
        put(ws, row, 6, "", F(9), FILL(C_EMPH), AL("center", "center"))
        put(ws, row, 7, "", F(9), FILL(C_WHITE), AL())
    ws.auto_filter.ref = f"A{header - 1}:G{r - 1}"
    ws.freeze_panes = f"A{header}"
    finish(ws, last)


def _history(wb, images):
    ws = new_sheet(wb, "33_改定履歴", C_BLACK)
    last = 7
    r = banner(ws, last, "履歴", "この手順書の版", "画面の名前やオプションは、下の公式ページを 2026-10-04 に確認して書いた。配布が 5.6.3 から進んだら、バージョンの行を直してから使う。")
    r = headers(ws, r, ["日付", "版", "内容", "確認した一次情報", "", "", ""])
    rows = [
        ["2026-10-04", "1.0", "初版。Windows 11 でシナリオを作り、Linux で CLI 実行し、Excel で結果と見込みを書く流れ。", "https://jmeter.apache.org/download_jmeter.cgi", "", "", ""],
        ["2026-10-04", "1.0", "Java の要件。5.6.3 は Java 8 以上。推奨は 17。次メジャーは 17 以上が必須、という Changes の記載に合わせ、新規導入は 17。", "https://jmeter.apache.org/changes.html", "", "", ""],
        ["2026-10-04", "1.0", "インストール、setenv、CLI の -n -t -l -j -e -o、ヒープ既定 1g。", "https://jmeter.apache.org/usermanual/get-started.html", "", "", ""],
        ["2026-10-04", "1.0", "結果ファイルと -l、-j。summary の見方。", "https://jmeter.apache.org/usermanual/listeners.html", "", "", ""],
        ["2026-10-04", "1.0", "HTML ダッシュボードの必要列、APDEX 既定、親サンプルはオフを推奨、という記載。", "https://jmeter.apache.org/usermanual/generating-dashboard.html", "", "", ""],
        ["2026-10-04", "1.1", "Cognito の /oauth2/token へ client_credentials。Authorization は Basic と Base64(client_id:client_secret)。本体は form。", "https://docs.aws.amazon.com/cognito/latest/developerguide/token-endpoint.html", "", "", ""],
        ["2026-10-04", "1.1", "Base64 は java.util.Base64。JMeter 5.6.3 の関数一覧に Base64 関数は無い。-q は重複可。", "https://jmeter.apache.org/usermanual/functions.html", "", "", ""],
        ["2026-10-04", "1.1", "setUp は通常スレッドより前。ヘッダは同名なら後のものが置き換わる。JSON Extractor の Match No. 空欄は 0（ランダム）。正規表現のテンプレートは $1$。", "https://jmeter.apache.org/usermanual/component_reference.html", "", "", ""],
        ["2026-10-04", "1.1", "API キーの公式例はヘッダ X-API-Key。手順の記入例 X-API-ID は合意がある現場用。名前だけ変える。", "https://docs.aws.amazon.com/apigateway/latest/developerguide/api-gateway-api-key-call.html", "", "", ""],
    ]
    r = table(ws, r, rows, [16, 12, 55, 55, 16, 14, 14])
    r += 1
    r = note(ws, r, last, "プラグイン（Plugins Manager）は Apache JMeter 本体の外にある。使うときだけ 11 の手順に戻り、入れた jar のファイル名をこの履歴の次の行へ手で追記する。\n日本語のメニュー名は、JMeter 5.6.3 の messages_ja.properties の、ファイル、編集、実行、オプション、追加、スレッドグループ、結果をツリーで表示、統計レポート、応答アサーション、定数タイマ、一様乱数タイマ、トランザクションコントローラ、一度だけ実行されるコントローラ、を確認した。HTTP リクエストなど一部は画面で英語のままのことがあるので、手順は英語名を併記している。")
    finish(ws, last)


def build_rest(wb, images) -> None:
    _cat(wb, images)
    _scenario(wb, images)
    _data(wb, images)
    _gui_run(wb, images)
    _gui_results(wb, images)
    _checks(wb, images)
    _transfer(wb, images)
    _linux(wb, images)
    _cli(wb, images)
    _watch(wb, images)
    _jtl(wb, images)
    _back(wb, images)
    _excel_in(wb, images)
    _analyze(wb, images)
    _judge(wb, images)
    _report(wb, images)
    _calc(wb, images)
    _trouble(wb, images)
    _check(wb, images)
    _commands(wb, images)
    _settings(wb, images)
    from sheets_auth import build_auth
    build_auth(wb, images)
    _images(wb, images)
    _history(wb, images)
    # タブは番号順。画像台帳は認証の図を取り込んだあとなので、中身はそのまま並びだけ直す。
    auth_tabs = [
        "34_認証の全体像", "35_Cognitoトークン", "36_BASE64ヘッダ",
        "37_トークンの保持", "38_正規表現抽出", "39_APIキーヘッダ",
    ]
    front = [name for name in wb.sheetnames if name not in auth_tabs]
    wb._sheets = [wb[name] for name in front + auth_tabs]
