# outlook-mcp

Hotmail / Outlook.com のメールボックスを、AIエージェントから自然言語で検索・整理できるようにする MCP サーバです。

```
あなた: 先月のAmazonからのメール、未読のやつを「買い物」フォルダにまとめて
Claude: (search_messages → list_folders → move_messages を自動で呼ぶ)
        7月のAmazonからの未読12件を「受信トレイ/買い物」へ移動しました。
```

Microsoft Graph API を使います。**メールの送信はできません。完全削除もできません。**

---

## できること / できないこと

| | |
|---|---|
| ✅ 検索 | 件名・本文・差出人・期間・未読で絞り込み |
| ✅ 読む | 本文の取得(HTMLメールは平文に変換) |
| ✅ 整理 | フォルダ移動、アーカイブ、既読/未読の切り替え |
| ✅ 一括整理 | 条件に一致するメールをまとめて移動(下見つき) |
| ✅ 棚の再編 | フォルダの作成・改名・移動・削除 |
| ✅ 自動振分 | Outlook側のルールを作る(MCPが起動していなくても効く) |
| ✅ 捨てる | ゴミ箱へ移動(**元に戻せます**) |
| ❌ 送信 | 実装していません。`Mail.Send` 権限を要求しません |
| ❌ 完全削除 | 実装していません。取り返しがつかないため |
| ❌ 添付の取得 | 未実装(添付の有無は 📎 で表示されます) |

要求する権限は **`Mail.ReadWrite`** と **`MailboxSettings.ReadWrite`**(振分ルール用)の2つです。

> **v0.2 から権限が1つ増えました。** それ以前にログイン済みの場合、古いトークンには
> `MailboxSettings.ReadWrite` が入っておらず、ルール系ツールだけが 403 になります。
> `python login.py --logout` → `python login.py` で同意し直してください。
> `check_config` が「振分ルール: 使えません」と表示したらこれが原因です。

---

## 設計方針

**AIに受信箱を触らせる**以上、壊せる範囲を先に狭めておくべきだと考えました。

- **完全削除の手段を置かない。** ゴミ箱へ移すだけ。ツールとして存在しない操作は、どう指示されても起きません。
- **一度に触れるのは25件まで。** 誤爆したときの被害を有限にします。大量処理は `move_by_search` に分け、**下見を既定**にしました。
- **無条件の一括操作を拒否する。** 絞り込み条件のない `move_by_search` はエラーになります。
- **システムフォルダを守る。** 受信トレイ・迷惑メール等は改名・移動・削除できません。
- **空でないフォルダは `force` なしに消せない。** 退避(`move_folder`)で足りるなら、そちらを促します。
- **`OUTLOOK_READONLY=true` で書き込みを全面停止。** 読み取り専用サーバとして動かせます。
- **送信権限を要求しない。** 「AIが勝手にメールを出す」経路を原理的に作りません。
- **`destructive_hint` を正しく申告する。** 対応クライアントは `move_to_trash` を他と区別して扱えます。

---

## セットアップ

**必要なもの**: Python 3.10以上、Microsoftアカウント、Claude Code(または他のMCPクライアント)

### 1. Azure でアプリを登録する

無料です。Azureのサブスクリプション契約は要りません。

ただし**アプリ登録はテナント(ディレクトリ)の中にしか置けません**。個人のMicrosoftアカウント(Hotmail / Outlook.com / 手持ちのメールアドレスで作ったもの)は既定ではテナントに属さないため、そのまま [Microsoft Entra 管理センター](https://entra.microsoft.com/) を開くとこう言われます。

```
選択されたユーザー アカウントは、テナント 'Microsoft Services' に存在しないため、
そのテナントのアプリケーション '...' にアクセスできません。
```

**これは「個人アカウントではこのサーバを使えない」という意味ではありません。** サインインもメールの読み書きも個人アカウントのままで動きます。足りないのは *クライアントIDというGUIDを1個発行する場所* だけで、それも最初の一度きりです。

#### テナントを用意する

**職場・学校アカウントを持っているなら**、それでサインインするのが最短です。**登録先のテナントと、実際にメールを読むアカウントは一致していなくて構いません。** 下の手順3で個人アカウントを含めておけば、会社のテナントで発行したIDのまま個人のHotmailにサインインできます。

**持っていないなら** [Azure 無料アカウント](https://azure.microsoft.com/free/) を作ります。サインアップするとディレクトリが同時に作られ、管理センターに入れるようになります。

- 本人確認でクレジットカードの登録を求められますが、明示的にアップグレードしない限り従量課金には移行しません
- アプリ登録と Microsoft Graph の呼び出しは、そもそも課金対象外です
- 登録が済めばAzure側は放置で構いません。トークンは個人アカウントに紐づくので、以後ログインする場面はありません

> ブラウザで複数のMicrosoftアカウントにログイン済みだと、意図しない方が選ばれて同じエラーに
> なることがあります。InPrivateウィンドウで開き直すと切り分けられます。

#### 登録する

1. [Microsoft Entra 管理センター](https://entra.microsoft.com/) → **アプリの登録** → **新規登録**
2. 名前は任意(例 `outlook-mcp`)
3. **サポートされているアカウントの種類**:
   - Hotmail/Outlook.com を使うなら → 「**任意の組織ディレクトリ内のアカウントと個人の Microsoft アカウント**」
4. リダイレクトURIはこの画面では**空のまま**で登録（手順7で追加します）
5. 登録後の概要画面から「**アプリケーション (クライアント) ID**」をコピー
6. 左メニュー **認証** → 一番下の「詳細設定」 →
   「**パブリック クライアント フローを許可する**」を **はい** にして保存
   （device code flow に必須。ここを忘れると `login.py` が失敗します）
7. 同じ **認証** の画面で **プラットフォームを追加** →
   **モバイル アプリケーションとデスクトップ アプリケーション** →
   次の2つにチェックを入れて保存
   ```
   https://login.microsoftonline.com/common/oauth2/nativeclient
   https://login.live.com/oauth20_desktop.srf     ← 個人アカウントに必須
   ```
8. 左メニュー **API のアクセス許可** → **アクセス許可の追加** →
   **Microsoft Graph** → **委任されたアクセス許可** → `Mail.ReadWrite` を追加

> **手順7を飛ばすと、個人アカウントでのログインが必ず失敗します。**
> ブラウザでコードを入力した直後に、こう出ます:
>
> ```
> invalid_request: The provided request must include a 'redirect_uri' input parameter.
> ```
>
> 手順6の「パブリック クライアント フローを許可する」は**必要条件ですが十分ではありません**。
> あれは `allowPublicClient` を立てるだけで、個人アカウント側の認証サーバ(`login.live.com`)は
> 実体としてのリダイレクトURIが登録されていることを別途要求します。職場・学校アカウントだけなら
> 空でも通るため、見落としやすい箇所です。device code flow で実際にブラウザがこのURIへ
> 飛ぶことはありません。「登録されている」という事実だけが要求されます。

> 画面の文言はAzure側の更新でしばしば変わります。見つからないときは英語UIの
> "Allow public client flows" / "Delegated permissions" で探すと当たります。
> 認証画面が新しいプレビューUI（**Authentication (Preview)**）になっていて
> 「プラットフォームを追加」が見当たらない場合は、画面上部の
> "To switch to the old experience, please click here." で旧UIに戻せます。

クライアントシークレットは**作らないでください**。このサーバはパブリッククライアントとして動くため不要です。

### 2. インストール

**macOS / Linux**

```bash
git clone <このリポジトリ>
cd outlook-mcp

python3 -m venv .venv
.venv/bin/pip install -r requirements.txt

cp .env.example .env
```

**Windows (PowerShell)**

```powershell
git clone <このリポジトリ>
cd outlook-mcp

py -3 -m venv .venv
.venv\Scripts\pip install -r requirements.txt

copy .env.example .env
```

> 以降このREADMEでは `.venv/bin/python` と書きますが、**Windows では
> `.venv\Scripts\python.exe` に読み替えてください**（`bin` ではなく `Scripts`、拡張子つき）。

`.env` を開いて、さっきコピーしたIDを書きます。

```ini
OUTLOOK_CLIENT_ID=12345678-abcd-...
OUTLOOK_TENANT=common
OUTLOOK_READONLY=false
```

### 3. ログイン(最初の一度だけ)

```bash
.venv/bin/python login.py            # macOS / Linux
.venv\Scripts\python.exe login.py    # Windows
```

表示されたURLをブラウザで開き、コードを入力して、Microsoftアカウントでサインインします。
**パスワードをこのツールに渡すことはありません。**

> **「通常は表示されないページにアクセスしました」という警告が出ますが、正常です。**
> device code flow は「別の端末で出たコードをブラウザで入力させる」仕組みのため、
> 他人に自分の作ったコードを入力させるフィッシングに悪用されてきました。そのため
> Microsoft はこのページに来た全員へ警告を出します。**同意画面に出るアプリ名が、
> 自分で登録した名前(例 `outlook-mcp`)であることだけは必ず確認してください。**
> 見覚えのない名前が出たら、そのコードは他人由来です。中止してください。

> **アカウントの選択画面が出たら「個人のアカウント」を選びます。**
> Azureのテナントを自分で作った場合、同じメールアドレスが「個人のMicrosoftアカウント」と
> 「職場または学校アカウント(テナントの所有者)」の2つの人格を持ちます。**メールボックスが
> あるのは個人の方だけ**で、職場側を選ぶとサインインは通るのに Graph が
> `MailboxNotEnabledForRESTAPI` を返して失敗します。
> `.env` で `OUTLOOK_TENANT=consumers` にすると、個人アカウント専用の経路に固定でき、
> この取り違えが起きなくなります。

成功すると `token_cache.json` ができ、以後サーバが無人でトークンを更新します。

> `token_cache.json` はメールボックスの鍵に相当します。`.gitignore` 済みですが、共有・コミットしないでください。
> アカウントを切り替えるときは `login.py --logout`（上と同じ python を指定）してからやり直します。

### 4. クライアントに登録する

**絶対パス**で指定してください。相対パスだと、別のディレクトリから起動したときに動きません。

**macOS / Linux**

```bash
claude mcp add outlook -- /abs/path/outlook-mcp/.venv/bin/python /abs/path/outlook-mcp/outlook_server.py
claude mcp list   # ✓ connected を確認
```

**Windows (PowerShell)**

```powershell
claude mcp add outlook -- "C:\abs\path\outlook-mcp\.venv\Scripts\python.exe" "C:\abs\path\outlook-mcp\outlook_server.py"
claude mcp list   # √ connected を確認
```

パスに空白や日本語が含まれる場合に備え、Windowsでは**二重引用符で囲ってください**。

登録は**この1回だけ**です。以後 `.env` を書き換えたら `/mcp` から再接続すれば反映されます。

> 既定では**そのプロジェクト内でだけ**有効な登録（local スコープ）になります。
> どのディレクトリからでも使いたいときは `-s user` を付けてください。
>
> ```
> claude mcp add -s user outlook -- <python の絶対パス> <outlook_server.py の絶対パス>
> ```
>
> 登録先を変えるときは、先に `claude mcp remove outlook` で古い方を消します。

> MCPサーバは `/mcp` に出ます。**`/plugin` には出ません**（プラグインは別の仕組みです）。

つながったら、まず `check_config` を呼ばせてみてください。

```
あなた: Outlookつながってる?
Claude: OK: 受信トレイ 3,412件(未読 87件)にアクセスできました。
```

---

## ツール

| ツール | 種別 | 説明 |
|---|---|---|
| `check_config` | 読取 | 設定・認証・接続の診断 |
| `list_folders` | 読取 | フォルダ一覧(件数・未読数つき) |
| `search_messages` | 読取 | 検索。キーワード / 差出人 / 期間 / 未読 / フォルダ |
| `get_message` | 読取 | 1通の本文と宛先を読む |
| `list_rules` | 読取 | 設定済みの振分ルールを一覧 |
| `create_folder` | 書込 | フォルダを作る |
| `rename_folder` | 書込 | フォルダを改名する(中身は動かない) |
| `move_folder` | 書込 | フォルダを別の親の下へ移す(中身ごと) |
| `move_messages` | 書込 | 指定フォルダへ移動(最大25件) |
| `move_by_search` | 書込 | 条件に一致するメールを一括移動(最大2000件) |
| `archive_messages` | 書込 | アーカイブへ移動 |
| `mark_messages_read` | 書込 | 既読 / 未読の切り替え |
| `create_rule` | 書込 | 自動振分ルールを作る |
| `move_to_trash` | 破壊 | ゴミ箱へ移動(元に戻せる) |
| `delete_folder` | 破壊 | フォルダを削除(空でなければ `force` が要る) |
| `delete_rule` | 破壊 | 振分ルールを削除(メールは動かない) |

### 一括移動 — `move_by_search`

`move_messages` は1回25件です。数千通を動かすために、条件に一致するものをまとめて処理します。

**既定は下見(`dry_run=True`)で、何件動くかを数えるだけです。**

```
move_by_search(dest="99_Archive", folder="Music")
  → 元: Music / 走査 6,214件 → 該当 6,214件
    【下見のみ・まだ動かしていません】

move_by_search(dest="99_Archive", folder="Music", dry_run=False)
  → 6,214件を「99_Archive」へ移動しました。
```

内部では Graph の `/$batch` に20件ずつ束ねます。1通ずつ POST すると往復回数が現実的でないためです。

- 絞り込み条件を1つも指定しない呼び出しは**拒否されます**(メールボックス全体を無条件に動かす事故を防ぐため)
- 1回の上限は2000件。超える場合は同じ呼び出しを繰り返します
- 差出人・件名の部分一致は手元で判定します(`$filter` が `contains()` を受け付けないため)

> **棚ごと動かせるなら `move_folder` のほうが速いです。** メールを1通も動かさずに階層だけ変わります。

### 自動振分ルール — `create_rule`

Outlook のサーバ側に保存されるルールです。**このMCPが起動していなくても24時間効きます。**

```
create_rule(name="BLOCKDAG を捨てる", from_contains="blockdag.network", to_trash=True)
create_rule(name="Meta広告の領収書", from_contains="facebookmail.com",
            subject_contains="領収書", move_to="04_Work/広告代行", mark_read=True)
```

- 条件(`from_contains` / `subject_contains` / `body_contains`)は複数指定すると **AND**。値はカンマ区切りで複数渡せます
- **既に届いているメールには適用されません。** 過去分は `move_by_search` で別途動かします
- `to_trash` はゴミ箱へ移すだけで、完全削除ではありません

### 短縮ID

`search_messages` の各行は `#12` のような番号から始まります。Graph のメッセージIDは150文字前後あり、50件返すとそれだけで文脈を食い潰すためです。整理系ツールにはこの番号をそのまま渡します。

```
#12 2026-08-09 14:03 ●📎 Amazon.co.jp | ご注文の確認 | 受信トレイ
```

番号はサーバのプロセスが生きている間だけ有効で、**使い回されません**（再検索で `#3` の指す先が変わると事故になるため）。生のGraph IDも受け付けます。

> **メールを移動すると、その短縮IDは失効します。** Graph は移動時にメッセージIDを再発行するためです。
>
> ```
> move_messages("#1,#2", "領収書")   → 成功
> mark_messages_read("#1")           → エラー: 対象が見つかりません。
> ```
>
> 「移動してから既読にする」のような連続操作をするときは、**移動後に `search_messages` を引き直して
> 新しい番号を取り直してください**。エラーは案内文字列で返るため処理は止まりません。

---

## 既知の制約

**キーワード検索と厳密な新着順は両立しません。** Microsoft Graph は `$search` と `$filter` / `$orderby` を併用できない仕様です。このサーバは:

- キーワードや差出人の指定があるとき → `$search`(KQL)で検索し、**関連度順**で最大100件取得してから手元で日付順に並べ替える
- 指定がないとき → `$filter` + `$orderby` で**確実に新着順**

という切り替えをしています。前者で該当が100件を超える場合、古いものが取りこぼされる可能性があります。その旨は結果の末尾に表示されます。

**差出人はKQL側で処理されます。** Graph の messages に対する `$filter` は `contains()` を受け付けないため、「`amazon.co.jp` を含む」のような部分一致は `$search` 経由になります。

---

## 開発 / テスト

```bash
# macOS / Linux
.venv/bin/pip install pytest
.venv/bin/pytest -q              # 単体テスト 35項目
.venv/bin/python smoke_test.py   # スモークテスト 10項目
```

```powershell
# Windows (PowerShell)
.venv\Scripts\pip install pytest
.venv\Scripts\pytest -q
.venv\Scripts\python.exe smoke_test.py
```

どちらも **Microsoft Graph に接続せず、メールボックスを変更しません**。認証情報も不要で、ログイン前の状態のまま実行できます。

- `test_outlook.py` — 関数単位。Graph呼び出しはスタブに差し替え、「どんなリクエストを組み立てたか」を検証します
- `smoke_test.py` — サーバを実際に stdio で起動し、MCPクライアントから見える外形（ツール一覧・入力スキーマ・`destructive_hint`・エラー時の応答）を検証します

テスト項目の一覧と実行エビデンスは **[docs/TEST.md](docs/TEST.md)** にあります。

実アカウント（Hotmail、264フォルダ）に対する結合テストも実施済みです（読み取り9項目 + 書き込み10項目）。書き込み側は**既存のメールに触れず**、テスト用のフォルダとメッセージを自作して往復させ、最後に自分で作ったものだけを片付ける形で確認しています。未検証のまま残っている範囲も同ドキュメントに明記しています。

---

## ライセンス

MIT
