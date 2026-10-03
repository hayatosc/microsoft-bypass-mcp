# セットアップと開発

[README](../README.md) に戻る · [正式な仕様](../SPEC.md) · [フロー更新手順](../power-automate/microsoft-bypass-flow/README.md)

このサーバーは、利用を許可された Microsoft 365 接続を個人で運用するためのものです。組織の承認を代替せず、権限や契約上の制約を回避しません。まずローカルのオフライン検証を行い、Microsoft / Cloudflare 側の変更や実データでの確認は、別途許可された作業として進めます。

## 1. 利用条件を確認する

- 組織がメール・ファイルの読み取り、下書き保存、外部 AI / MCP クライアントへのデータ提供を認めていること。
- 接続所有者が Outlook メールボックスと OneDrive for Business を利用できること。個人向け OneDrive、共有 SharePoint ドライブ、共有リンクは対象ではありません。
- 対象テナント・環境のライセンスで、Power Automate の HTTP Request トリガー、Office 365 Outlook の `HttpRequest`、使用する OneDrive for Business 操作が利用できること。DLP、管理者の同意、コネクタ制限、呼び出し割り当ても確認してください。プランを問わず使えるという保証はありません。
- 既存 Outlook 接続が固定の読み取り・下書き操作を許可すること。下書き操作には Microsoft の API 仕様上、委任された `Mail.ReadWrite` が必要です。ここで新しい同意・権限付与を行う手順は提供しません。
- Cloudflare Workers / Access を利用でき、必要な CPU・メモリ・リクエスト枠を確認できること。
- MCP over Streamable HTTP と、選択した外部 OAuth 認証に対応するクライアントがあること。ChatGPT / Claude などの名前だけで接続可否は判断できません。

## 2. ローカルで依存関係と検証を用意する

CI は `.github/workflows/ci.yml` で **Bun 1.3.14** を指定しています。Python 3 はフローの生成確認とオフラインテストに使います。初回の依存取得には公開パッケージへのネットワーク接続が必要です。

```sh
bun install --frozen-lockfile
bun run typecheck
bun run lint
bun run lint:types
bun run format:check
bun run test
bun run test:flow
python3 scripts/build_attachment_flow.py --check
```

`bun run test` は Vitest と公式 `@cloudflare/vitest-plugin` を使い、`vitest.config.ts` から `wrangler.jsonc` を読みます。Hono の `app.request`、MCP ツール、文書パーサーを Workers ランタイム上でテストし、フロー応答はモックします。`bun run test:flow` は Python の unittest による生成ソース・契約・秘匿化の検証です。いずれもライブフロー、メールボックス、OneDrive、デプロイを必要としません。

`--check` は生成 JSON を書き換えず、現在の生成結果との一致を調べます。生成ソースを意図的に変更するときだけ、次を使います。

```sh
python3 scripts/build_attachment_flow.py
python3 scripts/build_attachment_flow.py --check
bun run test:flow
```

フィクスチャは変更せず、生成 JSON を直接編集しないでください。コードの整形が必要なら既存コマンド `bun run format` を使います。`format` / `format:check` の対象は `src` で、Markdown の校正やリンク確認の代わりにはなりません。

## 3. Power Automate を手動で配線する

詳細は [既存フローの更新手順](../power-automate/microsoft-bypass-flow/README.md)を参照してください。公開ソースは既存フローを拡張するための WDL 定義であり、インポート用 ZIP や Dataverse ソリューションではありません。

**同梱されていないもの:** 新規環境への初期フロー作成手順、実接続、認証資格情報、コールバック URL、Cloudflare / OAuth の構築一式。既存フローを持っていない場合、このリポジトリだけで即時導入できるとは考えず、許可された担当者と初期構築を別途設計してください。

既存フローでは、次の配線を確認します。

| 箇所 | 必要な設定 |
| --- | --- |
| HTTP トリガー | `manual`（`Request` / `Http`）。既存の認証方式・トリガー条件を保持 |
| ゲートウェイキー | フローの `McpGatewayKey`（`SecureString`）を非公開で設定。Worker の `POWER_AUTOMATE_GATEWAY_KEY` と一致させる |
| 操作分岐 | `スイッチ` が `@triggerBody()?['operation']` を参照。14 操作とトリガーの enum を一致させる |
| Outlook | `shared_office365` を既存の Office 365 Outlook 接続へ。固定 `HttpRequest` 操作のみ |
| OneDrive | `shared_onedriveforbusiness` を所有者の既存 OneDrive for Business 接続へ |
| OneDrive 検索 | `OneDriveSearchRootId` と `OneDriveSearchMode` の Compose 入力を、対象デザイナーで確認した固定値へ |
| 実行履歴 | トリガーとデータを扱うアクションの安全な入出力設定を保持 |

Worker はフロー URL に POST し、キーを `X-MCP-Gateway-Key` ヘッダーで送ります。トリガー条件はキーが空でないことと一致を確認してからコネクタ処理を許可します。公開ソースの `McpGatewayKey` は空のため、そのままでは処理を許可しません。認証方式 `triggerAuthenticationType: "All"` の維持は、無認証公開を推奨する意味ではありません。URL に含まれる署名等も資格情報として扱います。

`OneDriveSearchRootId` / `OneDriveSearchMode` も初期値は空です。どちらかが未設定なら検索はコネクタに接続する前に HTTP 503 / `ONEDRIVE_SEARCH_NOT_CONFIGURED` で停止します。`FindFiles.findMode` の機械値を推測したり、ルート ID の代わりに URL を入れたりしないでください。これらはフロー内の設定で、Worker の環境変数ではありません。

## 4. Worker の 4 バインディングを設定する

実装が読むバインディングは次の 4 つだけです。

| 名前 | 必須条件 | 値と役割 |
| --- | --- | --- |
| `POWER_AUTOMATE_URL` | 常に必要 | 保存後に非公開で確認した HTTP トリガーの完全なコールバック URL |
| `POWER_AUTOMATE_GATEWAY_KEY` | 常に必要 | フローの `McpGatewayKey` と同じキー |
| `TEAM_DOMAIN` | 本番で必要 | Access のチーム URL。例: `https://<team>.cloudflareaccess.com` |
| `POLICY_AUD` | 本番で必要 | 対象 Access アプリケーションの AUD タグ |

ローカルでは `.dev.vars.example` を参考に、非公開の `.dev.vars` を用意します。サンプルには最初の 2 つだけが載っています。実 URL・キー・Access 設定を Git に追加しないでください。本番の値は Wrangler の secrets / vars 管理で非公開に設定します。公開 `wrangler.jsonc` には実際のバインディングや Access 構築は含まれません。

`TEAM_DOMAIN` と `POLICY_AUD` の**両方を省略すると、Worker 内の認証ミドルウェアは通過を許可します**。これは Access が前段にないローカル開発用です。本番で省略しないでください。片方だけの設定は `Incomplete Cloudflare Access config: set both TEAM_DOMAIN and POLICY_AUD` で失敗します。実装は実行環境を自動判別して本番の設定漏れを防ぐわけではありません。

## 5. Cloudflare Access とクライアント認証を別途準備する

`/mcp` の前段に Cloudflare Access のアプリケーションとポリシーを設け、許可した利用者だけが到達できるようにします。外部 OAuth 認証、クライアント登録・必要なメタデータや接続設定は、選択した Access / OAuth 構成と MCP クライアントに合わせて別途確認する必要があります。リポジトリにそのインフラ定義、OAuth サーバー、クライアント設定のひな形、自動登録処理はありません。

Worker 側が実装するのは、多層防御としての **`Cf-Access-Jwt-Assertion` の検証**です。

- `TEAM_DOMAIN` の `/cdn-cgi/access/certs` から公開署名鍵を取得。
- JWT の署名、issuer（`TEAM_DOMAIN`）、audience（`POLICY_AUD`）と有効期限を検証。
- ヘッダーがない、または検証に失敗した場合は HTTP 401 / `Unauthorized`。

一般的な OAuth の Bearer トークンを送るだけでこの条件を満たすとは限りません。クライアントから Access を経由したリクエストに有効な Access JWT が渡る構成であることを確認してください。接続できないからといって Access や JWT 検証を外したり、アプリ独自の認証を追加したりしないでください。Worker の別ホスト名・公開経路で前段保護を迂回できないかも確認します。

`GET /` はサーバー名・バージョン・ツール名を返す情報エンドポイントです。Worker の JWT ミドルウェアの対象は `/mcp` で、`GET /` の応答だけでは MCP 認証や Microsoft 側の接続成功を確認できません。

## 6. 起動と、公開前の確認

ローカル起動は次のとおりです。

```sh
bun run dev
```

**起動後のツール呼び出しは、設定した実フローと Microsoft 365 に接続し得ます。** `bun run test` のモックとは違います。特に下書きツールは実データを書き込みます。ローカル起動や公開を、ライブ実行の許可と混同しないでください。

既存の公開コマンドは次ですが、Cloudflare への書き込みを伴います。認証・フロー・契約条件のレビュー後、公開が許可された運用者が実行するものです。

```sh
bun run deploy
```

公開前には、次を別途確認してください。

- 4 バインディング、Access ポリシー、クライアントの OAuth 条件、直接到達経路の保護。
- フロー保存後の URL、キー、接続所有者、OneDrive の固定設定、実行履歴の秘匿化。
- PDF.js を含む処理の CPU・メモリ・割り当て。[解析の制約](attachments.md#resource-limits)に従い、計画した用途に見合う試験を許可された範囲で行う。
- 下書き作成・添付のユーザー承認と、曖昧な失敗時に Outlook を確認する運用。
- MCP ホスト / AI サービスと Microsoft 側のデータ保持・監査・外部提供の扱い。

既存資料のデザイナー確認や合成テストの成功記録は、過去の限定的な確認です。本番接続や大規模・代表的なファイル群の検証を保証するものではありません。
