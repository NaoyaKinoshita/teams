# Teams Recording Monitor

Teams 会議の録画開始・停止を検知し、録画ファイルを Azure Blob Storage に自動保存するアプリ。

- 専用ユーザーアカウント不要（Bot アプリ ID のみ）
- Teams Premium 不要（Microsoft 365 E3 で動作）
- 管理センターの App Setup Policy でセキュリティグループ単位に配布可能

---

## ディレクトリ構成

```
teams/
├── meeting-app/    # Teams アプリ本体（Bot + サイドパネルタブ）
└── functions/      # 録画受信・Blob Storage 保存サーバー（FastAPI + Container Apps）
```

---

## アーキテクチャ

```
管理者が App Setup Policy で対象グループにアプリを配布
    ↓
対象ユーザーが会議を作成
    ↓
バックエンドが /chats 購読で会議チャットを自動検知
    ↓
会議チャットへアプリをインストール + チャットメッセージを購読
    ↓
録画が開始（Graph API 変更通知）
    ↓
会議サイドパネルのタブがポーリングで録画開始を検知
    ↓
タブ内に「Azure に保存する」「スキップ」ボタンを表示
    ├── 「Azure に保存する」クリック → 同意を記録
    └── 「スキップ」クリック → 保存しない
    ↓
録画が OneDrive に保存完了 → 同意済みの場合のみ recording-functions へ Webhook 通知
    ↓
recording-functions が SharePoint から録画ファイルを取得し Azure Blob Storage へ保存
```

---

## 技術スタック

| コンポーネント | 技術 |
|---|---|
| Bot サーバー | Python / FastAPI + Bot Framework SDK |
| 録画検知 | Microsoft Graph API 変更通知（`chats/{id}/messages`） |
| 会議チャット自動検知 | Graph API `/chats` 購読 |
| コンセント UI | Teams JS SDK + HTML タブ（サイドパネル・コンテンツバブル） |
| 録画保存 | Azure Blob Storage（Graph API 経由でストリーム取得） |
| デプロイ | Azure Container Apps + Azure Container Registry |

---

## 環境情報

| 項目 | 値 |
|---|---|
| テナント | PersonalDev189.onmicrosoft.com |
| Bot 名 | teams-recording-bot |
| App ID | `c12e5c31-696d-454f-b246-287b96b06632` |
| Teams アプリ名 | Recording Monitor |

---

## セットアップ

詳細は各ディレクトリの README を参照：

- [meeting-app/README.md](meeting-app/README.md) — Teams アプリ本体のセットアップ手順
- [functions/README.md](functions/README.md) — recording-functions のセットアップ手順

---

## 設計上の経緯・判断

### なぜ専用ユーザーアカウントを使わないのか

900 人規模での利用を想定しており、特定のユーザーアカウントに紐づく実装は運用上の負担が大きい。
Bot アプリ ID（クライアントクレデンシャル）のみで動作する設計を採用。

### なぜ Compliance Recording を使わないのか

Compliance Recording（Teams Premium 機能）はポリシーベースの自動録画に特化しており、
ユーザーが任意に録画した会議を検知する用途には適合しない。
また E3 ライセンスでは利用不可。

### なぜチャットへのメッセージ送信をやめたのか

アプリ権限（クライアントクレデンシャル）で会議チャットにメッセージを送信するには
`Teamwork.Migrate.All` という高権限が必要。これはテナント全体のデータ移行を許可する権限であり、
セキュリティ上の理由から使用を避けた。代わりにサイドパネルタブ内の UI で代替している。
