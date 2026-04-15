# Meeting Recording Monitor

Teams アプリとして会議に自動追加され、録画の開始・停止を検知して Azure に通知する Bot。

## bot/ との違い

| | bot/ | meeting-app/ |
|---|---|---|
| 会議への参加方法 | カレンダー購読 → Bot が参加 | アプリが自動インストール（参加不要） |
| 専用ユーザーアカウント | 必要（RECORDING_BOT_USER_ID） | **不要** |
| Teams Premium | 不要 | **不要** |
| 対象ユーザーの指定 | カレンダー購読対象のみ | **セキュリティグループ単位で管理可** |
| 配布方法 | コード設定 | **管理センターのポリシーで配布** |

---

## 動作フロー

```
管理者が App Setup Policy で対象グループにアプリを配布
    ↓
対象ユーザーが会議を作成
    ↓
アプリが会議チャットに自動インストール
    ↓
Bot が installationUpdate を受信 → チャットメッセージを購読
    ↓
録画が開始 → Adaptive Card を送信（Azure 連携するか確認）
    ├── 「Azure に連携する」押下 → 同意を記録
    └── 「スキップ」押下 → 通知のみ
    ↓
録画が停止 → 同意済みなら Azure Webhook に通知
```

---

## セットアップ手順

### 1. 環境変数の設定

```bash
cp .env.example .env
# .env を編集（bot/ と同じ APP_ID / APP_SECRET / TENANT_ID を使用可）
```

### 2. Teams アプリマニフェストの編集

[manifest/manifest.json](manifest/manifest.json) の `<NGROK_DOMAIN>` を実際のドメインに置き換える：

```bash
# 例: femur-mousiness-calamity.ngrok-free.dev
sed -i 's/<NGROK_DOMAIN>/femur-mousiness-calamity.ngrok-free.dev/g' manifest/manifest.json
```

### 3. アイコンファイルの準備

`manifest/` に以下の 2 ファイルを追加（PNG 形式）：

| ファイル | サイズ | 用途 |
|---|---|---|
| `color.png` | 192×192 px | カラーアイコン |
| `outline.png` | 32×32 px（透過） | アウトラインアイコン |

任意の画像で可。[公式サンプル](https://github.com/OfficeDev/Microsoft-Teams-Samples/tree/main/assets) からダウンロードも可。

### 4. アプリパッケージの作成

```bash
cd manifest
zip ../recording-monitor.zip manifest.json color.png outline.png
```

### 5. Teams 管理センターへのアップロード

1. [admin.teams.microsoft.com](https://admin.teams.microsoft.com) にアクセス
2. 「Teams アプリ」→「アプリの管理」→「アップロード」
3. `recording-monitor.zip` をアップロード

### 6. App Setup Policy の設定

1. 「Teams アプリ」→「セットアップ ポリシー」→「追加」
2. ポリシー名を入力（例: `RecordingMonitorPolicy`）
3. 「インストール済みアプリ」→「アプリの追加」→「Recording Monitor」を追加
4. 「会議拡張機能」→「アプリの追加」→「Recording Monitor」を追加
5. 保存

### 7. ポリシーをグループに割り当て

1. 作成したポリシーを選択
2. 「グループ ポリシーの割り当て」→「グループの追加」
3. 対象のセキュリティグループを選択して保存

> **注意**: ポリシーの反映には最大 24 時間かかる場合があります。

### 8. ローカル起動

```bash
cd meeting-app
uv run python app.py
```

---

## 環境変数

```env
MICROSOFT_APP_ID=<アプリケーション（クライアント）ID>
MICROSOFT_APP_SECRET=<クライアントシークレット>
TENANT_ID=<ディレクトリ（テナント）ID>
NOTIFICATION_URL=https://<ngrok または本番 URL>
AZURE_WEBHOOK_URL=<録画停止通知先の Azure Webhook URL>
PORT=3980
```

---

## 必要な API 権限

既存の Entra ID アプリ登録に以下が付与されていること（管理者同意済み）：

| 権限 | 用途 |
|------|------|
| `Chat.ReadWrite.All` | 会議チャットへのメッセージ送信・購読 |

---

## マニフェストの id について

[manifest/manifest.json](manifest/manifest.json) の `id` フィールド（`3f8a7c2d-...`）は Teams アプリの一意な ID です。
本番利用時は新しい GUID を生成して使用してください：

```python
import uuid
print(uuid.uuid4())
```
