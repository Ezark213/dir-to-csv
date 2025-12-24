# ディレクトリスキャナー

フォルダ構造をスキャンしてCSV形式で出力するツールです。

## 使い方

### ダブルクリックで起動

`ディレクトリスキャナー.hta` をダブルクリックするだけ！

### 3ステップで完了

1. **フォルダを指定** - パスを入力するか「参照」ボタンで選択
2. **スキャン開始** - ボタンをクリック
3. **結果を確認** - CSVファイルが生成されます

## 対応するパス

| 種類 | 例 |
|------|-----|
| ローカルフォルダ | `C:\Users\...` |
| 社内サーバ | `\server\share` |
| Dropbox | `C:\Users\...\Dropbox\...` |
| OneDrive | `C:\Users\...\OneDrive\...` |
| 外付けHDD | `D:\...` |

## ファイル構成

### 必須ファイル（配布時に必要）

| ファイル | 必須 | 説明 |
|---------|:----:|------|
| `ディレクトリスキャナー.hta` | ○ | アプリ本体（これをダブルクリック） |
| `directory_scanner.ps1` | ○ | スキャン処理スクリプト |

```
配布フォルダ/
├── ディレクトリスキャナー.hta  ← 必須
└── directory_scanner.ps1       ← 必須
```

### 自動生成されるファイル（配布不要）

以下のファイルはスキャン実行時に自動生成されます。**配布時には含めないでください。**

| ファイル | 説明 | 削除可能 |
|---------|------|:--------:|
| `config.json` | 設定ファイル（毎回上書き） | ○ |
| `directory_structure.csv` | スキャン結果 | ○ |
| `error.log` | エラーログ | ○ |

> **Note:** `config.json`はGUIで設定した内容をPowerShellに渡すために自動生成されます。手動で作成・編集する必要はありません。

## 出力形式（CSV）

```csv
level1,level2,level3,level4,level5,level6,level7,level8,filename,fullPath
営業部,2024年度,報告書,,,,,売上報告.xlsx,\\server\営業部\2024年度\報告書\売上報告.xlsx
```

- 最大8階層まで対応
- Excelで開いても文字化けしません（BOM付きUTF-8）

## 詳細設定

アプリ画面の「詳細設定を表示」から設定できます：

| 項目 | 説明 | 初期値 |
|-----|------|-------|
| 最大階層数 | スキャンする深さ | 8 |
| 除外フォルダ | スキップするフォルダ名 | .git, node_modules等 |
| 出力ファイル名 | CSVファイルの名前 | directory_structure.csv |

## config.json について

`config.json`はスキャン実行時に自動生成される設定ファイルです。

### 構造

```json
{
    "rootPath": "C:\Users\example\\Documents",
    "outputPath": "directory_structure.csv",
    "errorLogPath": "error.log",
    "maxDepth": 8,
    "excludeDirs": [".git", "node_modules", "__pycache__", ".venv"],
    "excludeFiles": [".DS_Store", "Thumbs.db"],
    "includeHidden": false,
    "encoding": "UTF8",
    "logLevel": "Info"
}
```

### 各フィールドの説明

| フィールド | 必須 | 型 | 説明 |
|-----------|:----:|-----|------|
| `rootPath` | ○ | string | スキャン対象のルートフォルダパス |
| `outputPath` | ○ | string | 出力CSVファイル名 |
| `errorLogPath` | - | string | エラーログファイル名 |
| `maxDepth` | - | number | 最大階層数（1〜8） |
| `excludeDirs` | - | string[] | 除外するフォルダ名のリスト |
| `excludeFiles` | - | string[] | 除外するファイル名のリスト |
| `includeHidden` | - | boolean | 隠しファイル/フォルダを含めるか |
| `encoding` | - | string | 出力エンコーディング |
| `logLevel` | - | string | ログレベル（Debug/Info/Warning/Error） |

### 注意

- **通常利用ではconfig.jsonの手動編集は不要です**（GUIから設定可能）
- PowerShellスクリプトを直接実行する場合のみ、事前にconfig.jsonを作成してください

## 配布方法

以下の2ファイルをセットで配布してください：

- `ディレクトリスキャナー.hta`
- `directory_scanner.ps1`

## 動作環境

- Windows 10 / 11
- 管理者権限は不要

## トラブルシューティング

### 「スクリプトの実行中にエラーが発生しました」と表示される

エラー画面にデバッグ情報が表示されます。以下を確認してください：

| 確認項目 | 対処法 |
|---------|--------|
| `appPath`が空または不完全 | HTAファイルを別の場所にコピーして実行 |
| `psPath`のファイルが存在しない | `directory_scanner.ps1`が同じフォルダにあるか確認 |
| ネットワークドライブから実行 | ローカルにコピーしてから実行 |

### SharePoint / OneDrive 同期フォルダから実行する場合

同期フォルダ内で実行するとパス取得に失敗することがあります。その場合：

1. フォルダごとローカル（例：`C:\Tools\ディレクトリスキャナ`）にコピー
2. コピー先から実行

### PowerShellスクリプトがブロックされる

ダウンロードしたファイルはWindowsにブロックされる場合があります：

1. `directory_scanner.ps1`を右クリック → プロパティ
2. 「ブロックの解除」にチェック → OK

## 注意事項

- アクセス権限のないフォルダはスキップされます
- 大量のファイルがある場合、処理に時間がかかることがあります
