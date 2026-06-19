# mask-pii

日本語テキストの個人情報（氏名・住所・部門名）を自動でマスキングするツールです。  
Excelファイルの問い合わせ内容をまとめて処理することができます。

> 📌 このREADMEは **Pythonを使ったことがない方** を対象に書いています。  
> 上から順番に進めていけばそのまま動かせます。

---

## 目次

1. [このツールでできること](#1-このツールでできること)
2. [Pythonのインストール](#2-pythonのインストール)
3. [ライブラリのインストール](#3-ライブラリのインストール)
4. [ファイルの準備](#4-ファイルの準備)
5. [使い方](#5-使い方)
6. [出力結果の見方](#6-出力結果の見方)
7. [トラブルシューティング](#7-トラブルシューティング)

---

## 1. このツールでできること

顧客からの問い合わせテキストに含まれる個人情報を、AIを使わず**完全ローカル**でマスキングします。

**入力例:**
```
お世話になっております。A部門の佐藤です。
神奈川県横浜市に在住なのですが、NW速度が遅く、ご相談したいです。
```

**出力例:**
```
お世話になっております。[部門名]の[氏名]です。
[住所]に在住なのですが、NW速度が遅く、ご相談したいです。
```

**マスク対象:**

| 種別 | 例 |
|------|----|
| 氏名 | 佐藤、田中太郎 |
| 住所 | 神奈川県横浜市〇〇 |
| 部門名 | 営業部、情報システム部 |
| 電話番号 | 090-1234-5678 |
| メールアドレス | taro@example.com |
| 郵便番号 | 〒123-4567 |
| 生年月日 | 昭和60年4月1日、1985/04/01 |

---

## 2. Pythonのインストール

Python とは、このツールを動かすために必要なプログラムです。  
まず、すでに入っているか確認します。

### 確認方法

スタートメニューから「PowerShell」を開いて以下を実行してください。

```powershell
python --version
```

`Python 3.x.x` と表示されれば**インストール済みです**。[手順3](#3-ライブラリのインストール)へ進んでください。

> ⚠️ 「python」と打つと **Microsoft Storeが開く場合** は未インストールです。次の手順へ進んでください。

---

### インストール方法

#### 方法A: winget（推奨・管理者権限不要）

PowerShell で以下を1行実行するだけです。

```powershell
winget install Python.Python.3.12
```

完了後、**PowerShellを一度閉じて再度開いてください**（設定を反映するため）。

---

#### 方法B: 公式インストーラー（GUIで操作したい場合）

1. ブラウザで `https://www.python.org/downloads/` を開く
2. 「Download Python 3.x.x」をクリックしてダウンロード
3. ダウンロードした `.exe` ファイルを実行
4. **最初の画面で「Add Python to PATH」に必ずチェック** ✅  
   ⚠️ ここを忘れると `python` コマンドが動きません
5. 「Install Now」をクリック
6. 完了後、**PowerShellを一度閉じて再度開く**

---

### インストール確認

```powershell
python --version
pip --version
```

以下のように表示されれば成功です。

```
Python 3.12.x
pip 24.x.x from C:\Users\<ユーザー名>\AppData\Local\Programs\Python\...
```

---

## 3. ライブラリのインストール

ライブラリとは、ツールが使う追加機能のことです。  
以下のコマンドを PowerShell で**一度だけ**実行してください。

### 基本ライブラリ

```powershell
pip install openpyxl ginza ja-ginza
```

| ライブラリ | 用途 |
|-----------|------|
| `openpyxl` | Excelファイルの読み書き |
| `ginza` | 日本語NLPフレームワーク |
| `ja-ginza` | GiNZA 日本語モデル（標準） |

> ⏳ `ja-ginza` は日本語解析モデルのため、ダウンロードに数分かかります。

---

### 高精度モデルのインストール（推奨）

氏名・住所の検出精度が上がります（約400MB、時間がかかります）。

```powershell
pip install ja-ginza-electra
```

---

### インストール確認

```powershell
python -c "import spacy; spacy.load('ja_ginza'); print('OK')"
```

`OK` と表示されれば準備完了です。

---

## 4. ファイルの準備

以下のファイルを**同じフォルダ**に置いてください。

```
任意のフォルダ/
├── mask_pii.py           ← メインスクリプト（このリポジトリからDL）
├── 問い合わせ.xlsx        ← 処理したいExcelファイル
└── 部門一覧.csv           ← 部門名リスト（任意）
```

### Excelファイルの形式

**A列**に問い合わせ内容を1行1件で入力してください。  
セル内の改行（Alt+Enter）はそのまま処理されます。

| A列（問い合わせ内容） |
|-----------------------|
| お世話になっております。A部門の佐藤です。... |
| 田中と申します。東京都渋谷区に... |

### 部門一覧CSVの形式

1列目に部門名を1行1件で記載します（ヘッダー行は不要）。

```
営業部
経理部
情報システム部
カスタマーサポート部
```

---

## 5. 使い方

PowerShell を開き、スクリプトを置いたフォルダに移動してから実行します。

### フォルダへの移動方法

```powershell
# Cドライブの「work」フォルダに置いた場合
cd C:\work\mask-pii
```

> 💡 エクスプローラーでフォルダを開き、アドレスバーに `powershell` と入力してEnterを押すと、そのフォルダで PowerShell が開きます。

---

### パターン1: Excelファイルをまとめて処理（最もよく使う）

```powershell
python mask_pii.py --excel 問い合わせ.xlsx
```

部門名CSVも使う場合:

```powershell
python mask_pii.py --excel 問い合わせ.xlsx --dept-csv 部門一覧.csv
```

出力ファイルは自動的に `masked_問い合わせ.xlsx` という名前で同じフォルダに保存されます。

---

### パターン2: 文章を直接入力して試す

```powershell
python mask_pii.py --text "田中太郎（営業部）神奈川県横浜市1-2-3"
```

---

### パターン3: 出力先ファイル名を指定する

```powershell
python mask_pii.py --excel 問い合わせ.xlsx --excel-out 結果_20260620.xlsx
```

---

### オプション一覧

| オプション | 説明 |
|------------|------|
| `--excel <ファイル>` | 入力Excelファイルを指定 |
| `--excel-out <ファイル>` | 出力Excelファイル名を指定（省略可） |
| `--dept-csv <ファイル>` | 部門名CSVを指定（省略可） |
| `--text "テキスト"` | テキストを直接指定して処理 |
| `--no-ner` | AI解析をオフにする（高速・低精度） |
| `--quiet` | ログ出力を抑制する |

---

## 6. 出力結果の見方

### 画面への出力

```
3 行を処理します...
  行   2: 3 件マスク
  行   3: 2 件マスク
  行   4: 1 件マスク

マスク済みExcelを保存しました: masked_問い合わせ.xlsx
完了: 合計 6 件をマスクしました。
```

### 出力Excelの構成

| 列 | 内容 |
|----|------|
| A列 | 元の問い合わせ（変更なし） |
| B列 | マスク済みテキスト |
| C列 | マスクした件数 |

---

## 7. トラブルシューティング

### `python` が認識されない

```
'python' は、内部コマンドまたは外部コマンドとして認識されていません
```

**対処:** PowerShellを閉じて再度開く。それでもダメな場合は再インストール。

```powershell
winget install Python.Python.3.12
```

公式インストーラーを使った場合は **「Add Python to PATH」** のチェックを確認してください。

---

### `pip` が認識されない

```powershell
python -m pip install --upgrade pip
```

---

### ライブラリのインストールで権限エラーが出る

```powershell
pip install --user openpyxl ginza ja-ginza
```

---

### GiNZAモデルが見つからないエラー

```
OSError: [E050] Can't find model 'ja_ginza'
```

```powershell
pip install ginza ja-ginza
```

---

### PowerShellでスクリプトの実行がブロックされる

```
このシステムではスクリプトの実行が無効になっています
```

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

確認メッセージが出たら `Y` を入力してEnterを押します。

---

## 動作環境

- Python 3.10 以上
- Windows 11

---

## 参考リンク

- Python公式: `https://www.python.org/`
- GiNZA（日本語NLP）: `https://megagonlabs.github.io/ginza/`
- openpyxl: `https://openpyxl.readthedocs.io/`
