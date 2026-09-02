# WorkMonitor

Teamsから出力した稼働環境一覧を解析し、
Excelファイルへ日次履歴として保存するツールです。

## 概要

入力テキストファイルから
- 拠点
- 環境名<br>
を取得し、
`WorkList.xlsx`
へ日付単位のシートを作成します。<br>
さらに前日データと比較して、
- 初回稼働日
- 連続稼働日数<br>
を自動計算します。

---

## フォルダ構成

WorkMonitor<br> 
├─ main.py<br> 
├─ README.md<br> 
├─ .gitignore<br> 
│ ├─ input<br> 
│ └─ teams_yyyymmdd.txt<br> 
│ └─ data └─ WorkList.xlsx<br>

---

## 入力ファイル

### teams_yyyymmdd.txt

例

2025年09月02日 稼働環境一覧<br>
東京 <br>
環境A<br>
大阪 <br>
環境B<br>
名古屋 <br>
環境C<br>

形式

日付<br>
拠点<br> 
環境名<br>
拠点<br> 
環境名<br>
・・・

---

## 出力ファイル

### WorkList.xlsx

作成されるシート例

input 20250902 20250901 20250831

### 日付シート構成

|列|内容       |
|--|-----------|
|A |日付       |
|B |拠点       |
|C |環境名     |
|D |開始日     |
|E |連続稼働日数|

例

|日付    |拠点 |環境名|開始日   |連続稼働日数|
|--------|----|------|--------|-----------|
|20250902|東京|環境A  |20250901|          2|
|20250902|大阪|環境B  |20250902|          1|

---

## 動作仕様

### 新規環境

前日のシートに存在しない場合

- 開始日 = 当日
- 連続稼働日数 = 1

### 継続環境

前日のシートに存在する場合

- 開始日を引き継ぐ
- 連続稼働日数を +1

---

## アラート判定

連続稼働日数が3日以上の環境を集計します。

実行結果例

処理日: 20250902 <br>
前日シート発見: 20250901<br>
WorkList.xlsx 更新完了 <br>
シート名 : 20250902 <br>
件数 : 12 <br>
3日以上連続稼働: 4件<br>

---

## 実行方法

### 必要パッケージ
pip install openpyxl

### 実行
python main.py

---

## GitHub Codespaces
Codespacesでも実行可能です。<br>
python main.py実行前に
input/teams_yyyymmdd.txt
を配置してください。

---

## 開発環境

- Python 3.x
- openpyxl
- GitHub
- GitHub Codespaces
- Visual Studio Code

---

## ライセンス

個人利用向け
