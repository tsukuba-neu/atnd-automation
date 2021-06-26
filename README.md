# 予定調査自動化ツール

![](https://img.shields.io/badge/tsukuba-neu-blue?style=flat-square)

Google スプレッドシートで構築する予定調査の更新を自動化するツール

## Feature

### 日付ベースの行を自動更新（`updateDateRows()`）

A 列に日付を含むシートについて、過去の日付の行を削除し、`MAX_DAYS_LENGTH`（コード内で定義する変数）日先までの行を追加する。

## Usage

[Google Apps Script](https://developers.google.com/apps-script?hl=ja) 上で動作する。トランスパイルとデプロイに Node.js および[`@google/clasp`](https://github.com/google/clasp)を使用する。

### 依存パッケージのインストール

```sh
yarn
```

### Apps Script プロジェクトの準備

Apps Script プロジェクトを準備し、ローカルのリポジトリを接続する。

```sh
clasp settings scriptId [PROJECT_ID]
clasp settings rootDir src
```

プロジェクト ID は以下の手順で確認できる。

1. スプレッドシートの編集画面から **ツール -> スクリプト エディタ** と進み、Apps Script プロジェクトを開く。
2. URL `https://script.google.com/home/projects/[PROJECT_ID]` よりプロジェクト ID を確認する。

### コードのアップロード

```sh
clasp push
```

### 自動実行の設定

必要な機能の関数に[トリガーを設定](https://developers.google.com/apps-script/guides/triggers/installable#managing_triggers_manually)することで、時間ベースで自動的に実行させることができる。
