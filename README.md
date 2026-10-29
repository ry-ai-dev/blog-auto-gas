# blog-auto-gas

## 目的

Google スプレッドシートにキーワードを入力すると、Claude API で記事本文を生成し、
WordPress へ自動投稿するための GAS (Google Apps Script) プロジェクト。

スプレッドシートの各行がシート「記事」の1件分のタスクに対応し、ステータス列を
使って「未生成 → 生成中 → 生成済み → 投稿済み」の流れを管理する。

## フォルダ構成

```
blog-auto-gas/
├── README.md          このファイル
├── .gitignore          認証情報・ローカル生成物を除外
├── docs/
│   └── dev-notes.md    開発の進め方・操作メモ(段階ごとの記録)
├── src/
│   └── Code.gs         GAS 本体スクリプト
└── test/
    └── keywords.csv     動作確認用のキーワードサンプル(スプレッドシートにインポート)
```

## 手順

(後で追記する)
