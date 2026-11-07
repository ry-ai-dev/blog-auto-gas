# blog-auto-gas

## 概要

Google スプレッドシートにキーワードを入力すると、Claude API で記事本文を生成し、
内容を人が確認した上で、WordPress に「下書き」として保存できる GAS (Google Apps Script) ツールです。

- WordPress への投稿は常に「下書き(draft)」です。即時公開はできない仕様です。
- 「記事生成」と「WordPressへの下書き保存」は別の操作です。生成後は必ず人が内容を確認してから下書き保存を行ってください。

## 全体の流れ

```
①キーワードを入力       ②記事を生成            ③人が内容を確認       ④WordPressに下書き保存
(シート「記事」の         (メニュー「記事生成」    (タイトル・本文・       (メニュー「WordPressに
 A列・B列に入力)          →「選択行の記事を       文字数・エラー欄       下書き保存(選択行)」)
                          生成」)                 を確認)
```

1. シート「記事」に、メインキーワード(A列)とサブキーワード(B列、省略可)を入力する
2. 記事にしたい行を選択し、メニュー「記事生成」→「選択行の記事を生成」を実行する
3. 生成されたタイトル・本文・文字数・ステータス・エラー欄を人が確認する
4. 問題がなければ、同じ行を選択したまま「WordPressに下書き保存(選択行)」を実行する

## ドキュメント一覧

| ドキュメント | 内容 |
| --- | --- |
| [docs/setup-guide.md](docs/setup-guide.md) | 初期設定手順(スプレッドシートのコピー、APIキー登録、WordPress側の準備など) |
| [docs/user-manual.md](docs/user-manual.md) | 日々の使い方(最短手順) |
| [docs/troubleshooting.md](docs/troubleshooting.md) | トラブルシューティング(症状→原因→対処) |
| [docs/delivery-checklist.md](docs/delivery-checklist.md) | 納品前チェックリスト |
| [docs/dev-notes.md](docs/dev-notes.md) | 開発時の作業メモ(開発者向け、実測値など) |

## フォルダ構成

```
blog-auto-gas/
├── README.md                  このファイル
├── .gitignore                  認証情報・ローカル生成物を除外
├── docs/
│   ├── dev-notes.md            開発の進め方・操作メモ(段階ごとの記録、開発者向け)
│   ├── setup-guide.md          初期設定手順書
│   ├── user-manual.md          日々の使い方
│   ├── troubleshooting.md      トラブルシューティング
│   └── delivery-checklist.md   納品前チェックリスト
├── src/
│   └── Code.gs                  GAS 本体スクリプト
└── test/
    └── keywords.csv              動作確認用のキーワードサンプル(スプレッドシートにインポート)
```

## 技術メモ(使用技術・制限)

**使用技術**
- Google Apps Script(スプレッドシートに紐づくスクリプト)
- Claude API(Anthropic社のMessages API。記事本文の生成に使用)
- WordPress REST API(`/wp-json/wp/v2/posts`。下書き作成にのみ使用。アプリケーションパスワードによるBasic認証)

**制限事項**
- 1回の操作で処理されるのは、選択している1行のみです。複数行をまとめて処理する機能はありません(大量の行を一度に処理すると、GASの実行時間の上限に達する可能性があるため)。
- 本文の文字数が2,000〜3,000字の範囲に収まらない場合、ステータスが「エラー」になり「生成済み」にはなりません。内容を見直して再実行してください。
- WordPressへの投稿は常に下書き(draft)です。公開(publish)はコード上できない仕様になっています。
- 料金目安: 1記事あたり約10円(入力約1,900トークン / 出力約6,000トークンでの概算。実際の単価はAnthropic公式サイトの料金表で要確認)。月10本なら約100円、月50本なら約500円が目安です。
