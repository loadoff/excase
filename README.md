excase
======

Excelファイルにテストケースと結果を出力するためのライブラリ

[![godoc](https://godoc.org/github.com/loadoff/excl?status.svg)](https://godoc.org/github.com/loadoff/excase)
[![CircleCI](https://circleci.com/gh/loadoff/excl.svg?style=svg)](https://circleci.com/gh/loadoff/excase)
[![go report](https://goreportcard.com/badge/github.com/loadoff/excl)](https://goreportcard.com/report/github.com/loadoff/excase)

## Description

Excelファイルにテストケースと結果を出力するために使用する
大項目、中項目、小項目、テストケース、合格条件、合否、実行日時などを設定すると
テストケースと合否を出力することができる

## Usage

```go
// ケースの作成準備
ex := excase.InitExCase()
// ケースの作成
sec := ex.OpenSection("テストシート名")
// 大項目のセット
sec.Large("大項目1")
// 中項目のセット
sec.Middle("中項目1-1")
// 小項目のセット
sec.Small("小項目1-1-1")
// テストの内容と合格条件をセット
test := sec.Test("テスト内容1-1-1-1", "合格条件1-1-1-1")
// 合格をセット
test.Passed()
// 次の行に小項目とテスト内容と合格条件のみセット
sec.Small("小項目1-1-2")
test = sec.Test("テスト内容1-1-1-2", "合格条件1-1-1-2")
// 不合格をセット
test.Failed()
// チェーンして書くことも可能
sec.Middle("中項目1-2").Small("小項目1-2-1").Test("テスト内容1-2-1-1", "合格条件1-2-1-1").Passed()
// セクションを閉じる
sec.CloseSection()
// テスト出力終了
ex.Close()
```
出力イメージ
![top-page](https://raw.githubusercontent.com/loadoff/excase/images/screen1.png)

作成されるファイルのパスを確認する[ex.FilePath]に保管されてる
```go
ex := excase.InitExCase()
fmt.Println(ex.FilePath)
ex.Close()
```

## Install

```bash
$ go get github.com/loadoff/excl
$ go get github.com/loadoff/excase
```

## Licence

[MIT](https://github.com/loadoff/excase/LICENCE)

## Author

[YuIwasaki](https://github.com/loadoff)