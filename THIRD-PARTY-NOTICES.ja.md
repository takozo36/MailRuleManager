# サードパーティに関する表示

作成日: 2026-09-29
使用AIエージェント: Claude Code
使用モデル: Claude Opus 5.5 / claude-opus-5-5

[English](THIRD-PARTY-NOTICES.md) | **日本語**

> これは [THIRD-PARTY-NOTICES.md](THIRD-PARTY-NOTICES.md)（英語）の日本語版です。ライセンス本文は英語の原文が有効です。

このプロジェクトは、次のサードパーティの成果を利用しています。
アプリ本体（`MailRuleManager.exe`）には、サードパーティのライブラリのバイナリを含みません。
Outlook（クラシック）は COM の遅延バインディングで呼び出しており、Outlook の相互運用アセンブリ
（`Microsoft.Office.Interop.Outlook`）は参照も再配布もしていません。

## Outlook のルール形式（C# で再実装）

`src/OutlookRuleManager.Core/RulesStream.cs` は、Outlook のルールのまとめデータ（`PR_RW_RULES_STREAM`。
`.rwz` ファイルと同じ形式）を読みます。この形式は Microsoft が公開していないため、次の 2 つのプロジェクトの
形式の解析とパーサーをもとに C# で実装しました。どちらも MIT License です。

- [asklar/rwzreader](https://github.com/asklar/rwzreader) — Copyright (c) 2021 Alexander Sklar
- [hughbe/OutlookRulesReader](https://github.com/hughbe/OutlookRulesReader) — Copyright (c) 2021 Hugh Bellamy

ライセンス本文（英語の原文。許諾文の日本語参考訳は [LICENSE.ja.md](LICENSE.ja.md) を参照）:

```
MIT License

Copyright (c) 2021 Alexander Sklar
Copyright (c) 2021 Hugh Bellamy

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## テスト専用のパッケージ

次のパッケージは `tests/OutlookRuleManager.Core.Tests` だけで使い、アプリには含まれません。

| パッケージ | バージョン | ライセンス |
|---|---|---|
| [xunit](https://github.com/xunit/xunit) | 2.9.3 | Apache-2.0 |
| [xunit.runner.visualstudio](https://github.com/xunit/visualstudio.xunit) | 2.8.2 | Apache-2.0 |
| [Microsoft.NET.Test.Sdk](https://github.com/microsoft/vstest) | 17.11.1 | MIT |

## 実行環境

アプリは .NET 8 デスクトップランタイム（Windows フォーム）の上で動きます。ランタイムは利用者が別途インストールするもので、
ランタイムを同梱しない配布形態（既定）では再配布していません。.NET と Windows フォームは MIT License です。
ランタイムを同梱する形（self-contained）で配布する場合は、その配布物に .NET ランタイム自身のライセンスと
サードパーティ表示が適用されます。

## Microsoft Outlook

実行には Microsoft Outlook が必要ですが、このプロジェクトでは配布していません。
Microsoft および Outlook は Microsoft グループの商標です。
