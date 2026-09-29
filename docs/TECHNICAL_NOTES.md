# 技術ノート（Outlook のオブジェクトモデルについて実測で分かったこと）

作成日: 2026-09-29
使用AIエージェント: Claude Code
使用モデル: Claude Opus 5.5 / claude-opus-5-5

ここに書いたことは、次の環境で観測した挙動です。すべての Outlook 環境で同じになるとは限りません。

- Windows 11 x64、Outlook（クラシック）x64（Microsoft 365 版 16.0.20430）
- POP / IMAP の .pst ストア、クライアント側ルール約 340 件

## 保存の仕組みと順序

- `Store.GetRules()` で得た `Rules` への変更（追加・削除・並べ替え・プロパティ変更）は、`Rules.Save()` を呼ぶまで Outlook に保存されない。
  `Rules` を取り直すと保存前の状態に戻っている。途中で失敗したときに Save を呼ばなければ何も残らない。
- **`Rules.Remove()` で消したルールのオブジェクトに触ると、Outlook が異常終了する**（access violation。実際に落ちた）。
- `Rules.Create()` で作ったルールは実行順 1（先頭）に入る。
- `Rule.ExecutionOrder` を変えると、`Rules` の添字もすぐに並び替わる。
- 同じルールを `Rules[i]` で取り直すと同じ RCW（.NET 側の COM ラッパー）が返るので、`ReferenceEquals` で同一性を判定できる。

これらを踏まえて、保存は次の順で行う（`OutlookRuleGateway.Apply`）。

1. `GetRules()` を取り直し、件数と各位置の名前が読み込み時と一致するか確認する（違えば中止）
2. 既存ルールの名前・移動先・有効/無効を変える
3. 複製を `Rules.Create()` で作り、条件・例外・処理を書き写す。**読み戻して元のルールと中身を照合**する
4. 残すルールを 1 番から順に `ExecutionOrder` で並べる（削除するルールは末尾に押し出される）
5. 末尾のルールが削除対象であることを `ReferenceEquals` で確かめてから `Remove()`。以後そのオブジェクトには触らない
6. 並びを `ReferenceEquals` で最終確認する
7. `Rules.Save()`

`Apply` の `dryRun` 引数を true にすると 7 だけを行わない。実機で 1〜6 が通ること（複製 3 件の照合を含む）を確認した。

## 遅延バインディング（dynamic）で呼ぶ

相互運用アセンブリ（NuGet の `Microsoft.Office.Interop.Outlook`）は使わず、`Type.GetTypeFromProgID("Outlook.Application")` と
`dynamic` で呼んでいる。このパッケージは個人が再パッケージしたもので、説明に「サポート対象外・ライセンスなし」とあるため。

- 読み込み速度は早期バインディングと同等だった（20 件の平均で 早期 208 ms/件、dynamic 196 ms/件）。
- **オブジェクトや配列を値に持つプロパティへの代入は、`dynamic` だと失敗する**
  （例: `rule.Actions.MoveToFolder.Folder = folder` が「操作は失敗しました」0x80020009）。
  `Type.InvokeMember(名前, BindingFlags.SetProperty, ...)` で IDispatch の PROPERTYPUT を直接呼ぶと成功する（`OutlookCom.SetProperty`）。
  早期バインディングでは同じ代入が成功する。文字列・真偽値・数値の代入は `dynamic` のままで問題ない。

## 読み込み: オブジェクトモデルで 1 件ずつ読むと遅い

- 1 ルールあたり、条件 31・例外 31・処理 28 個のオブジェクトがあり、どれが有効かは 1 つずつ問い合わせないと分からない。
- オブジェクトを初めて触るときが遅い（条件 31 個の走査: 初回 111.5 ms/件、2 回目 39 ms/件）。
- 呼び出し側を早期/遅延バインディング、STA/MTA に変えても 1 件あたり 0.2〜0.35 秒で縮まらなかった（342 件で約 100 秒）。

## 読み込み: ルールのまとめデータを直接読む（高速読み込み）

Outlook は全ルールを、受信トレイの隠しメッセージ（メッセージクラス `IPM.RuleOrganizer`）の
`PR_RW_RULES_STREAM`（`0x68020102`、バイナリ）にまとめて保存している。

- 取り出し方: 受信トレイの `GetTable("", olHiddenItems)` で `IPM.RuleOrganizer` の EntryID を探し、
  `Namespace.GetItemFromID(EntryID, StoreID)`（StorageItem が返る）の `PropertyAccessor.GetProperty` で読む。
  342 件分 535,714 バイトを約 1 秒で取得できた。
  - `Folder.GetStorage("IPM.RuleOrganizer", olIdentifyByMessageClass)` は「見つからない」になった。
  - EntryID を渡した `GetStorage` は既存のものを返さず、空の新規 StorageItem を返した（保存しなければ残らない）。
  - テーブルの列にこのプロパティを加えると Outlook が「メモリ不足」を返した。
- 形式はルールのエクスポート（.rwz）と同じ。同じ時点の .rwz と比べると、サイズが同じで、
  違いはルールごとの 4 バイト（ルール間の区切りとルールの署名。.rwz は `0x00140000`、まとめデータは `0x060F4240`）だけだった。
- 形式は非公開だが、[asklar/rwzreader](https://github.com/asklar/rwzreader) と
  [hughbe/OutlookRulesReader](https://github.com/hughbe/OutlookRulesReader)（どちらも MIT）の解析結果をもとに C# で実装した（`RulesStream.cs`）。
  342 件の解析は 16 ms。
- 要素には長さの情報がないため、知らない要素があるとそれ以降を読めない。その場合は 1 件ずつの読み込みに切り替える。

まとめデータだけでは決まらないものは Outlook に問い合わせる（1 件あたり数 ms）。

| 項目 | 理由 |
|---|---|
| 各ルールの名前・有効/無効 | まとめデータが Outlook 側と食い違っていないかの照合。1 件でも違えば 1 件ずつの読み込みに切り替える |
| `Rule.IsLocalRule` | 「このコンピューターのみ」の条件が無くても、サウンド・新着通知の処理を含むと true になる |
| 「このコンピューターのみ」がこのPCか別のPCか | まとめデータにはコンピューターの識別子しかない。`Conditions.OnLocalMachine.Enabled` で確かめる |
| 移動・コピー先のフォルダーのパス | まとめデータには EntryID とフォルダー名しかない。`GetFolderFromID` で解決し、見つからなければエラーのルール（同じフォルダーは 1 回だけ問い合わせる） |

条件・例外・処理の並びは、オブジェクトモデルが返す順（種類ごとに固定）に並べ直している。

実機での突き合わせ（342 件）: 高速読み込み 2.3 秒、1 件ずつの読み込み 99.4 秒。
名前・有効/無効・種類・クライアントのみ・条件・例外・処理（宛先のアドレス、語句、移動先の EntryID とパス）・
宛先の解決状態・診断結果のすべてが一致した。

## エラーの判定

- 移動・コピーの処理が有効なのに `MoveToFolder.Folder`（`CopyToFolder.Folder`）が null または例外になる → 移動先のフォルダーが消えている。
  Outlook の画面で「エラー」と表示されるルールの典型的な原因。
- Exchange のサーバー側ルールには `PR_RULE_MSG_STATE` の `ST_ERROR` フラグがあるが、.pst のクライアント側ルールには無いため使っていない。

## 「実行されないルール」の判定

先に実行される有効なルール A が「処理を中止」を持ち、例外が無く、A のすべての条件が後のルール B の条件に包含されるなら、
B に一致するメールは必ず A で止まるので B は実行されない。包含の判定:

| 条件の種類 | A が B を包含する条件 |
|---|---|
| 差出人・宛先（アドレスのいずれかに一致） | B のアドレス ⊆ A のアドレス |
| 件名・本文などの語（いずれかを含む） | B の各語が A のいずれかの語を含む |
| 分類項目 | B の項目 ⊆ A の項目 |
| 値のない条件（添付ファイルあり など） | 同じ種類が B にもある |
| その他 | 値の集合が一致 |

「このコンピューターのみ」は実行場所の指定なので比較から外す。A が差出人 1 条件だけのときは、B の差出人の一部だけが重なる場合も個別に警告する。

## ストアの種類

- ルールを扱えるのはルールに対応したストア（今回は POP/IMAP の .pst）のみ。予定表だけのストアなどは
  `GetRules()` が「このストアではルールがサポートされていません」になる。
