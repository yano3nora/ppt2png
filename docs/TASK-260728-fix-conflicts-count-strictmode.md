# TASK-260728: 競合検出 0 件時に StrictMode で .Count が throw する不具合の修正

260728 fix conflicts count strictmode
===

## asis

Windows ユーザーから「git bash 経由で pwsh を叩くと動かない」という報告と修正版スクリプトを受領。差分を調査した結果、git bash は無関係で、以下の合わせ技による不具合と特定した (pwsh 7.6.4 で再現確認済み)。

- `Get-ConflictingOutputFiles` は関数内で `return @(...)` しているが、PowerShell の関数出力はパイプラインで列挙されるため、競合 0 件では呼び出し側に `$null`、1 件ではスカラー (FileInfo 単体) が届く
- 本体は `Set-StrictMode -Version 2.0` (`Export-PptxPng.ps1:105`) を設定しており、StrictMode 2 以上では `$null` / スカラーへの intrinsic な `.Count` が隠されるため、`$conflicts.Count` が `The property 'Count' cannot be found on this object` で throw する
- 発生条件: `-ExistingFile` がデフォルトの `Error` かつ競合ファイルが 0 件 (= 初回実行の happy path)。競合ちょうど 1 件でも同様に throw する
- 既存の Pester テストは呼び出し側で `@()` ラップ済みの形でしか検証しておらず、本体呼び出し箇所のラップ漏れを検出できていなかった

## tobe

- デフォルト引数の初回実行 (競合 0 件) が StrictMode 下でもエラーにならず完走する
- 配列アンロールの契約 (0 件 → `$null` / 1 件 → スカラー、呼び出し側 `@()` ラップ必須) が Pester テストで固定されている

## todo

- [x] 再現テストを Pester へ追加 (`tests/Export-PptxPng.Tests.ps1` の `Get-ConflictingOutputFiles` へ Context「戻り値の配列アンロール」を追加)
- [x] 呼び出し側を `@()` ラップ (`Export-PptxPng.ps1` の ExistingFilePolicy = Error 判定箇所、ユーザー修正版と同一の変更 + 意図コメント)
- [x] 他の `.Count` 使用箇所を監査 → すべて `$null` ガード / `@()` ラップ / Mandatory 型付きパラメーター済みで問題なし
- [ ] Windows 実機での受け入れ再確認 (デフォルト引数・空の出力先での実行) — 人間実施

## testcases

- [x] 競合 0 件の戻り値が `$null` になる (関数内 `@()` が呼び出し側を守らないことの固定)
- [x] 競合 1 件の戻り値がスカラー (FileInfo 単体) になる
- [x] StrictMode 2.0 下でラップなしの `.Count` が throw する
- [x] StrictMode 2.0 下でも `@()` ラップで 0 件 / 1 件とも安全に数えられ、`[0].Name` も取れる
- [x] 呼び出し箇所を直接守るテスト: `Invoke-Ppt2PngExport` を `New-PowerPointApplication` の Mock (sentinel throw) 付きで実行し、競合 0 件で COM 初期化まで到達 / 競合 1 件で InvalidArgument (2) + ファイル名例示を確認 (Codex レビュー指摘対応)
- [x] ミューテーション確認: 本体の `@()` を一時的に外すと上記 2 テストが本番同様のエラーで fail することを確認
- [x] 既存テスト含む全 37 件パス (pwsh 7.6.4 / macOS)

## notes

- 修正はユーザー提供版 (`@()` ラップ 1 行) をそのまま採用し、意図コメントのみ追記した
- 関数側の `return @(...)` は残置。除去しても挙動は同じだが、変更を最小にするため触らない
- **教訓**: この不具合はデフォルト引数の初回実行で必ず落ちるため、受け入れテスト (SPEC AT-001〜010) がデフォルトポリシー + 空出力先の happy path をカバーできていなかった可能性が高い。AT 実施手順の見直しを BACKLOG へ追加した
- 参考: PowerShell の配列アンロールと StrictMode での intrinsic member 抑制は言語仕様。`about_Set-StrictMode` 参照
