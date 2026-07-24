# TASK-260723: Export-PptxPng.ps1 本体実装

## 目的
[SPEC-0001](SPEC-0001-pptx-png-export.md) に準拠した `Export-PptxPng.ps1` を実装し、COM 非依存の純粋関数を Pester unit test で検証する。

## 実装方針
- 純粋関数 (サイズ計算・検証・命名) と COM 操作関数を分離し、前者のみ macOS (pwsh) でテスト
- テストは `PPT2PNG_SKIP_MAIN=1` + dot-source で関数定義のみ読み込む
- 終了コードは `Exception.Data['Ppt2PngExitCode']` で伝搬し、end ブロックで exit (dot-source 時は exit しない)
- `-ExistingFile Error` の起動前チェックは、総スライド数が未確定 (padding 不明) のため `slide-\d{3,}\.png` パターンで検出

## 成果物
- `Export-PptxPng.ps1`
- `tests/Export-PptxPng.Tests.ps1` / `tests/fixtures/README.md`
- `examples/export-retina.ps1`
- `LICENSE` (MIT)

## 検証
- [x] pwsh 7.6.4 + Pester 6.0.1 で unit test 31 件通過 (macOS)
- [x] codex exec によるコードレビュー・指摘反映 (2 巡 + 最終 LGTM)
- [ ] Windows 実機の受け入れテスト AT-001〜010 は BACKLOG 管理のまま (人間作業)

## レビューでの主要な発見と対応
- **Critical**: PowerPoint は single-instance (Multiuse) COM サーバーで、`New-Object` が既存インスタンスへ接続し得る (「新規プロセスを必ず生成する」という当初 ADR の前提が誤り)。→ 同一セッションの POWERPNT プロセス検出による共有モード (Quit しない / AutomationSecurity 復元) を実装し、ADR-0001 / SPEC-0001 の前提を修正
- 専有モードでも Quit 直前に他 Presentation の有無を確認 (検出後に PowerPoint が起動される競合への緩和。被害をプロセス残留側へ倒す)。厳密なプロセス ID 追跡はスコープ外のまま (Future considerations)
- Presentations コレクションの RCW 解放漏れを修正
- PowerShell バインド段階のエラー (ParameterSet 不整合など) は exit code 2 にならない旨を SPEC へ注記 (仕様化)
- `-AllowAspectRatioMismatch` を ExplicitSize セット専用へ制限
- パイプライン複数ファイル処理の「一部成功 (6)」判定を全体集計へ変更
- `-ExistingFile Error` の起動前チェックを `-Slides` 指定時は対象番号のみに限定
