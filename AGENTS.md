# AGENTS - Development Guide
## Overview
- この repo は **ppt2png** — Windows 版 Microsoft PowerPoint の COM Automation を利用し、PPT / PPTX の各スライドを任意解像度の PNG へ書き出す単一 PowerShell スクリプト `Export-PptxPng.ps1` を開発・GitHub 配布するプロジェクト
- 技術スタック: PowerShell 5.1 互換構文 (PowerShell 7 でも動作) / PowerPoint COM Automation / Pester (開発時テストのみ)。実行環境は Windows 10・11 + デスクトップ版 PowerPoint
- 最重要ドキュメント: `docs/ADR-0001-powerpoint-com-export.md` (なぜ COM Automation か)、`docs/SPEC-0001-pptx-png-export.md` (CLI 仕様・FR/NFR・受け入れテスト)
- 関連 repo / 参照実装: なし。参照 API は Microsoft Learn の `Slide.Export` / `Presentations.Open`

### 🎯 Role & Objective
あなたはエキスパートソフトウェアエンジニアとして、この repo の設計・実装・テストを行うこと。

### 🚨 CRITICAL: Architecture
- **レンダリングは PowerPoint 本体へ委譲する**: PPTX の構造・描画内容を独自解析・独自レンダリングしない。互換性問題 (フォント・SmartArt・テーマ等) を PowerPoint 自身に任せるのが本ツールの存在理由 (ADR-0001)
- **依存境界**: 外部 PowerShell モジュールへ依存しない単一 `.ps1`。Pester は開発時依存のみ。PowerShell 7 固有機能は使わない (5.1 互換)
- **状態管理の原則**: COM オブジェクトは自身が生成したものだけを扱い、取得と逆順で確実に解放する (`try/catch/finally` 必須)。既存 PowerPoint インスタンスへの接続 (`GetActiveObject`) と、`Stop-Process -Name POWERPNT` 等の名前指定 kill は**禁止**
- **失敗モード**: PowerPoint 未インストール / ダイアログ・アドインによる妨害 / COM 例外 / POWERPNT.EXE 残留を前提に設計する。入力ファイルは読み取り専用で開き、絶対に変更・保存しない。マクロは実行しない
- **YAGNI / 過剰設計禁止**: GUI・PDF/JPEG 出力・並列変換・サーバー実行・独自レンダラー・プロセス強制終了の方向へ広げない。範囲式パーサー等も SPEC の Out of scope を確認してから

### 📂 Code Organization Constraints
- **`Export-PptxPng.ps1`**: 本体。単一スクリプト。主要処理は Verb-Noun 規則の関数へ分割する (SPEC NFR-001)
- **`docs/`**: ADR / SPEC / TASK / BACKLOG。フラット構成 (`docs/ADR-XXXX-*.md` 形式)
- **`tests/`**: Pester テストと fixtures。fixtures は架空情報のみ
- **`examples/`**: 利用例スクリプト
- **型 / 境界**: 純粋計算 (出力サイズ計算・縦横比検証・ファイル名生成・パラメーター検証) は COM 非依存の関数へ分離し、PowerPoint なしで unit test 可能にする

### 🛠️ Workflow & Development Rules
- **Secrets**: 企業名・製品名・機密情報などがあった場合、コード上に残らないように汎用・一般名称に差し替えること。テスト用 PPTX / PNG は架空の内容だけで作成する
- **Commit**: `git commit` は基本的には人間判断で行うため、指示されたとき以外はコミットせず人間に判断を委ねること。
- **Push / Publish**: `github push` や `npm publish` など、外部へ公開・配布する操作は Agent が実行しない。人間が判断して実行する。
- **Testing**: タスク完了前に実行する検証を書く
    - linter / formatter: PSScriptAnalyzer 推奨設定に準拠する (導入は BACKLOG 参照)
    - unit test: Pester。COM 非依存の純粋関数 (サイズ計算・検証・命名) を対象とし、macOS 上でも実行可能に保つ
    - integration / e2e: PowerPoint 入り Windows 実機での受け入れテスト (SPEC AT-001〜010)。CI では実行しない
    - bugfix 時: 再現ケースを Pester テストとして先に追加してから修正する
- **Documentation**:
    - 技術的な意思決定や検討は `docs/ADR-XXXX-*.md` に記録し、大きな変更の前には既存 ADR を確認する
    - 設計・仕様の検討・決定事項は `docs/SPEC-XXXX-*.md` に記録する
    - 原則、全開発タスクが適切な粒度で `docs/TASK-YYMMDD-*.md` に残るようにする
    - 未着手・保留・トリガー待ちのタスクは `docs/BACKLOG.md` に一元管理し、着手時に TASK として切り出す
    - 画像などは `docs/assets/` へ配置してリンクする
- **Versioning / Release**: SemVer。タグ作成・GitHub Release 公開は人間が実施する。ライセンスは MIT

## Domains
- `Slide Export`
    - PowerPoint COM の `Slide.Export(path, "PNG", width, height)` によるスライド単位の画像書き出し。出力後にファイル存在と 0 バイト超を検証する
- `Scale / BaseWidth`
    - 出力サイズ指定の基準。デフォルト基準幅 1920px に倍率を掛け、高さはスライド縦横比から算出する (16:9 × Scale 2 = 3840 × 2160)
- `MsoTriState`
    - Office COM の三値型。`msoTrue = -1` / `msoFalse = 0`。`Presentations.Open` の ReadOnly / WithWindow 指定に使う
- `COM lifecycle`
    - Application → Presentation → Slides → Slide の順に取得し、逆順で `FinalReleaseComObject` する後始末規約。詳細は ADR-0001 / SPEC-0001 参照
