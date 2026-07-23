# TASK-260723: ADR / SPEC 策定とプロジェクト初期化

## 目的
ドラフト (adr.md / spec.md) をレビューし、repo 規約に沿った正式な ADR / SPEC として確定する。あわせて AGENTS.md の TODO を埋め、開発着手できる状態にする。

## 決定事項
- ツール名は repo 名に合わせ **ppt2png** に統一 (スクリプト名は PowerShell 規約どおり `Export-PptxPng.ps1`)
- ライセンスは **MIT**
- docs はドラフトの `docs/adr/` 構成ではなく、repo 既存規約のフラット構成 (`docs/ADR-XXXX-*.md` / `docs/SPEC-XXXX-*.md`) に統一
- 初期対応拡張子は `.ppt` / `.pptx` / `.pptm` (スライドショー形式は将来候補として BACKLOG へ)
- ドラフトの矛盾を解消: ADR 利用例の `-Size 3840x2160` を削除 (SPEC に存在しないパラメーター)、`-BaseSize` への言及を `-BaseWidth` に統一

## 成果物
- `docs/ADR-0001-powerpoint-com-export.md`
- `docs/SPEC-0001-pptx-png-export.md`
- `AGENTS.md` (TODO 埋め)
- `docs/BACKLOG.md` (初期リリースまでのタスク・将来候補を追記)

## 残項目
- BACKLOG へ集約済み (本体実装、テスト、LICENSE、README、実機受け入れテストなど)
