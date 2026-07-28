# Backlogs — 未解決／積み残しタスク

> **Status: 常設 (クローズしない)**。未着手・保留・トリガー待ちのタスクを一元管理する唯一の置き場。
> 各 TASK の残項目はここに集約済みなので、過去 TASK を漁る必要はない。

## 運用ルール
1. 次の作業を始めるときは、ここから 1 件 pick して `TASK-YYMMDD-<slug>.md` を新規作成する
2. pick した項目は新 TASK へのリンクに差し替え、完了したらリンクごと項目を削除する
3. 新しい未解決事項が出たら、他の TASK には「BACKLOGへ追加」だけ書いてここへ追記する
4. ADR, SPEC の Open Questions と重複する項目は、決着時に ADR, SPEC 側も更新すること

## 初期リリースまでに必要
- [x] `Export-PptxPng.ps1` 本体実装 → [TASK-260723-implement-export-script](TASK-260723-implement-export-script.md) で完了
- [ ] PSScriptAnalyzer 導入と設定 (mise.toml のタスク整備、powershell-core の tools 登録含む)
- [ ] Windows 実機での受け入れテスト AT-001〜010 実施 (人間作業)。特に共有モード (PowerPoint 起動中の実行) は AT-008 で重点確認する
    - デフォルト引数 + 空の出力先 (競合 0 件) の happy path を必ず含めること。[TASK-260728](TASK-260728-fix-conflicts-count-strictmode.md) の不具合 (StrictMode 下で `.Count` throw) はこの経路で必ず再現するのに AT で検出されなかったため、AT 手順のカバレッジ見直しも合わせて行う
- [ ] Windows PowerShell 5.1 実機での動作確認 (開発は pwsh 7 のみで検証済みのため)
- [ ] 初回リリース: タグ作成 + GitHub Release 作成・asset 添付 (人間作業、手順は README の Deployment 参照)

## 将来候補 (SPEC-0001 Future considerations)
- [ ] スライドショー形式 (`.pps` / `.ppsx` / `.ppsm`) の入力対応
- [ ] スライド範囲式 `1,3-5` / `-ExcludeHiddenSlides` / JPEG 対応
- [ ] JSON 形式の実行結果 / 出力 SHA-256 / タイムアウト処理 / 生成プロセス ID 追跡
- [ ] C# 単一 EXE 化・コード署名・`.psm1` モジュール化・PowerShell Gallery 公開
- [ ] Explorer コンテキストメニュー / ドラッグ＆ドロップランチャー
- [ ] Scoop / WinGet 配布、Windows Sandbox での検証
