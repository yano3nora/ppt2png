ppt2png
===

Windows 版 Microsoft PowerPoint を使い、PPT / PPTX の各スライドを任意の解像度で PNG へ書き出す PowerShell CLI です。

PowerPoint 標準の PNG 書き出し UI では出力ピクセルサイズを指定できません。本ツールはインストール済みの PowerPoint 本体をレンダラーとして利用し (COM Automation)、`-Scale` や `-Width` によるサイズ指定を CLI として補完します。

- 設計判断: [docs/ADR-0001-powerpoint-com-export.md](docs/ADR-0001-powerpoint-com-export.md)
- 詳細仕様: [docs/SPEC-0001-pptx-png-export.md](docs/SPEC-0001-pptx-png-export.md)

> [!WARNING]
> 出力PNGには、元のPowerPoint資料に含まれる会社名、顧客名、製品名、個人情報、非公開情報がそのまま描画される可能性があります。公開リポジトリのIssueやテストデータへ、実際の業務資料またはその出力画像を添付しないでください。

## Getting Started
### Requirements
- Windows 10 / 11
- デスクトップ版 Microsoft PowerPoint (レンダリングに使用。PowerPoint for the Web は不可)
- Windows PowerShell 5.1 以上、または PowerShell 7

macOS には対応していません (PowerPoint COM Automation が Windows 専用のため)。

### Installation
[Releases](https://github.com/yano3nora/ppt2png/releases) から最新の `Export-PptxPng.ps1` をダウンロードして、任意のフォルダへ配置してください。

gh CLI を使う場合:

```powershell
gh release download -R yano3nora/ppt2png -p Export-PptxPng.ps1
```

ブラウザでダウンロードした場合、Windows が実行をブロックすることがあります (Mark of the Web)。その場合は次で解除してください。

```powershell
Unblock-File .\Export-PptxPng.ps1
```

実行ポリシーでブロックされる場合は、ポリシーを恒久変更せず、プロセス単位の実行を推奨します。社内のセキュリティポリシーがある場合はそちらを優先してください。

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass `
  -File .\Export-PptxPng.ps1 .\presentation.pptx -Scale 2
```

### Usage
```powershell
# 2倍 (Retina) 出力。16:9 スライドなら 3840 x 2160
.\Export-PptxPng.ps1 .\presentation.pptx -Scale 2

# 幅指定。高さはスライドの縦横比から自動計算 (16:9 なら 2560 x 1440)
.\Export-PptxPng.ps1 .\presentation.pptx -Width 2560

# 幅・高さの明示指定 (スライドと縦横比が合わない場合はエラー)
.\Export-PptxPng.ps1 .\presentation.pptx -Width 3840 -Height 2160

# 一部スライドのみ出力
.\Export-PptxPng.ps1 .\presentation.pptx -Scale 2 -Slides 1,3,5

# 出力先の指定 / 既存ファイルの上書き
.\Export-PptxPng.ps1 .\presentation.pptx -Scale 2 -OutputDirectory .\dist -ExistingFile Overwrite

# 複数ファイルの一括処理 (逐次実行)
Get-ChildItem .\slides\*.pptx | .\Export-PptxPng.ps1 -Scale 2 -ExistingFile Overwrite
```

出力はデフォルトで `<入力ファイル名>-png` フォルダへ、`slide-001.png` 形式の連番で保存されます。

```text
C:\slides\presentation.pptx
C:\slides\presentation-png\slide-001.png
C:\slides\presentation-png\slide-002.png
```

対応形式は `.ppt` / `.pptx` / `.pptm` です (`.pptm` でもマクロは実行しません)。

### Known limitations
- 実行中、PowerPoint が非表示ウィンドウでバックグラウンド起動します (処理後に自動終了)
- 異常終了時に `POWERPNT.EXE` が残ることがあります。タスクマネージャーから該当プロセスを手動終了してください (編集中の資料を巻き込むため、名前指定の一括終了は行わないでください)
- パスワード保護ファイル・修復ダイアログが出る破損ファイルは非対応です
- PC に無いフォントや外部リンク画像は正確に再現されません (PowerPoint 本体の挙動に準じます)
- Windows サービスや CI などの非対話環境での実行はサポート対象外です

## Development
### Structure
```
.
├ Export-PptxPng.ps1    … Main script
├ docs/                 … ADR / SPEC / TASK / BACKLOG
├ examples/             … Usage examples
├ tests/                … Pester tests (COM 非依存の unit test)
└ mise.toml             … Toolchain
```

### Depends
- mise 2026+
- powershell-core (mise 経由。macOS 上で unit test を実行するために使用)

### Getting Started
```sh
# Uncomment required tools/tasks in mise.toml first.
mise install

# If you enable env / hooks / tasks, trust the config explicitly.
mise trust
mise run provision
```

### Commands
```sh
# Unit tests (COM 非依存関数のみ。macOS でも実行可)
pwsh -Command "Invoke-Pester ./tests"

# Integration / acceptance tests (Windows + PowerPoint 実機のみ)
# docs/SPEC-0001-pptx-png-export.md の AT-001〜010 を参照
```

## Deployment
リリース (tag 作成・push・Release 公開) は人間が実施する。Agent は行わない ([AGENTS.md](AGENTS.md))。

```sh
# 1. 検証: unit test 通過 + Windows 実機で受け入れテスト (SPEC AT-001〜010)
# 2. タグ作成・push
git tag vX.Y.Z
git push origin main --tags

# 3. GitHub Release 作成。Export-PptxPng.ps1 を asset として添付する
gh release create vX.Y.Z Export-PptxPng.ps1 --title "vX.Y.Z" --generate-notes
```

利用者は Releases の asset を取得する運用のため、**Release 作成時の asset 添付を忘れないこと**。

### Resources
- [Microsoft Learn: Slide.Export method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.export)
- [Microsoft Learn: Presentations.Open method](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentations.open)

## License
MIT
