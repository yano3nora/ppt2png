# SPEC-0001: ppt2png — PowerPoint PNG Exporter

## Overview

ppt2png は、Windows 版 Microsoft PowerPoint を使用して、PPT / PPTX の各スライドを任意の解像度で PNG 画像へ書き出す CLI ツールである。PowerPoint 標準 UI に不足している、ピクセルサイズおよび倍率指定を CLI として補完する。

- 主要スクリプト: `Export-PptxPng.ps1`
- 意思決定の背景: [ADR-0001](ADR-0001-powerpoint-com-export.md)

### Target environment

- Windows 10 または Windows 11
- Windows PowerShell 5.1 以上、または PowerShell 7
- デスクトップ版 Microsoft PowerPoint
- 対話ログイン済みユーザーのローカルセッション

PowerPoint for the Web は対象外とする。

### Design concept

PowerPoint をヘッドレスな変換ライブラリとして扱うのではなく、非表示ウィンドウで起動するローカル Automation アプリケーションとして扱う。PPTX の内容は解析せず、次の処理だけを行う。

1. 入力値を検証する
2. PowerPoint Application を新規作成する
3. Presentation を読み取り専用・ウィンドウ非表示で開く
4. スライドサイズを取得する
5. 出力ピクセルサイズを計算する
6. 各 Slide の `Export` を呼び出す
7. Presentation と PowerPoint を確実に終了する

## Goals

- 各スライドを任意のピクセルサイズ / 倍率で PNG 書き出しできること
- 元スライドの縦横比を維持できること
- 複数スライドの一括処理と対象スライド選択ができること
- 入力ファイルを一切変更しないこと
- 自身が生成した PowerPoint プロセスを確実に後始末すること

## Non-Goals (Out of scope)

初期バージョンでは次を実装しない。

- macOS 対応 / PowerPoint 未導入環境への対応 / PPTX 独自レンダリング
- GUI / Explorer 右クリック統合 / ドラッグ＆ドロップ UI
- PDF / JPEG / SVG / 動画出力、アニメーションの各フレーム出力
- ノートページ出力 / コメント出力 / 透過背景の保証
- 複数ファイルの並列変換 / Windows サービス実行 / サーバーサイド実行
- PowerPoint プロセスの強制終了
- GitHub Actions による実画像変換テスト
- 自動アップデート / GitHub Release への自動公開
- スライドショー形式 (`.pps` / `.ppsx` / `.ppsm`) の入力対応 (将来候補)

## Terms

- `Slide.Export`
    - PowerPoint COM API のスライド画像書き出しメソッド。出力幅・高さをピクセル単位で受け取る
- `MsoTriState`
    - Office COM の三値型。一般に `msoTrue = -1` / `msoFalse = 0` を用いる
- `BaseWidth`
    - Scale モードの基準となるピクセル幅。デフォルト 1920px
- `COM Automation`
    - Windows アプリケーションを外部プロセスから操作する仕組み。本ツールは PowerPoint 本体をレンダラーとして利用する

## Repository structure

```text
ppt2png/
├── Export-PptxPng.ps1
├── README.md
├── LICENSE                                  # MIT
├── docs/
│   ├── ADR-0001-powerpoint-com-export.md
│   └── SPEC-0001-pptx-png-export.md
├── examples/
│   └── export-retina.ps1
└── tests/
    ├── Export-PptxPng.Tests.ps1
    └── fixtures/
        └── README.md
```

実在企業・顧客・製品・人物の情報を含む PPTX や PNG をリポジトリへコミットしてはならない。

## API / Interface

### Basic syntax

```powershell
.\Export-PptxPng.ps1 <InputPath> [options]
```

### Parameter set A: scale

```powershell
.\Export-PptxPng.ps1 .\slides.pptx -Scale 2
.\Export-PptxPng.ps1 .\slides.pptx -BaseWidth 1280 -Scale 2
```

基準となるピクセルサイズに倍率を掛ける。基準サイズは次の優先順位で決定する。

1. `-BaseWidth` が指定されている場合、その値
2. 指定がない場合、スライド比率を維持しながら幅 1920px を基準とする

16:9 スライドに `-Scale 2` → `3840 × 2160`。4:3 スライドに `-Scale 2` → `3840 × 2880`。

### Parameter set B: width

```powershell
.\Export-PptxPng.ps1 .\slides.pptx -Width 3840
```

幅を指定し、高さはスライドの縦横比から計算する。高さは整数ピクセルへ四捨五入する。

### Parameter set C: explicit size

```powershell
.\Export-PptxPng.ps1 .\slides.pptx -Width 3840 -Height 2160
```

幅と高さを明示的に指定する。縦横比がスライドと一致しない場合、デフォルトではエラーにする。強制的に指定サイズで出力する場合のみ `-AllowAspectRatioMismatch` を許可する。PowerPoint 側の挙動によっては引き伸ばしや余白が発生し得るため、通常利用は推奨しない。

### Output directory

デフォルト: `<入力ファイルのディレクトリ>\<入力ファイル名>-png`

```text
C:\slides\presentation.pptx
C:\slides\presentation-png\
```

`-OutputDirectory` で任意指定できる。

### Slide selection

```powershell
.\Export-PptxPng.ps1 .\presentation.pptx -Scale 2              # 全スライド
.\Export-PptxPng.ps1 .\presentation.pptx -Scale 2 -Slides 3    # 単一
.\Export-PptxPng.ps1 .\presentation.pptx -Scale 2 -Slides 1,3,5 # 複数
```

範囲文字列 (`1,3-5,8`) は初期バージョンでは実装しない。PowerShell 標準の配列で十分なため、初期段階で独自パーサーを追加しない。

### Parameters

```powershell
param(
    [Parameter(
        Mandatory = $true,
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias("FullName")]
    [string] $InputPath,

    [Parameter(ParameterSetName = "Scale")]
    [ValidateRange(0.1, 10.0)]
    [double] $Scale = 1.0,

    [Parameter(ParameterSetName = "Scale")]
    [ValidateRange(1, 32767)]
    [int] $BaseWidth = 1920,

    [Parameter(
        Mandatory = $true,
        ParameterSetName = "Width"
    )]
    [Parameter(
        Mandatory = $true,
        ParameterSetName = "ExplicitSize"
    )]
    [ValidateRange(1, 32767)]
    [int] $Width,

    [Parameter(
        Mandatory = $true,
        ParameterSetName = "ExplicitSize"
    )]
    [ValidateRange(1, 32767)]
    [int] $Height,

    [string] $OutputDirectory,

    [ValidateNotNullOrEmpty()]
    [int[]] $Slides,

    [switch] $AllowAspectRatioMismatch,

    [ValidateSet("Error", "Overwrite", "Skip")]
    [string] $ExistingFile = "Error",

    [switch] $OpenOutputDirectory,

    [switch] $PassThru
)
```

実装時、PowerShell の ParameterSet 制約によって扱いにくい場合は調整してよい。ただし、次の組み合わせは必ず排他制御する。

- `Scale` と `Width`
- `Scale` と `Height`
- `BaseWidth` と明示的な `Width`
- `Height` のみの指定

### Exit codes

```text
0: 全スライドの書き出し成功
1: 実行時エラー
2: 引数または入力値エラー
3: PowerPointが利用できない
4: 入力ファイルを開けない
5: PNG書き出し失敗
6: 一部スライドのみ成功
```

PowerShell の標準的な例外処理を維持しつつ、スクリプト直接実行時には終了コードを設定する。dot-source された場合に呼び出し元セッションを終了しないよう注意する。

### PassThru output

`-PassThru` 指定時は、出力ファイルごとにオブジェクトを返す。

```powershell
[pscustomobject]@{
    InputPath   = $resolvedInputPath
    SlideNumber = $slideNumber
    OutputPath  = $outputFilePath
    Width       = $outputWidth
    Height      = $outputHeight
    Status      = "Exported"  # Exported | Skipped
}
```

通常実行時は、パイプラインへ文字列や COM オブジェクトを漏らさない。COM メソッドの戻り値は必要に応じて `[void]` で破棄する。

## Behavior

### Output size calculation

PowerPoint から `Presentation.PageSetup.SlideWidth` / `SlideHeight` を取得する。単位は比率計算にのみ利用し、DPI 変換には利用しない。

```text
aspectRatio = SlideWidth / SlideHeight

# Width mode
outputWidth  = Width
outputHeight = round(Width / aspectRatio)

# Scale mode
outputWidth  = round(BaseWidth × Scale)
outputHeight = round(outputWidth / aspectRatio)

# Explicit-size mode
outputWidth  = Width
outputHeight = Height
# 縦横比の許容誤差: abs((Width / Height) - aspectRatio) <= 0.001
# 誤差超過かつ AllowAspectRatioMismatch なしの場合、終了コード 2 で失敗
```

幅・高さはそれぞれ 1〜32767px に制限する。ただし PowerPoint のバージョンや画像フィルターによって実際の最大値がこれより小さい可能性がある。PowerPoint の Export 失敗は捕捉し、指定サイズを含むエラーとして表示する。

### File naming

```text
slide-001.png
slide-002.png
```

桁数は総スライド数に応じて最低 3 桁 (`padding = max(3, length(totalSlideCount))`)。スライド番号は PowerPoint 上のインデックスに対応する。

スライドタイトルはファイル名へ含めない。理由: 使用禁止文字への対応が必要 / タイトル重複 / タイトル変更で成果物名が変わる / 長すぎるファイル名 / ファイル名からの機密情報漏えい。

### Existing files

`-ExistingFile` により制御する。

- `Error` (デフォルト): 同名ファイルが 1 つでも存在する場合、PowerPoint を起動する前に終了する
- `Overwrite`: 既存ファイルを上書きする
- `Skip`: 存在するファイルを処理対象から除外する。一部成功になる可能性があるため、スキップ件数を標準出力へ表示する

### PowerPoint automation

必ず新しい COM Application を作成する。既存インスタンスを取得する `GetActiveObject` は使用しない。

```powershell
$powerPoint = New-Object -ComObject PowerPoint.Application
```

Presentation は ReadOnly: true / Untitled: false / WithWindow: false で開く。名前付き引数は PowerShell の COM 呼び出しで扱いにくいため、位置引数を利用してよい。

```powershell
# msoTrue = -1, msoFalse = 0
$presentation = $powerPoint.Presentations.Open(
    $resolvedInputPath,
    -1,   # ReadOnly
    0,    # Untitled
    0     # WithWindow
)
```

可能であれば、ファイルを開く前に `Application.AutomationSecurity` を利用しマクロを無効化する。PowerPoint のバージョン差などにより設定できない場合は、警告を出すか、明確なエラーとして終了する。`.pptm` を受け入れる場合でも、マクロを実行してはならない。

各スライドに対して次を呼び、出力後にファイル存在と 0 バイト超を検証する。

```powershell
$slide.Export($outputFilePath, "PNG", $outputWidth, $outputHeight)
```

- 非表示スライドも通常スライドと同様に書き出す (`-ExcludeHiddenSlides` は将来候補)
- ノート、コメント、発表者メモは画像へ含めない。スライド表示領域のみを書き出す
- アニメーションは実行しない。編集画面で表示される静的な初期状態を PNG 化する

### Resource cleanup

COM オブジェクトは取得した順序と逆順 (Slide → Slides collection → Presentation → Application) に解放する。実装は必ず `try/catch/finally` を利用する。

```powershell
$powerPoint = $null
$presentation = $null
$slides = $null
$slide = $null

try {
    # PowerPoint起動、ファイルオープン、書き出し
}
catch {
    throw
}
finally {
    if ($presentation -ne $null) {
        try { $presentation.Close() }
        catch { Write-Warning "Failed to close presentation: $($_.Exception.Message)" }
    }

    if ($powerPoint -ne $null) {
        try { $powerPoint.Quit() }
        catch { Write-Warning "Failed to quit PowerPoint: $($_.Exception.Message)" }
    }

    foreach ($comObject in @($slide, $slides, $presentation, $powerPoint)) {
        if ($null -ne $comObject -and
            [System.Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject(
                $comObject
            )
        }
    }

    $slide = $null
    $slides = $null
    $presentation = $null
    $powerPoint = $null

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
```

COM コレクションをパイプラインへ流したり、`foreach` の中で不要な中間 COM 参照を生成したりしない。COM 参照が暗黙生成されるコードを最小化する。

### Input validation

PowerPoint を起動する前に、最低限次を検証する。

- 入力パスが存在する / ファイルである / 拡張子がサポート対象である
- 出力先が入力ファイルと衝突しない / 出力先を作成できる
- Width、Height、Scale が有効範囲内である
- Slides の番号が 1 以上、かつ総スライド数以下である
- 既存ファイルポリシーに違反しない

サポート拡張子 (初期バージョン):

```text
.ppt
.pptx
.pptm
```

スライドショー形式 (`.pps` / `.ppsx` / `.ppsm`) は将来対応とする。

### Path handling

すべてのパスは PowerPoint へ渡す前に絶対パスへ解決する。

```powershell
$resolvedInputPath = (
    Resolve-Path -LiteralPath $InputPath -ErrorAction Stop
).Path
```

出力先では `-LiteralPath` を利用し、角括弧などをワイルドカードとして解釈しない。空白 / 日本語 / `[` `]` を含むパス、UNC パス、OneDrive 同期フォルダ、長いパスを考慮する。OneDrive のオンライン専用ファイルを自動ダウンロードする責任は持たない。

### Console output

```text
Input       : C:\slides\presentation.pptx
Slides      : 12
Output size : 3840 x 2160 px
Output      : C:\slides\presentation-png

[1/12] slide-001.png
[2/12] slide-002.png
...
[12/12] slide-012.png

Exported 12 slides.
```

`Write-Host` を多用せず、意味に応じて `Write-Verbose` / `Write-Information` / `Write-Warning` / `Write-Error` を使い分ける。進捗の最低限の表示は標準出力へ出してよい。

### Error messages

エラーには対象入力パス / 対象スライド番号 / 出力パス / 指定した幅と高さ / COM 例外の Message / 必要であれば HRESULT を含める。

```text
Failed to export slide 4.

Input : C:\slides\presentation.pptx
Output: C:\slides\presentation-png\slide-004.png
Size  : 3840 x 2160
Error : Slide.Export failed: 0x80004005
```

スタックトレースはデフォルトでは表示せず、`-Verbose` または `-Debug` 時に確認できるようにする。

## Invariants

- 入力ファイルを変更・保存しない (読み取り専用で開く)
- マクロを実行しない
- 出力先以外へファイルを書き込まない
- 自身が生成した PowerPoint Application 以外に接続・終了しない
- `Stop-Process -Name POWERPNT` / `taskkill /IM POWERPNT.EXE` を実行しない (既存プロセスを巻き込むため)。強制終了が必要な設計になった場合は、生成プロセス ID を厳密に特定する別 ADR を作成する
- パイプラインによる複数ファイル処理は逐次実行する。並列実行しない

## Functional requirements

- FR-001: 入力した PowerPoint ファイルを読み取り専用で開けること
- FR-002: PowerPoint のウィンドウを原則として表示しないこと
- FR-003: 全スライドを PNG として出力できること
- FR-004: 幅をピクセル単位で指定できること
- FR-005: 幅のみ指定した場合、元スライドの縦横比を維持して高さを計算すること
- FR-006: 基準幅に対する倍率を指定できること
- FR-007: 16:9 スライドで基準幅 1920、倍率 2 を指定した場合、3840 × 2160 で出力すること
- FR-008: 出力対象スライド番号を指定できること
- FR-009: 入力ファイルを変更または保存しないこと
- FR-010: 既存出力ファイルの扱いを Error、Overwrite、Skip から選択できること
- FR-011: 失敗したスライド番号と出力パスをエラーに含めること
- FR-012: 処理終了後に、自身が作成した PowerPoint Application を終了すること
- FR-013: ユーザーが既に起動している PowerPoint を終了しないこと
- FR-014: 出力 PNG が存在し、0 バイトでないことを検証すること
- FR-015: 日本語および空白を含むパスを扱えること

## Non-functional requirements

- NFR-001 (Maintainability): 主要処理を Verb-Noun 規則の関数へ分割する — `Resolve-ExportConfiguration` / `Get-OutputDimensions` / `Get-TargetSlideNumbers` / `Assert-InputFile` / `Assert-OutputDirectory` / `New-PowerPointApplication` / `Open-PowerPointPresentation` / `Export-PowerPointSlides` / `Release-ComObject`
- NFR-002 (Safety): 入力ファイルは読み取り専用とし、保存処理を実行しない
- NFR-003 (Compatibility): PowerShell 5.1 で動作する構文を基本とし、PowerShell 7 固有機能は使用しない
- NFR-004 (Observability): `-Verbose` 指定時に、解決済み入力パス / 出力先 / スライドサイズ / 計算された出力サイズ / 対象スライド / PowerPoint 起動 / Presentation オープン / COM 解放処理を確認できること
- NFR-005 (Idempotency): `-ExistingFile Overwrite` で同じ入力とオプションを繰り返した場合、同名の出力結果へ置き換わること
- NFR-006 (No additional dependencies): 外部 PowerShell モジュールを要求しない。Pester はテスト実行時の開発依存としてのみ利用してよい

## Testing strategy

### Unit tests

PowerPoint を起動せずテストできる処理を関数へ分離し、Pester でテストする。対象: 出力サイズ計算 / 縦横比検証 / スライド番号検証 / 出力ファイル名生成 / 既存ファイルポリシー / パラメーター競合 / パス解決。

### Integration tests

PowerPoint がインストールされた Windows PC で実行する。最低限のテスト資料: 16:9 / 4:3 / 日本語テキスト / 英数字テキスト / 図形 / 画像 / SVG / グラフ / 非表示スライド / 複数スライド。テスト資料には架空情報だけを含める。

### Acceptance tests

- AT-001: 16:9 PPTX + `-Scale 2` (BaseWidth 1920) → 3840 × 2160
- AT-002: 4:3 PPTX + `-Scale 2` → 3840 × 2880
- AT-003: 16:9 PPTX + `-Width 2560` → 2560 × 1440
- AT-004: 5 枚の PPTX → `slide-001.png`〜`slide-005.png` の 5 ファイル生成
- AT-005: 5 枚の PPTX + `-Slides 2,4` → `slide-002.png` と `slide-004.png` のみ生成
- AT-006: 実行前後で入力ファイルの更新日時とハッシュが変わらない
- AT-007: 正常終了時・エラー時とも、本ツールが生成した PowerPoint Application が残留しない
- AT-008: ユーザーが開いている別の PowerPoint と資料が終了されない
- AT-009: `C:\資料\発表資料 2026.pptx` のような日本語・空白パスで正常に PNG 出力できる
- AT-010: 出力先に同名 PNG があり `-ExistingFile Error` の場合、PowerPoint を起動する前に失敗する

## Edge Cases / Failure Modes

- PowerPoint 未インストール → 終了コード 3
- パスワード保護・破損ファイル (修復ダイアログ) → サポート対象外。開けない場合は終了コード 4
- Export 途中の COM 例外 → スライド番号・出力パス・サイズを含むエラーを表示し、終了コード 5 または 6
- PowerPoint のダイアログやアドインによる処理妨害 → サポート対象外だが、異常終了時も finally で後始末する
- 巨大サイズ指定 → 32767px 以内でも PowerPoint 側の上限で失敗し得る。失敗は捕捉して指定サイズを含めて報告する
- OneDrive オンライン専用ファイル → 自動ダウンロードしない

## Trouble Shooting

- **実行ポリシーで起動できない**: 恒久的な緩和は推奨せず、プロセス単位の実行を案内する。社内セキュリティポリシーがある場合はそちらを優先する

    ```powershell
    powershell.exe -NoProfile -ExecutionPolicy Bypass `
      -File .\Export-PptxPng.ps1 .\presentation.pptx -Scale 2
    ```

- **POWERPNT.EXE が残留した**: タスクマネージャーから該当プロセスを手動終了する。名前指定の一括 kill はユーザーの編集中資料を巻き込むため行わない
- **出力 PNG が 0 バイト / 生成されない**: 出力後検証で検出しエラー表示する。サイズ指定を下げて再実行する

## README requirements

README には次を記載する: ツールの目的 / Windows 専用であること / デスクトップ版 PowerPoint が必要であること / インストール方法 / 基本的な実行例 / 3840 × 2160 の実行例 / PowerShell 実行ポリシー / 出力先とファイル名 / 対応形式 / 既知の制約 / PowerPoint が裏で起動すること / 異常終了時の対処 / 機密情報に関する警告 / ライセンス (MIT)。

次の注意を明記する。

> 出力PNGには、元のPowerPoint資料に含まれる会社名、顧客名、製品名、個人情報、非公開情報がそのまま描画される可能性があります。公開リポジトリのIssueやテストデータへ、実際の業務資料またはその出力画像を添付しないでください。

## Future considerations

運用実績に応じて次を検討する (詳細は `docs/BACKLOG.md` で管理)。

- C# 単一 EXE への移行 / コード署名 / `.psm1` モジュール化 / PowerShell Gallery 公開
- Explorer コンテキストメニュー / ドラッグ＆ドロップランチャー
- JPEG 対応 / スライド範囲式 `1,3-5` / 非表示スライド除外 / `.pps` 系入力対応
- JSON 形式の実行結果 / 出力ファイルの SHA-256 生成 / タイムアウト処理
- 作成した PowerPoint プロセス ID の追跡 / Windows Sandbox での検証
- Scoop または WinGet による配布

## Open Questions

- なし (未着手事項は `docs/BACKLOG.md` を参照)
