<#
.SYNOPSIS
    Export PowerPoint slides to PNG at an arbitrary resolution.

.DESCRIPTION
    Windows 版 Microsoft PowerPoint を COM Automation で操作し、各スライドを
    指定ピクセルサイズの PNG として書き出す。レンダリングは PowerPoint 本体へ
    委譲するため、デスクトップ版 PowerPoint のインストールが必須。

    仕様: docs/SPEC-0001-pptx-png-export.md / 設計判断: docs/ADR-0001-powerpoint-com-export.md

.PARAMETER InputPath
    入力する PowerPoint ファイル (.ppt / .pptx / .pptm)。パイプライン入力可。

.PARAMETER Scale
    基準幅 (BaseWidth) に掛ける倍率。0.1〜10.0。

.PARAMETER BaseWidth
    Scale モードの基準ピクセル幅。デフォルト 1920。

.PARAMETER Width
    出力ピクセル幅。Height 省略時は縦横比から高さを自動計算する。

.PARAMETER Height
    出力ピクセル高さ。Width とセットで明示指定する。

.PARAMETER OutputDirectory
    出力先ディレクトリ。省略時は「<入力ディレクトリ>\<入力ファイル名>-png」。

.PARAMETER Slides
    出力対象のスライド番号 (1 始まり)。省略時は全スライド。

.PARAMETER AllowAspectRatioMismatch
    Width/Height 明示指定時、スライドと縦横比が合わなくても強制出力する。

.PARAMETER ExistingFile
    既存出力ファイルの扱い。Error (既定) / Overwrite / Skip。

.PARAMETER OpenOutputDirectory
    書き出し完了後に出力先ディレクトリを Explorer で開く。

.PARAMETER PassThru
    出力ファイルごとの結果オブジェクトをパイプラインへ返す。

.EXAMPLE
    .\Export-PptxPng.ps1 .\presentation.pptx -Scale 2
    16:9 スライドなら 3840 x 2160 で全スライドを書き出す。

.EXAMPLE
    .\Export-PptxPng.ps1 .\presentation.pptx -Width 2560
    幅 2560px、高さはスライド縦横比から自動計算。

.EXAMPLE
    Get-ChildItem .\slides\*.pptx | .\Export-PptxPng.ps1 -Scale 2 -ExistingFile Overwrite
    複数ファイルを逐次処理する (並列実行はしない)。
#>
#Requires -Version 5.1
[CmdletBinding(DefaultParameterSetName = 'Scale')]
param(
    [Parameter(
        Mandatory = $true,
        Position = 0,
        ValueFromPipeline = $true,
        ValueFromPipelineByPropertyName = $true
    )]
    [Alias('FullName')]
    [string] $InputPath,

    [Parameter(ParameterSetName = 'Scale')]
    [ValidateRange(0.1, 10.0)]
    [double] $Scale = 1.0,

    [Parameter(ParameterSetName = 'Scale')]
    [ValidateRange(1, 32767)]
    [int] $BaseWidth = 1920,

    [Parameter(Mandatory = $true, ParameterSetName = 'Width')]
    [Parameter(Mandatory = $true, ParameterSetName = 'ExplicitSize')]
    [ValidateRange(1, 32767)]
    [int] $Width,

    [Parameter(Mandatory = $true, ParameterSetName = 'ExplicitSize')]
    [ValidateRange(1, 32767)]
    [int] $Height,

    [string] $OutputDirectory,

    # ValidateRange は配列の各要素へ適用されるため、1 未満は起動前に弾かれる (SPEC §Input validation)
    [ValidateNotNullOrEmpty()]
    [ValidateRange(1, 2147483647)]
    [int[]] $Slides,

    [Parameter(ParameterSetName = 'ExplicitSize')]
    [switch] $AllowAspectRatioMismatch,

    [ValidateSet('Error', 'Overwrite', 'Skip')]
    [string] $ExistingFile = 'Error',

    [switch] $OpenOutputDirectory,

    [switch] $PassThru
)

begin {
    Set-StrictMode -Version 2.0

    # SPEC §Exit codes
    $script:ExitCode = @{
        Success               = 0
        RuntimeError          = 1
        InvalidArgument       = 2
        PowerPointUnavailable = 3
        CannotOpenInput       = 4
        ExportFailed          = 5
        PartialSuccess        = 6
    }

    $script:SupportedExtensions = @('.ppt', '.pptx', '.pptm')

    # MsoTriState / MsoAutomationSecurity (Office COM 定数)
    $script:MsoTrue = -1
    $script:MsoFalse = 0
    $script:MsoAutomationSecurityForceDisable = 3

    #region Pure functions (COM 非依存。unit test 対象)

    function New-ExportError {
        <#
            終了コードを Exception.Data へ載せて throw 用の例外を作る。
            PowerShell 5.1 互換のため独自例外クラスは使わない。
        #>
        param(
            [Parameter(Mandatory = $true)]
            [string] $Message,

            [Parameter(Mandatory = $true)]
            [int] $ExitCode,

            [System.Exception] $InnerException
        )
        $exception = if ($InnerException) {
            New-Object System.InvalidOperationException($Message, $InnerException)
        }
        else {
            New-Object System.InvalidOperationException($Message)
        }
        $exception.Data['Ppt2PngExitCode'] = $ExitCode
        return $exception
    }

    function Get-ExportErrorExitCode {
        param([System.Exception] $Exception)
        if ($null -ne $Exception -and $Exception.Data.Contains('Ppt2PngExitCode')) {
            return [int]$Exception.Data['Ppt2PngExitCode']
        }
        return $script:ExitCode.RuntimeError
    }

    function Get-OutputDimensions {
        <#
            出力ピクセルサイズを決定する (SPEC §Output size calculation)。
            SlideWidth / SlideHeight はポイント値だが、比率計算にのみ使用する。
        #>
        param(
            [Parameter(Mandatory = $true)]
            [ValidateSet('Scale', 'Width', 'ExplicitSize')]
            [string] $Mode,

            [double] $Scale = 1.0,
            [int] $BaseWidth = 1920,
            [int] $Width,
            [int] $Height,

            [Parameter(Mandatory = $true)]
            [double] $SlideWidth,

            [Parameter(Mandatory = $true)]
            [double] $SlideHeight,

            [switch] $AllowAspectRatioMismatch
        )
        if ($SlideWidth -le 0 -or $SlideHeight -le 0) {
            throw (New-ExportError -Message "Invalid slide size: $SlideWidth x $SlideHeight" -ExitCode $script:ExitCode.RuntimeError)
        }
        $aspectRatio = $SlideWidth / $SlideHeight

        switch ($Mode) {
            'Scale' {
                # 四捨五入 (banker's rounding 回避のため AwayFromZero を明示)
                $outputWidth = [int][Math]::Round($BaseWidth * $Scale, [MidpointRounding]::AwayFromZero)
                $outputHeight = [int][Math]::Round($outputWidth / $aspectRatio, [MidpointRounding]::AwayFromZero)
            }
            'Width' {
                $outputWidth = $Width
                $outputHeight = [int][Math]::Round($Width / $aspectRatio, [MidpointRounding]::AwayFromZero)
            }
            'ExplicitSize' {
                # スライドと異なる縦横比は既定でエラー (許容誤差 0.001)
                $ratioDifference = [Math]::Abs(($Width / $Height) - $aspectRatio)
                if ($ratioDifference -gt 0.001 -and -not $AllowAspectRatioMismatch) {
                    $mismatchMessage = ('Aspect ratio mismatch: requested {0} x {1} (ratio {2:0.####}) but slide ratio is {3:0.####}. ' -f $Width, $Height, ($Width / $Height), $aspectRatio) +
                        'Use -AllowAspectRatioMismatch to force the requested size.'
                    throw (New-ExportError -Message $mismatchMessage -ExitCode $script:ExitCode.InvalidArgument)
                }
                $outputWidth = $Width
                $outputHeight = $Height
            }
        }

        foreach ($value in @($outputWidth, $outputHeight)) {
            if ($value -lt 1 -or $value -gt 32767) {
                throw (New-ExportError -ExitCode $script:ExitCode.InvalidArgument -Message (
                    'Computed output size {0} x {1} px is out of range (1..32767).' -f $outputWidth, $outputHeight
                ))
            }
        }

        return [pscustomobject]@{
            Width  = $outputWidth
            Height = $outputHeight
        }
    }

    function Get-SlideFileName {
        # 桁数は総スライド数に応じて最低 3 桁 (SPEC §File naming)
        param(
            [Parameter(Mandatory = $true)]
            [int] $SlideNumber,

            [Parameter(Mandatory = $true)]
            [int] $TotalSlideCount
        )
        $padding = [Math]::Max(3, $TotalSlideCount.ToString().Length)
        return 'slide-{0}.png' -f $SlideNumber.ToString().PadLeft($padding, '0')
    }

    function Get-TargetSlideNumbers {
        param(
            [int[]] $Slides,

            [Parameter(Mandatory = $true)]
            [int] $TotalSlideCount
        )
        if ($null -eq $Slides -or $Slides.Count -eq 0) {
            return @(1..$TotalSlideCount)
        }
        foreach ($slideNumber in $Slides) {
            if ($slideNumber -gt $TotalSlideCount) {
                throw (New-ExportError -ExitCode $script:ExitCode.InvalidArgument -Message (
                    'Slide number {0} exceeds total slide count {1}.' -f $slideNumber, $TotalSlideCount
                ))
            }
        }
        return @($Slides | Sort-Object -Unique)
    }

    function Assert-InputFile {
        # PowerPoint へ渡す前に絶対パスへ解決する (SPEC §Path handling)
        param(
            [Parameter(Mandatory = $true)]
            [string] $Path
        )
        if (-not (Test-Path -LiteralPath $Path)) {
            throw (New-ExportError -Message "Input file not found: $Path" -ExitCode $script:ExitCode.InvalidArgument)
        }
        if (Test-Path -LiteralPath $Path -PathType Container) {
            throw (New-ExportError -Message "Input path is a directory, not a file: $Path" -ExitCode $script:ExitCode.InvalidArgument)
        }
        # UNC パス対応のため PSDrive 表記ではなく ProviderPath を使う
        $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
        $extension = [System.IO.Path]::GetExtension($resolved).ToLowerInvariant()
        if ($script:SupportedExtensions -notcontains $extension) {
            throw (New-ExportError -ExitCode $script:ExitCode.InvalidArgument -Message (
                "Unsupported file extension '{0}'. Supported: {1}" -f $extension, ($script:SupportedExtensions -join ', ')
            ))
        }
        return $resolved
    }

    function Get-DefaultOutputDirectory {
        param(
            [Parameter(Mandatory = $true)]
            [string] $ResolvedInputPath
        )
        $directory = [System.IO.Path]::GetDirectoryName($ResolvedInputPath)
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($ResolvedInputPath)
        return (Join-Path -Path $directory -ChildPath ($baseName + '-png'))
    }

    function Resolve-OutputDirectoryPath {
        param(
            [string] $OutputDirectory,

            [Parameter(Mandatory = $true)]
            [string] $ResolvedInputPath
        )
        try {
            if ([string]::IsNullOrWhiteSpace($OutputDirectory)) {
                return (Get-DefaultOutputDirectory -ResolvedInputPath $ResolvedInputPath)
            }
            if ([System.IO.Path]::IsPathRooted($OutputDirectory)) {
                return [System.IO.Path]::GetFullPath($OutputDirectory)
            }
            # 相対パスは PowerShell のカレントロケーション基準で解決する
            # (.NET の CurrentDirectory は PS のカレントと一致しないため)
            $base = (Get-Location).ProviderPath
            return [System.IO.Path]::GetFullPath((Join-Path -Path $base -ChildPath $OutputDirectory))
        }
        catch {
            throw (New-ExportError -ExitCode $script:ExitCode.InvalidArgument -InnerException $_.Exception -Message (
                'Invalid output directory path: {0} ({1})' -f $OutputDirectory, $_.Exception.Message
            ))
        }
    }

    function Get-ConflictingOutputFiles {
        <#
            既存出力ファイルの起動前検出 (SPEC §Existing files / AT-010)。
            正確なファイル名は総スライド数 (= 連番の桁数) が確定するまで決まらないため、
            パターン一致で検出する。-Slides 指定時は該当スライド番号のみを競合対象にする
            (例: TargetSlides 2 なら slide-002.png / slide-0002.png のみマッチ)。
        #>
        param(
            [Parameter(Mandatory = $true)]
            [string] $OutputDirectoryPath,

            [int[]] $TargetSlides
        )
        if (-not (Test-Path -LiteralPath $OutputDirectoryPath -PathType Container)) {
            return @()
        }
        $pattern = if ($null -ne $TargetSlides -and $TargetSlides.Count -gt 0) {
            # 桁数 (padding) が未確定でも番号が一致すれば競合とみなす。
            # lookahead で「3 桁以上の連番」形式のみに限定する
            '^slide-(?=\d{3,}\.png$)0*(' + (($TargetSlides | ForEach-Object { $_.ToString() }) -join '|') + ')\.png$'
        }
        else {
            '^slide-\d{3,}\.png$'
        }
        return @(
            Get-ChildItem -LiteralPath $OutputDirectoryPath -File |
                Where-Object { $_.Name -match $pattern }
        )
    }

    #endregion

    #region COM functions (Windows + PowerPoint 必須)

    function New-PowerPointApplication {
        <#
            PowerPoint は single-instance (Multiuse) の COM サーバーであり、
            New-Object はユーザーが既に起動している PowerPoint と同じインスタンスを
            返し得る。そのまま Quit() するとユーザーのセッションを巻き込むため、
            起動前に POWERPNT プロセスの有無を記録し、共有モードかどうかを判定する
            (ADR-0001 COM lifecycle policy)。
        #>
        $isSharedInstance = $false
        try {
            # COM の接続先になり得るのは同一セッションの PowerPoint だけなので、
            # 別ユーザー / 別 RDP セッションの POWERPNT は共有判定から除外する
            $currentSessionId = (Get-Process -Id $PID).SessionId
            $isSharedInstance = @(
                Get-Process -Name POWERPNT -ErrorAction SilentlyContinue |
                    Where-Object { $_.SessionId -eq $currentSessionId }
            ).Count -gt 0
        }
        catch {
            Write-Verbose "Failed to check existing PowerPoint processes: $($_.Exception.Message)"
        }

        try {
            $application = New-Object -ComObject PowerPoint.Application
        }
        catch {
            throw (New-ExportError -ExitCode $script:ExitCode.PowerPointUnavailable -InnerException $_.Exception -Message (
                'PowerPoint is not available. Desktop version of Microsoft PowerPoint is required. ({0})' -f $_.Exception.Message
            ))
        }

        if ($isSharedInstance) {
            Write-Warning 'PowerPoint is already running. The tool will reuse the existing instance and will NOT quit it after export.'
        }

        # ファイルを開く前にマクロを強制無効化する (SPEC §Macro security)。
        # 共有モードではユーザーセッションの設定を変えるため、元の値を保持して後で復元する
        $originalAutomationSecurity = $null
        try {
            $originalAutomationSecurity = [int]$application.AutomationSecurity
            $application.AutomationSecurity = $script:MsoAutomationSecurityForceDisable
        }
        catch {
            Write-Warning "Failed to set AutomationSecurity (macros may not be force-disabled): $($_.Exception.Message)"
        }

        return [pscustomobject]@{
            Application                = $application
            IsSharedInstance           = $isSharedInstance
            OriginalAutomationSecurity = $originalAutomationSecurity
        }
    }

    function Open-PowerPointPresentation {
        param(
            [Parameter(Mandatory = $true)]
            $PresentationsCollection,

            [Parameter(Mandatory = $true)]
            [string] $ResolvedInputPath
        )
        try {
            # 位置引数: FileName, ReadOnly=msoTrue, Untitled=msoFalse, WithWindow=msoFalse
            return $PresentationsCollection.Open(
                $ResolvedInputPath,
                $script:MsoTrue,
                $script:MsoFalse,
                $script:MsoFalse
            )
        }
        catch {
            throw (New-ExportError -ExitCode $script:ExitCode.CannotOpenInput -InnerException $_.Exception -Message (
                'Failed to open input file: {0} ({1})' -f $ResolvedInputPath, $_.Exception.Message
            ))
        }
    }

    function Release-ComObject {
        # 関数名は SPEC NFR-001 指定 (Release は unapproved verb だが仕様準拠を優先)
        param($ComObject)
        if ($null -ne $ComObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
        }
    }

    function Export-PowerPointSlides {
        param(
            [Parameter(Mandatory = $true)]
            $SlidesCollection,

            [Parameter(Mandatory = $true)]
            [int[]] $TargetSlideNumbers,

            [Parameter(Mandatory = $true)]
            [int] $TotalSlideCount,

            [Parameter(Mandatory = $true)]
            [string] $OutputDirectoryPath,

            [Parameter(Mandatory = $true)]
            [int] $OutputWidth,

            [Parameter(Mandatory = $true)]
            [int] $OutputHeight,

            [Parameter(Mandatory = $true)]
            [string] $ExistingFilePolicy,

            [Parameter(Mandatory = $true)]
            [string] $ResolvedInputPath
        )
        $results = New-Object System.Collections.Generic.List[object]
        $exportedCount = 0
        $progressIndex = 0
        $targetCount = $TargetSlideNumbers.Count

        foreach ($slideNumber in $TargetSlideNumbers) {
            $progressIndex++
            $fileName = Get-SlideFileName -SlideNumber $slideNumber -TotalSlideCount $TotalSlideCount
            $outputFilePath = Join-Path -Path $OutputDirectoryPath -ChildPath $fileName

            if ($ExistingFilePolicy -eq 'Skip' -and [System.IO.File]::Exists($outputFilePath)) {
                Write-Host ('[{0}/{1}] {2} (skipped)' -f $progressIndex, $targetCount, $fileName)
                $results.Add([pscustomobject]@{
                    InputPath   = $ResolvedInputPath
                    SlideNumber = $slideNumber
                    OutputPath  = $outputFilePath
                    Width       = $OutputWidth
                    Height      = $OutputHeight
                    Status      = 'Skipped'
                })
                continue
            }

            $slide = $null
            try {
                $slide = $SlidesCollection.Item($slideNumber)
                [void]$slide.Export($outputFilePath, 'PNG', $OutputWidth, $OutputHeight)
            }
            catch {
                # 1 枚でも成功済みなら PartialSuccess として報告する (SPEC §Exit codes)
                $failureCode = if ($exportedCount -gt 0) { $script:ExitCode.PartialSuccess } else { $script:ExitCode.ExportFailed }
                $hresult = ''
                if ($_.Exception -is [System.Runtime.InteropServices.COMException]) {
                    $hresult = ' (HRESULT: 0x{0:X8})' -f $_.Exception.HResult
                }
                throw (New-ExportError -ExitCode $failureCode -InnerException $_.Exception -Message (
                    "Failed to export slide {0}.`n`nInput : {1}`nOutput: {2}`nSize  : {3} x {4}`nError : {5}{6}" -f
                        $slideNumber, $ResolvedInputPath, $outputFilePath, $OutputWidth, $OutputHeight, $_.Exception.Message, $hresult
                ))
            }
            finally {
                Release-ComObject -ComObject $slide
                $slide = $null
            }

            # 出力検証: 存在すること、0 バイトでないこと (FR-014)
            $outputFile = New-Object System.IO.FileInfo($outputFilePath)
            if (-not $outputFile.Exists -or $outputFile.Length -le 0) {
                $failureCode = if ($exportedCount -gt 0) { $script:ExitCode.PartialSuccess } else { $script:ExitCode.ExportFailed }
                throw (New-ExportError -ExitCode $failureCode -Message (
                    "Exported PNG is missing or empty.`n`nInput : {0}`nOutput: {1}`nSize  : {2} x {3}" -f
                        $ResolvedInputPath, $outputFilePath, $OutputWidth, $OutputHeight
                ))
            }

            $exportedCount++
            Write-Host ('[{0}/{1}] {2}' -f $progressIndex, $targetCount, $fileName)
            $results.Add([pscustomobject]@{
                InputPath   = $ResolvedInputPath
                SlideNumber = $slideNumber
                OutputPath  = $outputFilePath
                Width       = $OutputWidth
                Height      = $OutputHeight
                Status      = 'Exported'
            })
        }

        return $results
    }

    function Invoke-Ppt2PngExport {
        # 1 入力ファイル分の変換処理。COM の生成〜解放まで完結させる
        param(
            [Parameter(Mandatory = $true)]
            [string] $InputPath,

            [Parameter(Mandatory = $true)]
            [string] $Mode,

            [double] $Scale,
            [int] $BaseWidth,
            [int] $Width,
            [int] $Height,
            [string] $OutputDirectory,
            [int[]] $Slides,
            [switch] $AllowAspectRatioMismatch,
            [string] $ExistingFilePolicy,
            [switch] $OpenOutputDirectory
        )
        $resolvedInputPath = Assert-InputFile -Path $InputPath
        Write-Verbose "Resolved input path: $resolvedInputPath"

        $outputDirectoryPath = Resolve-OutputDirectoryPath -OutputDirectory $OutputDirectory -ResolvedInputPath $resolvedInputPath
        Write-Verbose "Output directory: $outputDirectoryPath"

        if ($ExistingFilePolicy -eq 'Error') {
            $conflicts = Get-ConflictingOutputFiles -OutputDirectoryPath $outputDirectoryPath -TargetSlides $Slides
            if ($conflicts.Count -gt 0) {
                throw (New-ExportError -ExitCode $script:ExitCode.InvalidArgument -Message (
                    "Output file(s) already exist in {0} (e.g. {1}). Use -ExistingFile Overwrite or Skip." -f
                        $outputDirectoryPath, $conflicts[0].Name
                ))
            }
        }

        if (Test-Path -LiteralPath $outputDirectoryPath -PathType Leaf) {
            throw (New-ExportError -Message "Output directory path exists as a file: $outputDirectoryPath" -ExitCode $script:ExitCode.InvalidArgument)
        }
        try {
            # New-Item はパス中の [ ] をワイルドカード解釈する場合があるため .NET API を使う
            [void][System.IO.Directory]::CreateDirectory($outputDirectoryPath)
        }
        catch {
            throw (New-ExportError -ExitCode $script:ExitCode.InvalidArgument -InnerException $_.Exception -Message (
                'Cannot create output directory: {0} ({1})' -f $outputDirectoryPath, $_.Exception.Message
            ))
        }

        $powerPointContext = $null
        $powerPoint = $null
        $presentationsCollection = $null
        $presentation = $null
        $slidesCollection = $null
        $results = $null

        try {
            Write-Verbose 'Starting PowerPoint application...'
            $powerPointContext = New-PowerPointApplication
            $powerPoint = $powerPointContext.Application

            Write-Verbose 'Opening presentation (ReadOnly, WithWindow=false)...'
            $presentationsCollection = $powerPoint.Presentations
            $presentation = Open-PowerPointPresentation -PresentationsCollection $presentationsCollection -ResolvedInputPath $resolvedInputPath

            $slidesCollection = $presentation.Slides
            $totalSlideCount = [int]$slidesCollection.Count
            if ($totalSlideCount -lt 1) {
                throw (New-ExportError -Message "Presentation has no slides: $resolvedInputPath" -ExitCode $script:ExitCode.CannotOpenInput)
            }

            $targetSlideNumbers = Get-TargetSlideNumbers -Slides $Slides -TotalSlideCount $totalSlideCount

            # PageSetup も COM 参照のため、値の取得後すぐ解放する
            $pageSetup = $null
            try {
                $pageSetup = $presentation.PageSetup
                $slideWidth = [double]$pageSetup.SlideWidth
                $slideHeight = [double]$pageSetup.SlideHeight
            }
            finally {
                Release-ComObject -ComObject $pageSetup
                $pageSetup = $null
            }
            Write-Verbose ("Slide size: {0} x {1} pt" -f $slideWidth, $slideHeight)

            $dimensions = Get-OutputDimensions `
                -Mode $Mode `
                -Scale $Scale `
                -BaseWidth $BaseWidth `
                -Width $Width `
                -Height $Height `
                -SlideWidth $slideWidth `
                -SlideHeight $slideHeight `
                -AllowAspectRatioMismatch:$AllowAspectRatioMismatch
            Write-Verbose ("Output size: {0} x {1} px" -f $dimensions.Width, $dimensions.Height)
            Write-Verbose ("Target slides: {0}" -f ($targetSlideNumbers -join ', '))

            Write-Host ''
            Write-Host ('Input       : {0}' -f $resolvedInputPath)
            Write-Host ('Slides      : {0}' -f $totalSlideCount)
            Write-Host ('Output size : {0} x {1} px' -f $dimensions.Width, $dimensions.Height)
            Write-Host ('Output      : {0}' -f $outputDirectoryPath)
            Write-Host ''

            $results = Export-PowerPointSlides `
                -SlidesCollection $slidesCollection `
                -TargetSlideNumbers $targetSlideNumbers `
                -TotalSlideCount $totalSlideCount `
                -OutputDirectoryPath $outputDirectoryPath `
                -OutputWidth $dimensions.Width `
                -OutputHeight $dimensions.Height `
                -ExistingFilePolicy $ExistingFilePolicy `
                -ResolvedInputPath $resolvedInputPath

            $exportedCount = @($results | Where-Object { $_.Status -eq 'Exported' }).Count
            $skippedCount = @($results | Where-Object { $_.Status -eq 'Skipped' }).Count
            Write-Host ''
            Write-Host ('Exported {0} slides.' -f $exportedCount)
            if ($skippedCount -gt 0) {
                Write-Host ('Skipped {0} existing file(s).' -f $skippedCount)
            }
        }
        finally {
            # 成功・失敗を問わず必ず後始末する (ADR-0001 COM lifecycle policy)
            if ($null -ne $presentation) {
                try {
                    $presentation.Close()
                }
                catch {
                    Write-Warning "Failed to close presentation: $($_.Exception.Message)"
                }
            }
            if ($null -ne $powerPoint) {
                # 共有モードではユーザーが使用中の PowerPoint を終了させない (FR-013)。
                # 専有モードでも、プロセス検出と Application 生成の間に PowerPoint が
                # 起動された競合に備え、他の Presentation が開いていれば Quit しない
                # (被害は「ユーザーの資料を閉じる」ではなく「プロセス残留」側に倒す)
                $shouldQuit = -not $powerPointContext.IsSharedInstance
                if ($shouldQuit) {
                    try {
                        if ($null -ne $presentationsCollection -and [int]$presentationsCollection.Count -gt 0) {
                            $shouldQuit = $false
                            Write-Warning 'Another presentation is open in this PowerPoint instance. Skipping quit to avoid closing it.'
                        }
                    }
                    catch {
                        Write-Verbose "Failed to check remaining presentations: $($_.Exception.Message)"
                    }
                }

                if ($shouldQuit) {
                    try {
                        $powerPoint.Quit()
                    }
                    catch {
                        Write-Warning "Failed to quit PowerPoint: $($_.Exception.Message)"
                    }
                }
                elseif ($null -ne $powerPointContext.OriginalAutomationSecurity) {
                    # PowerPoint が存続する場合は、変更した AutomationSecurity を必ず元へ戻す
                    try {
                        $powerPoint.AutomationSecurity = $powerPointContext.OriginalAutomationSecurity
                    }
                    catch {
                        Write-Warning "Failed to restore AutomationSecurity: $($_.Exception.Message)"
                    }
                }
            }

            # 取得と逆順で解放: Slides collection -> Presentation -> Presentations -> Application
            Release-ComObject -ComObject $slidesCollection
            Release-ComObject -ComObject $presentation
            Release-ComObject -ComObject $presentationsCollection
            Release-ComObject -ComObject $powerPoint
            $slidesCollection = $null
            $presentation = $null
            $presentationsCollection = $null
            $powerPoint = $null

            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            [GC]::Collect()
            [GC]::WaitForPendingFinalizers()
            Write-Verbose 'Released COM objects.'
        }

        if ($OpenOutputDirectory) {
            Invoke-Item -LiteralPath $outputDirectoryPath
        }

        return $results
    }

    #endregion

    # dot-source 時は exit で呼び出し元セッションを終了させない (SPEC §Exit codes)
    $script:isDotSourced = $MyInvocation.InvocationName -eq '.'
    $script:overallExitCode = $script:ExitCode.Success
    # パイプラインで複数ファイルを処理した場合の「一部成功」判定用 (SPEC §Exit codes)
    $script:totalExportedCount = 0
}

process {
    # テスト時は関数定義のみ読み込む (tests/Export-PptxPng.Tests.ps1 参照)
    if ($env:PPT2PNG_SKIP_MAIN -eq '1') {
        return
    }

    try {
        $results = Invoke-Ppt2PngExport `
            -InputPath $InputPath `
            -Mode $PSCmdlet.ParameterSetName `
            -Scale $Scale `
            -BaseWidth $BaseWidth `
            -Width $Width `
            -Height $Height `
            -OutputDirectory $OutputDirectory `
            -Slides $Slides `
            -AllowAspectRatioMismatch:$AllowAspectRatioMismatch `
            -ExistingFilePolicy $ExistingFile `
            -OpenOutputDirectory:$OpenOutputDirectory

        $script:totalExportedCount += @($results | Where-Object { $_.Status -eq 'Exported' }).Count

        if ($PassThru) {
            $results
        }
    }
    catch {
        $failureCode = Get-ExportErrorExitCode -Exception $_.Exception
        if ($script:overallExitCode -eq $script:ExitCode.Success) {
            $script:overallExitCode = $failureCode
        }
        # スタックトレースは -Verbose 時のみ (SPEC §Error messages)
        Write-Verbose $_.ScriptStackTrace
        Write-Error -Message $_.Exception.Message
    }
}

end {
    if ($env:PPT2PNG_SKIP_MAIN -eq '1' -or $script:isDotSourced) {
        return
    }
    # 先行ファイルで書き出し成功があれば、後続の書き出し失敗は全体として一部成功 (6) とする
    if ($script:overallExitCode -eq $script:ExitCode.ExportFailed -and $script:totalExportedCount -gt 0) {
        $script:overallExitCode = $script:ExitCode.PartialSuccess
    }
    exit $script:overallExitCode
}
