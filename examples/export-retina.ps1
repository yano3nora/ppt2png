<#
    Retina (2x) 想定の書き出し例。
    16:9 スライドなら 3840 x 2160 で出力し、既存ファイルは上書きする。
#>
param(
    [Parameter(Mandatory = $true)]
    [string] $InputPath
)

& (Join-Path $PSScriptRoot '..\Export-PptxPng.ps1') $InputPath -Scale 2 -ExistingFile Overwrite
