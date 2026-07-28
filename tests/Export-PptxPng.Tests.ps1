<#
    COM 非依存の純粋関数のみを対象とした unit test (Pester 5)。
    macOS / Linux の pwsh でも実行できる。

    実行例:
        pwsh -Command "Invoke-Pester ./tests"

    PowerPoint 実機が必要な受け入れテスト (AT-001〜010) は
    docs/SPEC-0001-pptx-png-export.md を参照 (手動実施)。
#>

BeforeAll {
    # main を実行させず関数定義だけを読み込む (InputPath は Mandatory のためダミーを渡す)
    $env:PPT2PNG_SKIP_MAIN = '1'
    . (Join-Path $PSScriptRoot '..' 'Export-PptxPng.ps1') -InputPath 'dummy'
}

AfterAll {
    Remove-Item Env:PPT2PNG_SKIP_MAIN -ErrorAction SilentlyContinue
}

Describe 'Get-OutputDimensions' {
    Context 'Scale mode' {
        It '16:9 スライドで BaseWidth 1920 x Scale 2 -> 3840 x 2160 (AT-001 相当)' {
            $result = Get-OutputDimensions -Mode Scale -Scale 2 -BaseWidth 1920 -SlideWidth 960 -SlideHeight 540
            $result.Width | Should -Be 3840
            $result.Height | Should -Be 2160
        }

        It '4:3 スライドで Scale 2 -> 3840 x 2880 (AT-002 相当)' {
            $result = Get-OutputDimensions -Mode Scale -Scale 2 -BaseWidth 1920 -SlideWidth 720 -SlideHeight 540
            $result.Width | Should -Be 3840
            $result.Height | Should -Be 2880
        }

        It 'デフォルト (Scale 1, BaseWidth 1920) で 16:9 -> 1920 x 1080' {
            $result = Get-OutputDimensions -Mode Scale -SlideWidth 960 -SlideHeight 540
            $result.Width | Should -Be 1920
            $result.Height | Should -Be 1080
        }

        It '計算結果が 32767 を超える場合は InvalidArgument (2) で失敗する' {
            { Get-OutputDimensions -Mode Scale -Scale 1 -BaseWidth 32000 -SlideWidth 540 -SlideHeight 960 } |
                Should -Throw -ErrorId '*'
            try {
                Get-OutputDimensions -Mode Scale -Scale 1 -BaseWidth 32000 -SlideWidth 540 -SlideHeight 960
            }
            catch {
                $_.Exception.Data['Ppt2PngExitCode'] | Should -Be 2
            }
        }
    }

    Context 'Width mode' {
        It '16:9 スライドで Width 2560 -> 2560 x 1440 (AT-003 相当)' {
            $result = Get-OutputDimensions -Mode Width -Width 2560 -SlideWidth 960 -SlideHeight 540
            $result.Width | Should -Be 2560
            $result.Height | Should -Be 1440
        }

        It '高さの端数 .5 は四捨五入 (away from zero) される' {
            # aspect = 2.0, Width 5 -> 2.5 -> banker's rounding なら 2 になってしまうケース
            $result = Get-OutputDimensions -Mode Width -Width 5 -SlideWidth 800 -SlideHeight 400
            $result.Height | Should -Be 3
        }
    }

    Context 'ExplicitSize mode' {
        It '縦横比が一致すればそのまま出力する' {
            $result = Get-OutputDimensions -Mode ExplicitSize -Width 3840 -Height 2160 -SlideWidth 960 -SlideHeight 540
            $result.Width | Should -Be 3840
            $result.Height | Should -Be 2160
        }

        It '許容誤差 0.001 以内の縦横比は許可する' {
            # 1778/1000 = 1.778 vs 16:9 = 1.77778 (差 約0.0002)
            $result = Get-OutputDimensions -Mode ExplicitSize -Width 1778 -Height 1000 -SlideWidth 960 -SlideHeight 540
            $result.Width | Should -Be 1778
        }

        It '縦横比不一致はデフォルトで InvalidArgument (2) エラーになる' {
            try {
                Get-OutputDimensions -Mode ExplicitSize -Width 3840 -Height 2000 -SlideWidth 960 -SlideHeight 540
                throw 'Expected an error but none was thrown.'
            }
            catch {
                $_.Exception.Message | Should -Match 'Aspect ratio mismatch'
                $_.Exception.Data['Ppt2PngExitCode'] | Should -Be 2
            }
        }

        It 'AllowAspectRatioMismatch 指定時は不一致でも出力する' {
            $result = Get-OutputDimensions -Mode ExplicitSize -Width 3840 -Height 2000 -SlideWidth 960 -SlideHeight 540 -AllowAspectRatioMismatch
            $result.Width | Should -Be 3840
            $result.Height | Should -Be 2000
        }
    }
}

Describe 'Get-SlideFileName' {
    It '総数 5 のスライド 1 -> slide-001.png (最低 3 桁)' {
        Get-SlideFileName -SlideNumber 1 -TotalSlideCount 5 | Should -Be 'slide-001.png'
    }

    It '総数 12 のスライド 12 -> slide-012.png' {
        Get-SlideFileName -SlideNumber 12 -TotalSlideCount 12 | Should -Be 'slide-012.png'
    }

    It '総数 1000 のスライド 7 -> slide-0007.png (桁数は総数に追従)' {
        Get-SlideFileName -SlideNumber 7 -TotalSlideCount 1000 | Should -Be 'slide-0007.png'
    }
}

Describe 'Get-TargetSlideNumbers' {
    It 'Slides 未指定なら全スライド番号を返す' {
        Get-TargetSlideNumbers -TotalSlideCount 3 | Should -Be @(1, 2, 3)
    }

    It '指定番号は昇順・重複排除して返す' {
        Get-TargetSlideNumbers -Slides 3, 1, 3 -TotalSlideCount 5 | Should -Be @(1, 3)
    }

    It '総スライド数を超える番号は InvalidArgument (2) エラーになる' {
        try {
            Get-TargetSlideNumbers -Slides 6 -TotalSlideCount 5
            throw 'Expected an error but none was thrown.'
        }
        catch {
            $_.Exception.Message | Should -Match 'exceeds total slide count'
            $_.Exception.Data['Ppt2PngExitCode'] | Should -Be 2
        }
    }
}

Describe 'Assert-InputFile' {
    It '存在しないパスはエラーになる' {
        { Assert-InputFile -Path (Join-Path $TestDrive 'missing.pptx') } | Should -Throw
    }

    It 'ディレクトリはエラーになる' {
        $directory = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'a-dir.pptx')
        { Assert-InputFile -Path $directory.FullName } | Should -Throw
    }

    It '非対応拡張子はエラーになる' {
        $file = New-Item -ItemType File -Path (Join-Path $TestDrive 'document.txt')
        try {
            Assert-InputFile -Path $file.FullName
            throw 'Expected an error but none was thrown.'
        }
        catch {
            $_.Exception.Message | Should -Match 'Unsupported file extension'
            $_.Exception.Data['Ppt2PngExitCode'] | Should -Be 2
        }
    }

    It '対応拡張子は解決済み絶対パスを返す (大文字拡張子も許可)' {
        $file = New-Item -ItemType File -Path (Join-Path $TestDrive 'deck.PPTX')
        Assert-InputFile -Path $file.FullName | Should -Be $file.FullName
    }
}

Describe 'Get-DefaultOutputDirectory / Resolve-OutputDirectoryPath' {
    It 'デフォルトは「入力ディレクトリ/入力ファイル名-png」' {
        $inputPath = Join-Path $TestDrive 'presentation.pptx'
        Get-DefaultOutputDirectory -ResolvedInputPath $inputPath |
            Should -Be (Join-Path $TestDrive 'presentation-png')
    }

    It 'OutputDirectory 未指定ならデフォルトへ解決する' {
        $inputPath = Join-Path $TestDrive 'presentation.pptx'
        Resolve-OutputDirectoryPath -ResolvedInputPath $inputPath |
            Should -Be (Join-Path $TestDrive 'presentation-png')
    }

    It '絶対パス指定はそのまま利用する' {
        $inputPath = Join-Path $TestDrive 'presentation.pptx'
        $absolute = Join-Path $TestDrive 'dist'
        Resolve-OutputDirectoryPath -OutputDirectory $absolute -ResolvedInputPath $inputPath |
            Should -Be $absolute
    }

    It '相対パス指定はカレントロケーション基準で解決する' {
        $inputPath = Join-Path $TestDrive 'presentation.pptx'
        Push-Location $TestDrive
        try {
            Resolve-OutputDirectoryPath -OutputDirectory 'dist' -ResolvedInputPath $inputPath |
                Should -Be (Join-Path (Get-Location).ProviderPath 'dist')
        }
        finally {
            Pop-Location
        }
    }
}

Describe 'Get-ConflictingOutputFiles' {
    It '出力先が存在しなければ空を返す' {
        $result = @(Get-ConflictingOutputFiles -OutputDirectoryPath (Join-Path $TestDrive 'no-such-dir'))
        $result.Count | Should -Be 0
    }

    It 'slide-NNN.png 形式のファイルを検出する' {
        $directory = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'out1')
        $null = New-Item -ItemType File -Path (Join-Path $directory.FullName 'slide-001.png')
        $null = New-Item -ItemType File -Path (Join-Path $directory.FullName 'slide-0042.png')
        $result = @(Get-ConflictingOutputFiles -OutputDirectoryPath $directory.FullName)
        $result.Count | Should -Be 2
    }

    It '無関係なファイルは検出しない' {
        $directory = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'out2')
        $null = New-Item -ItemType File -Path (Join-Path $directory.FullName 'notes.png')
        $null = New-Item -ItemType File -Path (Join-Path $directory.FullName 'slide-01.png')  # 2 桁は対象外
        $result = @(Get-ConflictingOutputFiles -OutputDirectoryPath $directory.FullName)
        $result.Count | Should -Be 0
    }

    It 'TargetSlides 指定時は該当番号のみ検出する (padding 差も吸収)' {
        $directory = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'out3')
        $null = New-Item -ItemType File -Path (Join-Path $directory.FullName 'slide-001.png')
        $null = New-Item -ItemType File -Path (Join-Path $directory.FullName 'slide-0002.png')
        @(Get-ConflictingOutputFiles -OutputDirectoryPath $directory.FullName -TargetSlides 2).Count | Should -Be 1
        @(Get-ConflictingOutputFiles -OutputDirectoryPath $directory.FullName -TargetSlides 3).Count | Should -Be 0
    }

    It 'TargetSlides の番号が別番号の部分文字列でも誤検出しない (12 vs 102)' {
        $directory = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'out4')
        $null = New-Item -ItemType File -Path (Join-Path $directory.FullName 'slide-102.png')
        @(Get-ConflictingOutputFiles -OutputDirectoryPath $directory.FullName -TargetSlides 12).Count | Should -Be 0
    }

    Context '戻り値の配列アンロール (regression: 競合 0 件で StrictMode により .Count が throw した不具合)' {
        # 関数内で return @(...) しても PowerShell は出力を列挙するため、
        # 競合 0 件では呼び出し側に $null が届く。本体の StrictMode 2.0 下では
        # $null.Count が throw するので、呼び出し側の @() ラップが必須という契約を固定する。
        It '競合 0 件の戻り値は $null になる (関数内の @() は呼び出し側を守らない)' {
            $directory = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'unroll-empty')
            $raw = Get-ConflictingOutputFiles -OutputDirectoryPath $directory.FullName
            $null -eq $raw | Should -BeTrue
        }

        It '競合 1 件の戻り値は配列でなくスカラー (FileInfo 単体) になる' {
            $directory = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'unroll-single')
            $null = New-Item -ItemType File -Path (Join-Path $directory.FullName 'slide-001.png')
            $raw = Get-ConflictingOutputFiles -OutputDirectoryPath $directory.FullName
            $raw -is [System.IO.FileInfo] | Should -BeTrue
        }

        It 'StrictMode 2.0 下でラップなしの .Count は throw する (@() ラップが必須な理由)' {
            $directory = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'unroll-strict')
            Set-StrictMode -Version 2.0
            $raw = Get-ConflictingOutputFiles -OutputDirectoryPath $directory.FullName
            { $raw.Count } | Should -Throw
        }

        It 'StrictMode 2.0 下でも @() ラップすれば 0 件 / 1 件とも安全に数えられる' {
            $emptyDir = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'unroll-safe0')
            $singleDir = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'unroll-safe1')
            $null = New-Item -ItemType File -Path (Join-Path $singleDir.FullName 'slide-001.png')
            Set-StrictMode -Version 2.0
            @(Get-ConflictingOutputFiles -OutputDirectoryPath $emptyDir.FullName).Count | Should -Be 0
            $single = @(Get-ConflictingOutputFiles -OutputDirectoryPath $singleDir.FullName)
            $single.Count | Should -Be 1
            $single[0].Name | Should -Be 'slide-001.png'
        }
    }
}

Describe 'Invoke-Ppt2PngExport' {
    Context '起動前検証 (COM 初期化手前まで)' {
        It '競合 0 件 + StrictMode 2.0 でも .Count で落ちず COM 初期化まで到達する (regression)' {
            # 本体の呼び出し箇所 (@() ラップ) を直接守るテスト。ラップを外すと
            # COM 初期化前の $conflicts.Count で throw し、sentinel に到達しなくなる。
            # COM は macOS に存在しないため、到達確認用の sentinel 例外へ差し替える
            Mock New-PowerPointApplication { throw 'SENTINEL-COM-INIT-REACHED' }
            $inputFile = New-Item -ItemType File -Path (Join-Path $TestDrive 'deck.pptx')
            Set-StrictMode -Version 2.0
            { Invoke-Ppt2PngExport -InputPath $inputFile.FullName -Mode Scale -ExistingFilePolicy Error } |
                Should -Throw 'SENTINEL-COM-INIT-REACHED'
        }

        It '競合 1 件なら COM 初期化前に InvalidArgument (2) で失敗し、ファイル名を例示する' {
            Mock New-PowerPointApplication { throw 'SENTINEL-COM-INIT-REACHED' }
            $inputFile = New-Item -ItemType File -Path (Join-Path $TestDrive 'single.pptx')
            $outputDir = New-Item -ItemType Directory -Path (Join-Path $TestDrive 'single-png')
            $null = New-Item -ItemType File -Path (Join-Path $outputDir.FullName 'slide-001.png')
            Set-StrictMode -Version 2.0
            try {
                Invoke-Ppt2PngExport -InputPath $inputFile.FullName -Mode Scale -ExistingFilePolicy Error
                throw 'Expected an error but none was thrown.'
            }
            catch {
                # 競合 1 件はスカラー化するため、修正前は .Count / [0].Name も壊れていた
                $_.Exception.Message | Should -Match 'slide-001\.png'
                $_.Exception.Data['Ppt2PngExitCode'] | Should -Be 2
            }
        }
    }
}

Describe 'New-ExportError / Get-ExportErrorExitCode' {
    It '例外の Data へ終了コードを載せて取り出せる' {
        $exception = New-ExportError -Message 'test' -ExitCode 5
        Get-ExportErrorExitCode -Exception $exception | Should -Be 5
    }

    It '終了コードを持たない例外は RuntimeError (1) とみなす' {
        Get-ExportErrorExitCode -Exception (New-Object System.Exception('plain')) | Should -Be 1
    }
}
