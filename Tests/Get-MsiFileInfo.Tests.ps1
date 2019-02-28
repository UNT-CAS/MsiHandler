[string]           $projectDirectoryName = 'MsiHandler'
[IO.FileInfo]      $pesterFile = [io.fileinfo] ([string] (Resolve-Path -Path $MyInvocation.MyCommand.Path))
[IO.DirectoryInfo] $projectRoot = Split-Path -Parent $pesterFile.Directory
[IO.DirectoryInfo] $projectDirectory = Join-Path -Path $projectRoot -ChildPath $projectDirectoryName -Resolve
[IO.FileInfo]      $testFile = Join-Path -Path $projectDirectory -ChildPath (Join-Path -Path 'Public' -ChildPath ($pesterFile.Name -replace '\.Tests\.', '.')) -Resolve
. $testFile

$testMsiUrls = (Import-PowerShellDataFile ('{0}\TestMsiUrls.psd1' -f $PSScriptRoot)).URLs
[Collections.ArrayList] $testMsiFiles = @()

[IO.DirectoryInfo] $destPath = '{0}/dev/testMSIs' -f $projectRoot
if (-not $destPath.Exists) {
    New-Item -ItemType Directory -Path $destPath -Force
    $destPath.Refresh()
}

foreach ($testMSI in $testMsiUrls) {
    [uri] $testMSI = $testMSI
    [IO.FileInfo] $destFile = Join-Path $destPath $testMsi.Segments[-1]

    if (-not $destFile.Exists) {
        Invoke-WebRequest -Uri $testMSI -OutFile $destFile
        $destFile.Refresh()
    }

    [void] $testMsiFiles.Add($destFile)
}

if ($testMsiUrls.Count -ne $testMsiFiles.Count) {
    Write-Warning 'All of the MSI Files may not have been downloaded.'
}

Describe 'Get-MsiFileInfo' {
    for ($i = 0; $i -lt 3; $i++) {
        Context ('Default Properties; Path From Pipeline (i:{0})' -f $i) {
            switch ($i) {
                0 {
                    It 'Default Properties: *Should Not Error*' {
                        { $script:results = Get-ChildItem -Path $destPath | Get-MsiFileInfo } | Should Not Throw
                    }
                }
                1 {
                    [IO.FileInfo] $tmp = New-TemporaryFile | Move-Item -Destination $destPath -Verbose

                    It 'Default Properties: *Should Not Error*' {
                        { $script:results = Get-ChildItem -Path $destPath | Get-MsiFileInfo } | Should Not Throw
                    }
                }
                2 {
                    $tmp = New-TemporaryFile | Move-Item -Destination $destPath -PassThru -Verbose
                    Rename-Item -LiteralPath $tmp.FullName -NewName ('00{0}' -f $tmp.Name) -Confirm:$false -Force -PassThru -Verbose

                    It 'Default Properties: *Should Not Error*' {
                        { $script:results = Get-ChildItem -Path $destPath | Get-MsiFileInfo } | Should Not Throw
                    }
                }
            }
    
            Write-Host "Results: $($script:results | Out-String)"
            
            It 'Default Properties: .IO.FileInfo' {
                $script:results.'.IO.FileInfo' | Should BeOfType [System.IO.FileInfo]
            }

            if ($i -eq 0) {
                foreach ($exists in $script:results.'.IO.FileInfo'.Exists) {
                    It 'Default Properties: .IO.FileInfo Exists' {
                        $exists | Should Be $true
                    }
                }
            }
            
            if ($i -eq 0) {
                It 'Default Properties: Manufacturer String' {
                    $script:results.Manufacturer | Should BeOfType [System.String]
                }
            } else {
                It 'Default Properties: Manufacturer String[]' {
                    $script:results.Manufacturer.GetType().FullName.GetType().FullName | Should Be 'System.String'
                }
            }

            # The null check fails when there's a tmp file that gives null results
            if ($i -eq 0) {
                It 'Default Properties: Manufacturer Not Null' {
                    $script:results.Manufacturer | Should Not BeNullOrEmpty
                }
            }
            
            if ($i -eq 0) {
                It 'Default Properties: ProductCode String' {
                    $script:results.ProductCode | Should BeOfType [System.String]
                }

                foreach ($productCode in $script:results.ProductCode) {
                    It ('Default Properties: ProductCode Guid {0}' -f $productCode) {
                        $productCode -as [guid] | Should BeOfType [System.Guid]
                    }
                }
            } else {
                It 'Default Properties: ProductCode Object[]' {
                    $script:results.ProductCode.GetType().FullName | Should Be 'System.Object[]'
                }
            }
            
            if ($i -eq 0) {
                It 'Default Properties: ProductLanguage String' {
                    $script:results.ProductLanguage | Should BeOfType [System.String]
                }
            } else {
                It 'Default Properties: ProductLanguage Object[]' {
                    $script:results.ProductLanguage.GetType().FullName | Should Be 'System.Object[]'
                }
            }
            
            if ($i -eq 0) {
                It 'Default Properties: ProductName String' {
                    $script:results.ProductName | Should BeOfType [System.String]
                }
            } else {
                It 'Default Properties: ProductName Object[]' {
                    $script:results.ProductName.GetType().FullName | Should Be 'System.Object[]'
                }
            }
            
            if ($i -eq 0) {
                It 'Default Properties: ProductVersion String' {
                    $script:results.ProductVersion | Should BeOfType [System.String]
                }

                foreach ($productVersion in $script:results.ProductVersion) {
                    It ('Default Properties: ProductVersion Version {0}' -f $productVersion) {
                        $productVersion -as [version] | Should BeOfType [System.Version]
                    }
                }
            } else {
                It 'Default Properties: ProductVersion Object[]' {
                    $script:results.ProductVersion.GetType().FullName | Should Be 'System.Object[]'
                }
            }
        }
    }

    Remove-Variable -Scope 'Script' -Name 'results' -Force -ErrorAction Ignore

    Context ('[{0}] Non-MSI Files Are Skipped With Warning' -f $testMsiFile.BaseName) {
        It 'Default Properties: *Should Not Error*' {
            { $script:results = Get-ChildItem -Path $destPath | Get-MsiFileInfo } | Should Not Throw
        }

        It 'Default Properties: *Should Warning*' {
            { $script:results = Get-ChildItem -Path $destPath | Get-MsiFileInfo -WarningAction Stop } | Should Throw
        }

        It 'Default Properties: Warning Should be correct type' {
            # $Error[0] | Should BeOfType [System.Management.Automation.ActionPreferenceStopException]
            $Error[0] | Should BeLike 'The running command stopped because the preference variable "WarningPreference" or common parameter is set to Stop: *'
        }
    }

    Remove-Variable -Scope 'Script' -Name 'results' -Force -ErrorAction Ignore

    Context ('[{0}] Non-MSI Files Can Be Returned With Warning' -f $testMsiFile.BaseName) {
        It 'Default Properties: *Should Not Error*' {
            { $script:results = Get-ChildItem -Path $destPath | Get-MsiFileInfo -IncludeNonMsiFileInfo } | Should Not Throw
        }

        It 'Default Properties: *Should Warning*' {
            { [void] (Get-ChildItem -Path $destPath | Get-MsiFileInfo -IncludeNonMsiFileInfo -WarningAction Stop) } | Should Throw
        }

        It 'Default Properties: Warning Should be correct type' {
            $Error[0] | Should BeLike 'The running command stopped because the preference variable "WarningPreference" or common parameter is set to Stop: *'
        }

        $script:fileCount = (Get-ChildItem -Path $destPath | Measure-Object).Count

        It 'Default Properties: .IO.FileInfo Count' {
            ($script:results.'.IO.FileInfo' | Measure-Object).Count | Should Be $script:fileCount
        }

        [System.Collections.ArrayList] $testCases = @()
        foreach ($fullName in ($script:results.'.IO.FileInfo' | ForEach-Object { $_.GetType().FullName })) {
            $testCases.Add(@{
                fileInfoTypeFullName = $fullName
            })
        }

        It 'Default Properties: .IO.FileInfo' -TestCases $testCases {
            param($fileInfoTypeFullName)

            $fileInfoTypeFullName | Should Be 'System.IO.FileInfo'
        }
        
        [System.Collections.ArrayList] $testCases = @()
        foreach ($exists in $script:results.'.IO.FileInfo'.Exists) {
            $testCases.Add(@{
                    fileInfoExists = $exists
                })
        }
        
        It 'Default Properties: .IO.FileInfo Exists' -TestCases $testCases {
            param($fileInfoExists)

            $fileInfoExists | Should Be $true
        }
    }

    Remove-Variable -Scope 'Script' -Name 'fileCount' -Force -ErrorAction Ignore
    Get-ChildItem -LiteralPath $destPath -Filter '*.tmp' | Remove-Item -Force -Verbose

    foreach ($testMsiFile in $testMsiFiles) {
        Remove-Variable -Scope 'Script' -Name 'results' -Force -ErrorAction Ignore

        Context ('[{0}] Default Properties' -f $testMsiFile.BaseName) {
            It 'Default Properties: *Should Not Error*' {
                { $script:results = Get-MsiFileInfo -Path $testMsiFile.FullName } | Should Not Throw
            }

            Write-Host "Results: $($script:results | Out-String)"
            
            It 'Default Properties: .IO.FileInfo' {
                $script:results.'.IO.FileInfo' | Should BeOfType [System.IO.FileInfo]
            }
            
            It 'Default Properties: .IO.FileInfo Exists' {
                $script:results.'.IO.FileInfo'.Exists | Should Be $true
            }
            
            It 'Default Properties: Manufacturer' {
                $script:results.Manufacturer | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode String' {
                $script:results.ProductCode | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode Guid' {
                $script:results.ProductCode -as [guid] | Should BeOfType [System.Guid]
            }
            
            It 'Default Properties: ProductLanguage' {
                $script:results.ProductLanguage | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductName' {
                $script:results.ProductName | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion' {
                $script:results.ProductVersion | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion Version' {
                $script:results.ProductVersion -as [version] | Should BeOfType [System.Version]
            }
        }
        
        Remove-Variable -Scope 'Script' -Name 'results' -Force -ErrorAction Ignore

        Context ('[{0}] One Property' -f $testMsiFile.BaseName) {
            It 'Default Properties: *Should Not Error*' {
                { $script:results = Get-MsiFileInfo -Path $testMsiFile.FullName -Properties UpgradeCode } | Should Not Throw
            }

            Write-Host "Results: $($script:results | Out-String)"
            
            It 'Default Properties: .IO.FileInfo' {
                $script:results.'.IO.FileInfo' | Should BeOfType [System.IO.FileInfo]
            }
            
            It 'Default Properties: .IO.FileInfo Exists' {
                $script:results.'.IO.FileInfo'.Exists | Should Be $true
            }
            
            It 'Default Properties: Manufacturer' {
                $script:results.Manufacturer | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode String' {
                $script:results.ProductCode | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode Guid' {
                $script:results.ProductCode -as [guid] | Should BeOfType [System.Guid]
            }
            
            It 'Default Properties: ProductLanguage' {
                $script:results.ProductLanguage | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductName' {
                $script:results.ProductName | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion' {
                $script:results.ProductVersion | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion Version' {
                $script:results.ProductVersion -as [version] | Should BeOfType [System.Version]
            }
            
            It 'Default Properties: UpgradeCode' {
                $script:results.UpgradeCode | Should BeOfType [System.String]
            }
            
            It 'Default Properties: UpgradeCode Guid' {
                $script:results.UpgradeCode -as [guid] | Should BeOfType [System.Guid]
            }
        }
        
        Remove-Variable -Scope 'Script' -Name 'results' -Force -ErrorAction Ignore

        Context ('[{0}] Two Or More Properties' -f $testMsiFile.BaseName) {
            It 'Default Properties: *Should Not Error*' {
                { $script:results = Get-MsiFileInfo -Path $testMsiFile.FullName -Properties UpgradeCode,ALLUSERS } | Should Not Throw
            }

            Write-Host "Results: $($script:results | Out-String)"
            
            It 'Default Properties: .IO.FileInfo' {
                $script:results.'.IO.FileInfo' | Should BeOfType [System.IO.FileInfo]
            }
            
            It 'Default Properties: .IO.FileInfo Exists' {
                $script:results.'.IO.FileInfo'.Exists | Should Be $true
            }
            
            It 'Default Properties: Manufacturer' {
                $script:results.Manufacturer | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode String' {
                $script:results.ProductCode | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode Guid' {
                $script:results.ProductCode -as [guid] | Should BeOfType [System.Guid]
            }
            
            It 'Default Properties: ProductLanguage' {
                $script:results.ProductLanguage | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductName' {
                $script:results.ProductName | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion' {
                $script:results.ProductVersion | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion Version' {
                $script:results.ProductVersion -as [version] | Should BeOfType [System.Version]
            }
            
            if ($script:results.UpgradeCode) {
                It 'Default Properties: UpgradeCode' {
                    $script:results.UpgradeCode | Should BeOfType [System.String]
                }
                
                It 'Default Properties: UpgradeCode Guid' {
                    $script:results.UpgradeCode -as [guid] | Should BeOfType [System.Guid]
                }
            }
            
            if ($script:results.ALLUSERS) {
                It 'Default Properties: ALLUSERS' {
                    $script:results.ALLUSERS | Should BeOfType [System.String]
                }
            }
        }
        
        Remove-Variable -Scope 'Script' -Name 'results' -Force -ErrorAction Ignore

        Context ('[{0}] Property That Does Not Exist' -f $testMsiFile.BaseName) {
            $script:property = 'PropertyThatDoesNotExist-{0}' -f (New-Guid)
            It 'Default Properties: *Should Not Error*' {
                { $script:results = Get-MsiFileInfo -Path $testMsiFile.FullName -Properties $property } | Should Not Throw
            }

            Write-Host "Results: $($script:results | Out-String)"
            
            It 'Default Properties: .IO.FileInfo' {
                $script:results.'.IO.FileInfo' | Should BeOfType [System.IO.FileInfo]
            }
            
            It 'Default Properties: .IO.FileInfo Exists' {
                $script:results.'.IO.FileInfo'.Exists | Should Be $true
            }
            
            It 'Default Properties: Manufacturer' {
                $script:results.Manufacturer | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode String' {
                $script:results.ProductCode | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode Guid' {
                $script:results.ProductCode -as [guid] | Should BeOfType [System.Guid]
            }
            
            It 'Default Properties: ProductLanguage' {
                $script:results.ProductLanguage | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductName' {
                $script:results.ProductName | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion' {
                $script:results.ProductVersion | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion Version' {
                $script:results.ProductVersion -as [version] | Should BeOfType [System.Version]
            }
            
            It ('Default Properties: {0}' -f $script:property) {
                $script:results.$script:property | Should BeNullOrEmpty
            }
        }

        Remove-Variable -Scope 'Script' -Name 'results' -Force -ErrorAction Ignore
        Remove-Variable -Scope 'Script' -Name 'property' -Force -ErrorAction Ignore

        Context ('[{0}] All Properties' -f $testMsiFile.BaseName) {
            It 'Default Properties: *Should Not Error*' {
                { $script:results = Get-MsiFileInfo -Path $testMsiFile.FullName -Properties * } | Should Not Throw
            }

            Write-Host "Results: $($script:results.GetType() | Out-String)"
            Write-Host "Results: $($script:results | Out-String)"
            
            It 'Default Properties: .IO.FileInfo' {
                $script:results.'.IO.FileInfo' | Should BeOfType [System.IO.FileInfo]
            }
            
            It 'Default Properties: .IO.FileInfo Exists' {
                $script:results.'.IO.FileInfo'.Exists | Should Be $true
            }
            
            It 'Default Properties: Manufacturer' {
                $script:results.Manufacturer | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode String' {
                $script:results.ProductCode | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode Guid' {
                $script:results.ProductCode -as [guid] | Should BeOfType [System.Guid]
            }
            
            It 'Default Properties: ProductLanguage' {
                $script:results.ProductLanguage | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductName' {
                $script:results.ProductName | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion' {
                $script:results.ProductVersion | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion Version' {
                $script:results.ProductVersion -as [version] | Should BeOfType [System.Version]
            }

            It 'Default Properties: Count' {
                ($script:results.PSObject.Properties | Measure-Object).Count | Should BeGreaterThan 5
            }
        }

        Remove-Variable -Scope 'Script' -Name 'results' -Force -ErrorAction Ignore

        Context ('[{0}] Public Properties' -f $testMsiFile.BaseName) {
            It 'Default Properties: *Should Not Error*' {
                { $script:results = Get-MsiFileInfo -Path $testMsiFile.FullName -Properties * } | Should Not Throw
            }

            Write-Host "Results: $($script:results.GetType() | Out-String)"
            Write-Host "Results: $($script:results | Out-String)"
            
            It 'Default Properties: .IO.FileInfo' {
                $script:results.'.IO.FileInfo' | Should BeOfType [System.IO.FileInfo]
            }
            
            It 'Default Properties: .IO.FileInfo Exists' {
                $script:results.'.IO.FileInfo'.Exists | Should Be $true
            }
            
            It 'Default Properties: Manufacturer' {
                $script:results.Manufacturer | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode String' {
                $script:results.ProductCode | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode Guid' {
                $script:results.ProductCode -as [guid] | Should BeOfType [System.Guid]
            }
            
            It 'Default Properties: ProductLanguage' {
                $script:results.ProductLanguage | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductName' {
                $script:results.ProductName | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion' {
                $script:results.ProductVersion | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion Version' {
                $script:results.ProductVersion -as [version] | Should BeOfType [System.Version]
            }
            
            $script:publicPropertiesCount = 0
            foreach ($property in $script:results.PSObject.Properties) {
                if ($property.Name -cnotmatch '[a-z]') {
                    $script:publicPropertiesCount++

                    It ('Public Property: {0}' -f $property.Name) {
                        $property.Value | Should BeOfType [System.String]
                    }
                }
            }

            It 'Public Property: Count' {
                $script:publicPropertiesCount | Should BeGreaterThan 0
            }
        }

        Remove-Variable -Scope 'Script' -Name 'results' -Force -ErrorAction Ignore

        Context ('[{0}] Do Not Include File Info' -f $testMsiFile.BaseName) {
            It 'Default Properties: *Should Not Error*' {
                { $script:results = Get-MsiFileInfo -Path $testMsiFile.FullName -DoNotIncludeFileInfo } | Should Not Throw
            }

            It 'Default Properties: .IO.FileInfo Not Exists' {
                $script:results.PSObject.Properties -notcontains '.IO.FileInfo' | Should Be $true
            }
            
            It 'Default Properties: Manufacturer' {
                $script:results.Manufacturer | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode String' {
                $script:results.ProductCode | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductCode Guid' {
                $script:results.ProductCode -as [guid] | Should BeOfType [System.Guid]
            }
            
            It 'Default Properties: ProductLanguage' {
                $script:results.ProductLanguage | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductName' {
                $script:results.ProductName | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion' {
                $script:results.ProductVersion | Should BeOfType [System.String]
            }
            
            It 'Default Properties: ProductVersion Version' {
                $script:results.ProductVersion -as [version] | Should BeOfType [System.Version]
            }
        }

        Remove-Variable -Scope 'Script' -Name 'results' -Force -ErrorAction Ignore

        Context ('[{0}] MSI Does Not Exist' -f $testMsiFile.BaseName) {
            It 'Default Properties: *Should Not Error*' {
                { $script:results = Get-MsiFileInfo -Path ('{0}.msi' -f ((New-Guid).Guid -replace '{}', '')) -DoNotIncludeFileInfo } | Should Not Throw
            }

            Write-Host "Results: $($script:results.GetType() | Out-String)"
            Write-Host "Results: $($script:results | Out-String)"
            
            It 'Default Properties: .IO.FileInfo Not Exists' {
                $script:results.'.IO.FileInfo' | Should BeNullOrEmpty
            }
            
            It 'Default Properties: .IO.FileInfo Exists' {
                $script:results.'.IO.FileInfo'.Exists | Should BeNullOrEmpty
            }
            
            It 'Default Properties: Manufacturer' {
                $script:results.Manufacturer | Should BeNullOrEmpty
            }
            
            It 'Default Properties: ProductCode' {
                $script:results.ProductCode | Should BeNullOrEmpty
            }
            
            It 'Default Properties: ProductLanguage' {
                $script:results.ProductLanguage | Should BeNullOrEmpty
            }
            
            It 'Default Properties: ProductName' {
                $script:results.ProductName | Should BeNullOrEmpty
            }
            
            It 'Default Properties: ProductVersion' {
                $script:results.ProductVersion | Should BeNullOrEmpty
            }
        }
    }
}