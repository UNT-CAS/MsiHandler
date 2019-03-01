[string]           $projectDirectoryName = 'MsiHandler'
[IO.FileInfo]      $pesterFile = [io.fileinfo] ([string] (Resolve-Path -Path $MyInvocation.MyCommand.Path))
[IO.DirectoryInfo] $projectRoot = Split-Path -Parent $pesterFile.Directory
[IO.DirectoryInfo] $projectDirectory = Join-Path -Path $projectRoot -ChildPath $projectDirectoryName -Resolve
[IO.FileInfo]      $testFile = Join-Path -Path $projectDirectory -ChildPath (Join-Path -Path 'Public' -ChildPath ($pesterFile.Name -replace '\.Tests\.', '.')) -Resolve
. "${projectDirectory}\Private\Add-MsiType.ps1"
. "${projectDirectory}\Public\Get-MsiFileInfo.ps1"
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

Describe 'New-MsiTransformFile' {
    foreach ($testMsiFile in $testMsiFiles) {
        Context ('[{0}] Create MSTs' -f $testMsiFile.BaseName) {
            $script:newMsiTransformFile = @{
                MsiPath    = $testMsiFile
                MstPath    = 'dev\{0}.mst' -f $testMsiFile.BaseName
                Properties = @{
                    ALLUSERS           = 'ValueChanged'
                    MSIRMSHUTDOWN      = 'ValueChanged'
                    SOMETHINGTOADD     = 'ValueAdded'
                    SOMETHINGELSETOADD = 'ValueAdded'
                }
            }

            It 'New-MsiTransformFile: Should Work' {
                { $script:results = New-MsiTransformFile @script:newMsiTransformFile } | Should Not Throw
            }

            Write-Host "Results: $($script:results | Out-String)" -ForegroundColor Cyan
            Remove-Variable -Scope 'Script' -Name 'results' -Force -ErrorAction Ignore
            
            $script:newMsiTransformFile.MsiPath = New-TemporaryFile | Move-Item -Destination 'dev\TestMSIs\tmp.msi' -Force -Verbose
            
            It 'New-MsiTransformFile: Should Error' {
                { $script:results = New-MsiTransformFile @script:newMsiTransformFile } | Should Throw
            }
            
            Write-Host "Results: $($script:results | Out-String)" -ForegroundColor Cyan
            Remove-Variable -Scope 'Script' -Name 'results' -Force -ErrorAction Ignore
        }
    }
}