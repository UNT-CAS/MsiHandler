function Get-MsiFileInfo {
    [OutputType([hashtable])]
    param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeLine = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [ValidateNotNullOrEmpty()]
        [IO.FileInfo] $Path,
 
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string[]] $Properties = @('Manufacturer', 'ProductName', 'ProductVersion', 'ProductCode', 'ProductLanguage', 'FullVersion'),

        [parameter(Mandatory = $false)]
        [switch] $GetPublicProperties,

        [parameter(Mandatory = $false)]
        [switch] $DoNotIncludeFileInfo,

        [parameter(Mandatory = $false)]
        [switch] $IncludeNonMsiFileInfo
    )
    Begin {
        $alwaysGetProperties = @('Manufacturer', 'ProductName', 'ProductVersion', 'ProductCode', 'ProductLanguage', 'FullVersion')
        
        if (-not $Properties -or $Properties -eq '*' -or $GetPublicProperties.IsPresent) {
            $query = 'SELECT * FROM Property'
        } else {
            $queryWhere = (($Properties + $alwaysGetProperties | Select-Object -Unique) | Foreach-Object { 'Property = ''{0}''' -f $_ }) -join ' OR '
            $query = 'SELECT Property, Value FROM Property WHERE {0}' -f $queryWhere
        }
        
        Write-Verbose "[Get-MsiFileInfo] MSI Query: ${query}"
    }
    
    Process {
        Write-Verbose "[Get-MsiFileInfo] Path: ${Path}"

        if (-not $Path.Exists) {
            $resolvedPath = (Resolve-Path $Path -ErrorAction 'Ignore').ProviderPath
            Write-Verbose "[Get-MsiFileInfo] ResolvedPath: ${resolvedPath}"
            if ($resolvedPath) {
                [IO.FileInfo] $Path = $resolvedPath
            }
        }

        [hashtable] $msiProperties = @{}
        if ($IncludeNonMsiFileInfo.IsPresent -or -not $DoNotIncludeFileInfo.IsPresent) {
            $msiProperties.Add('.IO.FileInfo', $Path)
        }

        if ($Path.Exists) {
            $windowsInstaller = New-Object -ComObject windowsInstaller.Installer
            try {
                $msiDatabase = $windowsInstaller.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstaller, @($Path.FullName, 0))
                $view = $msiDatabase.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $msiDatabase, ($query))
                [void] $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)
    
                do {
                    $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)
    
                    if (-not [string]::IsNullOrEmpty($record)) {
                        $addMember = $false
    
                        # Return the value
                        $name = $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1)
                        Write-Debug "  [Get-MsiFileInfo] 1 (name): ${name}"
                        $value = $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 2)
                        Write-Debug "  [Get-MsiFileInfo] 2 (value): ${value}"
                        if ($GetPublicProperties.IsPresent) {
                            if ($alwaysGetProperties -contains $name) {
                                $addMember = $true
                            } elseif ($name -cnotmatch '[a-z]') {
                                $addMember = $true
                            }
                        } else {
                            $addMember = $true
                        }
                        
                        if ($addMember) {
                            Write-Debug "  [Get-MsiFileInfo] Adding to return set."
                            [void] $msiProperties.Add($name, $value)
                        }
                    }
                } until ([string]::IsNullOrEmpty($record))

                # Commit database and close view
                [void] $msiDatabase.GetType().InvokeMember('Commit', 'InvokeMethod', $null, $msiDatabase, $null)
                [void] $view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null)           
            } catch {
                Write-Debug ('[Get-MsiFileInfo] Error Caught' -f $_.Exception.Message)
                Write-Warning ('Unable to open MSI database; it''s either not an MSI file or the file is corrupted: {0}' -f $Path.FullName)
            } finally {
                $view = $null
                $msiDatabase = $null
                [void] [System.Runtime.Interopservices.Marshal]::ReleaseComObject($windowsInstaller)
                $windowsInstaller = $null
            }
        }

        Write-Debug ('msiProperties: {0}' -f ($msiProperties | Out-String))
        # Write-Output $msiProperties
        Write-Output (New-Object PSObject -Property $msiProperties)
    }
    
    End {
        # Run garbage collection and release ComObject
        if ($windowsInstaller) {
            [void] [System.Runtime.Interopservices.Marshal]::ReleaseComObject($windowsInstaller)
        }
        [void] [System.GC]::Collect()
    }
}