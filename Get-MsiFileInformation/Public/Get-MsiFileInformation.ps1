function Get-MsiFileInformation {
    param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeLine = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [ValidateNotNullOrEmpty()]
        [ValidateScript( {
            if (([System.IO.FileInfo]$_).Extension -eq '.msi') {
                $true
            }
            else {
                throw 'Path must be a Windows Installer Database (*.msi) file.'
            }
        })]
        [System.IO.FileInfo]$Path,
 
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        # [ValidateSet('ProductCode', 'ProductVersion', 'ProductName', 'Manufacturer', 'ProductLanguage', 'FullVersion')]
        [string[]] $Properties
    )
    Begin {
        if (-not $Properties) {
            $query = 'SELECT * FROM Property'
        } else {
            $queryWhere = ($Properties | Foreach-Object { 'Property = "{0}"' -f $_ }) -join ' AND '
            $query = 'SELECT Value FROM Property WHERE {0}' -f $queryWhere
        }
    }

    Process {
        try {
            # Read property from MSI database
            $windowsInstaller = New-Object -ComObject windowsInstaller.Installer
            $msiDatabase = $windowsInstaller.GetType().InvokeMember('OpenDatabase', 'InvokeMethod', $null, $windowsInstaller, @($Path.FullName, 0))
            $view = $msiDatabase.GetType().InvokeMember('OpenView', 'InvokeMethod', $null, $msiDatabase, ($query))
            $view.GetType().InvokeMember('Execute', 'InvokeMethod', $null, $view, $null)

            $msiProperties = New-Object PSObject

            do {
                $record = $view.GetType().InvokeMember('Fetch', 'InvokeMethod', $null, $view, $null)

                if (-not [string]::IsNullOrEmpty($record)) {
                    # Return the value
                    $name = $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 1)
                    $value = $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 2)
                    $msiProperties | Add-Member -MemberType NoteProperty -Name $name -Value $value
                }
            } until ([string]::IsNullOrEmpty($record))

            Write-Output $msiProperties
 
            # Commit database and close view
            $msiDatabase.GetType().InvokeMember('Commit', 'InvokeMethod', $null, $msiDatabase, $null)
            $view.GetType().InvokeMember('Close', 'InvokeMethod', $null, $view, $null)           
            $msiDatabase = $null
            $view = $null
        } 
        catch {
            Write-Warning -Message $_.Exception.Message ; break
        }
    }
    
    End {
        # Run garbage collection and release ComObject
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($windowsInstaller) | Out-Null
        [System.GC]::Collect()
    }
}