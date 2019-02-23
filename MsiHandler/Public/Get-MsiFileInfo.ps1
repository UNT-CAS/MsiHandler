function Get-MsiFileInfo {
    [OutputType([PSObject])]
    param(
        [Parameter(
            Mandatory = $true,
            ValueFromPipeLine = $true,
            ValueFromPipelineByPropertyName = $true
        )]
        [ValidateNotNullOrEmpty()]
        [System.IO.FileInfo] $Path,
 
        [parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string[]] $Properties
    )
    Begin {
        if (-not $Properties) {
            $query = 'SELECT * FROM Property'
        } else {
            $queryWhere = ($Properties | Foreach-Object { 'Property = ''{0}''' -f $_ }) -join ' OR '
            $query = 'SELECT Property, Value FROM Property WHERE {0}' -f $queryWhere
        }
        Write-Verbose "[Get-MsiFileInfo] MSI Query: ${query}"
    }
    
    Process {
        Write-Verbose "[Get-MsiFileInfo] Path: ${Path}"
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
                    Write-Verbose "[Get-MsiFileInfo] 1 (name): ${name}"
                    $value = $record.GetType().InvokeMember('StringData', 'GetProperty', $null, $record, 2)
                    Write-Verbose "[Get-MsiFileInfo] 2 (value): ${value}"
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

$VerbosePreference = 'c'

# Get-ChildItem -LiteralPath 'C:\Users\verti\Downloads' -Filter '*.msi' -Recurse | Get-MsiFileInfo
# Get-MsiFileInfo -Path 'C:\Users\verti\Downloads\Rapid7_x64\agentInstaller-x86_64.msi'
# Get-MsiFileInfo -Path 'C:\Users\verti\Downloads\Rapid7_x64\agentInstaller-x86_64.msi' -Properties ProductName
# Get-MsiFileInfo -Path 'C:\Users\verti\Downloads\Rapid7_x64\agentInstaller-x86_64.msi' -Properties ProductName,ProductVersion
Get-MsiFileInfo -Path 'C:\Users\verti\Downloads\Rapid7_x64\agentInstaller-x86_64.msi' -Properties ProductNamesss