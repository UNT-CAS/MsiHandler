function New-MsiTransformFile {
    param(
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [ValidateScript({ Test-Path $_ })]
        [IO.FileInfo] $MsiPath,
        
        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [IO.FileInfo] $MstPath,

        [parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [hashtable] $Properties
    )

    Write-Verbose "[New-MsiTransformFile] Param MsiPath: ${MsiPath}"
    Write-Verbose "[New-MsiTransformFile] Param MstPath: ${MstPath}"
    
    [Msi]::ClearReplacements()
    [Msi]::ClearAdditions()

    $msiFileInfo = Get-MsiFileInfo $MsiPath -Properties *
    $existingProperties = $msiFileInfo.PSObject.Properties.Name

    foreach ($property in $Properties.GetEnumerator()) {
        if ($existingProperties -contains $property.Name) {
            [Msi]::AddReplacement($property.Name, $property.Value)
        } else {
            [Msi]::AddAddition($property.Name, $property.Value)
        }
    }
    
    Write-Verbose "[New-MsiTransformFile] MsiPath: $($msiFileInfo.'.IO.FileInfo')"
    Write-Verbose "[New-MsiTransformFile] MstPath: ${MstPath}"

    # Create a new STDOUT for catching assembly output
    $writer = New-Object IO.StringWriter
    [Console]::SetOut($writer)

    # This doesn't write to a PowerShell stream
    #   It writes, but to something else.
    #   Redirect all (*>&1) doesn't catch the output.
    [Msi]::CreateTransform($msiFileInfo.'.IO.FileInfo', $MstPath)

    # Store the output and bring back real STDOUT
    $result = $writer.ToString()
    $standardOutput = New-Object IO.StreamWriter([Console]::OpenStandardOutput())
    $standardOutput.AutoFlush = $true
    [Console]::SetOut($standardOutput)
    
    Write-Verbose "[New-MsiTransformFile] MSI CreateTransform: ${result}"
    
    [Collections.ArrayList] $msg = $result.Split(' ')

    if ($msg[0] -eq 'ERROR') {
        $msg.RemoveAt(0)
        Throw ($msg -join ' ')
    }

    # $MstPath.Refresh()
    return $MstPath
}


# $script:newMsiTransformFile = @{
#     MsiPath    = '.\dev\testMSIs\Firefox%20Setup%2065.0.msi'
#     MstPath    = '.\dev\foo.mst'
#     Properties = @{
#         ALLUSERS           = 'ValueChanged'
#         MSIRMSHUTDOWN      = 'ValueChanged'
#         SOMETHINGTOADD     = 'ValueAdded'
#         SOMETHINGELSETOADD = 'ValueAdded'
#     }
# }

# New-MsiTransformFile @script:newMsiTransformFile