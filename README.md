[![Build status](https://ci.appveyor.com/api/projects/status/2dv8llmk58ix7i83?svg=true)](https://ci.appveyor.com/project/VertigoRay/msihandler)
[![codecov](https://codecov.io/gh/UNT-CAS/MsiHandler/branch/master/graph/badge.svg)](https://codecov.io/gh/UNT-CAS/MsiHandler)
[![version](https://img.shields.io/powershellgallery/v/MsiHandler.svg)](https://www.powershellgallery.com/packages/MsiHandler)
[![downloads](https://img.shields.io/powershellgallery/dt/MsiHandler.svg?label=downloads)](https://www.powershellgallery.com/stats/packages/MsiHandler?groupby=Version)

This modules consists of several functions for handling MSI files.

# Quick Start

```powershell
Install-Module MsiHandler
Import-Module MsiHandler
```

# Available Functions

## Get-MsiFileInfo

This function will read the properties out of an MSI file without locking the specified file.
The MSI file's `[IO.FileInfo]` is returned as the `.IO.FileInfo` property.
Here's the default usage:

```powershell
Get-MsiFileInfo -Path 'C:\Temp\SurfaceBook_Win10_15063_1802100_0.msi'
```

**Output:**

```text
.IO.FileInfo    : C:\Temp\SurfaceBook_Win10_15063_1802100_0.msi
Manufacturer    : Microsoft
ProductCode     : {1CD69D1F-0C2D-46A0-89A9-F29582DC718F}
ProductLanguage : 1033
ProductName     : SurfaceBook Update 18_021_00 (64 bit)
ProductVersion  : 18.021.18206.0
```

[Full usage documentations available in the wiki](https://github.com/UNT-CAS/MsiHandler/wiki/Get-MsiFileInfo).

## New-MsiTransformFile

Create an MST file on the fly.

```powershell
$newMsiTransformFile = @{
    MsiPath      = 'dev\TestMSIs\7z1900-x64.msi'
    MstPath      = 'dev\test.mst'
    Properties = @{
        ALLUSERS           = 'ValueChanged'
        MSIRMSHUTDOWN      = 'ValueChanged'
        SOMETHINGTOADD     = 'ValueAdded'
        SOMETHINGELSETOADD = 'ValueAdded'
    }
}

New-MsiTransformFile @newMsiTransformFile
```

**Output:**

```text
Mode                LastWriteTime         Length Name
----                -------------         ------ ----
-a----        2/28/2019   3:46 AM          20480 test.mst
```

[Full usage documentations available in the wiki](https://github.com/UNT-CAS/MsiHandler/wiki/New-MsiTransformFile).
