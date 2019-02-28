This modules consists of several functions for handling MSI files.

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

Additional parameters available:

- Properties
- GetPublicProperties
- DoNotIncludeFileInfo

[Full usage documentations available in the wiki](https://github.com/UNT-CAS/MsiHandler/wiki/Get-MsiFileInfo).