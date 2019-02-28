$signature = @' 
using System.Text;
using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;

public static class Msi
{
    [DllImport("msi.dll", SetLastError = true)]
    private static extern uint MsiOpenDatabase(string szDatabasePath, IntPtr phPersist, out IntPtr phDatabase);
    
    [DllImport("msi.dll", CharSet = CharSet.Unicode)]
    private static extern uint MsiDatabaseOpenView(IntPtr hDatabase, [MarshalAs(UnmanagedType.LPWStr)] string szQuery, out IntPtr phView);

    [DllImport("msi.dll", CharSet = CharSet.Unicode)]
    private static extern uint MsiViewFetch(IntPtr hView, out IntPtr hRecord);

    [DllImport("msi.dll", CharSet = CharSet.Unicode)]
    private static extern uint MsiViewExecute(IntPtr hView, IntPtr hRecord);

    [DllImport("msi.dll", CharSet = CharSet.Unicode)]
    private static extern uint MsiRecordGetString(IntPtr hRecord, int iField, System.Text.StringBuilder szValueBuf, ref int pcchValueBuf);

    [DllImport("msi.dll", CharSet = CharSet.Unicode)]
    private static extern uint MsiRecordSetString(IntPtr hRecord, int iField, string szValue);

    [DllImport("msi.dll", ExactSpelling = true)]
    private static extern IntPtr MsiCreateRecord(uint cParams);

    [DllImport("msi.dll", ExactSpelling = true)]
    private static extern uint MsiViewModify(IntPtr hView, uint eModifyMode, IntPtr hRecord);

    [DllImport("msi.dll", SetLastError = true)]
    private static extern uint MsiDatabaseGenerateTransform(IntPtr hDatabase, IntPtr hDatabaseReference, string szTransformFile, uint iReserved1, uint iReserved2);

    [DllImport("msi.dll", ExactSpelling = true)]
    private static extern uint MsiCloseHandle(IntPtr hAny);

    [DllImport("msi.dll")]
    private static extern uint MsiCreateTransformSummaryInfo(IntPtr hDatabase, IntPtr hDatabaseReference, string szTransformFile, int iErrorConditions, int iValidation);

    private static readonly Dictionary<string,string> replacements = new Dictionary<string,string>();
    private static readonly Dictionary<string,string> additions = new Dictionary<string,string>();
    private static readonly List<IntPtr> openHandles = new List<IntPtr>();

    private static string tempFile = System.IO.Path.GetTempFileName();

    public static void ClearReplacements()
    {
        replacements.Clear();
    }
    
    public static void AddReplacement(string key, string value)
    {
        replacements.Add(key, value);
    }
    
    public static void ClearAdditions()
    {
        additions.Clear();
    }
    
    public static void AddAddition(string key, string value)
    {
        additions.Add(key, value);
    }
    
    public static void CreateTransform(string msiPath, string mstPath)
    {		
        IntPtr msiInputHandle;
        IntPtr msiAlteredHandle;
        IntPtr msiDatabaseView;
        IntPtr currentRecord;
        uint result;
        bool shouldContinue = true;
        
        System.IO.File.Delete(tempFile);
        System.IO.File.Copy(msiPath, tempFile);
        
        if ((result = MsiOpenDatabase(msiPath, (IntPtr)0, out msiInputHandle)) != 0)
        {
            CleanUp("ERROR MsiOpenDatabase " + result);			
            return;
        }

        openHandles.Add(msiInputHandle);
        
        if ((result = MsiOpenDatabase(tempFile, (IntPtr)2, out msiAlteredHandle)) != 0)
        {
            CleanUp("ERROR MsiOpenDatabase " + result);			
            return;
        }
        
        openHandles.Add(msiAlteredHandle);
        
        if ((result = MsiDatabaseOpenView(msiAlteredHandle, "SELECT Property, Value FROM Property", out msiDatabaseView)) != 0)
        {
            CleanUp("ERROR MsiDatabaseOpenView " + result);			
            return;
        }		
        
        openHandles.Add(msiDatabaseView);
        
        if ((result = MsiViewExecute(msiDatabaseView, IntPtr.Zero)) != 0)
        {
            CleanUp("ERROR MsiViewExecute " + result);			
            return;
        }
        
        while (shouldContinue)
        {
            result = MsiViewFetch(msiDatabaseView, out currentRecord);
            
            if (result == 259)
            {
                shouldContinue = false;
            }
            else if (result == 0)
            {
                openHandles.Add(currentRecord);
                
                StringBuilder builder = new System.Text.StringBuilder(256);
                int count = builder.Capacity;

                if ((result = MsiRecordGetString(currentRecord, 1, builder, ref count)) != 0)
                {
                    CleanUp("ERROR MsiRecordGetString " + result);
                    return;
                }

                string key = builder.ToString().Trim();

                if (replacements.ContainsKey(key))
                {
                    if ((result = MsiRecordSetString(currentRecord, 2, replacements[key])) != 0)
                    {
                        CleanUp("ERROR MsiRecordSetString " + result);
                        return;
                    }
                    
                    if ((result = MsiViewModify(msiDatabaseView, 2, currentRecord)) != 0)
                    {
                        CleanUp("ERROR MsiViewModify " + result);
                        return;
                    }                        						
                }                
                                
                MsiCloseHandle(currentRecord);
                openHandles.Remove(currentRecord);
            }
            else
            {
                CleanUp("ERROR MsiViewFetch " + result);
                return;
            }
        }
        
        foreach (KeyValuePair<string,string> item in additions)
        {
            IntPtr newRecord = MsiCreateRecord(2);
            openHandles.Add(newRecord);

            if ((result = MsiRecordSetString(newRecord, 1, item.Key)) != 0)
            {
                CleanUp("ERROR MsiRecordSetString " + result);
                return;
            }

            if ((result = MsiRecordSetString(newRecord, 2, item.Value)) != 0)
            {
                CleanUp("ERROR MsiRecordSetString " + result);
                return;
            }

            if ((result = MsiViewModify(msiDatabaseView, 1, newRecord)) != 0)
            {
                CleanUp("ERROR MsiViewModify " + result);
                return;
            }

            MsiCloseHandle(newRecord);
            openHandles.Remove(newRecord);
        }
        
        if ((result = MsiDatabaseGenerateTransform(msiAlteredHandle, msiInputHandle, mstPath, 0, 0)) != 0)
        {
            CleanUp("ERROR MsiDatabaseGenerateTransform " + result);
            return;
        }
        
        if ((result = MsiCreateTransformSummaryInfo(msiAlteredHandle, msiInputHandle, mstPath, 0, 0)) != 0)
        {
            CleanUp("ERROR MsiCreateTransformSummaryInfo " + result);
            return;
        }

        CleanUp("OK CreatedTransformFile " + mstPath);
    }
    
    private static void CleanUp(string message)
    {
        Console.WriteLine(message);
        
        foreach(IntPtr handle in openHandles)
        {
            MsiCloseHandle(handle);
        }
        
        openHandles.Clear();
        System.IO.File.Delete(tempFile);
    }
}
'@
Add-Type -TypeDefinition $signature


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

    [Msi]::ClearReplacements()
    [Msi]::ClearAdditions()

    $existingProperties = (Get-MsiFileInfo $MsiPath -Properties *).PSObject.Properties.Name

    foreach ($property in $Properties.GetEnumerator()) {
        if ($existingProperties -contains $property.Name) {
            [Msi]::AddReplacement($property.Name, $property.Value)
        } else {
            [Msi]::AddAddition($property.Name, $property.Value)
        }
    }

    # Create a new STDOUT for catching assembly output
    $writer = New-Object IO.StringWriter
    [Console]::SetOut($writer)

    # This doesn't write to a PowerShell stream
    #   It writes, but to something else.
    #   Redirect all (*>&1) doesn't catch the output.
    [Msi]::CreateTransform($msiPath, $mstPath)

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

    return $mstPath
}