# Windows Troubleshooting Guide for AAD Connect XML Parser

## Common Issues and Solutions

### 1. PowerShell Execution Policy Error
**Error:** "Execution of scripts is disabled on this system"

**Solution:**
```powershell
# Run PowerShell as Administrator and execute:
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

Or use the batch file `Run-Parser.bat` which bypasses the execution policy.

### 2. Unicode/Encoding Issues
**Error:** "The string is missing the terminator" or weird characters

**Solutions:**
- Use `Run-XMLParser-Windows.ps1` instead of the original script
- Ensure your XML file is saved with UTF-8 encoding
- Run the script from PowerShell ISE instead of command prompt

### 3. Module Loading Issues
**Error:** "Failed to load module"

**Solutions:**
- Ensure all files are in the same directory
- Right-click the PowerShell script and select "Run with PowerShell"
- Copy all files to `C:\Temp\AADCON\` and run from there

### 4. Path Issues
**Error:** "XML file not found" or "Path not found"

**Solutions:**
- Place your `today-RunStep.xml` file in the same folder as the scripts
- Use the `Run-XMLParser-Windows.ps1` script which uses relative paths
- Avoid spaces in folder names

### 5. System.Web Assembly Issues
**Error:** "Could not load System.Web"

**Solution:**
The script now includes fallback HTML decoding. If you still get errors:
```powershell
# Install .NET Framework 4.7.2 or later
# Or use this alternative decode method (already included in the script)
```

## Recommended Setup Steps for Windows

1. **Create a dedicated folder:**
   ```
   C:\AADConnect-Parser\
   ```

2. **Copy these files to the folder:**
   - `AADConnect-PowerBI-Functions.psm1`
   - `Parse-RunStepXML.ps1`
   - `Run-XMLParser-Windows.ps1`
   - `Run-Parser.bat`
   - `today-RunStep.xml` (your data file)

3. **Run the parser:**
   - Double-click `Run-Parser.bat` for easiest execution
   - Or right-click `Run-XMLParser-Windows.ps1` → "Run with PowerShell"

## PowerShell Version Requirements

**Minimum:** PowerShell 5.0 (Windows 10/Server 2016)
**Recommended:** PowerShell 5.1 or PowerShell 7.x

Check your version:
```powershell
$PSVersionTable.PSVersion
```

## File Locations After Processing

The script creates these files in the same directory:
- `RunStepResults.csv`
- `ExportErrors.csv`
- `EnhancedExportErrors.csv`
- `ErrorSummary.csv`
- `AggregatedMetrics.csv`
- `ConnectionInfo.csv`
- `AADConnect_RunStep_Data.json`

## Alternative Method: PowerShell ISE

If the command-line methods fail:

1. Open PowerShell ISE as Administrator
2. Open `Parse-RunStepXML.ps1`
3. Modify the file paths at the top if needed:
   ```powershell
   $XmlFilePath = "C:\Path\To\Your\today-RunStep.xml"
   $OutputPath = "C:\Path\To\Output\"
   ```
4. Press F5 to run

## Power BI Import Tips for Windows

1. **Open Power BI Desktop**
2. **Get Data** → **Text/CSV**
3. **Navigate to your output folder**
4. **Import each CSV file as a separate table**
5. **In Power BI Model view, create relationships:**
   - `RunStepResults[RunHistoryId]` ↔ `ExportErrors[RunHistoryId]`

## Contact Information

If you continue to have issues:
1. Check the PowerShell error messages carefully
2. Ensure you have the latest version of PowerShell
3. Try running the script from a simple path like `C:\Temp\`
4. Consider using PowerShell 7.x which has better cross-platform compatibility

## Quick Test

To verify everything works, run this simple test:
```powershell
# Test PowerShell and module loading
Get-Module -ListAvailable
Test-Path ".\today-RunStep.xml"
```

Both commands should work without errors before running the main script.
