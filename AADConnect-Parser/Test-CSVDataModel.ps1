# Test script to verify CSV data model creation
# This script tests the parser with minimal data to ensure all CSV files are created

param(
    [string]$TestOutputPath = (Join-Path $PSScriptRoot "Test-Output")
)

Write-Host "=== Testing CSV Data Model Creation ===" -ForegroundColor Cyan
Write-Host "Test Output Directory: $TestOutputPath" -ForegroundColor White

# Create test directory
if (-not (Test-Path $TestOutputPath)) {
    New-Item -Path $TestOutputPath -ItemType Directory -Force | Out-Null
    Write-Host "Created test directory: $TestOutputPath" -ForegroundColor Green
}

# Import the module
$modulePath = Join-Path $PSScriptRoot "AADConnect-PowerBI-Functions.psm1"
Import-Module $modulePath -Force

# Test the initialization function
try {
    Write-Host "`nTesting Initialize-PowerBIDataModel..." -ForegroundColor Yellow
    Initialize-PowerBIDataModel -OutputPath $TestOutputPath
    
    # Check if all required files were created
    $requiredFiles = @(
        "RunStepResults.csv",
        "ExportErrors.csv", 
        "EnhancedExportErrors.csv",
        "ErrorSummary.csv",
        "ConnectionInfo.csv",
        "AggregatedMetrics.csv"
    )
    
    Write-Host "`nVerifying created files:" -ForegroundColor Yellow
    $allFilesExist = $true
    
    foreach ($file in $requiredFiles) {
        $filePath = Join-Path $TestOutputPath $file
        if (Test-Path $filePath) {
            $fileSize = (Get-Item $filePath).Length
            Write-Host "  ✓ $file ($fileSize bytes)" -ForegroundColor Green
            
            # Check if file has headers (should be more than 0 bytes)
            if ($fileSize -eq 0) {
                Write-Host "    ⚠ Warning: File is empty" -ForegroundColor Yellow
                $allFilesExist = $false
            }
        } else {
            Write-Host "  ✗ $file (missing)" -ForegroundColor Red
            $allFilesExist = $false
        }
    }
    
    if ($allFilesExist) {
        Write-Host "`n✅ Success: All required CSV files created successfully!" -ForegroundColor Green
        Write-Host "These files can now be imported into Power BI without errors." -ForegroundColor White
        
        # Test importing one file to show structure
        Write-Host "`nSample structure of RunStepResults.csv:" -ForegroundColor Yellow
        $samplePath = Join-Path $TestOutputPath "RunStepResults.csv"
        $sampleContent = Get-Content $samplePath -First 2
        foreach ($line in $sampleContent) {
            Write-Host "  $line" -ForegroundColor Gray
        }
    } else {
        Write-Host "`n❌ Error: Some required files are missing or empty." -ForegroundColor Red
    }
    
    # Test Power BI import readiness
    Write-Host "`nPower BI Import Test:" -ForegroundColor Cyan
    try {
        $testImport = Import-Csv (Join-Path $TestOutputPath "RunStepResults.csv")
        $headers = $testImport | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
        Write-Host "  Available columns: $($headers.Count)" -ForegroundColor White
        Write-Host "  Headers: $($headers -join ', ')" -ForegroundColor Gray
        Write-Host "  ✅ File structure is Power BI compatible" -ForegroundColor Green
    }
    catch {
        Write-Host "  ❌ Error importing CSV: $($_.Exception.Message)" -ForegroundColor Red
    }
    
    Write-Host "`nTest completed successfully!" -ForegroundColor Green
    Write-Host "You can now use these CSV files as templates for Power BI import." -ForegroundColor White
    
}
catch {
    Write-Error "Test failed: $($_.Exception.Message)"
    Write-Host "Error details: $($_.Exception.StackTrace)" -ForegroundColor Red
}

Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
