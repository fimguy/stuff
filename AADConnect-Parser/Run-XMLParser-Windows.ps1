# Execute AAD Connect XML parsing for Power BI (Windows Compatible Version)
# Run this script to parse your XML file and generate Power BI ready data files

# Set encoding for Windows compatibility
$OutputEncoding = [console]::InputEncoding = [console]::OutputEncoding = New-Object System.Text.UTF8Encoding

# Import the functions module with error handling
try {
    $modulePath = Join-Path $PSScriptRoot "AADConnect-PowerBI-Functions.psm1"
    Import-Module $modulePath -Force
    Write-Host "Module loaded successfully" -ForegroundColor Green
}
catch {
    Write-Error "Failed to load module: $($_.Exception.Message)"
    Write-Host "Please ensure AADConnect-PowerBI-Functions.psm1 is in the same directory" -ForegroundColor Red
    exit 1
}

# Configuration
$xmlFilePath = Join-Path $PSScriptRoot "today-RunStep.xml"
$outputDirectory = $PSScriptRoot

Write-Host "=== AAD Connect XML to Power BI Parser ===" -ForegroundColor Cyan
Write-Host "XML File: $xmlFilePath" -ForegroundColor White
Write-Host "Output Directory: $outputDirectory" -ForegroundColor White
Write-Host ""

# Check if XML file exists
if (-not (Test-Path $xmlFilePath)) {
    Write-Error "XML file not found: $xmlFilePath"
    Write-Host "Please ensure today-RunStep.xml is in the same directory as this script" -ForegroundColor Red
    exit 1
}

# Create output directory if it doesn't exist
if (-not (Test-Path $outputDirectory)) {
    try {
        New-Item -Path $outputDirectory -ItemType Directory -Force | Out-Null
        Write-Host "Created output directory: $outputDirectory" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to create output directory: $($_.Exception.Message)"
        exit 1
    }
}

try {
    # Run the complete parsing with advanced metrics
    Write-Host "Starting XML parsing..." -ForegroundColor Yellow
    Invoke-CompleteXMLParsing -XmlFilePath $xmlFilePath -OutputPath $outputDirectory -IncludeAdvancedMetrics
    
    Write-Host "`n=== Files Created for Power BI ===" -ForegroundColor Green
    
    # List all created files with better error handling
    try {
        $createdFiles = Get-ChildItem -Path $outputDirectory -Filter "*.csv" -ErrorAction SilentlyContinue | 
                       Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) }
        $jsonFiles = Get-ChildItem -Path $outputDirectory -Filter "*.json" -ErrorAction SilentlyContinue | 
                    Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) }
        
        if ($createdFiles) {
            Write-Host "`nCSV Files:" -ForegroundColor Yellow
            foreach ($file in $createdFiles) {
                $size = [math]::Round($file.Length / 1KB, 2)
                Write-Host "  $($file.Name) ($size KB)" -ForegroundColor White
            }
        }
        
        if ($jsonFiles) {
            Write-Host "`nJSON Files:" -ForegroundColor Yellow
            foreach ($file in $jsonFiles) {
                $size = [math]::Round($file.Length / 1KB, 2)
                Write-Host "  $($file.Name) ($size KB)" -ForegroundColor White
            }
        }
    }
    catch {
        Write-Warning "Could not list created files: $($_.Exception.Message)"
    }
    
    Write-Host "`n=== Power BI Import Guide ===" -ForegroundColor Cyan
    Write-Host "STEP 1: Open Power BI Desktop" -ForegroundColor White
    Write-Host "STEP 2: Get Data > Text/CSV" -ForegroundColor White
    Write-Host "STEP 3: Import these files as separate tables:" -ForegroundColor White
    Write-Host "  - RunStepResults.csv (main summary data)" -ForegroundColor Gray
    Write-Host "  - ExportErrors.csv (error details)" -ForegroundColor Gray
    Write-Host "  - EnhancedExportErrors.csv (errors with user classification)" -ForegroundColor Gray
    Write-Host "  - ErrorSummary.csv (aggregated error counts)" -ForegroundColor Gray
    Write-Host "  - ConnectionInfo.csv (connection status)" -ForegroundColor Gray
    Write-Host "  - AggregatedMetrics.csv (key performance indicators)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "STEP 4: Create relationships in Model view:" -ForegroundColor White
    Write-Host "  - RunStepResults[RunHistoryId] -> ExportErrors[RunHistoryId]" -ForegroundColor Gray
    Write-Host "  - RunStepResults[StepHistoryId] -> ExportErrors[StepHistoryId]" -ForegroundColor Gray
    Write-Host ""
    Write-Host "STEP 5: Create visualizations using:" -ForegroundColor White
    Write-Host "  - Time series: Use Start/End date columns" -ForegroundColor Gray
    Write-Host "  - Error analysis: Use ErrorType, UserType columns" -ForegroundColor Gray
    Write-Host "  - Performance: Use DurationSeconds, operation counts" -ForegroundColor Gray
    Write-Host ""
    Write-Host "=== Suggested Power BI Visuals ===" -ForegroundColor Cyan
    Write-Host "Line Chart: Steps over time (StartDate vs StepNumber)" -ForegroundColor White
    Write-Host "Bar Chart: Error types distribution (ErrorType vs Count)" -ForegroundColor White
    Write-Host "Card Visuals: Total errors, successful steps, avg duration" -ForegroundColor White
    Write-Host "Table: Latest errors with user names and retry counts" -ForegroundColor White
    Write-Host "Pie Chart: User types affected by errors" -ForegroundColor White
    Write-Host "Matrix: Errors by month/day pattern" -ForegroundColor White
    Write-Host ""
    Write-Host "Parsing completed successfully!" -ForegroundColor Green
    
    # Pause for user to read output on Windows
    Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    
}
catch {
    Write-Error "Failed to parse XML: $($_.Exception.Message)"
    Write-Host "Error details:" -ForegroundColor Red
    Write-Host $_.Exception.StackTrace -ForegroundColor Red
    Write-Host "`nPress any key to continue..." -ForegroundColor Yellow
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit 1
}
