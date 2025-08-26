# Execute AAD Connect XML parsing for Power BI
# Run this script to parse your XML file and generate Power BI ready data files

# Import the functions module
Import-Module "/Users/steady/Desktop/AADCON/AADConnect-PowerBI-Functions.psm1" -Force

# Configuration
$xmlFilePath = "/Users/steady/Desktop/AADCON/today-RunStep.xml"
$outputDirectory = "/Users/steady/Desktop/AADCON/"

Write-Host "=== AAD Connect XML to Power BI Parser ===" -ForegroundColor Cyan
Write-Host "XML File: $xmlFilePath" -ForegroundColor White
Write-Host "Output Directory: $outputDirectory" -ForegroundColor White
Write-Host ""

# Check if XML file exists
if (-not (Test-Path $xmlFilePath)) {
    Write-Error "XML file not found: $xmlFilePath"
    exit 1
}

# Create output directory if it doesn't exist
if (-not (Test-Path $outputDirectory)) {
    New-Item -Path $outputDirectory -ItemType Directory -Force
    Write-Host "Created output directory: $outputDirectory" -ForegroundColor Green
}

try {
    # Run the complete parsing with advanced metrics
    Invoke-CompleteXMLParsing -XmlFilePath $xmlFilePath -OutputPath $outputDirectory -IncludeAdvancedMetrics
    
    Write-Host "`n=== Files Created for Power BI ===" -ForegroundColor Green
    
    # List all created files
    $createdFiles = Get-ChildItem -Path $outputDirectory -Filter "*.csv" | Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) }
    $jsonFiles = Get-ChildItem -Path $outputDirectory -Filter "*.json" | Where-Object { $_.LastWriteTime -gt (Get-Date).AddMinutes(-5) }
    
    Write-Host "`nCSV Files:" -ForegroundColor Yellow
    foreach ($file in $createdFiles) {
        $size = [math]::Round($file.Length / 1KB, 2)
        Write-Host "  $($file.Name) ($size KB)" -ForegroundColor White
    }
    
    Write-Host "`nJSON Files:" -ForegroundColor Yellow
    foreach ($file in $jsonFiles) {
        $size = [math]::Round($file.Length / 1KB, 2)
        Write-Host "  $($file.Name) ($size KB)" -ForegroundColor White
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
    
}
catch {
    Write-Error "Failed to parse XML: $($_.Exception.Message)"
    Write-Host "Error details: $($_.Exception.StackTrace)" -ForegroundColor Red
}
