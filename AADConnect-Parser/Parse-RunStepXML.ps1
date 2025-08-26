# PowerShell script to parse AAD Connect RunStep XML for Power BI import
# This script extracts key data from the Microsoft Identity Management XML export

param(
    [Parameter(Mandatory=$false)]
    [string]$XmlFilePath = (Join-Path $PSScriptRoot "today-RunStep.xml"),
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath = $PSScriptRoot,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportToCSV = $true,
    
    [Parameter(Mandatory=$false)]
    [switch]$ExportToJSON = $true
)

Write-Host "Starting XML parsing for Power BI import..." -ForegroundColor Green

# Validate input file exists
if (-not (Test-Path $XmlFilePath)) {
    throw "XML file not found: $XmlFilePath"
}

# Ensure output directory exists
if (-not (Test-Path $OutputPath)) {
    New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    Write-Host "Created output directory: $OutputPath" -ForegroundColor Yellow
}

try {
    # Load the XML file
    Write-Host "Loading XML file: $XmlFilePath" -ForegroundColor Yellow
    [xml]$xmlContent = Get-Content -Path $XmlFilePath -Raw
    
    # Initialize collections for different data types
    $runStepResults = @()
    $syncErrors = @()
    $exportErrors = @()
    
    # Parse RunStepResult objects
    Write-Host "Parsing RunStep results..." -ForegroundColor Yellow
    
    foreach ($obj in $xmlContent.Objs.Obj) {
        if ($obj.TN.T -contains "Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepResult") {
            
            $runStep = [PSCustomObject]@{
                RunHistoryId = $obj.Props.G | Where-Object { $_.N -eq "RunHistoryId" } | Select-Object -ExpandProperty '#text'
                StepHistoryId = $obj.Props.G | Where-Object { $_.N -eq "StepHistoryId" } | Select-Object -ExpandProperty '#text'
                StepNumber = $obj.Props.I32 | Where-Object { $_.N -eq "StepNumber" } | Select-Object -ExpandProperty '#text'
                StepResult = $obj.Props.S | Where-Object { $_.N -eq "StepResult" } | Select-Object -ExpandProperty '#text'
                StartDate = $obj.Props.DT | Where-Object { $_.N -eq "StartDate" } | Select-Object -ExpandProperty '#text'
                EndDate = $obj.Props.DT | Where-Object { $_.N -eq "EndDate" } | Select-Object -ExpandProperty '#text'
                StageNoChange = $obj.Props.I32 | Where-Object { $_.N -eq "StageNoChange" } | Select-Object -ExpandProperty '#text'
                StageAdd = $obj.Props.I32 | Where-Object { $_.N -eq "StageAdd" } | Select-Object -ExpandProperty '#text'
                StageUpdate = $obj.Props.I32 | Where-Object { $_.N -eq "StageUpdate" } | Select-Object -ExpandProperty '#text'
                StageRename = $obj.Props.I32 | Where-Object { $_.N -eq "StageRename" } | Select-Object -ExpandProperty '#text'
                StageDelete = $obj.Props.I32 | Where-Object { $_.N -eq "StageDelete" } | Select-Object -ExpandProperty '#text'
                StageDeleteAdd = $obj.Props.I32 | Where-Object { $_.N -eq "StageDeleteAdd" } | Select-Object -ExpandProperty '#text'
                StageFailure = $obj.Props.I32 | Where-Object { $_.N -eq "StageFailure" } | Select-Object -ExpandProperty '#text'
                ExportFailure = $obj.Props.I32 | Where-Object { $_.N -eq "ExportFailure" } | Select-Object -ExpandProperty '#text'
                ConnectorFlow = $obj.Props.I32 | Where-Object { $_.N -eq "ConnectorFlow" } | Select-Object -ExpandProperty '#text'
                FlowFailure = $obj.Props.I32 | Where-Object { $_.N -eq "FlowFailure" } | Select-Object -ExpandProperty '#text'
            }
            
            # Calculate duration
            if ($runStep.StartDate -and $runStep.EndDate) {
                $startDateTime = [DateTime]::Parse($runStep.StartDate)
                $endDateTime = [DateTime]::Parse($runStep.EndDate)
                $runStep | Add-Member -MemberType NoteProperty -Name DurationSeconds -Value ($endDateTime - $startDateTime).TotalSeconds
            }
            
            $runStepResults += $runStep
            
            # Extract sync errors from this run step
            $syncErrorsXml = $obj.Props.Obj | Where-Object { $_.N -eq "SyncErrors" }
            if ($syncErrorsXml) {
                $syncErrorsText = $syncErrorsXml.Props.S | Where-Object { $_.N -eq "SyncErrorsXml" } | Select-Object -ExpandProperty '#text'
                
                if ($syncErrorsText -and $syncErrorsText -ne "&lt;synchronization-errors&gt;&lt;/synchronization-errors&gt;") {
                    # Decode HTML entities with error handling
                    try {
                        # Load System.Web if not already loaded
                        Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
                        $decodedXml = [System.Web.HttpUtility]::HtmlDecode($syncErrorsText)
                    }
                    catch {
                        Write-Warning "Failed to decode HTML entities, using raw text: $($_.Exception.Message)"
                        $decodedXml = $syncErrorsText -replace '&lt;', '<' -replace '&gt;', '>' -replace '&amp;', '&'
                    }
                    
                    try {
                        [xml]$syncXml = $decodedXml
                        
                        foreach ($exportError in $syncXml.'synchronization-errors'.'export-error') {
                            $errorObj = [PSCustomObject]@{
                                RunHistoryId = $runStep.RunHistoryId
                                StepHistoryId = $runStep.StepHistoryId
                                StepNumber = $runStep.StepNumber
                                CsGuid = $exportError.'cs-guid'
                                DistinguishedName = $exportError.dn
                                DateOccurred = $exportError.'date-occurred'
                                FirstOccurred = $exportError.'first-occurred'
                                RetryCount = $exportError.'retry-count'
                                ErrorType = $exportError.'error-type'
                                ErrorCode = $exportError.'cd-error'.'error-code'
                                ErrorLiteral = $exportError.'cd-error'.'error-literal'
                                ServerErrorDetail = $exportError.'cd-error'.'server-error-detail'
                            }
                            
                            $exportErrors += $errorObj
                        }
                    }
                    catch {
                        Write-Warning "Failed to parse sync errors XML for step $($runStep.StepNumber): $($_.Exception.Message)"
                    }
                }
            }
        }
    }
    
    # Summary statistics
    $totalSteps = $runStepResults.Count
    $totalErrors = $exportErrors.Count
    $stepsWithErrors = ($runStepResults | Where-Object { $_.ExportFailure -gt 0 }).Count
    
    Write-Host "`nParsing Summary:" -ForegroundColor Green
    Write-Host "Total Run Steps: $totalSteps" -ForegroundColor White
    Write-Host "Steps with Errors: $stepsWithErrors" -ForegroundColor White
    Write-Host "Total Export Errors: $totalErrors" -ForegroundColor White
    
    # Export data for Power BI
    if ($ExportToCSV) {
        Write-Host "`nExporting to CSV files..." -ForegroundColor Yellow
        
        # Export run step results - always create the file
        $runStepCsvPath = Join-Path $OutputPath "RunStepResults.csv"
        if ($runStepResults.Count -eq 0) {
            # Create empty CSV with headers for Power BI compatibility
            $emptyRunStep = [PSCustomObject]@{
                RunHistoryId = $null
                StepHistoryId = $null
                StepNumber = $null
                StepResult = $null
                StartDate = $null
                EndDate = $null
                StageNoChange = $null
                StageAdd = $null
                StageUpdate = $null
                StageRename = $null
                StageDelete = $null
                StageDeleteAdd = $null
                StageFailure = $null
                ExportFailure = $null
                ConnectorFlow = $null
                FlowFailure = $null
                DurationSeconds = $null
            }
            @($emptyRunStep) | Export-Csv -Path $runStepCsvPath -NoTypeInformation -Encoding UTF8
        } else {
            $runStepResults | Export-Csv -Path $runStepCsvPath -NoTypeInformation -Encoding UTF8
        }
        Write-Host "Run Step Results exported to: $runStepCsvPath" -ForegroundColor Green
        
        # Export export errors - always create the file
        $errorsCsvPath = Join-Path $OutputPath "ExportErrors.csv"
        if ($exportErrors.Count -eq 0) {
            # Create empty CSV with headers for Power BI compatibility
            $emptyError = [PSCustomObject]@{
                RunHistoryId = $null
                StepHistoryId = $null
                StepNumber = $null
                CsGuid = $null
                DistinguishedName = $null
                DateOccurred = $null
                FirstOccurred = $null
                RetryCount = $null
                ErrorType = $null
                ErrorCode = $null
                ErrorLiteral = $null
                ServerErrorDetail = $null
            }
            @($emptyError) | Export-Csv -Path $errorsCsvPath -NoTypeInformation -Encoding UTF8
        } else {
            $exportErrors | Export-Csv -Path $errorsCsvPath -NoTypeInformation -Encoding UTF8
        }
        Write-Host "Export Errors exported to: $errorsCsvPath" -ForegroundColor Green
        
        # Create error summary - always create the file
        $summaryCsvPath = Join-Path $OutputPath "ErrorSummary.csv"
        if ($exportErrors.Count -gt 0) {
            $errorSummary = $exportErrors | Group-Object ErrorType | Select-Object @{
                Name = 'ErrorType'
                Expression = { $_.Name }
            }, @{
                Name = 'Count'
                Expression = { $_.Count }
            }, @{
                Name = 'Percentage'
                Expression = { [math]::Round(($_.Count / $totalErrors) * 100, 2) }
            }
            $errorSummary | Export-Csv -Path $summaryCsvPath -NoTypeInformation -Encoding UTF8
        } else {
            # Create empty error summary CSV with headers
            $emptyErrorSummary = [PSCustomObject]@{
                ErrorType = $null
                Count = $null
                Percentage = $null
            }
            @($emptyErrorSummary) | Export-Csv -Path $summaryCsvPath -NoTypeInformation -Encoding UTF8
        }
        Write-Host "Error Summary exported to: $summaryCsvPath" -ForegroundColor Green
    }
    
    if ($ExportToJSON) {
        Write-Host "`nExporting to JSON files..." -ForegroundColor Yellow
        
        # Create comprehensive data structure for Power BI
        $powerBIData = @{
            Metadata = @{
                ExportDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
                SourceFile = Split-Path $XmlFilePath -Leaf
                TotalSteps = $totalSteps
                TotalErrors = $totalErrors
                StepsWithErrors = $stepsWithErrors
            }
            RunStepResults = $runStepResults
            ExportErrors = $exportErrors
            ErrorSummary = $errorSummary
        }
        
        $jsonPath = Join-Path $OutputPath "AADConnect_RunStep_Data.json"
        $powerBIData | ConvertTo-Json -Depth 10 | Out-File -FilePath $jsonPath -Encoding UTF8
        Write-Host "Complete dataset exported to: $jsonPath" -ForegroundColor Green
    }
    
    Write-Host "`nData parsing completed successfully!" -ForegroundColor Green
    Write-Host "`nFor Power BI import:" -ForegroundColor Cyan
    Write-Host "1. Use the CSV files for individual table imports" -ForegroundColor White
    Write-Host "2. Use the JSON file for hierarchical data relationships" -ForegroundColor White
    Write-Host "3. Key metrics: RunStepResults table contains step-level statistics" -ForegroundColor White
    Write-Host "4. Error details: ExportErrors table contains detailed error information" -ForegroundColor White
    
    # Return the parsed data for further processing if needed
    return @{
        RunStepResults = $runStepResults
        ExportErrors = $exportErrors
        ErrorSummary = $errorSummary
    }
}
catch {
    Write-Error "Failed to parse XML file: $($_.Exception.Message)"
    Write-Error "Stack trace: $($_.Exception.StackTrace)"
    throw
}
