# Enhanced PowerShell functions for AAD Connect XML parsing and Power BI preparation
# Additional utility functions for data analysis and visualization preparation

# Function to parse connection information
function Get-ConnectionInfo {
    param([xml]$XmlContent)
    
    $connectionInfo = @()
    
    foreach ($obj in $XmlContent.Objs.Obj) {
        if ($obj.TN.T -contains "Microsoft.IdentityManagement.PowerShell.ObjectModel.RunStepResult") {
            $connInfoXml = $obj.Props.S | Where-Object { $_.N -eq "ConnectorConnectionInformationXml" } | Select-Object -ExpandProperty '#text'
            
            if ($connInfoXml) {
                try {
                    $decodedXml = [System.Web.HttpUtility]::HtmlDecode($connInfoXml)
                    [xml]$connXml = $decodedXml
                    
                    $runHistoryId = $obj.Props.G | Where-Object { $_.N -eq "RunHistoryId" } | Select-Object -ExpandProperty '#text'
                    $stepNumber = $obj.Props.I32 | Where-Object { $_.N -eq "StepNumber" } | Select-Object -ExpandProperty '#text'
                    
                    $connObj = [PSCustomObject]@{
                        RunHistoryId = $runHistoryId
                        StepNumber = $stepNumber
                        ConnectionResult = $connXml.'connection-result'
                        Server = $connXml.server
                        ConnectionDate = $connXml.'connection-log'.incident.date
                        ConnectionServer = $connXml.'connection-log'.incident.server
                    }
                    
                    $connectionInfo += $connObj
                }
                catch {
                    Write-Warning "Failed to parse connection info for step $stepNumber"
                }
            }
        }
    }
    
    return $connectionInfo
}

# Function to create Power BI friendly date/time measures
function Add-DateTimeMeasures {
    param([array]$Data)
    
    foreach ($item in $Data) {
        if ($item.StartDate) {
            try {
                $startDateTime = [DateTime]::Parse($item.StartDate)
                $item | Add-Member -MemberType NoteProperty -Name StartYear -Value $startDateTime.Year -Force
                $item | Add-Member -MemberType NoteProperty -Name StartMonth -Value $startDateTime.Month -Force
                $item | Add-Member -MemberType NoteProperty -Name StartDay -Value $startDateTime.Day -Force
                $item | Add-Member -MemberType NoteProperty -Name StartHour -Value $startDateTime.Hour -Force
                $item | Add-Member -MemberType NoteProperty -Name StartDayOfWeek -Value $startDateTime.DayOfWeek.ToString() -Force
                $item | Add-Member -MemberType NoteProperty -Name StartWeekOfYear -Value (Get-Date $startDateTime -UFormat %V) -Force
            }
            catch {
                Write-Warning "Failed to parse StartDate for item: $($item.RunHistoryId)"
            }
        }
        
        if ($item.EndDate) {
            try {
                $endDateTime = [DateTime]::Parse($item.EndDate)
                $item | Add-Member -MemberType NoteProperty -Name EndYear -Value $endDateTime.Year -Force
                $item | Add-Member -MemberType NoteProperty -Name EndMonth -Value $endDateTime.Month -Force
                $item | Add-Member -MemberType NoteProperty -Name EndDay -Value $endDateTime.Day -Force
                $item | Add-Member -MemberType NoteProperty -Name EndHour -Value $endDateTime.Hour -Force
            }
            catch {
                Write-Warning "Failed to parse EndDate for item: $($item.RunHistoryId)"
            }
        }
        
        if ($item.DateOccurred) {
            try {
                $occurredDateTime = [DateTime]::Parse($item.DateOccurred)
                $item | Add-Member -MemberType NoteProperty -Name OccurredYear -Value $occurredDateTime.Year -Force
                $item | Add-Member -MemberType NoteProperty -Name OccurredMonth -Value $occurredDateTime.Month -Force
                $item | Add-Member -MemberType NoteProperty -Name OccurredDay -Value $occurredDateTime.Day -Force
                $item | Add-Member -MemberType NoteProperty -Name OccurredHour -Value $occurredDateTime.Hour -Force
                $item | Add-Member -MemberType NoteProperty -Name OccurredDayOfWeek -Value $occurredDateTime.DayOfWeek.ToString() -Force
            }
            catch {
                Write-Warning "Failed to parse DateOccurred for item: $($item.CsGuid)"
            }
        }
        
        if ($item.FirstOccurred) {
            try {
                $firstDateTime = [DateTime]::Parse($item.FirstOccurred)
                $item | Add-Member -MemberType NoteProperty -Name FirstOccurredYear -Value $firstDateTime.Year -Force
                $item | Add-Member -MemberType NoteProperty -Name FirstOccurredMonth -Value $firstDateTime.Month -Force
                $item | Add-Member -MemberType NoteProperty -Name FirstOccurredDay -Value $firstDateTime.Day -Force
            }
            catch {
                Write-Warning "Failed to parse FirstOccurred for item: $($item.CsGuid)"
            }
        }
    }
    
    return $Data
}

# Function to create aggregated metrics for Power BI dashboards
function Get-AggregatedMetrics {
    param(
        [array]$RunStepResults,
        [array]$ExportErrors
    )
    
    $metrics = @{
        # Step-level metrics
        TotalSteps = $RunStepResults.Count
        SuccessfulSteps = ($RunStepResults | Where-Object { $_.StepResult -eq "completed-success" }).Count
        StepsWithErrors = ($RunStepResults | Where-Object { $_.ExportFailure -gt 0 }).Count
        
        # Error metrics
        TotalErrors = $ExportErrors.Count
        UniqueErrorTypes = ($ExportErrors | Select-Object -Unique ErrorType).Count
        MostCommonError = ($ExportErrors | Group-Object ErrorType | Sort-Object Count -Descending | Select-Object -First 1).Name
        
        # Performance metrics
        AverageStepDuration = ($RunStepResults | Where-Object { $_.DurationSeconds } | Measure-Object -Property DurationSeconds -Average).Average
        MaxStepDuration = ($RunStepResults | Where-Object { $_.DurationSeconds } | Measure-Object -Property DurationSeconds -Maximum).Maximum
        MinStepDuration = ($RunStepResults | Where-Object { $_.DurationSeconds } | Measure-Object -Property DurationSeconds -Minimum).Minimum
        
        # Operation counts
        TotalAdds = ($RunStepResults | Measure-Object -Property StageAdd -Sum).Sum
        TotalUpdates = ($RunStepResults | Measure-Object -Property StageUpdate -Sum).Sum
        TotalDeletes = ($RunStepResults | Measure-Object -Property StageDelete -Sum).Sum
        TotalNoChanges = ($RunStepResults | Measure-Object -Property StageNoChange -Sum).Sum
    }
    
    return $metrics
}

# Function to extract user information from Distinguished Names
function Parse-DistinguishedNames {
    param([array]$ExportErrors)
    
    foreach ($error in $ExportErrors) {
        if ($error.DistinguishedName) {
            # Extract CN (Common Name)
            if ($error.DistinguishedName -match "CN=([^,]+)") {
                $error | Add-Member -MemberType NoteProperty -Name UserName -Value $matches[1] -Force
            }
            
            # Extract OU (Organizational Unit)
            if ($error.DistinguishedName -match "OU=([^,]+)") {
                $error | Add-Member -MemberType NoteProperty -Name OrganizationalUnit -Value $matches[1] -Force
            }
            
            # Extract DC (Domain Component) - usually the domain
            $dcMatches = [regex]::Matches($error.DistinguishedName, "DC=([^,]+)")
            if ($dcMatches.Count -gt 0) {
                $domain = ($dcMatches | ForEach-Object { $_.Groups[1].Value }) -join "."
                $error | Add-Member -MemberType NoteProperty -Name Domain -Value $domain -Force
            }
            
            # Classify user type based on naming patterns
            $userName = $error.UserName
            if ($userName) {
                $userType = switch -Regex ($userName) {
                    "^_.*" { "Service Account" }
                    ".*T[0-9]$" { "Test Account" }
                    ".*\s.*" { "Regular User" }
                    default { "Other" }
                }
                $error | Add-Member -MemberType NoteProperty -Name UserType -Value $userType -Force
            }
        }
    }
    
    return $ExportErrors
}

# Function to ensure all required CSV files exist with proper headers for Power BI
function Initialize-PowerBIDataModel {
    param([string]$OutputPath)
    
    Write-Host "Ensuring all Power BI data model files exist..." -ForegroundColor Yellow
    
    # Define all required CSV files with their expected headers
    $csvTemplates = @{
        "RunStepResults.csv" = [PSCustomObject]@{
            RunHistoryId = $null; StepHistoryId = $null; StepNumber = $null; StepResult = $null
            StartDate = $null; EndDate = $null; StageNoChange = $null; StageAdd = $null
            StageUpdate = $null; StageRename = $null; StageDelete = $null; StageDeleteAdd = $null
            StageFailure = $null; ExportFailure = $null; ConnectorFlow = $null; FlowFailure = $null
            DurationSeconds = $null; StartYear = $null; StartMonth = $null; StartDay = $null
            StartHour = $null; StartDayOfWeek = $null; StartWeekOfYear = $null
            EndYear = $null; EndMonth = $null; EndDay = $null; EndHour = $null
        }
        
        "ExportErrors.csv" = [PSCustomObject]@{
            RunHistoryId = $null; StepHistoryId = $null; StepNumber = $null; CsGuid = $null
            DistinguishedName = $null; DateOccurred = $null; FirstOccurred = $null; RetryCount = $null
            ErrorType = $null; ErrorCode = $null; ErrorLiteral = $null; ServerErrorDetail = $null
        }
        
        "EnhancedExportErrors.csv" = [PSCustomObject]@{
            RunHistoryId = $null; StepHistoryId = $null; StepNumber = $null; CsGuid = $null
            DistinguishedName = $null; DateOccurred = $null; FirstOccurred = $null; RetryCount = $null
            ErrorType = $null; ErrorCode = $null; ErrorLiteral = $null; ServerErrorDetail = $null
            OccurredYear = $null; OccurredMonth = $null; OccurredDay = $null; OccurredHour = $null
            OccurredDayOfWeek = $null; FirstOccurredYear = $null; FirstOccurredMonth = $null; FirstOccurredDay = $null
            UserName = $null; OrganizationalUnit = $null; Domain = $null; UserType = $null
        }
        
        "ErrorSummary.csv" = [PSCustomObject]@{
            ErrorType = $null; Count = $null; Percentage = $null
        }
        
        "ConnectionInfo.csv" = [PSCustomObject]@{
            RunHistoryId = $null; StepNumber = $null; ConnectionResult = $null
            Server = $null; ConnectionDate = $null; ConnectionServer = $null
        }
        
        "AggregatedMetrics.csv" = [PSCustomObject]@{
            Metric = "NoData"; Value = 0
        }
    }
    
    # Create any missing CSV files
    foreach ($csvFile in $csvTemplates.Keys) {
        $csvPath = Join-Path $OutputPath $csvFile
        if (-not (Test-Path $csvPath)) {
            Write-Host "Creating empty data model file: $csvFile" -ForegroundColor Gray
            @($csvTemplates[$csvFile]) | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        }
    }
    
    Write-Host "Power BI data model initialization complete." -ForegroundColor Green
}

# Main execution function that combines all parsing operations
function Invoke-CompleteXMLParsing {
    param(
        [string]$XmlFilePath,
        [string]$OutputPath,
        [switch]$IncludeAdvancedMetrics
    )
    
    Write-Host "Starting comprehensive XML parsing..." -ForegroundColor Green
    
    # Initialize Power BI data model files first
    Initialize-PowerBIDataModel -OutputPath $OutputPath
    
    # Load and parse basic data - use relative path
    $parseScriptPath = Join-Path $PSScriptRoot "Parse-RunStepXML.ps1"
    $basicData = & $parseScriptPath -XmlFilePath $XmlFilePath -OutputPath $OutputPath
    
    if ($IncludeAdvancedMetrics) {
        Write-Host "Adding advanced metrics and analysis..." -ForegroundColor Yellow
        
        # Load XML again for additional parsing
        [xml]$xmlContent = Get-Content -Path $XmlFilePath -Raw
        
        # Get connection information
        $connectionInfo = Get-ConnectionInfo -XmlContent $xmlContent
        
        # Enhance data with date/time measures
        $enhancedRunSteps = Add-DateTimeMeasures -Data $basicData.RunStepResults
        $enhancedErrors = Add-DateTimeMeasures -Data $basicData.ExportErrors
        
        # Parse distinguished names for user classification
        $enhancedErrors = Parse-DistinguishedNames -ExportErrors $enhancedErrors
        
        # Generate aggregated metrics
        $aggregatedMetrics = Get-AggregatedMetrics -RunStepResults $enhancedRunSteps -ExportErrors $enhancedErrors
        
        # Export enhanced data - ensure CSV files are created even with no data
        $connectionCsvPath = Join-Path $OutputPath "ConnectionInfo.csv"
        if ($connectionInfo.Count -eq 0) {
            # Create empty CSV with headers for Power BI compatibility
            $emptyConnection = [PSCustomObject]@{
                RunHistoryId = $null
                StepNumber = $null
                ConnectionResult = $null
                Server = $null
                ConnectionDate = $null
                ConnectionServer = $null
            }
            @($emptyConnection) | Export-Csv -Path $connectionCsvPath -NoTypeInformation -Encoding UTF8
        } else {
            $connectionInfo | Export-Csv -Path $connectionCsvPath -NoTypeInformation -Encoding UTF8
        }
        
        $metricsCsvPath = Join-Path $OutputPath "AggregatedMetrics.csv"
        if ($aggregatedMetrics.Count -eq 0) {
            # Create empty metrics CSV with headers
            $emptyMetrics = [PSCustomObject]@{
                Metric = "NoData"
                Value = 0
            }
            @($emptyMetrics) | Export-Csv -Path $metricsCsvPath -NoTypeInformation -Encoding UTF8
        } else {
            $aggregatedMetrics.GetEnumerator() | Select-Object @{Name='Metric';Expression={$_.Key}}, @{Name='Value';Expression={$_.Value}} | 
                Export-Csv -Path $metricsCsvPath -NoTypeInformation -Encoding UTF8
        }
        
        # Enhanced errors with user classification
        $enhancedErrorsCsvPath = Join-Path $OutputPath "EnhancedExportErrors.csv"
        if ($enhancedErrors.Count -eq 0) {
            # Create empty enhanced errors CSV with all expected headers
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
                OccurredYear = $null
                OccurredMonth = $null
                OccurredDay = $null
                OccurredHour = $null
                OccurredDayOfWeek = $null
                FirstOccurredYear = $null
                FirstOccurredMonth = $null
                FirstOccurredDay = $null
                UserName = $null
                OrganizationalUnit = $null
                Domain = $null
                UserType = $null
            }
            @($emptyError) | Export-Csv -Path $enhancedErrorsCsvPath -NoTypeInformation -Encoding UTF8
        } else {
            $enhancedErrors | Export-Csv -Path $enhancedErrorsCsvPath -NoTypeInformation -Encoding UTF8
        }
        
        Write-Host "Enhanced analysis files created:" -ForegroundColor Green
        Write-Host "- Connection Info: $connectionCsvPath" -ForegroundColor White
        Write-Host "- Aggregated Metrics: $metricsCsvPath" -ForegroundColor White
        Write-Host "- Enhanced Errors: $enhancedErrorsCsvPath" -ForegroundColor White
    }
    
    Write-Host "`nPower BI Import Instructions:" -ForegroundColor Cyan
    Write-Host "1. Import CSV files into separate tables" -ForegroundColor White
    Write-Host "2. Create relationships using RunHistoryId and StepHistoryId" -ForegroundColor White
    Write-Host "3. Use the date/time columns for time-based analysis" -ForegroundColor White
    Write-Host "4. Create measures using the aggregated metrics" -ForegroundColor White
    Write-Host "5. Use UserType and ErrorType for categorical analysis" -ForegroundColor White
}

# Export functions for module use
Export-ModuleMember -Function Get-ConnectionInfo, Add-DateTimeMeasures, Get-AggregatedMetrics, Parse-DistinguishedNames, Invoke-CompleteXMLParsing, Initialize-PowerBIDataModel, Initialize-PowerBIDataModel
