# Power BI Import Guide for AAD Connect XML Data

## Overview
This guide walks you through importing the parsed AAD Connect XML data into Power BI Desktop to create insightful dashboards and reports.

## Prerequisites
- Power BI Desktop (latest version recommended)
- Parsed CSV files from the PowerShell script

## Step 1: Data Import

### 1.1 Open Power BI Desktop
- Launch Power BI Desktop
- Select "Get data" from the Home ribbon

### 1.2 Import CSV Files
Import the following files as separate tables:

**Required Tables:**
1. **RunStepResults.csv** - Main execution summary
2. **ExportErrors.csv** - Detailed error information  
3. **EnhancedExportErrors.csv** - Errors with user classification
4. **ErrorSummary.csv** - Aggregated error statistics
5. **AggregatedMetrics.csv** - Key performance indicators
6. **ConnectionInfo.csv** - Connection status data

**Import Process:**
- Home → Get Data → Text/CSV
- Navigate to each CSV file
- Click "Load" (or "Transform Data" if cleaning needed)
- Repeat for all files

## Step 2: Data Model Setup

### 2.1 Create Relationships
In Model view, create these relationships:

```
RunStepResults[RunHistoryId] ←→ ExportErrors[RunHistoryId]
RunStepResults[StepHistoryId] ←→ ExportErrors[StepHistoryId]
```

### 2.2 Set Data Types
Verify these column data types:

**RunStepResults Table:**
- StartDate, EndDate: Date/Time
- DurationSeconds: Decimal Number
- All count fields (StageAdd, StageUpdate, etc.): Whole Number

**ExportErrors Table:**
- DateOccurred, FirstOccurred: Date/Time
- RetryCount: Whole Number

### 2.3 Create Calculated Columns
Add these calculated columns if needed:

**In RunStepResults:**
```DAX
Duration Minutes = RunStepResults[DurationSeconds] / 60
Has Errors = IF(RunStepResults[ExportFailure] > 0, "Yes", "No")
```

**In ExportErrors:**
```DAX
Days Since First Error = DATEDIFF(ExportErrors[FirstOccurred], ExportErrors[DateOccurred], DAY)
```

## Step 3: Create Measures

Copy the measures from `PowerBI-DAX-Measures.txt` into your Power BI model:

### Essential Measures:
1. Total Export Errors
2. Error Rate %
3. Average Step Duration (Seconds)
4. Users Affected
5. Success Rate %

### How to Add Measures:
1. Right-click on any table in Fields pane
2. Select "New measure"
3. Paste the DAX formula
4. Click checkmark to save

## Step 4: Build Visualizations

### 4.1 Executive Summary Page

**Key Metrics Cards:**
- Total Export Errors
- Success Rate %
- Users Affected
- Average Step Duration

**Trend Charts:**
- Line chart: Errors over time (DateOccurred vs Total Export Errors)
- Column chart: Error types distribution

### 4.2 Error Analysis Page

**Error Details Table:**
Columns: UserName, ErrorType, RetryCount, DateOccurred, FirstOccurred
- Apply conditional formatting on RetryCount
- Add filters for ErrorType and UserType

**User Classification:**
- Pie chart: UserType distribution
- Bar chart: Errors by Domain

### 4.3 Performance Dashboard

**Step Performance:**
- Gauge: Success Rate %
- Bar chart: Duration by StepNumber
- Scatter plot: Duration vs Error Count

**Connection Analysis:**
- Table: ConnectionInfo with status indicators
- Card: Connection success rate

## Step 5: Advanced Features

### 5.1 Slicers and Filters
Add these slicers for interactive filtering:
- Date range (DateOccurred)
- ErrorType
- UserType
- Domain

### 5.2 Drill-Through Pages
Create drill-through from summary to details:
- Source: Error type in summary charts
- Target: Detailed error page with user-level data

### 5.3 Bookmarks
Create bookmarks for different views:
- "All Errors" - Shows all error data
- "Recent Errors" - Last 30 days only
- "Service Accounts" - Service account errors only
- "High Priority" - Errors with high retry counts

## Step 6: Formatting and Themes

### 6.1 Color Scheme
- **Success/Good**: Green (#00B04F)
- **Errors/Bad**: Red (#FF5F5F)
- **Warning**: Orange (#FFA500)
- **Information**: Blue (#0078D4)

### 6.2 Conditional Formatting
Apply to relevant columns:

**RetryCount:**
- \> 50,000: Red background
- \> 20,000: Orange background
- < 1,000: Green background

**UserType:**
- Service Account: Gray background
- Regular User: Default
- Test Account: Light blue background

## Step 7: Publishing and Sharing

### 7.1 Save and Publish
- File → Save As → Choose location
- File → Publish → Select workspace

### 7.2 Refresh Schedule
If data updates regularly:
- Set up scheduled refresh in Power BI Service
- Configure data source credentials

## Sample Dashboard Layout

```
┌─────────────────────────────────────────────────────────┐
│                Executive Summary                         │
├─────────────┬─────────────┬─────────────┬──────────────┤
│Total Errors │Success Rate │Users Affected│Avg Duration  │
│    [23]     │   [95.7%]   │    [15]     │  [0.127s]    │
├─────────────┴─────────────┴─────────────┴──────────────┤
│                                                         │
│  📈 Error Trend Over Time                              │
│                                                         │
├─────────────────────────────┬───────────────────────────┤
│                             │                           │
│  📊 Error Types             │  👥 User Types           │
│  Distribution               │  Affected                 │
│                             │                           │
└─────────────────────────────┴───────────────────────────┘
```

## Troubleshooting

### Common Issues:
1. **Date parsing errors**: Check regional date format settings
2. **Relationship errors**: Verify key column data types match
3. **Slow performance**: Consider aggregating large datasets

### Performance Tips:
- Use DirectQuery for very large datasets
- Create aggregation tables for better performance
- Remove unused columns from data model

## Next Steps
- Set up automated data refresh
- Create alerts for error thresholds
- Build additional drill-down reports
- Export insights to SharePoint or Teams

For additional help, refer to the Microsoft Power BI documentation or contact your BI administrator.
