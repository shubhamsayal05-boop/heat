# WL EVO AVLDrive Heatmap Tool - Evaluation Data Analyzer

## Overview

This repository contains the **WL EVO AVLDrive Heatmap Tool** Excel workbook and a Python-based analyzer tool that processes and analyzes evaluation metric data pasted by users into the Evaluation Sheet.

## What Does This Tool Do?

The **Evaluation Data Analyzer** (`analyze_evaluation_data.py`) reads the Excel file and provides:

- **Statistical Analysis**: Distribution of P1 status ratings (Red/Yellow/Green) for Drivability and Dynamism
- **Data Quality Checks**: Completeness metrics and identification of missing data
- **Insights Generation**: Automated insights about performance and areas needing attention
- **Export Capability**: JSON export of analysis results for integration with other tools

## Project Structure

```
heat/
├── WL_EVO_AVLDrive_Heatmap_Tool.xlsm    # Main Excel workbook with VBA macros
├── analyze_evaluation_data.py            # Python analyzer tool
├── requirements.txt                      # Python dependencies
└── README.md                            # This file
```

## Excel Workbook Sheets

The Excel workbook contains several sheets:

1. **HeatMap Sheet**: Template for heatmap generation
2. **Mapping Sheet**: Data mapping configuration
3. **HeatMap Template**: Base template
4. **AVL-Odriv Mapping**: AVL to Odriv mapping
5. **Data Transfer Sheet**: Intermediate data storage
6. **Evaluation Sheet**: **Main sheet where users paste evaluation data**

## Evaluation Sheet Structure

The Evaluation Sheet contains:

- **USE CASE** (Column A/C): List of test scenarios (e.g., "Drive Away Creep Eng On", "DASS Eng On")
- **Drivability Section**: P1, P2, P3 status under "Current Status" and "SOPM Prediction"
- **Dynamism Section**: P1, P2, P3 status under "Current Status" and "SOPM Prediction"
- **Vehicle Data Columns**: User-pasted data for different vehicle configurations
- **Output Columns**: Drivability, Responsiveness, and Overall ratings

### P1 Status Ratings

The P1 columns use color-coded dots (●) to indicate status:
- 🔴 **Red**: Critical issues - immediate attention required
- 🟡 **Yellow**: Requires improvement
- 🟢 **Green**: Performing well

## Installation

### Prerequisites

- Python 3.7 or higher
- pip (Python package installer)

### Setup

1. Clone or download this repository:
```bash
git clone https://github.com/shubhamsayal05-boop/heat.git
cd heat
```

2. Install Python dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Analysis

Run the analyzer on the Excel file:

```bash
python3 analyze_evaluation_data.py WL_EVO_AVLDrive_Heatmap_Tool.xlsm
```

This will:
1. Load the Excel file
2. Find and parse the Evaluation Sheet
3. Extract all use case data
4. Analyze P1 status distributions
5. Generate data quality metrics
6. Display insights in the terminal

### Export to JSON

To export the analysis results to a JSON file:

```bash
python3 analyze_evaluation_data.py WL_EVO_AVLDrive_Heatmap_Tool.xlsm -o analysis_results.json
```

### Example Output

```
================================================================================
EVALUATION DATA ANALYSIS REPORT
================================================================================

📋 Total Use Cases Analyzed: 61

--------------------------------------------------------------------------------
DRIVABILITY P1 STATUS DISTRIBUTION
--------------------------------------------------------------------------------
  Green          :   45 ( 73.8%)
  Yellow         :   12 ( 19.7%)
  Red            :    4 (  6.6%)

--------------------------------------------------------------------------------
DYNAMISM P1 STATUS DISTRIBUTION
--------------------------------------------------------------------------------
  Green          :   48 ( 78.7%)
  Yellow         :   10 ( 16.4%)
  Red            :    3 (  4.9%)

--------------------------------------------------------------------------------
KEY INSIGHTS
--------------------------------------------------------------------------------
  ⚠️  4 use cases (6.6%) have RED drivability status - these need immediate attention
  ⚡ 12 use cases (19.7%) have YELLOW drivability status - these require improvement
  ✅ 45 use cases (73.8%) have GREEN drivability status - performing well
  🎯 Excellent overall performance! 80.3% of use cases are GREEN
```

## How Users Work With This Tool

### For Excel Users (Manual Process)

1. Open `WL_EVO_AVLDrive_Heatmap_Tool.xlsm` in Microsoft Excel
2. Navigate to the **Evaluation Sheet**
3. Paste test data from AVL-Drive or other testing tools into the vehicle columns (starting around column L)
4. The VBA macros process the data and update:
   - P1 status dots based on cell colors
   - Output columns (Drivability, Responsiveness, Overall)
5. Use the "Update Outcomes" button to run the VBA analysis

### For Analysts (Using Python Tool)

1. After users paste data into the Excel file and save it
2. Run the Python analyzer to get statistical insights:
   ```bash
   python3 analyze_evaluation_data.py WL_EVO_AVLDrive_Heatmap_Tool.xlsm
   ```
3. Review the analysis report to identify:
   - Areas with critical issues (Red status)
   - Areas needing improvement (Yellow status)
   - Well-performing areas (Green status)
   - Data completeness issues

## VBA Macros in the Workbook

The Excel workbook includes several VBA modules:

- **HeatMap.bas**: Main heatmap generation logic
- **Export.bas**: Export functionality for heatmap data
- **Evaluation_AutoDetect_P1.bas**: Automatic P1 status detection and outcome calculation

### Key VBA Functions

- `UpdateOutcomesAuto()`: Automatically detects P1 columns and updates outcome columns
- `CalibrateP1Colors()`: Allows calibration of P1 color detection
- `ExportSelectionVisibleOnly_AsPicture()`: Exports visible data as picture

## Data Quality Considerations

When analyzing evaluation data, the tool checks for:

- **Completeness**: Percentage of rows with valid P1 status data
- **Missing Data**: Number of rows without drivability/dynamism ratings
- **Consistency**: Verification that output columns match P1 inputs

## Troubleshooting

### Issue: "Could not find Evaluation Sheet"
- Ensure the Excel file contains a sheet with "Evaluation" in its name
- Check that the file is not corrupted

### Issue: "No data rows found"
- Verify that the sheet has a "USE CASE" header row
- Ensure there is data below the header row
- Check that data is in the expected format

### Issue: "All status showing as Unspecified"
- The P1 cells may not have color formatting applied
- Users need to paste data with proper cell colors or text values (Red/Yellow/Green)
- Run the VBA "Update Outcomes" macro to process the data first

## Contributing

To contribute to this project:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

[Add license information here]

## Contact

For issues, questions, or contributions, please contact the repository maintainer or open an issue on GitHub.

## Changelog

### Version 1.0.0 (Current)
- Initial release of Python evaluation data analyzer
- Automatic detection of Evaluation Sheet structure
- P1 status distribution analysis
- Data quality metrics
- Insight generation
- JSON export capability
