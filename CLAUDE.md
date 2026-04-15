# Financial Modeling Project

## Overview
Fractional CFO financial modeling work. This project is separate from the fraud detection work in `/Users/chandlerclemons/dme-fraud/`.

## Conventions
- All financial models should use Python (pandas, numpy, openpyxl) unless the user specifies otherwise
- Currency values should be formatted with commas and 2 decimal places (e.g., $1,234,567.89)
- Percentages should be displayed to 1 decimal place unless more precision is needed
- Use consistent variable naming: snake_case for Python, descriptive names for spreadsheet tabs
- When building projections, always clearly label assumptions vs. calculated values
- Include sensitivity analysis where appropriate

## Output Formats
- Excel (.xlsx) is the default deliverable format for client-facing models
- Python scripts for data processing and model logic
- Markdown summaries for internal documentation

## File Organization
```
/financial-modeling/
  /models/        - Excel and Python model files
  /data/          - Input data files
  /outputs/       - Generated reports and exports
```

## Best Practices
- Always separate inputs/assumptions from calculations
- Build models that flow top-to-bottom, left-to-right
- Include a summary/dashboard tab in Excel workbooks
- Document key assumptions inline
- Use named ranges or clear cell references
- Error-check formulas with reasonableness tests
