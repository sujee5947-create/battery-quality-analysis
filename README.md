# Battery Quality Data Analysis

**Python | Battery Quality | Data Analysis | Report Automation**

## Overview

This project explores the use of Python to analyze synthetic lithium-ion battery cell measurement data and automate report generation.

Building on my experience in battery cell development and quality assurance, I created a small-scale workflow to compare OK and NG samples, examine relationships between measurement parameters, and visualize analysis results.

## Key Features

- Load and process synthetic battery measurement data.
- Compare measurement characteristics between OK and NG samples.
- Analyze correlations between measurement parameters.
- Visualize data distributions and analysis results.
- Calculate process capability indices (Cp/Cpk), where implemented.
- Generate HTML and PowerPoint reports using Python.

## Data Parameters

- Capacity (Ah)
- Voltage (V)
- Internal Resistance (mΩ)
- Thickness (mm)
- Weight (g)

## Tools

- Python
- Pandas
- NumPy, where applicable
- Data visualization libraries used in the implementation
- HTML and PowerPoint report generation

## How to Run

1. Clone or download this repository.
2. Ensure the required input files are available in the expected locations.
3. Install the Python packages required by the application.
4. Run the script from the project directory:

   ```bash
   python battery_quality_app_updated.py
   ```

5. Check the output directory for the generated reports.

**Note:** Execution steps may vary depending on the local environment, package dependencies, and input file paths.

## Data Disclaimer

All datasets used in this project are synthetic data created for learning and demonstration purposes. They do not contain actual manufacturing data or confidential company information.

## Project Scope

This is a personal learning and demonstration project focused on applying Python to battery quality data analysis and report automation.

The current implementation is a prototype and is not intended for production quality monitoring or manufacturing decisions.

## Future Improvements

- Expand data validation and error handling.
- Support datasets of different sizes.
- Explore SPC control charts and additional process capability analysis.
- Improve report visualization and usability.

## Author

Battery development and quality professional with experience in lithium-ion cell development, quality assurance, and root-cause analysis.
