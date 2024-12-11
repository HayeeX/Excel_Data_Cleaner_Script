# Excel Data Cleaner Script

## Overview
**Excel_Data_Cleaner_Script** is a VBA-powered tool designed to automate common data cleaning tasks in Excel. This utility is ideal for anyone working with messy datasets, helping to save time and improve data quality.

The script performs the following cleaning operations:
- Removes duplicate rows.
- Handles missing values by replacing blank cells with a placeholder (e.g., `"N/A"`).
- Standardizes text (e.g., trims whitespace and converts text to uppercase).
- Identifies and removes outliers in numerical data.
- Ensures consistent date formatting.

## Features
1. **Automated Duplicate Removal**: Quickly removes duplicate rows based on specific columns.
2. **Missing Value Handling**: Replaces blank cells with a user-friendly placeholder.
3. **Text Standardization**: Ensures consistent formatting by trimming spaces and applying uppercase transformation.
4. **Outlier Detection**: Identifies and removes outliers in numerical data using statistical thresholds.
5. **Date Formatting**: Standardizes date values in the dataset.
6. **User Feedback**: Notifies users upon successful completion of the cleaning process.

## Installation
1. Download the `Excel_Data_Cleaner_Script.xlsm` file from this repository.
2. Open the file in Microsoft Excel.
3. Enable macros by clicking **Enable Content** when prompted.

## Usage Instructions
### Running the Script
1. Open your dataset in the Excel workbook.
2. Navigate to the worksheet containing the data you want to clean.
3. Go to the **Developer** tab in Excel and click **Macros**.
4. Select `DataCleaningUtility` from the list of macros and click **Run**.
5. The script will clean your data and notify you when it’s complete.

### Assigning the Script to a Button (Optional)
1. Go to the **Developer** tab and click **Insert**.
2. Select a **Button (Form Control)** and draw it on the worksheet.
3. Right-click the button, select **Assign Macro**, and choose `DataCleaningUtility`.
4. Click **OK** and use the button to run the script with one click.

## How It Works
### Steps Performed by the Script
1. Identifies the dataset range dynamically based on the last row and column.
2. Cleans the data through the following operations:
   - Removes duplicate rows based on the first two columns.
   - Replaces blank cells with `"N/A"`.
   - Trims whitespace and converts text in column A to uppercase.
   - Calculates average and standard deviation for column B and removes outliers.
   - Formats column C as `"mm/dd/yyyy"` for consistent date representation.
3. Displays a message box upon successful completion.

## Customization
- **Target Columns**: Modify the script to clean specific columns by changing the column indices in the code.
- **Placeholder for Missing Values**: Replace `"N/A"` with a custom placeholder of your choice.
- **Outlier Thresholds**: Adjust the outlier detection logic by changing the multiplier (currently set to `2 * stdDev`).

## Requirements
- Microsoft Excel (2016 or later recommended)
- Enabled macros in Excel

## Contributing
Contributions are welcome! If you have ideas for additional features or enhancements, feel free to submit a pull request or open an issue.

## License
This project is licensed under the Haye_Tech License. See the LICENSE file for details.

## Contact
For questions, suggestions, or support, please contact Hayelom at haye.officiall@gmail.com.

