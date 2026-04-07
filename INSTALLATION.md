# Installation Instructions for Primavera P6 Excel Export

This document outlines the step-by-step installation and setup instructions for using the VBA and Power Query functionality within the Primavera P6 Excel Export project.

## Prerequisites
- Microsoft Excel (2016 or later recommended)
- Primavera P6 installed and configured

## Setup Instructions for VBA

1. **Open Excel**: Launch Microsoft Excel.
2. **Enable Developer Tab**:
   - Go to `File` > `Options` > `Customize Ribbon`.
   - Check the `Developer` option to enable the Developer tab.
3. **Open VBA Editor**:
   - Click on the `Developer` tab.
   - Click on `Visual Basic` or press `ALT + F11` to open the VBA editor.
4. **Import VBA Module**:
   - Right-click on any item in the `Project Explorer` pane.
   - Select `Import File...` and choose the appropriate `.bas` file from the project folder.
5. **Adjust Settings**:
   - Navigate to `Tools` > `References` in the VBA editor.
   - Make sure to add any required references which might be listed in the project's documentation.
6. **Run the Macro**:
   - Close the VBA editor and return to Excel.
   - Press `ALT + F8`, select the macro, and click `Run`.

## Setup Instructions for Power Query

1. **Open Excel**: Launch Microsoft Excel.
2. **Access Power Query**:
   - Go to the `Data` tab.
   - Look for the `Get & Transform Data` section.
3. **Import Data**:
   - Click on `Get Data` > `From File` > `From Workbook` and browse for the Excel file you want to connect to.
4. **Transform Data**: 
   - Use Power Query editor to filter, transform, and prepare the data as needed.
5. **Load Data**:
   - After transforming, click on `Close & Load` to load the data into your Excel sheet.
6. **Refresh Data**:
   - Use `Data` > `Refresh All` to refresh the data as necessary.

## Common Issues
- Ensure that macros are enabled in Excel if the VBA functionality does not work.
- Check for any missing references in the VBA editor if errors occur while running macros.

By following these instructions, you should be able to successfully set up and use the VBA and Power Query functionalities of the Primavera P6 Excel Export project.