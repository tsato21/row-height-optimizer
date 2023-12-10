---
layout: default
title: 'Automate Adjusting Row Heights in Google Sheet'
---

## About this Project

This project uses Google Apps Script that automates adjusting row heights in Google Sheet based on contents. In fact, row height can automatically be adjusted by Google Sheets' built-in function, `Resize row` => `Fit to data`. However, as shown in the screenshots below, this function results in rows with single lines having very small row heights and multiple lines not providing any padding at the top and bottom of the cell, which can lead to a cramped and visually unappealing spreadsheet.
<div style="margin-left: 30px">
  <img src="assets/images/fit-to-data-feature.png" alt="Image of Fit to Data Feature" width="400" height="200">
</div>

<div style="margin-left: 30px">
  <img src="assets/images/fit-to-data-result.png" alt="Image of Fit to Data Result" width="400" height="200">
</div>

Manually resizing each row to fit its content is a potential solution. Still, this can be a time-consuming process, especially for many rows and spreadsheets to deal with. 
This project automates this process, saving users valuable time and ensuring a consistent, readable, and aesthetically pleasing row height throughout the spreadsheet. It adjusts the row height based on the content of each cell and also adds some space to the top and bottom for better readability and visual appeal.

## Prerequisites

- A Google account with access to Google Sheets.
- Basic understanding of Google Sheets and Google Apps Script.
- Understanding of the functions used in the script.

## Setup

1. **Open Your Google Sheet**: Access <a href="https://docs.google.com/spreadsheets/d/1154I0kMvhbp9WZ4r6COu_oA4lPJZvh9MnFxa2G7HZck/edit#gid=0" target="_blank" rel="noopener noreferrer">Sample Google Sheet</a>.
2. **Copy the Google Sheet**: Copy the sheet to make your sheet.
3. **Conduct GAS Authorization**: Click `Show Authorization` button from `Custom Menu`. This enables you to go to the authorization page for Google Apps Script.
4. **Change a constant variable in Google Apps Script**: Go to `Extensions`>`Apps Script` and change the value for each constant variables to suit your needs.
  - `AVERAGE_CHART_WIDTH`: The average width of the charts in your sheet. This is used to calculate the width of the text in each cell. The pre-set value, "4.5", is the average width of the charts in the sample sheet which uses Times New Roman font with 10pt font size.
  - `BASE_ROW_HEIGHT`: The base height of a row with a single line of text. The pre-set value, "20", is the base height of a row with a single line of text in the sample sheet which uses Times New Roman font with 10pt font size.
  - `ADDITIONAL_ROW_HEIGHT`: The additional height for each additional line of text. The pre-set value, "20", is the additional height for each additional line of text in the sample sheet which uses Times New Roman font with 10pt font size.
<div style="margin-left: 30px">
  <img src="assets/images/customize-constant-variables.png" alt="Image of Constant Variables" width="300" height="50">
</div>

## Usage

1. **Adjust Row Heights**: Select `Custom Menu` > `Adjust Row Heights` in the Google Sheets UI. This initiates a process that adjusts the height of each row in a specified sheet, starting from a user-defined row number, based on the content of the cells in each row. Here's the step-by-step logic:

  - **Input Sheet Name and Starting Row**: A dialog prompts you to input the name of the sheet and the starting row number.
  
  - **Data Retrieval**: The script retrieves the data from the specified sheet and starting row, and gets the widths of all columns in the sheet.
  
  - **Row Height Calculation**: For each row, the script calculates the line count for each cell, taking into account the physical width of the text and any manual line breaks. It then determines the largest line count among all cells in the row.
  
  - **Height Adjustment**: The row height is adjusted based on the largest line count. The height is calculated using a base height for a single line and an additional height for each additional line.
  
  This function ensures that the row height accommodates the cell with the most content, providing a cleaner and more organized view of your data.

## Others

- **Customization**: Scripts can be customized to fit specific needs and workflows.
