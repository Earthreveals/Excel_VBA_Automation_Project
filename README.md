# Excel VBA Automation Project

This repository contains a VBA script that automates the process of analyzing stock data for a given year. The script loops through all the stocks, calculates key metrics, and applies conditional formatting to highlight important data points.

## Overview

The objective of this project is to use VBA (Visual Basic for Applications) to automate the analysis of stock data in Excel. The script performs the following tasks:
- Extracts the ticker symbol.
- Calculates the yearly change in stock price from the opening price at the beginning of the year to the closing price at the end of the year.
- Calculates the percentage change in stock price from the opening price at the beginning of the year to the closing price at the end of the year.
- Calculates the total stock volume for each ticker.
- Applies conditional formatting to the yearly change and percentage change columns to highlight positive and negative changes.

## Features

- **Ticker Symbol**: Identifies and extracts the ticker symbol for each stock.
- **Yearly Change**: Calculates the change in stock price from the beginning to the end of the year.
- **Percentage Change**: Calculates the percentage change in stock price over the year.
- **Total Stock Volume**: Sums the total volume of stocks traded throughout the year.
- **Conditional Formatting**: Applies conditional formatting to highlight positive changes in green and negative changes in red.
- **Greatest Percent Increase, Greatest Percent Decrease, Greatest Total Volume**: Identifies the stocks with the greatest percentage increase, greatest percentage decrease, and greatest total volume for the year.

## Files in the Repository

- **VBA_Challenge_Script.vb**: The VBA script that performs the stock data analysis.
- **Screenshots**: 
  - `Screenshot_1.png`: Example output showing the ticker symbols and yearly changes.
  - `Screenshot_2.png`: Example output showing the percentage changes.
  - `Screenshot_3.png`: Example output showing the total volume and conditional formatting.

## Project Structure
README.md
├── VBA_Challenge_Script.vb
├── Screenshot_1.png
├── Screenshot_2.png
└── Screenshot_3.png


## How to Use This Repository

1. **Clone the repository:**
   ```sh
   git clone https://github.com/yourusername/Excel_VBA_Automation_Project.git
   cd Excel_VBA_Automation_Project

2. Open the Excel file:

Open the Excel workbook where you want to run the VBA script.

3. Insert the VBA Script:

Open the VBA editor by pressing ALT + F11.
Insert a new module by clicking Insert > Module.
Copy the contents of VBA_Challenge_Script.vb and paste it into the new module.

4. Run the Script:

Close the VBA editor.
Run the script by pressing ALT + F8, selecting the script, and clicking Run

## License
This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgements
This project was completed with the help of the study group and tutoring sessions. Special thanks to the contributors who provided guidance and support throughout the development of this project.
