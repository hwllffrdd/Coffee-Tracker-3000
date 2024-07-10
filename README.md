# Coffee Tracker

A Python script to automate monthly coffee consumption tracking and billing for workplace coffee.

## Description

This script helps manage a workplace coffee system by:
- Loading the previous month's coffee consumption data
- Allowing updates to each employee's coffee count and payment status
- Providing options to add or remove employees
- Generating a new Excel sheet with the updated information

## Requirements

- Python 3.6+
- openpyxl library

## Installation

1. Clone this repository
2. Install the required library:

pip install openpyxl

## Usage

1. Ensure the previous month's Excel file is in the same directory as the script.
2. Rename the previous month's file to `Coffee_Sheet_[MONTH_YEAR].xlsx` (e.g., `Coffee_Sheet_June_2024.xlsx`).
3. Run the script with:

python coffee_tracker.py

4. Follow the prompts to update employee data, add new employees, or remove employees.

The script will generate a new Excel file named `Coffee_Sheet_[CURRENT_MONTH_YEAR].xlsx` in the same directory.

## Example File

An example of a previous month's sheet (`Coffee_Sheet_Example.xlsx`) is included in this repository. To use it:

1. Copy `Coffee_Sheet_Example.xlsx` to create a new file.
2. Rename the new file to match the previous month (e.g., `Coffee_Sheet_June_2024.xlsx` if the current month is July 2024).

## Important Note

The script relies on correct file naming to function properly. Always ensure that the previous month's file is named correctly (`Coffee_Sheet_[MONTH_YEAR].xlsx`) before running the script.
