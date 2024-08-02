# Expense Tracker

## Overview
The Expense Tracker is a Python-based application that helps you track your daily expenses. It prompts you to enter details about each expense, saves them to an Excel file, and keeps a running total of the amount spent. This project utilizes Object-Oriented Programming (OOP) concepts and the `pandas` library for data management.

## Features
- Add multiple expenses with details such as item, amount, date, and notes.
- Automatically assigns a serial number to each expense.
- Displays a summary of all expenses and the total amount spent.
- Saves all expenses to an Excel file (`expenses.xlsx`).
- Loads existing expenses from the Excel file on startup to avoid overwriting data.
- Prompts the user to close the Excel file if it's open when saving new data.

## Prerequisites
- Python 3.x
- `pandas` library

You can install the `pandas` library using pip:
```bash
pip install pandas
