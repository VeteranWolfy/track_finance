# Transaction Categorizer

A simple desktop application that helps you categorize financial transactions from a CSV file.

## Features

- Load transactions from a CSV file
- View transactions one at a time
- Categorize transactions using predefined categories or custom categories
- Navigate between transactions
- Save categorized data to a new CSV file

## Requirements

- Python 3.x
- pandas
- tkinter (usually comes with Python)

## Installation

1. Clone this repository
2. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Prepare your CSV file with the following columns:
   - date
   - description
   - cost

2. Run the application:
   ```
   python transaction_categorizer.py
   ```

3. Click "Select CSV File" to load your transaction data
4. For each transaction:
   - Click one of the predefined category buttons (1-8)
   - Or enter a custom category in the text field and click "Submit Category"
5. Use the "Previous" and "Next" buttons to navigate between transactions
6. Click "Save Categorized Data" when finished to save your categorized transactions

## Predefined Categories

1. Groceries
2. Transportation
3. Entertainment
4. Bills & Utilities
5. Dining Out
6. Shopping
7. Income
8. Other
