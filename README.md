# Trading Journal

This repository contains Python scripts for managing a trading journal, calculating realized profits and losses (PnL), and tracking active trades. The primary focus is on handling option and futures trades, recording necessary details, and exporting data for analysis.

## Overview

The code automates the process of generating a tradebook and calculating the following:

- **Realized PnL**: Profit or loss realized from closed trades.
- **Active Trades**: Tracking positions for each symbol, including costs and quantities.
- **Cumulative PnL**: Cumulative profit and loss from realized trades.

Note: The data included in this repository has been multiplied by a random factor to protect the privacy of the individuals involved

## Project Structure

- **`tradebook_builder.py`**: This script builds the initial tradebook DataFrame from your trading data.
- **`profit_calculator.py`**: This script processes the tradebook, calculates realized PnL, and tracks active trades.
- **`position_tracker.py`**: Manages the active trade positions, calculates PnL, and handles trade logic.
- **`active_trade_and_pnl.xlsx`**: The output file generated by the profit calculator, containing detailed trade information.

## Requirements

To run this project, ensure you have the following libraries installed:

- `pandas`: For data manipulation and analysis.
- `openpyxl`: For Excel file operations (if not included with pandas).

You can install the required libraries using pip:

```bash
pip install pandas openpyxl
