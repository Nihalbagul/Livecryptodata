
---

# Cryptocurrency Live Data Fetch and Analysis

## Overview

This project fetches live cryptocurrency data for the top 50 cryptocurrencies by market capitalization using the CoinGecko API. The data is then analyzed to extract useful insights, and it is continuously updated in an Excel sheet. The analysis includes:

- Top 5 cryptocurrencies by market capitalization.
- Average price of the top 50 cryptocurrencies.
- Highest and lowest 24-hour percentage price change among the top 50.

The project is written in Python and utilizes `xlwings` (or `pandas` for an alternative), to keep an Excel sheet updated with live cryptocurrency data every 5 minutes.

## Features

- Fetches live data including cryptocurrency name, symbol, current price, market capitalization, 24-hour trading volume, and 24-hour percentage price change.
- Calculates key insights and provides analysis of the top 50 cryptocurrencies.
- Updates an Excel sheet in real-time (every 5 minutes).

## Requirements

Before running the project, ensure you have the following installed:

- Python 3.x
- Virtual environment setup
- Required Python libraries: `requests`, `pandas`, `xlwings` (or `openpyxl` for an Excel-independent version)

### Required Libraries Installation

You can install the required libraries using `pip`:

```bash
pip install requests pandas xlwings openpyxl
```

## Setup

1. Clone or download the project to your local system.
   
2. Navigate to the project directory and create a virtual environment (optional but recommended):

    ```bash
    python -m venv venv
    ```

3. Activate the virtual environment:

    - **Windows**:  
      ```bash
      venv\Scripts\activate
      ```
    - **Linux/MacOS**:  
      ```bash
      source venv/bin/activate
      ```

4. Install the dependencies as mentioned above:

    ```bash
    pip install -r requirements.txt
    ```

## How to Run the Project

1. Make sure you have Excel installed on your system (for `xlwings`) or use the `openpyxl` version if you don't have Excel.

2. Run the script using the following command:

    ```bash
    python main.py
    ```

3. The script will automatically fetch the data, perform the analysis, and update the Excel sheet every 5 minutes.

4. The generated Excel file (`crypto_analysis.xlsx`) will store live data and key insights.

## Analysis Performed

- **Top 5 Cryptocurrencies**: Based on market capitalization.
- **Average Price**: The average price of the top 50 cryptocurrencies.
- **Highest and Lowest Percentage Change**: Identifies the cryptocurrencies with the highest and lowest 24-hour percentage price change.

## Output

- **Excel File**: The output data is written to an Excel file (`crypto_analysis.xlsx`) and updated every 5 minutes.
- **Report**: You can create a report summarizing the key insights from the analysis.

## Known Issues

- Ensure Microsoft Excel is installed and properly registered for COM access if using `xlwings`. Alternatively, use `openpyxl` to avoid this dependency.

## Contact

For any issues or queries, feel free to contact me at:

**Nihal Bagul**  
nihalbagul08120506@gmail.com  
+91 6355203029  

---
