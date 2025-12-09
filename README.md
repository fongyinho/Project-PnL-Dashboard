## Project P&L Dashboard

A financial performance dashboard built using Python, Plotly, and Dash for analysing project-level P&L, including revenue, budget, actual spending, and trend analytics.
This dashboard supports finance teams and project managers by providing a clear, interactive overview of profitability and cost structure.


## Features

- **Single-Project P&L Snapshot**
  - Project info card: project name, status, type (TDP/PDP), division, department, PICs and managers.
  - Summary KPIs: Total Budget, Total Actual, and Remaining Balance.

- **Budget Usage Donut**
  - Donut chart showing overall budget usage % for the selected project.
  - Quickly see how much budget has been consumed vs. remaining.

- **P&L by Category**
  - Table view by **expense category** (R&D, Admin, Fixed Assets / Intangible Assets, Labour Cost, etc.).
  - Columns: Budget Amount, Actual Amount, Usage %, Balance.
  - Bar chart “Budget vs Actual (By Category)” aligned with the table.

- **P&L by Subject**
  - Detailed breakdown by **expense subject** (Material Cost, Tooling & Fixture Cost, Internal Testing Cost, etc.).
  - Same structure: Budget, Actual, Usage %, Balance.

- **Monthly Expense Trend**
  - Line chart of actual expenses by month.
  - Helps identify spending spikes and low-spend periods across the project lifecycle.

- **Budget vs Actual by Stage**
  - Stage-based comparison (e.g. R0–R5 or M0–M6, depending on project type).
  - Visualises how the budget is consumed across different development stages.

- **Project Selector**
  - Dropdown at the top to switch between project codes (e.g. `2020X14`).
  - All cards, tables and charts are refreshed based on the selected project.

- **Clean, Excel-style Layout**
  - UI design follows a structured report layout similar to a finance dashboard in Excel.
  - Numbers formatted in kCNY for easier reading.
 

How to Run the Dashboard

1. Clone the repository
   
  git clone https://github.com/YOUR_USERNAME/Project-PnL-Dashboard.git

  cd Project-PnL-Dashboard


3. Install dependencies
   
  pip install -r requirements.txt


5. Run the Dash app
   
  python app.py


  The dashboard will start at:
  👉 http://127.0.0.1:8050/


Project Structure

Project-PnL-Dashboard/

│

├── app.py                 # Main Dash application

├── data/

│   └── sample_PnL.xlsx    # Depersonalized sample dataset

│

├── assets/

│   └── styles.css         # Custom CSS (optional)

│

├── requirements.txt       # Python dependencies

└── README.md              # Project documentation

