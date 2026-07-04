# Business Sales & Expense Tracker (Excel)

A ready-to-use Excel system I built for small business owners to track daily sales and expenses without needing pivot tables or any external software. You enter your data, the dashboard updates itself.

## Why I built this

I noticed most small shop owners and freelancers track their sales in random notebooks or messy spreadsheets. Basic questions like "what sold today" or "which expense is costing us the most" take way longer to answer than they should. So I built a single Excel file that handles all of it automatically using formulas — no manual recalculating, no separate reports.

## What it does

- Product entry is dropdown-based, so you're not retyping product names every time and making typos
- The dashboard shows total sales, total expenses, net profit, and profit margin, updated live
- There's a "Today's Sales" section that tells you exactly what sold today, product by product
- An all-time totals table so you can see which products have sold the most overall
- Auto-written insights — it tells you your best-selling product, your biggest expense category, and whether you're currently profitable or not
- A bar chart for monthly sales and a pie chart for expense categories
- An optional "custom field" column in both entry sheets in case someone wants to track extra info like customer name or payment mode, without messing up any formula

## How it's built

Everything runs on native Excel formulas — SUMIFS, SUMPRODUCT, INDEX/MATCH, IFERROR — combined with data validation for the dropdowns and native Excel charts for the visuals. I built it using Python (openpyxl) so I could iterate on the structure quickly and keep it error-checked, but the final file is a normal .xlsx that works in Excel or Google Sheets like any other.

## Files in this repo

- `Business_Sales_Expense_Tracker.xlsx` — the actual tracker
- `build_template.py`, `add_features.py`, `add_dashboard_upgrade.py` — the scripts used to generate it
- This README

## How to use it

Open the file and start with the Instructions tab — it walks through everything. Add your products in the Product List tab, log your sales/expenses daily, and check the Dashboard tab whenever you want a snapshot of how the business is doing.

## Who this is for

Small business owners, shop owners, or freelancers who want something simple to track their numbers without paying for accounting software they don't really need.
