# Google Sheets Expense Tracker - Setup Guide

This guide will walk you through setting up a complete expense tracking system in Google Sheets with automated calculations, charts, and reports.

## Quick Start

### Step 1: Create Your Google Sheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Click "+ Blank" to create a new spreadsheet
3. Name it "Office Expense Tracker"

### Step 2: Import the CSV Template

1. Click **File** → **Import**
2. Click the **Upload** tab
3. Drag and drop `expense-template.csv` or click **Browse** to select it
4. Choose **"Replace spreadsheet"** as the import location
5. Choose **"Comma"** as the separator type
6. Click **Import data**

### Step 3: Rename and Create Additional Sheets

1. Double-click the tab at the bottom that says "Sheet1"
2. Rename it to **"Expenses"**
3. Create three more sheets by clicking the "+" button at the bottom left:
   - **Summary Dashboard**
   - **Monthly Reports**
   - **Category Reference**

---

## Detailed Setup Instructions

### Part 1: Configure the Expenses Sheet

#### 1.1 Format the Header Row

1. Click on row 1 to select the entire row
2. Click **Format** → **Text wrapping** → **Wrap**
3. Make it bold: Click the **B** (Bold) button
4. Add background color: Click the fill color icon → choose a light color (e.g., light blue)
5. Freeze the header: Click **View** → **Freeze** → **1 row**

#### 1.2 Set Up Data Validation

**For the Date Column (Column A):**
1. Click on cell A2
2. Click **Data** → **Data validation**
3. In "Criteria", select **"Date"** → **"is valid date"**
4. Check **"Reject input"** for invalid data
5. Click **Done**
6. Copy cell A2, select cells A3:A1000, and paste (Ctrl+V / Cmd+V)

**For the Category Column (Column B):**
1. Click on cell B2
2. Click **Data** → **Data validation**
3. In "Criteria", select **"List of items"**
4. Enter: `Food, Travel, Supplies, Equipment, Other`
5. Check **"Show dropdown list in cell"**
6. Check **"Reject input"** for invalid data
7. Click **Done**
8. Copy cell B2, select cells B3:B1000, and paste

**For the Amount Column (Column D):**
1. Click on cell D2
2. Click **Data** → **Data validation**
3. In "Criteria", select **"Number"** → **"greater than"** → enter `0`
4. Check **"Reject input"** for invalid data
5. Click **Done**
6. Copy cell D2, select cells D3:D1000, and paste

#### 1.3 Add Formulas for Month and Year

The template already includes formulas in columns E and F. These formulas automatically extract the month and year from the date in column A.

**If you need to add them manually:**
- In cell E2, enter: `=TEXT(A2,"MMMM")`
- In cell F2, enter: `=YEAR(A2)`
- Copy these cells down to row 1000

#### 1.4 Set Up Conditional Formatting (Color Code by Category)

1. Select the range B2:B1000 (Category column)
2. Click **Format** → **Conditional formatting**
3. Under "Format rules", select **"Text is exactly"**
4. Enter `Food` and choose a color (e.g., light green)
5. Click **Done**
6. Repeat for other categories:
   - **Travel**: Light blue
   - **Supplies**: Light yellow
   - **Equipment**: Light orange
   - **Other**: Light gray

**Alternative: Color entire rows by category**
1. Select A2:F1000
2. Click **Format** → **Conditional formatting**
3. Under "Format rules", select **"Custom formula is"**
4. Enter: `=$B2="Food"`
5. Choose a color and click **Done**
6. Repeat for each category

#### 1.5 Format the Amount Column as Currency

1. Select column D (click the column header)
2. Click **Format** → **Number** → **Currency**

---

### Part 2: Create the Summary Dashboard

Go to the **"Summary Dashboard"** sheet.

#### 2.1 Set Up Summary Metrics

Create a clean layout with these metrics:

**Cell A1:** Type "EXPENSE SUMMARY"
- Make it bold, large (18pt), centered

**Cell A3:** Type "Total Expenses (All Time)"
**Cell B3:** Enter formula: `=SUM(Expenses!D:D)`

**Cell A4:** Type "Current Month Total"
**Cell B4:** Enter formula: `=SUMIFS(Expenses!D:D, Expenses!E:E, TEXT(TODAY(),"MMMM"), Expenses!F:F, YEAR(TODAY()))`

**Cell A5:** Type "Current Year Total"
**Cell B5:** Enter formula: `=SUMIFS(Expenses!D:D, Expenses!F:F, YEAR(TODAY()))`

**Cell A6:** Type "Number of Expenses"
**Cell B6:** Enter formula: `=COUNTA(Expenses!A:A)-1`

Format cells B3:B6 as currency.

#### 2.2 Category Breakdown

**Cell A8:** Type "SPENDING BY CATEGORY"

**Cell A10:** Type "Food"
**Cell B10:** Enter formula: `=SUMIF(Expenses!B:B, "Food", Expenses!D:D)`

**Cell A11:** Type "Travel"
**Cell B11:** Enter formula: `=SUMIF(Expenses!B:B, "Travel", Expenses!D:D)`

**Cell A12:** Type "Supplies"
**Cell B12:** Enter formula: `=SUMIF(Expenses!B:B, "Supplies", Expenses!D:D)`

**Cell A13:** Type "Equipment"
**Cell B13:** Enter formula: `=SUMIF(Expenses!B:B, "Equipment", Expenses!D:D)`

**Cell A14:** Type "Other"
**Cell B14:** Enter formula: `=SUMIF(Expenses!B:B, "Other", Expenses!D:D)`

Format cells B10:B14 as currency.

#### 2.3 Add a Pie Chart for Category Spending

1. Select cells A10:B14
2. Click **Insert** → **Chart**
3. In the Chart editor (right side):
   - Chart type: **Pie chart**
   - Customize the title: "Spending by Category"
4. Move and resize the chart as desired

#### 2.4 Add Top 5 Largest Expenses

**Cell D3:** Type "TOP 5 EXPENSES"

**Cell D5:** Type "Date"
**Cell E5:** Type "Description"
**Cell F5:** Type "Amount"

In cells D6:F10, manually reference the top expenses or use this approach:
1. Copy the expense data from the Expenses sheet
2. Sort by Amount (descending)
3. Display the top 5

(Alternatively, you can use a pivot table or query function for automatic updates)

---

### Part 3: Set Up Monthly Reports

Go to the **"Monthly Reports"** sheet.

#### 3.1 Create a Pivot Table

1. Click **Insert** → **Pivot table**
2. Select **"Expenses"** as the data range
3. Choose **"Existing sheet"** and select cell A1
4. Click **Create**

#### 3.2 Configure the Pivot Table

In the Pivot table editor (right side):

**Rows:**
- Add **"Month"**
- Add **"Category"**

**Values:**
- Add **"Amount"**
- Summarize by: **SUM**

**Columns:** (Optional)
- Add **"Year"** if you want to compare years side by side

This will create a breakdown of expenses by month and category.

---

### Part 4: Create Category Reference Sheet

Go to the **"Category Reference"** sheet.

**Cell A1:** Type "Category"
**Cell B1:** Type "Description"
**Cell C1:** Type "Color Code"

Fill in the categories:

| Category | Description | Color Code |
|----------|-------------|------------|
| Food | Team meals, snacks, catering | Green |
| Travel | Transportation, taxis, mileage | Blue |
| Supplies | Office supplies, paper, stationery | Yellow |
| Equipment | Hardware, software, tools | Orange |
| Other | Miscellaneous expenses | Gray |

This sheet serves as a reference and can be used for dropdown validation.

---

## Daily Usage

### Adding a New Expense

1. Go to the **Expenses** sheet
2. Find the next empty row
3. Enter:
   - **Date**: Click the cell and select a date from the calendar picker
   - **Category**: Select from the dropdown
   - **Description**: Type a brief description
   - **Amount**: Enter the amount (numbers only)
4. The Month and Year columns will auto-fill

### Viewing Reports

- **Quick Summary**: Go to **Summary Dashboard** to see totals and charts
- **Detailed Breakdown**: Go to **Monthly Reports** to see pivot table
- **Filter/Sort**: Use Google Sheets' built-in filter and sort features on the Expenses sheet

### Exporting Data

To export your expenses:
1. Click **File** → **Download**
2. Choose **"Comma-separated values (.csv)"** for Excel compatibility
3. Or choose **"Microsoft Excel (.xlsx)"**

---

## Tips and Best Practices

### Backup Your Data
- Google Sheets auto-saves, but you can make manual backups
- Use **File** → **Make a copy** to create snapshots

### Mobile Access
- Download the Google Sheets app for iOS/Android
- You can add expenses on the go

### Sharing with Team Members
- Click the **Share** button (top right)
- Add email addresses of team members
- Set permissions: **Editor** (can add/edit), **Viewer** (can only view)

### Keyboard Shortcuts
- `Ctrl + ;` (Windows) / `Cmd + ;` (Mac): Insert today's date
- `Tab`: Move to next cell
- `Enter`: Move down to next row

### Troubleshooting

**Problem: Formulas show as text**
- Solution: Make sure formulas start with `=`

**Problem: Data validation not working**
- Solution: Check that you applied validation to the entire column range (A2:A1000, etc.)

**Problem: Charts not updating**
- Solution: Charts should update automatically. If not, click the chart → three dots menu → **Refresh**

**Problem: Pivot table not showing new data**
- Solution: Click inside the pivot table → three dots menu → **Refresh**

---

## Customization Ideas

### Add Budget Tracking
- Add a "Budget" column in Category Reference
- Compare actual spending vs. budget in Summary Dashboard

### Add Monthly Trends Chart
1. Create a summary table with months and totals
2. Insert a line chart to visualize spending over time

### Add Receipt Links
- Add a "Receipt" column in Expenses sheet
- Store links to photos/scans of receipts (in Google Drive)

### Set Up Alerts
- Use Google Sheets' built-in notification rules
- Get email alerts when spending exceeds a threshold

---

## Need Help?

- [Google Sheets Help Center](https://support.google.com/docs#topic=1382883)
- [Data Validation Guide](https://support.google.com/docs/answer/186103)
- [Pivot Table Guide](https://support.google.com/docs/answer/1272900)

---

**Happy Tracking!** Your expense tracker is now ready to use. Start adding your office expenses and watch the reports update automatically.
