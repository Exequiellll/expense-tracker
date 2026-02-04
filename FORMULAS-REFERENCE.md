# Google Sheets Formulas Reference

This document contains all the formulas used in your Office Expense Tracker, with explanations and troubleshooting tips.

---

## Expenses Sheet Formulas

### Month Extraction Formula
**Location:** Column E (Month)
**Formula:** `=TEXT(A2,"MMMM")`

**What it does:** Converts the date in column A to the full month name (e.g., "February")

**How to use:**
- Enter this formula in cell E2
- Copy it down to all rows where you'll have data (e.g., E2:E1000)
- The formula will automatically update when you add dates

**Customization:**
- For abbreviated month: `=TEXT(A2,"MMM")` → "Feb"
- For month number: `=TEXT(A2,"MM")` → "02"

**Troubleshooting:**
- If it shows a number instead of month name, check that column A contains a valid date
- If it shows #VALUE!, make sure cell A2 isn't empty or contains invalid data

---

### Year Extraction Formula
**Location:** Column F (Year)
**Formula:** `=YEAR(A2)`

**What it does:** Extracts the 4-digit year from the date in column A (e.g., "2026")

**How to use:**
- Enter this formula in cell F2
- Copy it down to all rows where you'll have data
- Automatically updates when you add dates

**Troubleshooting:**
- If it shows #VALUE!, check that column A contains a valid date

---

## Summary Dashboard Formulas

### Total Expenses (All Time)
**Location:** Cell B3
**Formula:** `=SUM(Expenses!D:D)`

**What it does:** Adds up all amounts in column D of the Expenses sheet

**How to use:**
- This formula references the entire column D, so it automatically includes new expenses
- No need to update as you add more data

**Customization:**
- To sum only a specific range: `=SUM(Expenses!D2:D100)`

---

### Current Month Total
**Location:** Cell B4
**Formula:** `=SUMIFS(Expenses!D:D, Expenses!E:E, TEXT(TODAY(),"MMMM"), Expenses!F:F, YEAR(TODAY()))`

**What it does:** Sums all expenses where:
- The month (column E) matches the current month
- The year (column F) matches the current year

**How it works:**
- `SUMIFS` adds up values with multiple conditions
- `Expenses!D:D` is the range to sum (amounts)
- `Expenses!E:E` and `Expenses!F:F` are the ranges to check (month and year)
- `TEXT(TODAY(),"MMMM")` gets the current month name
- `YEAR(TODAY())` gets the current year

**Troubleshooting:**
- If it shows 0 but you have expenses this month, check that:
  - Column E formulas are working correctly
  - Your computer's date/time is set correctly
  - There are no extra spaces in the month names

**Customization:**
- For a specific month: Replace `TEXT(TODAY(),"MMMM")` with `"February"`
- For previous month: `=SUMIFS(Expenses!D:D, Expenses!E:E, TEXT(TODAY()-30,"MMMM"))`

---

### Current Year Total
**Location:** Cell B5
**Formula:** `=SUMIFS(Expenses!D:D, Expenses!F:F, YEAR(TODAY()))`

**What it does:** Sums all expenses where the year matches the current year

**Customization:**
- For a specific year: Replace `YEAR(TODAY())` with `2026`
- For previous year: `=SUMIFS(Expenses!D:D, Expenses!F:F, YEAR(TODAY())-1)`

---

### Number of Expenses
**Location:** Cell B6
**Formula:** `=COUNTA(Expenses!A:A)-1`

**What it does:** Counts how many expenses you've entered (counts non-empty cells in the Date column, minus 1 for the header)

**How it works:**
- `COUNTA` counts non-empty cells
- `-1` subtracts the header row

**Alternative:**
- To count only expenses with amounts: `=COUNTA(Expenses!D2:D1000)`

---

### Spending by Category
**Location:** Cells B10-B14
**Formulas:**
```
Food:      =SUMIF(Expenses!B:B, "Food", Expenses!D:D)
Travel:    =SUMIF(Expenses!B:B, "Travel", Expenses!D:D)
Supplies:  =SUMIF(Expenses!B:B, "Supplies", Expenses!D:D)
Equipment: =SUMIF(Expenses!B:B, "Equipment", Expenses!D:D)
Other:     =SUMIF(Expenses!B:B, "Other", Expenses!D:D)
```

**What it does:** Sums all expenses where the category matches the specified category

**How it works:**
- `SUMIF` adds up values based on one condition
- `Expenses!B:B` is the range to check (categories)
- `"Food"` is the condition to match
- `Expenses!D:D` is the range to sum (amounts)

**Troubleshooting:**
- If a category shows 0 but you have expenses, check:
  - Category names match exactly (case-sensitive)
  - No extra spaces in category names
  - Data validation is set up correctly

**Customization:**
- Add new categories: Copy the formula and change the category name
- Exclude certain expenses: Use SUMIFS with additional conditions

---

## Advanced Formula Examples

### Expenses in a Date Range
**Formula:** `=SUMIFS(Expenses!D:D, Expenses!A:A, ">=1/1/2026", Expenses!A:A, "<=1/31/2026")`

**What it does:** Sums expenses between two dates (January 1-31, 2026)

**How to use:**
- Replace the dates with your desired range
- Use this for custom reporting periods

---

### Average Expense Amount
**Formula:** `=AVERAGE(Expenses!D:D)`

**What it does:** Calculates the average of all expenses

**Alternative:**
- Average for a category: `=AVERAGEIF(Expenses!B:B, "Food", Expenses!D:D)`

---

### Largest Expense
**Formula:** `=MAX(Expenses!D:D)`

**What it does:** Finds the largest expense amount

**Pairing with details:**
```
Largest amount: =MAX(Expenses!D:D)
Description: =INDEX(Expenses!C:C, MATCH(MAX(Expenses!D:D), Expenses!D:D, 0))
```

---

### Count Expenses by Category
**Formula:** `=COUNTIF(Expenses!B:B, "Food")`

**What it does:** Counts how many expenses belong to a specific category

---

### Month-over-Month Change
**Formula:**
```
=SUMIFS(Expenses!D:D, Expenses!E:E, TEXT(TODAY(),"MMMM"), Expenses!F:F, YEAR(TODAY()))
-
SUMIFS(Expenses!D:D, Expenses!E:E, TEXT(TODAY()-30,"MMMM"), Expenses!F:F, YEAR(TODAY()))
```

**What it does:** Calculates the difference between current month and previous month spending

---

### Percentage by Category
**Formula:** `=B10/B3`

**What it does:** Shows what percentage of total spending is in a specific category (e.g., Food spending ÷ Total spending)

**Formatting:** Format as percentage (click **Format** → **Number** → **Percent**)

---

## Conditional Formatting Formulas

### Highlight High-Value Expenses
**Formula:** `=$D2>100`

**Where to apply:** Select range D2:D1000, then use this custom formula in conditional formatting

**What it does:** Highlights any expense over $100

**Customization:** Change `100` to your threshold amount

---

### Highlight Current Month Expenses
**Formula:** `=AND($E2=TEXT(TODAY(),"MMMM"), $F2=YEAR(TODAY()))`

**Where to apply:** Select range A2:F1000

**What it does:** Highlights all expenses from the current month

---

### Color Code by Amount Range
**Formulas for different ranges:**
- Low (0-50): `=$D2<=50`
- Medium (51-100): `=AND($D2>50, $D2<=100)`
- High (>100): `=$D2>100`

Apply different colors to each range

---

## Data Validation Formulas

### Custom Category List from Reference Sheet
Instead of typing categories manually, reference the Category Reference sheet:

**In Data Validation:**
1. Select **"List from a range"**
2. Enter range: `'Category Reference'!A2:A6`

This way, if you add/change categories in the reference sheet, the dropdown updates automatically

---

### Date Range Validation
To restrict dates to current year only:

**In Data Validation:**
1. Select **"Date"**
2. Choose **"is between"**
3. Enter: `1/1/2026` and `12/31/2026`

---

## Troubleshooting Common Issues

### Formula Shows as Text
**Problem:** Formula appears as `=SUM(A1:A10)` instead of calculating
**Solution:** Make sure the cell format is not set to "Plain text". Change to "Automatic" under **Format** → **Number**

### #REF! Error
**Problem:** Formula references a deleted cell or sheet
**Solution:** Check that all sheet names and cell references are correct

### #VALUE! Error
**Problem:** Formula is trying to perform math on non-numeric data
**Solution:** Check that all referenced cells contain the expected data types

### #DIV/0! Error
**Problem:** Formula is dividing by zero
**Solution:** Wrap formula with `IFERROR`: `=IFERROR(B10/B3, 0)`

### #N/A Error
**Problem:** Lookup function can't find the value
**Solution:** Check that the lookup value exists in the data range

### Circular Dependency Error
**Problem:** A formula references itself
**Solution:** Review the formula and break the circular reference

---

## Useful Google Sheets Functions

Here are additional functions you might find useful:

### Date Functions
- `TODAY()` - Current date
- `NOW()` - Current date and time
- `DATE(year, month, day)` - Create a date
- `EDATE(date, months)` - Date n months from start date
- `DATEDIF(start, end, unit)` - Difference between dates

### Math Functions
- `SUM(range)` - Add up numbers
- `AVERAGE(range)` - Calculate average
- `MIN(range)` - Find smallest value
- `MAX(range)` - Find largest value
- `ROUND(value, decimals)` - Round to decimals

### Conditional Functions
- `IF(condition, true_value, false_value)` - Conditional logic
- `SUMIF(range, criterion, sum_range)` - Sum with one condition
- `SUMIFS(sum_range, range1, criterion1, ...)` - Sum with multiple conditions
- `COUNTIF(range, criterion)` - Count with condition
- `AVERAGEIF(range, criterion, average_range)` - Average with condition

### Text Functions
- `TEXT(value, format)` - Format value as text
- `CONCATENATE(text1, text2, ...)` or `&` - Join text
- `LEFT(text, n)` - First n characters
- `RIGHT(text, n)` - Last n characters
- `TRIM(text)` - Remove extra spaces

### Lookup Functions
- `VLOOKUP(search_key, range, index, [is_sorted])` - Vertical lookup
- `INDEX(range, row, column)` - Get value at position
- `MATCH(search_key, range, [search_type])` - Find position of value

---

## Formula Best Practices

1. **Use Named Ranges:** Instead of `Expenses!D:D`, create a named range called "Amounts"
   - Select the range → **Data** → **Named ranges**
   - Then use: `=SUM(Amounts)`

2. **Document Complex Formulas:** Add a comment to cells with complex formulas
   - Right-click cell → **Insert comment**
   - Explain what the formula does

3. **Test Formulas:** Before copying down, test on a few rows to ensure correct results

4. **Use Absolute References:** When copying formulas, use `$` to fix references
   - `$A$1` - Fixed row and column
   - `$A1` - Fixed column only
   - `A$1` - Fixed row only

5. **Break Down Complex Formulas:** Use helper columns for intermediate calculations

6. **Error Handling:** Wrap formulas with `IFERROR` to display friendly messages
   - `=IFERROR(formula, "No data")`

---

## Need More Help?

- [Google Sheets Function List](https://support.google.com/docs/table/25273)
- [Formula Writing Tips](https://support.google.com/docs/answer/46977)
- [Common Formula Errors](https://support.google.com/docs/answer/3093359)

---

**Happy Calculating!** Use this reference anytime you need to modify or understand the formulas in your expense tracker.
