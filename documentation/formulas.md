Excel Formulas Used

This section explains the Excel formulas used in the Personal Finance Tracker to calculate income, expenses, balance, category-wise expenses, and expense transaction count.

**1. Total Income**

Purpose: Calculates the total income recorded in "TransactionSheet".

Excel Formula:
---
=SUMIF(TransactionSheet!D:D,"Income",TransactionSheet!E:E)
---

How it works:

- "TransactionSheet!D:D" — Checks the transaction type.
- ""Income"" — Selects transactions marked as Income.
- "TransactionSheet!E:E" — Adds the corresponding amounts.

Result: Displays the total income.

---

**2. Total Expenses**

Purpose: Calculates the total expenses recorded in "TransactionSheet".

Excel Formula:
---
=SUMIF(TransactionSheet!D:D,"Expense",TransactionSheet!E:E)
---

How it works:

- Checks column D for transactions marked as "Expense".
- Adds the corresponding amounts from column E.

Result: Displays the total expenses.

---

**3. Remaining Balance**

Purpose: Calculates the money remaining after deducting expenses from income.

Calculation:

=Total Income - Total Expenses

Example Excel Formula:

---

=B2-B3
---

Assumption: Cell "B2" contains Total Income and cell "B3" contains Total Expenses. Adjust the cell references to match your worksheet.

Result: Displays the remaining balance.

---

**4. Category-wise Expenses**

Purpose: Calculates the total expenses for each category.

Excel Formula:

---
=SUMIFS(TransactionSheet!E:E,TransactionSheet!C:C,A2,TransactionSheet!D:D,"Expense")
---

How it works:

- "TransactionSheet!E:E" — Amounts to sum.
- "TransactionSheet!C:C" — Category column.
- "A2" — The category to calculate.
- "TransactionSheet!D:D,"Expense"" — Includes only expense transactions.

Result: Displays the total expenses for the selected category.

Note: This formula uses "SUMIFS" to ensure that only transactions marked as "Expense" are included in each category's total.

---

**5. Expense Transaction Count**

Purpose: Counts the number of transactions marked as expenses.

Excel Formula:

---
=COUNTIF(TransactionSheet!D:D,"Expense")
---

How it works:

- Checks column D for the transaction type "Expense".
- Counts the matching transactions.

Result: Displays the total number of expense transactions.

---

TransactionSheet Column Structure

Column| Description
C| Category
D| Type (Income/Expense)
E| Amount

**Important Notes**

- Ensure the worksheet name is exactly "TransactionSheet".
- Verify that the column references match your actual Excel file.
- Enter transaction types consistently as "Income" or "Expense".
- Use the correct cell references for the Remaining Balance formula.
- Format financial amounts as currency for better readability.

Summary

These formulas automate financial calculations and help track income, expenses, category-wise spending, remaining balance, and expense transaction counts.
