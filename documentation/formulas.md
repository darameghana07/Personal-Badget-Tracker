Excel Formulas Used

1. Total Income

Calculates the total income recorded in the TransactionSheet.

=SUMIF(TransactionSheet!D:D,"Income",TransactionSheet!E:E)

2. Total Expenses

Calculates the total expenses recorded in the TransactionSheet.

=SUMIF(TransactionSheet!D:D,"Expense",TransactionSheet!E:E)

3. Remaining Balance

Calculates the money remaining after expenses.

Calculation: Total Income − Total Expenses

4. Category-wise Expenses

Calculates the total expenses for each category.

=SUMIF(TransactionSheet!C:C,A2,TransactionSheet!E:E)

5. Expense Transaction Count

Counts transactions marked as expenses.

=COUNTIF(TransactionSheet!D:D,"Expense")

Note: These formulas assume column C contains Category, column D contains Type, and column E contains Amount. Make sure these match your actual Excel sheet.