# How to add a LAMBDA function to a workbook

Each function page has a **Definition** section with the formula. To use a function, save that formula as a named function in your workbook.

1. Select **Formulas**, then **Name Manager**, then **New**.
2. In **Name**, enter the function name exactly as shown on its page, for example `IS_NULL_SUM`.
3. Leave **Scope** as **Workbook**.
4. In **Refers to**, paste the definition. It must start with `=LAMBDA(`.
5. Select **OK**, then close Name Manager.

You can now use the function in any cell of that workbook, like a built-in function.

## Using the function in another workbook

A named function belongs to the workbook where you saved it. In another workbook, repeat the steps above. Copying a sheet that uses the function may also bring the name across. If you try that, check Name Manager to confirm the name arrived.

## Requirements

LAMBDA and the functions each page lists must be available in your version of Excel. They are available in Microsoft 365 and Excel for the web. For other versions, check Microsoft's support pages for LAMBDA and for each function the page lists.
