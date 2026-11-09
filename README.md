# excel-lambda-functions
 
Custom Excel functions built with `LAMBDA`, for data tasks that need long nested formulas or that Excel's built-in functions do not cover. Each function is documented like a native Excel function, with its formula, examples and error behaviour.
 
## Why this exists
 
Excel's built-in functions cover a lot, but some data tasks need formulas that are long and hard to write, read and reuse. Checking missing values across many columns, turning survey labels into numerical codes, and converting timestamps are examples. A `LAMBDA` lets me write the logic once, give it a name and use it like any other Excel function. There is no VBA and no add-in, only a formula saved in the workbook.
 
These functions come from my series *Thinking in Excel Formulas and Functions*. They start from real data problems, from features I wish Excel had, and from ideas I think are useful in data analysis generally.
 
## What every function includes
 
- **A documentation page** in the style of Excel's own function pages: syntax, arguments, remarks, errors and examples.
- **The full formula,** ready to paste into a workbook, and a short explanation of how it works.
- **Defined behaviour for awkward input,** such as blank cells, text versus numbers, upper and lower case, and values that are not found.
- **Examples that have been run in Excel,** with the results shown on the page.
- **A repeatable result.** The same input always gives the same output.
## Functions
 
The [function index](docs/index.md) lists every function with a one-line description and its status, grouped by task. Each function has its own page in the `docs` folder.
 
 
## Requirements
 
An Excel version that supports `LAMBDA`, and the functions listed on each function's page. They are available in Microsoft 365 and Excel for the web. For other versions, check Microsoft's support pages.
 
## Using a function
 
1. Open the function's page in `docs` and copy its **Definition**.
2. In Excel, select **Formulas**, then **Name Manager**, then **New**.
3. Enter the function name, paste the definition into **Refers to**, and select **OK**.
The full steps are in [How to add a LAMBDA function to a workbook](docs/how-to-add-a-lambda.md).
 
## Earlier Gists
 
Some of these functions were first shared as GitHub Gists. This repository replaces them. The pages here are the current versions, and the Gists link back to this repository.
 

## Feedback
 
If a function gives a result you did not expect, please open an issue with the formula, the input and the result you expected or reach out me via [LinkedIn](https://www.linkedin.com/in/olajide-atanda-084756239). 
 

## About

- [LinkedIn](https://www.linkedin.com/in/olajide-atanda-084756239)
