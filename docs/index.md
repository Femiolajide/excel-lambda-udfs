# LAMBDA function index

Custom Excel functions written with `LAMBDA`. Each page follows the same layout: syntax, arguments, remarks, errors, examples, the formula to paste, and how it works.

To use one, follow [How to add a LAMBDA function to a workbook](how-to-add-a-lambda.md).

## Missing values

| Function | Description | 
| --- | --- | 
| [`IS_NULL_SUM`](IS_NULL_SUM.md) | Counts blank cells in each column | 
| [`GET_NULLS`](GET_NULLS.md) | Counts blank cells and one extra missing-value marker in each column | 

## Survey data preparation

| Function | Description 
| --- | --- | 
| [`UNIQUE_VALUES`](UNIQUE_VALUES.md) | Lists the distinct values in a range, sorted, with blanks shown as `#N/A` | 
| [`CODE_VALUES`](CODE_VALUES.md) | Converts text labels into sequential numerical codes |
| [`CUSTOM_CODE_VALUES`](CUSTOM_CODE_VALUES.md) | Replaces values using two matching lists, to recode or collapse categories 

## Dates and times

| Function | Description |
| --- | --- | 
| [`FROM_TIMESTAMP`](FROM_TIMESTAMP.md) | Converts a UNIX timestamp into `yyyy-mm-dd hh:mm:ss` text | 

## A typical data preparation workflow

1. Use `UNIQUE_VALUES` to list every distinct entry in a questionnaire column.
2. Find unexpected entries, inconsistent spellings and missing-value markers.
3. Use `IS_NULL_SUM` or `GET_NULLS` to see how much data is missing.
4. Clean and standardise the responses.
5. Use `CODE_VALUES` or `CUSTOM_CODE_VALUES` to turn the categories into the codes your analysis needs.

A function is marked **Experimental** until its examples have been checked in Excel.
