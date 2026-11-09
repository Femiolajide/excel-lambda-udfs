# IS_NULL_SUM function

Counts the blank cells in each column of a range and returns the counts as a two-column table.

## Syntax

```
=IS_NULL_SUM(arr)
```

## Arguments

| Argument | Required | Description |
| --- | --- | --- |
| `arr` | Yes | The data range, including the header row. |

## Returns

A table that spills into the cells below and to the right of the formula. It has two columns, `Columns` and `No_of_Missing_val`, and one row for each column of `arr`, plus a header row.

## Remarks

- The first row of `arr` must hold the column headers. They become the names in the first column of the result.
- A cell counts as blank when `COUNTBLANK` counts it. That includes empty cells and cells whose formula returns an empty string (`""`). A zero or a space is not blank.
- The header row is part of `arr`, so a blank header cell is counted as a missing value in its column.
- To treat text such as `NA` as missing too, use [`GET_NULLS`](GET_NULLS.md).

## Errors

| Condition | Result |
| --- | --- |
| The cells where the table would spill are not empty | `#SPILL!` |

## Examples

Cells `A1:D6` contain:

| Name | Age | Gender | Score |
| --- | ---: | --- | ---: |
| John | 25 | Male | 75 |
| Mary |  | Female | NA |
| Peter | 31 |  | 82 |
| Jane | 28 | Female |  |
| David | NA | Male | 91 |

| Formula | Result |
| --- | --- |
| `=IS_NULL_SUM(A1:D6)` | Name 0, Age 1, Gender 1, Score 1 |

`NA` is text, not a blank cell, so it is not counted. The result appears as a table:

| Columns | No_of_Missing_val |
| --- | ---: |
| Name | 0 |
| Age | 1 |
| Gender | 1 |
| Score | 1 |

## Definition

```
=LAMBDA(arr,VSTACK(HSTACK("Columns","No_of_Missing_val"),TRANSPOSE(VSTACK(TAKE(arr,1),BYCOL(arr,LAMBDA(x,COUNTBLANK(x)))))))
```

See [How to add a LAMBDA function to a workbook](how-to-add-a-lambda.md).

## How it works

| Function | Role |
| --- | --- |
| `BYCOL` | Applies `COUNTBLANK` to each column of `arr`. |
| `TAKE(arr,1)` | Takes the header row. |
| `VSTACK`, `TRANSPOSE` | Stack the headers above the counts, then turn the two rows into two columns. |
| `HSTACK`, `VSTACK` | Add the title row `Columns` and `No_of_Missing_val`. |

## Requirements

`LAMBDA`, `VSTACK`, `HSTACK`, `TAKE` and `BYCOL` must be available in your version of Excel.

## Design notes

- **Why it exists.** It is the Excel version of the pandas expression `df.isna().sum()`, so a quick missing-value check needs no trip to Python.
- **Why a LAMBDA.** The long nested formula is saved once and reused across workbooks.

## See also

[`GET_NULLS`](GET_NULLS.md), [`COUNTBLANK`](https://support.microsoft.com/en-us/excel/functions/countblank-function), [`BYCOL`](https://support.microsoft.com/en-us/excel/functions/bycol-function)
