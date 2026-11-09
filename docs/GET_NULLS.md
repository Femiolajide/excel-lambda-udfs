# GET_NULLS function

Counts the blank cells, and the cells equal to one extra value that you choose, in each column of a range.

## Syntax

```
=GET_NULLS(arr, null_value)
```

## Arguments

| Argument | Required | Description |
| --- | --- | --- |
| `arr` | Yes | The data range, including the header row. |
| `null_value` | Yes | Text to count as missing, in addition to blank cells, for example `"NA"`. To count blank cells only, leave this argument empty but keep the comma: `=GET_NULLS(A1:E50,)`. |

## Returns

A table that spills into the cells below and to the right of the formula. It has two columns, `Columns` and `No_of_Missing_val`, and one row for each column of `arr`, plus a header row.

## Remarks

- Both arguments must be present. `=GET_NULLS(A1:E50,)` is valid. `=GET_NULLS(A1:E50)` returns `#VALUE!`.
- Only one extra value can be counted. To count several representations of a missing value, nest or add the results of separate calls.
- The extra value is matched with `COUNTIF`. The match is not case-sensitive, so `"NA"` also counts `na`. The characters `*`, `?` and `~` are wildcards.
- A cell counts as blank when `COUNTBLANK` counts it, including cells whose formula returns an empty string. A blank header cell is counted as a missing value in its column.
- [`IS_NULL_SUM`](IS_NULL_SUM.md) counts blank cells only.

## Errors

| Condition | Result |
| --- | --- |
| The second argument is left out, with no comma | `#VALUE!` |
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

| Formula | Result | Explanation |
| --- | --- | --- |
| `=GET_NULLS(A1:D6,"NA")` | Name 0, Age 2, Gender 1, Score 2 | Counts blank cells and cells equal to `NA`. |
| `=GET_NULLS(A1:D6,)` | Name 0, Age 1, Gender 1, Score 1 | The empty second argument counts blank cells only. |

The first formula returns this table:

| Columns | No_of_Missing_val |
| --- | ---: |
| Name | 0 |
| Age | 2 |
| Gender | 1 |
| Score | 2 |

### Other missing-value markers

| Formula | Counts |
| --- | --- |
| `=GET_NULLS(A1:E50,"NULL")` | Blank cells and cells equal to `NULL` |
| `=GET_NULLS(A1:E50,"N/A")` | Blank cells and cells equal to `N/A` |
| `=GET_NULLS(A1:E50,"Unknown")` | Blank cells and cells equal to `Unknown` |

## Definition

```
=LAMBDA(arr,null_value,VSTACK(HSTACK("Columns","No_of_Missing_val"),TRANSPOSE(VSTACK(TAKE(arr,1),BYCOL(arr,LAMBDA(x,COUNTBLANK(x)+COUNTIF(x,null_value)))))))
```

See [How to add a LAMBDA function to a workbook](how-to-add-a-lambda.md).

## How it works

The layout is the same as in [`IS_NULL_SUM`](IS_NULL_SUM.md). The calculation for each column is:

```
COUNTBLANK(x)+COUNTIF(x,null_value)
```

`COUNTBLANK` counts the blank cells and `COUNTIF` counts the cells equal to `null_value`.

## Requirements

`LAMBDA`, `VSTACK`, `HSTACK`, `TAKE` and `BYCOL` must be available in your version of Excel.

## Design notes

- **Why it extends `IS_NULL_SUM`.** Missing values are not always empty cells. Exports often use `NA`, `NULL` or `Unknown`, and one extra argument covers the most common case.
- **The pandas link.** Like `IS_NULL_SUM`, it follows `df.isna().sum()`, adapted to data where missing values are written as text.

## See also

[`IS_NULL_SUM`](IS_NULL_SUM.md), [`COUNTIF`](https://support.microsoft.com/en-us/excel/get-started/use-the-countif-function-in-microsoft-excel), [`COUNTBLANK`](https://support.microsoft.com/en-us/excel/functions/countblank-function)
