# UNIQUE_VALUES function

Returns the distinct values in a range or array as one sorted column, with blank cells shown as `#N/A`.

## Syntax

```
=UNIQUE_VALUES(array)
```

## Arguments

| Argument | Required | Description |
| --- | --- | --- |
| `array` | Yes | The range or array to extract distinct values from. |

## Returns

A single column that spills downwards, sorted in ascending order.

## Remarks

- Blank cells are changed to `#N/A` before the distinct values are extracted, so an empty cell is not confused with a real `0`. Cells that contain an empty string (`""`) are not blank, and are listed as an empty entry.
- Errors already in the data are not removed.
- A range with several columns is flattened row by row into one column.
- `UNIQUE` is not case-sensitive, so `Agree` and `agree` count as the same value and only the first one appears. Entries that differ by a space, such as `Agree` and `Agree `, are listed separately.
- The function lists distinct values. It does not decide whether a value is correct. That depends on the variable and the study's coding rules.

## Errors

| Condition | Result |
| --- | --- |
| The cells where the list would spill are not empty | `#SPILL!` |

## Examples

Cells `A2:A10` contain `Male`, `Female`, `Male`, `Other`, `Female`, `Unknown`, `0`, `Male` and one blank cell.

| Formula | Result |
| --- | --- |
| `=UNIQUE_VALUES(A2:A10)` | `0`, `Female`, `Male`, `Other`, `Unknown`, `#N/A` |

The number sorts first, the text follows in alphabetical order, and the `#N/A` that stands for the blank cell comes last.

## Definition

```
=LAMBDA(array,SORT(UNIQUE(TOCOL(IF(ISBLANK(array),NA(),array)),FALSE,FALSE)))
```

See [How to add a LAMBDA function to a workbook](how-to-add-a-lambda.md).

## How it works

| Step | Formula | What it does |
| --- | --- | --- |
| 1 | `ISBLANK(array)` | Finds the genuinely blank cells. |
| 2 | `IF(ISBLANK(array),NA(),array)` | Replaces each blank with `#N/A` and keeps every other value. |
| 3 | `TOCOL(...)` | Flattens the result into one column. |
| 4 | `UNIQUE(...,FALSE,FALSE)` | Returns each distinct value once. The two `FALSE` arguments compare rows, not columns, and return all distinct values, not only those that occur once. |
| 5 | `SORT(...)` | Sorts the list in ascending order. |

## Requirements

`LAMBDA`, `TOCOL`, `UNIQUE` and `SORT` must be available in your version of Excel.

## Design notes

- **Why `#N/A` for blanks.** Using `0` as a marker for empty cells would hide real zeros, and zero can be a valid value in a dataset.
- **Where it fits.** It is a first step in cleaning: list every distinct entry, find the unexpected ones, correct them, then code the categories with [`CODE_VALUES`](CODE_VALUES.md).

## See also

[`CODE_VALUES`](CODE_VALUES.md), [`UNIQUE`](https://support.microsoft.com/en-us/excel/functions/unique-function), [`TOCOL`](https://support.microsoft.com/en-us/excel/functions/tocol-function)
