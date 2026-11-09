# CUSTOM_CODE_VALUES function

Replaces a value with the replacement you choose from two matching lists. Use it to recode categories, collapse several categories into fewer, or give labels new names.

## Syntax

```
=CUSTOM_CODE_VALUES(value, old_values, new_values)
```

## Arguments

| Argument | Required | Description |
| --- | --- | --- |
| `value` | Yes | The cell or value to transform. |
| `old_values` | Yes | The existing values as one text string, separated by commas. |
| `new_values` | Yes | The replacement values as one text string, separated by commas, in the same order as `old_values`. |

## Returns

A single value: the replacement for `value`. If the replacement can be read as a number, it is returned as a number. Otherwise it is returned as text.

## Remarks

- `old_values` and `new_values` must have the same number of entries, in matching order.
- Matching is exact, but not case-sensitive.
- Do not put spaces around the commas. A space becomes part of the entry and stops it matching.
- Entries that contain commas are not supported.
- Several old values can share one replacement, which collapses categories.
- A replacement that Excel can read as a number, percentage, date or time is converted. For example, `"3"` becomes the number 3 and `"5%"` becomes 0.05.
- The function does not correct spelling, remove spaces or check that the recoding suits the analysis. It applies only the mapping you supply.
- To give sequential codes without writing the replacements, use [`CODE_VALUES`](CODE_VALUES.md).

## Errors

| Condition | Result |
| --- | --- |
| `value` is not in `old_values` | `#N/A` |
| `new_values` has fewer entries than `old_values`, and `value` matches an entry with no replacement | `#N/A` |

## Examples

### Collapsing categories

| Formula | Result | Explanation |
| --- | --- | --- |
| `=CUSTOM_CODE_VALUES("Always","Always,Often,Sometimes,Rarely","Frequent,Frequent,Not frequent,Not frequent")` | `Frequent` | Two old values map to one new value. |
| `=CUSTOM_CODE_VALUES("Rarely","Always,Often,Sometimes,Rarely","Frequent,Frequent,Not frequent,Not frequent")` | `Not frequent` | |

With the value in a cell, use `=CUSTOM_CODE_VALUES(A2,"Always,Often,Sometimes,Rarely","Frequent,Frequent,Not frequent,Not frequent")`.

### Recoding labels as numbers

| Formula | Result | Explanation |
| --- | --- | --- |
| `=CUSTOM_CODE_VALUES("Agree","Strongly Disagree,Disagree,Agree,Strongly Agree","1,2,3,4")` | `3` | The replacement is returned as a number. |

### Value not found

| Formula | Result | Explanation |
| --- | --- | --- |
| `=CUSTOM_CODE_VALUES("Never","Always,Often,Sometimes,Rarely","Frequent,Frequent,Not frequent,Not frequent")` | `#N/A` | `Never` is not in `old_values`. |

## Definition

```
=LAMBDA(value,old_values,new_values,IFERROR(VALUE(HLOOKUP(value,VSTACK(TEXTSPLIT(old_values,","),TEXTSPLIT(new_values,",")),2,FALSE)),HLOOKUP(value,VSTACK(TEXTSPLIT(old_values,","),TEXTSPLIT(new_values,",")),2,FALSE)))
```

See [How to add a LAMBDA function to a workbook](how-to-add-a-lambda.md).

## How it works

| Step | What it does |
| --- | --- |
| `TEXTSPLIT` | Splits each comma-separated list into a row of entries. |
| `VSTACK` | Puts the old values above the new values, forming a two-row lookup table. |
| `HLOOKUP(...,2,FALSE)` | Finds `value` in the first row and returns the entry beneath it. |
| `VALUE(...)` | Tries to turn the replacement into a number. |
| `IFERROR(...,HLOOKUP(...))` | If the number conversion fails, returns the replacement as text. |

`IFERROR` handles a failed number conversion only. If `value` is not found, the second lookup also fails, so `#N/A` is returned.

## Requirements

`LAMBDA`, `TEXTSPLIT`, `VSTACK`, `HLOOKUP`, `VALUE` and `IFERROR` must be available in your version of Excel.

## Design notes

- **Why it exists.** [`CODE_VALUES`](CODE_VALUES.md) gives sequential numbers. Some recoding needs other labels, or fewer categories, and nested `IF` formulas for that are long and repetitive.
- **Why one function for numbers and text.** Trying `VALUE` first lets the same function return numeric codes for analysis and text labels for grouping.

## See also

[`CODE_VALUES`](CODE_VALUES.md), [`HLOOKUP`](https://support.microsoft.com/en-us/excel/functions/hlookup-function), [`TEXTSPLIT`](https://support.microsoft.com/en-us/excel/functions/textsplit-function)
