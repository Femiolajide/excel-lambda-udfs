# CODE_VALUES function

Converts a text label into a numerical code. The codes are generated from the order of the labels you supply.

## Syntax

```
=CODE_VALUES(label, all_labels, start_no, na_value)
```

## Arguments

| Argument | Required | Description |
| --- | --- | --- |
| `label` | Yes | The cell or value to convert. |
| `all_labels` | Yes | All possible labels as one text string, separated by commas, in the order you want them coded. |
| `start_no` | Yes | The code given to the first label. Each later label gets the previous code plus 1. |
| `na_value` | Yes | The value returned when `label` is not found in `all_labels`. |

## Returns

A single value: the code for `label`, or `na_value` if the label is not found.

## Remarks

- The order of `all_labels` decides the codes. `Very Low,Low,High,Very High` and `Very High,High,Low,Very Low` give opposite codes.
- Do not put spaces around the commas. A space becomes part of the label, so `"Agree, Disagree"` contains the label `" Disagree"`, which does not match `Disagree`.
- Matching is exact, but not case-sensitive.
- Labels that contain commas are not supported.
- Only `#N/A` is replaced by `na_value`. If `label` is itself an error other than `#N/A`, that error is returned.
- Choose `na_value` with care. If `0` is also a valid code, an unmatched label cannot be told apart from a real `0`.
- To choose the replacement for each label yourself, use [`CUSTOM_CODE_VALUES`](CUSTOM_CODE_VALUES.md).

## Examples

Cell `A2` contains `Strongly Agree`.

| Formula | Result | Explanation |
| --- | --- | --- |
| `=CODE_VALUES(A2,"Strongly Disagree,Disagree,Agree,Strongly Agree",1,0)` | `4` | The fourth label gets code 4. |
| `=CODE_VALUES(A2,"Strongly Disagree,Disagree,Agree,Strongly Agree",0,0)` | `3` | Coding starts at 0. |
| `=CODE_VALUES(A2,"Strongly Disagree,Disagree,Agree,Strongly Agree",10,0)` | `13` | Coding starts at 10. |
| `=CODE_VALUES(A2,"Strongly Agree,Agree,Disagree,Strongly Disagree",1,0)` | `1` | A reversed list reverses the codes. |

### Labels that are not found

Cell `A2` contains `Not Applicable`.

| Formula | Result | Explanation |
| --- | --- | --- |
| `=CODE_VALUES(A2,"Strongly Disagree,Disagree,Agree,Strongly Agree",1,0)` | `0` | The label is not in the list, so `na_value` is returned. |
| `=CODE_VALUES(A2,"Strongly Disagree,Disagree,Agree,Strongly Agree",1,-1)` | `-1` | A different `na_value`. |

### A five-point scale

For a rating of `Very Good` in `A2`:

| Formula | Result |
| --- | --- |
| `=CODE_VALUES(A2,"Very Poor,Poor,Good,Very Good,Excellent",1,0)` | `4` |

## Definition

```
=LAMBDA(label,all_labels,start_no,na_value,IFNA(HLOOKUP(label,VSTACK(TEXTSPLIT(all_labels,","),SEQUENCE(,COUNTA(TEXTSPLIT(all_labels,",")),start_no)),2,FALSE),na_value))
```

See [How to add a LAMBDA function to a workbook](how-to-add-a-lambda.md).

## How it works

| Function | Role |
| --- | --- |
| `TEXTSPLIT` | Splits `all_labels` into a row of labels. |
| `COUNTA`, `SEQUENCE` | Count the labels and build a row of codes starting at `start_no`. |
| `VSTACK` | Puts the labels above the codes, forming a two-row lookup table. |
| `HLOOKUP` | Finds `label` in the first row and returns the code from the second. |
| `IFNA` | Returns `na_value` when the label is not found. |

## Requirements

`LAMBDA`, `TEXTSPLIT`, `SEQUENCE`, `VSTACK` and `HLOOKUP` must be available in your version of Excel.

## Design notes

- **Why it exists.** Survey data from online forms arrives as text, such as `Strongly Agree`, and needs numerical codes before analysis. Building a coding table and lookup formulas for every questionnaire is repetitive.
- **Why the coding is generated.** The code scheme is defined inside the formula, so the coding follows the order of labels you write and needs no table.

## See also

[`CUSTOM_CODE_VALUES`](CUSTOM_CODE_VALUES.md), [`UNIQUE_VALUES`](UNIQUE_VALUES.md), `HLOOKUP`
