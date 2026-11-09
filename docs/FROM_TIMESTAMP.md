# FROM_TIMESTAMP function

Converts a UNIX timestamp, in seconds, into a date and time written as `yyyy-mm-dd hh:mm:ss`.

## Syntax

```
=FROM_TIMESTAMP(timestamp)
```

## Arguments

| Argument | Required | Description |
| --- | --- | --- |
| `timestamp` | Yes | A UNIX timestamp: the number of seconds since 00:00:00 on 1 January 1970 (UTC). |

## Returns

Text in the format `yyyy-mm-dd hh:mm:ss`.

## Remarks

- The timestamp must be in seconds. If yours is in milliseconds, divide it by 1,000 first: `=FROM_TIMESTAMP(A2/1000)`.
- The result is in UTC. The function does not convert to a time zone.
- The result is text, not an Excel date-time value. To use it in date calculations, convert it first, for example with `DATEVALUE` and `TIMEVALUE`, and check the outcome under your regional settings.
- Negative timestamps are not supported.
- The format codes `yyyy-mm-dd` and `hh:mm:ss` are English-language codes. In versions of Excel that use other format codes, edit them in the definition.
- An empty cell is treated as 0 and returns `1970-01-01 00:00:00`.

## Errors

| Condition | Result |
| --- | --- |
| `timestamp` is text that cannot be read as a number | `#VALUE!` |

## Examples

| Formula | Result | Explanation |
| --- | --- | --- |
| `=FROM_TIMESTAMP(1704067200)` | `2024-01-01 00:00:00` | Midnight at the start of 1 January 2024. |
| `=FROM_TIMESTAMP(1704110400)` | `2024-01-01 12:00:00` | 12 hours (43,200 seconds) later. |
| `=FROM_TIMESTAMP(1704153599)` | `2024-01-01 23:59:59` | One second before midnight. |
| `=FROM_TIMESTAMP(1704240000)` | `2024-01-03 00:00:00` | Two days after the first example. |
| `=FROM_TIMESTAMP(3665)` | `1970-01-01 01:01:05` | 1 hour, 1 minute and 5 seconds after the start. |

To convert a column, enter `=FROM_TIMESTAMP(A2)` in `B2` and fill down. You can try it with this [sample dataset](https://drive.google.com/file/d/1hQ3vsOIt_esE72v09RaGu_BSGUu08qx8/view?usp=sharing).

## Definition

```
=LAMBDA(timestamp,TEXT(TRUNC(timestamp/86400)+DATE(1970,1,1),"yyyy-mm-dd")&" "&TEXT(TIME(,TRUNC(TRUNC(MOD(timestamp,86400))/60),MOD(MOD(timestamp,86400),60)),"hh:mm:ss"))
```

See [How to add a LAMBDA function to a workbook](how-to-add-a-lambda.md).

## How it works

There are 86,400 seconds in a day, so the formula splits the timestamp into whole days and the seconds left over.

| Part | Formula | What it does |
| --- | --- | --- |
| Date | `TRUNC(timestamp/86400)+DATE(1970,1,1)` | Counts the whole days and adds them to 1 January 1970. |
| Seconds left | `MOD(timestamp,86400)` | The seconds after the whole days. |
| Minutes | `TRUNC(MOD(timestamp,86400)/60)` | The complete minutes in those seconds. |
| Seconds | `MOD(MOD(timestamp,86400),60)` | The seconds after the complete minutes. |
| Time | `TIME(,minutes,seconds)` | Builds the time. The hour argument is empty on purpose, because `TIME` turns more than 59 minutes into hours. For example, `TIME(,61,5)` is 01:01:05. |
| Output | `TEXT(...)&" "&TEXT(...)` | Formats the date and time and joins them with a space. |

## Requirements

`LAMBDA` must be available in your version of Excel. The rest of the formula uses long-standing functions (`TEXT`, `TRUNC`, `DATE`, `MOD`, `TIME`).

## Design notes

- **Why it exists.** Data exported from applications, databases and APIs often stores date-times as integers. They need converting before any time-based analysis.
- **Why in Excel.** The same conversion is easy with Python's `datetime` module or pandas. This function does it with Excel formulas, working from the arithmetic of days, minutes and seconds.

## See also

[`DATE`](https://support.microsoft.com/en-us/excel/functions/date-function), [`TIME`](https://support.microsoft.com/en-us/excel/functions/time-function), [`TEXT`](https://support.microsoft.com/en-us/excel/functions/text-function), [`MOD`](https://support.microsoft.com/en-us/excel/functions/mod-function), [`TRUNC`](https://support.microsoft.com/en-us/excel/functions/trunc-function)
