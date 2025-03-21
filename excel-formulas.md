## Excel Formulas Cheat Sheet

### Logical Functions
| Formula       | Description                                                                 | Usage Example                                                                 |
|---------------|-----------------------------------------------------------------------------|-------------------------------------------------------------------------------|
| `AND()`       | Returns TRUE if all arguments are TRUE.                                     | `=AND(A1>10, B1<20)`                                                          |
| `OR()`        | Returns TRUE if any argument is TRUE.                                       | `=OR(A1="Yes", B1="No")`                                                      |
| `NOT()`       | Reverses the logical value of its argument.                                 | `=NOT(A1=TRUE)`                                                               |
| `XOR()`       | Returns TRUE for an odd number of TRUE arguments.                           | `=XOR(A1>5, B1<10)`                                                           |
| `IF()`        | Returns one value if TRUE, another if FALSE.                                | `=IF(A1>50, "Pass", "Fail")`                                                  |
| `IFERROR()`   | Returns a custom value if the formula results in an error.                  | `=IFERROR(A1/B1, "Error")`                                                    |
| `IFNA()`      | Returns a custom value if the formula results in `#N/A`.                    | `=IFNA(VLOOKUP(A1, B:C, 2, 0), "Not Found")`                                  |

---

### IS Functions
| Formula          | Description                                                                 | Usage Example                                                                 |
|------------------|-----------------------------------------------------------------------------|-------------------------------------------------------------------------------|
| `ISBLANK()`      | Checks if a cell is empty.                                                  | `=ISBLANK(A1)`                                                                |
| `ISERR()`        | Checks if the value is an error (excludes `#N/A`).                         | `=ISERR(A1)`                                                                  |
| `ISERROR()`      | Checks if the value is any error.                                           | `=ISERROR(A1)`                                                                |
| `ISEVEN()`       | Checks if a number is even.                                                 | `=ISEVEN(4)` → `TRUE`                                                         |
| `ISODD()`        | Checks if a number is odd.                                                  | `=ISODD(5)` → `TRUE`                                                          |
| `ISFORMULA()`    | Checks if a cell contains a formula.                                        | `=ISFORMULA(A1)`                                                              |
| `ISLOGICAL()`    | Checks if the value is a logical (TRUE/FALSE).                              | `=ISLOGICAL(A1)`                                                              |
| `ISNA()`         | Checks if the value is `#N/A`.                                              | `=ISNA(A1)`                                                                   |
| `ISNUMBER()`     | Checks if the value is a number.                                            | `=ISNUMBER(A1)`                                                               |
| `ISREF()`        | Checks if the value is a valid cell reference.                              | `=ISREF(A1)`                                                                  |
| `ISTEXT()`       | Checks if the value is text.                                                 | `=ISTEXT(A1)`                                                                 |
| `ISNONTEXT()`    | Checks if the value is not text.                                             | `=ISNONTEXT(A1)`                                                              |

---

### Conditional Functions
| Formula          | Description                                                                 | Usage Example                                                                 |
|------------------|-----------------------------------------------------------------------------|-------------------------------------------------------------------------------|
| `COUNTIF()`      | Counts cells that meet a single condition.                                  | `=COUNTIF(A1:A10, ">20")`                                                     |
| `SUMIF()`        | Sums cells that meet a single condition.                                    | `=SUMIF(A1:A10, "Apple", B1:B10)`                                             |
| `AVERAGEIF()`    | Averages cells that meet a single condition.                                | `=AVERAGEIF(A1:A10, ">50")`                                                   |
| `COUNTIFS()`     | Counts cells that meet multiple conditions.                                 | `=COUNTIFS(A1:A10, ">20", B1:B10, "<100")`                                    |
| `SUMIFS()`       | Sums cells that meet multiple conditions.                                   | `=SUMIFS(C1:C10, A1:A10, "Apple", B1:B10, ">10")`                             |
| `AVERAGEIFS()`   | Averages cells that meet multiple conditions.                               | `=AVERAGEIFS(C1:C10, A1:A10, "Apple", B1:B10, ">10")`                         |

---

### Mathematical Functions
| Formula          | Description                                                                 | Usage Example                                                                 |
|------------------|-----------------------------------------------------------------------------|-------------------------------------------------------------------------------|
| `SUM()`          | Adds numbers.                                                               | `=SUM(A1:A10)`                                                                |
| `AVERAGE()`      | Calculates the average of numbers.                                          | `=AVERAGE(A1:A10)`                                                            |
| `AVERAGEA()`     | Averages numbers, treating text/TRUE as 1.                                  | `=AVERAGEA(A1:A10)`                                                           |
| `COUNT()`        | Counts cells containing numbers.                                            | `=COUNT(A1:A10)`                                                              |
| `COUNTA()`       | Counts non-empty cells.                                                     | `=COUNTA(A1:A10)`                                                             |
| `MEDIAN()`       | Finds the middle value in a dataset.                                        | `=MEDIAN(A1:A10)`                                                             |
| `SUMPRODUCT()`   | Multiplies arrays and returns the sum of products.                          | `=SUMPRODUCT(A1:A10, B1:B10)`                                                 |
| `SUMSQ()`        | Returns the sum of squares of arguments.                                    | `=SUMSQ(2,3)` → `13`                                                         |
| `COUNTBLANK()`   | Counts empty cells in a range.                                              | `=COUNTBLANK(A1:A10)`                                                         |
| `EVEN()`         | Rounds a number up to the nearest even integer.                             | `=EVEN(3.2)` → `4`                                                            |
| `ODD()`          | Rounds a number up to the nearest odd integer.                              | `=ODD(2.1)` → `3`                                                             |
| `INT()`          | Rounds a number down to the nearest integer.                                | `=INT(5.9)` → `5`                                                             |
| `LARGE()`        | Returns the k-th largest value in a dataset.                                | `=LARGE(A1:A10, 2)`                                                           |
| `SMALL()`        | Returns the k-th smallest value in a dataset.                               | `=SMALL(A1:A10, 2)`                                                            |
| `MAX()`          | Returns the largest value in a range.                                       | `=MAX(A1:A10)`                                                                |
| `MAXA()`         | Returns the largest value, including text/TRUE as 1.                        | `=MAXA(A1:A10)`                                                               |
| `MIN()`          | Returns the smallest value in a range.                                      | `=MIN(A1:A10)`                                                                |
| `MINA()`         | Returns the smallest value, including text/TRUE as 1.                       | `=MINA(A1:A10)`                                                               |
| `MOD()`          | Returns the remainder after division.                                       | `=MOD(10, 3)` → `1`                                                           |
| `RAND()`         | Generates a random number between 0 and 1.                                  | `=RAND()`                                                                     |
| `RANDBETWEEN()`  | Generates a random integer between two numbers.                             | `=RANDBETWEEN(1, 100)`                                                        |
| `SQRT()`         | Returns the square root of a number.                                        | `=SQRT(25)` → `5`                                                             |
| `SUBTOTAL()`     | Performs a calculation (e.g., SUM, AVERAGE) while ignoring hidden rows.     | `=SUBTOTAL(9, A1:A10)` (9 = SUM)                                              |

---

### Find & Search Functions
| Formula          | Description                                                                 | Usage Example                                                                 |
|------------------|-----------------------------------------------------------------------------|-------------------------------------------------------------------------------|
| `FIND()`         | Returns the position of text (case-sensitive).                              | `=FIND("n", "Apple")` → `4`                                                   |
| `SEARCH()`       | Returns the position of text (case-insensitive).                            | `=SEARCH("p", "Apple")` → `1`                                                 |
| `SUBSTITUTE()`   | Replaces specific text in a string.                                         | `=SUBSTITUTE("Hello World", "World", "Excel")` → `"Hello Excel"`               |
| `REPLACE()`      | Replaces text at a specific position.                                       | `=REPLACE("ABCDEF", 2, 3, "-")` → `"A-EF"`                                    |

---

### Lookup Functions
| Formula          | Description                                                                 | Usage Example                                                                 |
|------------------|-----------------------------------------------------------------------------|-------------------------------------------------------------------------------|
| `MATCH()`        | Returns the position of a value in a range.                                 | `=MATCH("Apple", A1:A10, 0)`                                                  |
| `LOOKUP()`       | Searches for a value in a range and returns a corresponding value.          | `=LOOKUP("Apple", A1:A10, B1:B10)`                                            |
| `HLOOKUP()`      | Searches horizontally for a value and returns a match.                      | `=HLOOKUP("Q1", A1:D4, 3, FALSE)`                                             |
| `VLOOKUP()`      | Searches vertically for a value and returns a match.                        | `=VLOOKUP("Apple", A1:B10, 2, FALSE)`                                         |

---

### Reference Functions
| Formula          | Description                                                                 | Usage Example                                                                 |
|------------------|-----------------------------------------------------------------------------|-------------------------------------------------------------------------------|
| `ADDRESS()`      | Creates a cell address from row and column numbers.                         | `=ADDRESS(3, 4)` → `"$D$3"`                                                  |
| `CHOOSE()`       | Returns a value from a list based on position.                              | `=CHOOSE(2, "Apple", "Banana")` → `"Banana"`                                 |
| `INDEX()`        | Returns a value from a specific position in a range.                        | `=INDEX(A1:C10, 2, 3)` → Value at row 2, column 3                             |
| `INDIRECT()`     | Converts a text string into a cell reference.                               | `=INDIRECT("A" & 5)` → Value of cell A5                                       |
| `OFFSET()`       | Returns a reference offset from a starting cell.                            | `=OFFSET(A1, 2, 3)` → Value 2 rows down and 3 columns right from A1           |

---

### Date & Time Functions
| Formula          | Description                                                                 | Usage Example                                                                 |
|------------------|-----------------------------------------------------------------------------|-------------------------------------------------------------------------------|
| `DATE()`         | Creates a date from year, month, day.                                       | `=DATE(2023, 12, 25)` → `12/25/2023`                                         |
| `DATEVALUE()`    | Converts a date string to a serial number.                                  | `=DATEVALUE("2023-12-25")` → Excel date code                                  |
| `TIME()`         | Creates a time from hours, minutes, seconds.                               | `=TIME(14, 30, 0)` → `2:30 PM`                                               |
| `TIMEVALUE()`    | Converts a time string to a serial number.                                  | `=TIMEVALUE("14:30")` → Excel time code                                       |
| `NOW()`          | Returns the current date and time.                                          | `=NOW()` → `2023-10-05 15:30`                                                |
| `TODAY()`        | Returns the current date.                                                   | `=TODAY()` → `2023-10-05`                                                    |
| `YEAR()`         | Extracts the year from a date.                                              | `=YEAR(A1)` → `2023`                                                         |
| `MONTH()`        | Extracts the month from a date.                                             | `=MONTH(A1)` → `12`                                                          |
| `DAY()`          | Extracts the day from a date.                                               | `=DAY(A1)` → `25`                                                            |
| `HOUR()`         | Extracts the hour from a time.                                              | `=HOUR(A1)` → `14`                                                           |
| `MINUTE()`       | Extracts the minute from a time.                                            | `=MINUTE(A1)` → `30`                                                         |
| `SECOND()`       | Extracts the second from a time.                                            | `=SECOND(A1)` → `0`                                                          |
| `WEEKDAY()`      | Returns the day of the week (1-7).                                          | `=WEEKDAY(A1)` → `1` (Sunday)                                                |
| `DAYS()`         | Returns the number of days between two dates.                               | `=DAYS("2023-12-31", "2023-12-25")` → `6`                                    |
| `NETWORKDAYS()`  | Returns workdays between two dates, excluding weekends.                     | `=NETWORKDAYS(A1, B1)`                                                       |
| `WORKDAY()`      | Adds workdays to a date, excluding weekends.                               | `=WORKDAY(A1, 5)`                                                            |

---

### Text Functions (Misc.)
| Formula          | Description                                                                 | Usage Example                                                                 |
|------------------|-----------------------------------------------------------------------------|-------------------------------------------------------------------------------|
| `CHAR()`         | Returns a character based on ASCII code.                                    | `=CHAR(65)` → `"A"`                                                          |
| `CODE()`         | Returns the ASCII code of a character.                                      | `=CODE("A")` → `65`                                                          |
| `CLEAN()`        | Removes non-printable characters from text.                                 | `=CLEAN(A1)`                                                                 |
| `TRIM()`         | Removes extra spaces from text.                                             | `=TRIM("  Excel  ")` → `"Excel"`                                             |
| `LEN()`          | Returns the length of text.                                                 | `=LEN("Excel")` → `5`                                                        |
| `EXACT()`        | Checks if two text strings are identical (case-sensitive).                  | `=EXACT("Apple", "apple")` → `FALSE`                                         |
| `FORMULATEXT()`  | Returns the formula in a cell as text.                                      | `=FORMULATEXT(A1)`                                                           |
| `LEFT()`         | Extracts leftmost characters from text.                                     | `=LEFT("Excel", 2)` → `"Ex"`                                                 |
| `RIGHT()`        | Extracts rightmost characters from text.                                    | `=RIGHT("Excel", 3)` → `"cel"`                                               |
| `MID()`          | Extracts characters from the middle of text.                                | `=MID("Excel", 2, 3)` → `"xce"`                                              |
| `LOWER()`        | Converts text to lowercase.                                                 | `=LOWER("EXCEL")` → `"excel"`                                                |
| `PROPER()`       | Converts text to title case.                                                | `=PROPER("excel")` → `"Excel"`                                               |
| `UPPER()`        | Converts text to uppercase.                                                 | `=UPPER("excel")` → `"EXCEL"`                                                |
| `REPT()`         | Repeats text a specified number of times.                                   | `=REPT("*", 5)` → `"*****"`                                                  |
| `VALUE()`        | Converts text to a number.                                                  | `=VALUE("123")` → `123`                                                      |

---

### Rank Functions
| Formula          | Description                                                                 | Usage Example                                                                 |
|------------------|-----------------------------------------------------------------------------|-------------------------------------------------------------------------------|
| `RANK()`         | Returns the rank of a number in a list (legacy).                            | `=RANK(85, A1:A10)`                                                          |
| `RANK.AVG()`     | Returns the average rank for tied numbers.                                  | `=RANK.AVG(85, A1:A10)`                                                      |
| `RANK.EQ()`      | Returns the top rank for tied numbers.                                      | `=RANK.EQ(85, A1:A10)`                                                       |