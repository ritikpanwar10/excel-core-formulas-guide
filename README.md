# Essential Excel Functions with Real-World Examples

A comprehensive reference guide for essential Excel functions—ranging from logical evaluations (`IF`, `AND`, `OR`) and aggregation criteria (`COUNTIF`, `SUMIF`, `AVERAGEIF`, `COUNTA`) to error handling (`IFERROR`) 

---

## Sample Reference Dataset

The examples below reference the following dummy employee dataset:
<img width="496" height="235" alt="p" src="https://github.com/user-attachments/assets/bd77bb25-165d-4ea0-9c00-33e14b9f220d" />

---

## 1. IF

* **Purpose:** Returns one value if a condition evaluates to `TRUE`, and another value if it evaluates to `FALSE`.
* **Syntax:**
  ```excel
  =IF(logical_test, value_if_true, value_if_false)
  ```
* **Example:** Check whether an employee has achieved a passing score ($\ge 50$):
 <img width="562" height="251" alt="o" src="https://github.com/user-attachments/assets/e81587ba-b650-4ab6-8d8d-a73d1fad06d6" />

* **Output:** `"Pass"` (since the score in cell D2 is 85).

---

## 2. AND

* **Purpose:** Returns `TRUE` only if all specified conditions evaluate to `TRUE`; returns `FALSE` if any single condition is not met.
* **Syntax:**
  ```excel
  =AND(logical1, [logical2], ...)
  ```
* **Example:** Verify whether an employee belongs to the "IT" department and has a Score greater than 80:
  <img width="583" height="301" alt="q" src="https://github.com/user-attachments/assets/3469c1c0-bfef-44c5-96a0-e76ffa7a317a" />

* **Output:** `TRUE` (both conditions are satisfied for Row 2).

---

## 3. OR

* **Purpose:** Returns `TRUE` if at least one of the specified arguments evaluates to `TRUE`.
* **Syntax:**
  ```excel
  =OR(logical1, [logical2], ...)
  ```
* **Example:** Check whether the department is either "HR" or "Sales":
 <img width="546" height="273" alt="1" src="https://github.com/user-attachments/assets/cd9ed026-aeb7-4133-8ea5-44d861eb609a" />

* **Output:** `FALSE` (cell C2 is "IT").

---

## 4. Nested Logic: IF with AND / OR

* **Purpose:** Combines multiple logical checks to return customized outputs.
* **Example:** If Score is 80 **and** Sales are > 40,000$, display "Eligible for Bonus"; otherwise, display "Not Eligible":
  <img width="739" height="270" alt="2" src="https://github.com/user-attachments/assets/299042af-8838-4fdb-97ee-ad287d1f9122" />

* **Output:** `"Eligible for Bonus"`

---

## 5. COUNTA

* **Purpose:** Counts all non-empty cells within a given range (including text, numbers, dates, booleans, and formula errors).
* **Syntax:**
  ```excel
  =COUNTA(value1, [value2], ...)
  ```
* **Example:** Count total registered employees where the name field is not empty:
<img width="482" height="287" alt="3" src="https://github.com/user-attachments/assets/63ac0c75-23c2-464d-98a6-70d7d3db6d6b" />

* **Output:** `4` (Row 6 is blank, so only 4 cells are counted).

---

## 6. COUNTIF

* **Purpose:** Counts the number of cells that meet a single specific condition.
* **Syntax:**
  ```excel
  =COUNTIF(range, criteria)
  ```
* **Example:** Count the total number of records for the "IT" department:
 <img width="478" height="294" alt="4" src="https://github.com/user-attachments/assets/25b9bc76-5755-47cc-b79f-c45dece3094f" />

* **Output:** `3`

---

## 7. SUMIF

* **Purpose:** Sums values in a range that satisfy a given criterion.
* **Syntax:**
  ```excel
  =SUMIF(range, criteria, [sum_range])
  ```
* **Example:** Calculate total sales generated exclusively by the "IT" department:
 <img width="488" height="295" alt="5" src="https://github.com/user-attachments/assets/0d2b225c-c364-41c2-909d-af3deb501b8b" />

* **Calculation:** $50000 + 20000 + 75000$
* **Output:** `145000`

---

## 8. AVERAGEIF

* **Purpose:** Computes the arithmetic mean of all cells that satisfy a specific condition.
* **Syntax:**
  ```excel
  =AVERAGEIF(range, criteria, [average_range])
  ```
* **Example:** Calculate the average score for the "IT" department:
 <img width="528" height="302" alt="6" src="https://github.com/user-attachments/assets/a4d84777-e20c-4408-aad5-f6044dce8ea3" />

* **Calculation:** $\frac{85 + 38 + 92}{3} \approx 71.67$
* **Output:** `71.67`

---

## 9. IFERROR

* **Purpose:** Traps standard Excel errors (`#DIV/0!`, `#N/A`, `#VALUE!`, `#REF!`) and displays a fallback value or clean message.
* **Syntax:**
  ```excel
  =IFERROR(value, value_if_error)
  ```
* **Example:** Suppress division-by-zero errors when dividing score by zero:
  ```excel
  =IFERROR(D2/0, "Division Not Possible")
  ```
* **Output:** `"Division Not Possible"`

---



