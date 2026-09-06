# Essential Excel Functions with Real-World Examples

Yeh repository basic-to-intermediate Excel functions (`IF`, `AND`, `OR`, `COUNTIF`, `SUMIF`, `AVERAGEIF`, `COUNTA`, `IFERROR`) ka hands-on reference guide hai.

---

### Dummy Data Table Reference

| Row | A (Emp ID) | B (Name) | C (Department) | D (Score) | E (Sales in INR) |
| --- | --- | --- | --- | --- | --- |
| **2** | 101 | Aman | IT | 85 | 50000 |
| **3** | 102 | Priya | HR | 45 | 30000 |
| **4** | 103 | Rohit | IT | 92 | 75000 |
| **5** | 104 | Neha | Sales | 60 | 40000 |
| **6** | 105 | (Blank) | IT | 38 | 20000 |

---

### 1. IF

* **Purpose:** Di gayi condition true hone par ek value aur false hone par doosri value deta hai.
* **Syntax:** `=IF(logical_test, value_if_true, value_if_false)`
* **Example:** Check karein ki employee ne pass score (Score $\ge$ 50) clear kiya ya nahi:
```excel
=IF(D2>=50, "Pass", "Fail")

```


* **Output:** `Pass` (kyunki Row 2 ka score 85 hai).

---

### 2. AND

* **Purpose:** Jab saari conditions TRUE ho tabhi TRUE return karta hai; agar ek bhi galat hui toh FALSE.
* **Syntax:** `=AND(logical1, [logical2], ...)`
* **Example:** Check karein ki employee IT department ka ho aur Score 80 se upar ho:
```excel
=AND(C2="IT", D2>80)

```


* **Output:** `TRUE` (Dono conditions match hoti hain).

---

### 3. OR

* **Purpose:** Agar di gayi conditions mein se koi **ek** bhi TRUE ho, toh TRUE return karta hai.
* **Syntax:** `=OR(logical1, [logical2], ...)`
* **Example:** Check karein ki department ya toh HR ho ya Sales:
```excel
=OR(C2="HR", C2="Sales")

```


* **Output:** `FALSE` (Row 2 IT department se hai).

---

### 4. Nesting: IF with AND / OR

* **Purpose:** Multiple conditions ko evaluate karke custom result show karna.
* **Example:** Score $\ge$ 80 ho **AUR** Sales > 40000 ho toh "Eligible for Bonus", warna "Not Eligible":
```excel
=IF(AND(D2>=80, E2>40000), "Eligible for Bonus", "Not Eligible")

```


* **Output:** `Eligible for Bonus`

---

### 5. COUNTA

* **Purpose:** Range ke andar un sabhi cells ko count karta hai jo khali (empty) nahi hain (text, numbers, symbols sab count karta hai).
* **Syntax:** `=COUNTA(value1, [value2], ...)`
* **Example:** Total registered employees count karna jinka name blank na ho:
```excel
=COUNTA(B2:B6)

```


* **Output:** `4` (Row 6 blank hai, isliye sirf 4 count honge).

---

### 6. COUNTIF

* **Purpose:** Kisi specific criteria ya condition ke aadhar par cells ko count karta hai.
* **Syntax:** `=COUNTIF(range, criteria)`
* **Example:** IT department mein total kitne records hain:
```excel
=COUNTIF(C2:C6, "IT")

```


* **Output:** `3`

---

### 7. SUMIF

* **Purpose:** Di gayi condition match hone par corresponding cells ka sum nikalta hai.
* **Syntax:** `=SUMIF(range, criteria, [sum_range])`
* **Example:** Sirf "IT" department ki total sales ka total:
```excel
=SUMIF(C2:C6, "IT", E2:E6)

```


* **Output:** `145000` ($50000 + 75000 + 20000$)

---

### 8. AVERAGEIF

* **Purpose:** Specific condition match hone wale cells ka average nikalta hai.
* **Syntax:** `=AVERAGEIF(range, criteria, [average_range])`
* **Example:** Sirf "IT" department ka average score calculate karna:
```excel
=AVERAGEIF(C2:C6, "IT", D2:D6)

```


* **Output:** `71.67` ($(85 + 92 + 38) / 3$)

---

### 9. IFERROR

* **Purpose:** Formula mein koi error (`#DIV/0!`, `#N/A`, `#VALUE!`) aane par error message ki jagah custom text ya clean value show karta hai.
* **Syntax:** `=IFERROR(value, value_if_error)`
* **Example:** Score ko Zero se divide karne par aane wale `#DIV/0!` error ko handle karna:
```excel
=IFERROR(D2/0, "Division Not Possible")

```


* **Output:** `Division Not Possible` (bina kisi system crash ya red flag ke).
