<div align="center">

# 📊 PR. 1 Fundamental Booster — Excel Essentials

### A practical Excel project covering formulas, lookups, conditional logic, text functions, date calculations, dynamic references, and data analysis.

**Prepared by:** Priyanshu Buhecha

<p>
  <img src="https://img.shields.io/badge/Microsoft%20Excel-Fundamental%20Booster-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white" alt="Microsoft Excel" />
  <img src="https://img.shields.io/badge/Project-PR.%201-1f6feb?style=for-the-badge" alt="Project" />
  <img src="https://img.shields.io/badge/Tasks-Essential%20Tasks-f5a623?style=for-the-badge" alt="Essential Tasks" />
  <img src="https://img.shields.io/badge/Status-Completed-2ea44f?style=for-the-badge" alt="Completed" />
</p>

<p><b>📘 Workbook:</b> PR. 1 Fundamental Booster &nbsp;•&nbsp; <b>📈 Focus:</b> Excel Formula & Data Analysis Skills</p>

</div>

---

## 📌 Project Overview

This project is an **Excel Fundamental Booster** practical workbook created to strengthen core spreadsheet and data-analysis skills through hands-on tasks.

The workbook focuses on using Excel formulas to solve realistic problems involving **student grades, sales records, and employee information**. Each essential task is documented below with the **original question number, the formula used, an explanation of the formula, and the final result** from the completed workbook.

### 🎥 Project Explanation Video

Watch the detailed walkthrough of this project:  
👉 **[Google Drive Video Link](https://drive.google.com/file/d/14oVZnh6JNXObl8Ig4DC08Xea8M7Fmt-7/view?usp=sharing)**

### 🎯 Main Objectives

- Build confidence with **relative and absolute cell references**.
- Apply **IF, AND, OR, and nested IF** logic.
- Perform conditional calculations with **COUNTIF/COUNTIFS, SUMIFS, and AVERAGEIF/AVERAGEIFS**.
- Use powerful lookup tools such as **VLOOKUP, INDEX/MATCH, XLOOKUP, and XMATCH**.
- Work with **TEXT functions** for practical data cleaning.
- Create flexible formulas using **INDIRECT and OFFSET**.
- Perform **date and time calculations**.
- Apply **ROUND, FLOOR, and CEILING** for numerical and financial analysis.
- Return multiple matching records with the **FILTER** function.

---

## 🗂️ Workbook Structure

| 📄 Sheet | 🔎 Purpose |
|:--|:--|
| **Student Grade** | Student IDs, names, marks, enrollment dates, grading logic, percentages, lookups, rounding, and FILTER analysis |
| **Sales Data** | Sales IDs, products, regions, salespersons, amounts, dates, discounts, conditional logic, aggregations, INDEX/MATCH, XMATCH, INDIRECT, OFFSET, and date-based analysis |
| **Employee Data** | Employee IDs, names, departments, salaries, joining dates, text manipulation, date calculations, XLOOKUP, and dynamic lookup analysis |

---

## 🧭 Essential Tasks — Quick Navigation

| # | Essential Task | Main Functions / Skills |
|:--:|:--|:--|
| 1 | Relative & Absolute Cell References, Formatting, and Data Input | Relative / Absolute References, Formatting |
| 2 | Using IF Formula & Nesting IF Formulas | `IF`, Nested `IF` |
| 3 | Using IF with AND/OR Formulas | `IF`, `AND`, `OR` |
| 4 | COUNTIFS, SUMIFS, AVERAGEIFS | `COUNTIF`, `SUMIFS`, `AVERAGEIF` |
| 5 | VLOOKUP Function | `VLOOKUP` |
| 6 | INDEX and MATCH Functions | `INDEX`, `MATCH`, `SUMIFS` |
| 7 | Excel TEXT Functions | `LEFT`, `FIND`, `UPPER` |
| 8 | XLOOKUP Function | `XLOOKUP` |
| 9 | XMATCH Function | `XMATCH` |
| 10 | INDIRECT Function | `INDIRECT`, `SUM` |
| 11 | OFFSET Function | `OFFSET`, `SUM` |
| 12 | Date and Time Functions | `DATEDIF`, `TODAY` |
| 13 | Math Functions | `ROUND`, `FLOOR`, `CEILING` |
| 14 | Return Multiple Values using FILTER Function | `FILTER` |

> **💡 Reading this README:** Every question follows the same pattern — **Task → Formula → Explanation → Final Answer**.

---

# ✅ Essential Tasks — Detailed Work

## 1. 📐 Relative & Absolute Cell References, Formatting, and Data Input

### 🔹 Question 7 — Enter and format data

**Task**  
Enter and format the required data using appropriate number, currency, date, and text formatting.

**Implementation**

The workbook contains structured **Student, Sales, and Employee** datasets with formatted headers and appropriate data types for marks, amounts, dates, salaries, IDs, and text fields.

**Key examples**

- Student marks are stored as numeric values.
- Enrollment and joining dates are stored as Excel dates.
- Sales amounts and salaries are stored as numeric financial values.
- Headers identify each dataset and calculation area.

**🏁 Final Result:** The three worksheets contain organized, analysis-ready data.

---

### 🔹 Question 8 — Use absolute and relative references

**Formula**

```excel
=E2*$H$2
```

**Explanation**

- `E2` is a **relative reference** and changes when the formula is copied down.
- `$H$2` is an **absolute reference** and always points to the same discount-rate cell.
- This allows one fixed rate to be applied to every sales amount.

**🏁 Final Answer:** The discount column calculates **10% of each sales amount** using the fixed rate in `H2`.

---

## 2. 🔀 Using IF Formula & Nesting IF Formulas

### 🔹 Question 11 — Classify students' grades based on their scores

**Formula**

```excel
=IF(C2>=90,"A",IF(C2>=75,"B",IF(C2>=50,"C","F")))
```

**Explanation**

The nested `IF` checks the Math score against multiple thresholds:

| Score | Grade |
|:--:|:--:|
| ≥ 90 | A |
| 75–89 | B |
| 50–74 | C |
| < 50 | F |

**🏁 Final Answer Example:** Student 1 has a Math score of **61**, so the result is **C**.

---

### 🔹 Question 12 — Compute discounts for sales based on price thresholds

**Formula**

```excel
=E2*$H$2
```

**Explanation**

The sales amount in `E2` is multiplied by the fixed discount rate stored in `$H$2`.

For example:

```text
7048 × 10% = 704.80
```

**🏁 Final Answer:** **704.80** discount for Sales ID **2001**.

---

## 3. 🧠 Using IF with AND/OR Formulas

### 🔹 Question 15 — Identify students who scored above 80 in both Math & Science

**Formula**

```excel
=IF(AND(C2>80,D2>80),"Above","Below")
```

**Explanation**

`AND` requires both conditions to be TRUE:

- Math score > 80
- Science score > 80

If both conditions are satisfied, the formula returns **Above**; otherwise it returns **Below**.

**🏁 Final Answer Example:** Student 13 has **Math = 98** and **Science = 90**, so the result is **Above**.

---

### 🔹 Question 16 — Determine product eligibility for discount using IF with OR

**Formula**

```excel
=IF(OR(B2="Laptop",E2>20000),"Eligible","Not Eligible")
```

**Explanation**

The `OR` function returns TRUE when at least one condition is met:

- Product is **Laptop**, or
- Sales amount is greater than **20,000**.

**🏁 Final Answer Example:** Sales ID **2003** has an amount of **23,672**, so it is **Eligible**.

---

## 4. 🧮 COUNTIFS, SUMIFS, AVERAGEIFS

### 🔹 Question 19 — Count students who scored above 50 in Math

**Formula used in the completed workbook**

```excel
=COUNTIF(C2:C21,">50")
```

**Explanation**

`COUNTIF` counts the number of Math scores in `C2:C21` that are greater than 50.

**🏁 Final Answer:** **19 students** scored above 50 in Math.

> **📝 Note:** Although this topic is grouped under COUNTIFS/SUMIFS/AVERAGEIFS in the task sheet, the completed workbook uses `COUNTIF` because only one condition is required for this question.

---

### 🔹 Question 20 — Sum total sales for a specific region and product

**Formula**

```excel
=SUMIFS(E2:E21,C2:C21,"East",B2:B21,"Keyboard")
```

**Explanation**

`SUMIFS` adds values from `E2:E21` only when **both** conditions are satisfied:

- Region = **East**
- Product = **Keyboard**

**🏁 Final Answer:** **80,349**

---

### 🔹 Question 21 — Calculate the average Math score for students who scored above 60

**Formula**

```excel
=AVERAGEIF(C2:C21,">60")
```

**Explanation**

`AVERAGEIF` calculates the average of Math scores in `C2:C21` only for students whose Math score is greater than 60.

**🏁 Final Answer:** **82.26315789** (approximately **82.26**)

---

## 5. 🔎 VLOOKUP Function

### 🔹 Question 24 — Fetch student names based on ID

**Formula**

```excel
=VLOOKUP(1003,A2:F21,2,FALSE)
```

**Explanation**

`VLOOKUP` searches for Student ID `1003` in the first column of `A2:F21` and returns the matching student name from the second column.

`FALSE` requests an **exact match**.

**🏁 Final Answer:** **Student 3**

---

### 🔹 Question 25 — Retrieve product prices based on product codes

**Implementation**

The completed **Sales Data** worksheet organizes product identifiers and sales amounts for lookup-based analysis. A separate Q25 formula is not stored as a standalone labeled answer cell in the inspected PR 1 workbook.

**🏁 Final Result:** Product and sales data are available in the Sales Data sheet for lookup operations.

---

## 6. 🧩 INDEX and MATCH Functions

### 🔹 Question 28 — Find sales values for a salesperson in a specific month

**Formula**

```excel
=SUMIFS(E2:E21,D2:D21,"Person 13",F2:F21,">=2023-06-01",F2:F21,"<=2023-06-30")
```

**Explanation**

`SUMIFS` adds sales values from column `E` when:

- Salesperson = **Person 13**
- Date is between **June 1, 2023** and **June 30, 2023**

**🏁 Final Answer:** **15,560**

### 🔹 INDEX/MATCH demonstration

**Formula**

```excel
=INDEX(A2:A21,MATCH("South",C2:C21,0))
```

**Explanation**

- `MATCH` finds the first row where the Region is **South**.
- `INDEX` returns the corresponding Sales ID from column A.
- `0` tells `MATCH` to use an exact match.

**🏁 Final Answer:** **2001**

---

### 🔹 Question 29 — Look up employee details dynamically

**Formula**

```excel
=XLOOKUP(A6,A2:A21,B2:B21)
```

**Explanation**

`XLOOKUP` uses the employee ID in `A6` as the lookup value, searches `A2:A21`, and returns the corresponding employee name from `B2:B21`.

**🏁 Final Answer:** **Employee 5**

---

## 7. ✍️ Excel TEXT Functions

### 🔹 Question 32 — Extract the first name using LEFT and FIND

**Formula**

```excel
=LEFT(B2,FIND(" ",B2)-1)
```

**Explanation**

- `FIND(" ",B2)` locates the first space in the full name.
- `LEFT` extracts all characters before that space.

**🏁 Final Answer:** **Employee**

---

### 🔹 Question 33 — Convert names to uppercase/lowercase

**Formula**

```excel
=UPPER(B2)
```

**Explanation**

`UPPER` converts every letter in the selected text to uppercase.

**🏁 Final Answer:** **EMPLOYEE 1**

---

## 8. 🔍 XLOOKUP Function

### 🔹 Question 36 — Search for salesperson performance using flexible criteria

**Implementation**

The **Sales Data** worksheet contains salesperson, region, amount, and date fields prepared for flexible lookup and performance analysis. A separate Q36 formula is not stored as a standalone labeled answer in the inspected PR 1 workbook.

**🏁 Final Result:** Salesperson performance data are available in the Sales Data sheet for flexible lookup analysis.

---

### 🔹 Question 37 — Find employee salaries without worrying about sorted data

**Formula**

```excel
=XLOOKUP("Employee 2",B2:B21,D2:D21)
```

**Explanation**

`XLOOKUP` searches for **Employee 2** in the Name column and returns the corresponding salary from the Salary column.

**🏁 Final Answer:** **31,943**

---

## 9. 📍 XMATCH Function

### 🔹 Question 40 — Find the position of a product in the sales list

**Formula 1**

```excel
=XMATCH("Mouse",B2:B21)
```

**🏁 Final Answer:** **3**

**Formula 2**

```excel
=XMATCH("Keyboard",B2:B21)
```

**🏁 Final Answer:** **1**

**Explanation**

`XMATCH` returns the relative position of the requested product inside the lookup range.

---

## 10. 🔗 INDIRECT Function

### 🔹 Question 43 — Dynamically reference a cell range

**Formula**

```excel
=SUM(INDIRECT("E2:E21"))
```

**Explanation**

`INDIRECT` converts the text `"E2:E21"` into a real Excel reference, which can then be passed to `SUM`.

**🏁 Final Answer:** **507,725** total sales.

---

## 11. ↔️ OFFSET Function

### 🔹 Question 46 — Create dynamic ranges for sales trends

**Formula**

```excel
=SUM(OFFSET(E1,1,0,3,1))
```

**Explanation**

`OFFSET` starts from `E1`, moves down one row, and creates a range with:

- **Height:** 3 rows
- **Width:** 1 column

The resulting range is `E2:E4`, and `SUM` adds those values.

**🏁 Final Answer:** **48,148**

---

## 12. 📅 Date and Time Functions

### 🔹 Question 49 — Calculate age / year difference from a date

**Formula**

```excel
=DATEDIF(E2,TODAY(),"Y")
```

**Explanation**

`DATEDIF` calculates the completed number of years between the joining date in `E2` and today's date.

**🏁 Final Answer:** **11 years** for the displayed workbook result.

---

### 🔹 Question 50 — Find the difference in days between two dates

**Formula**

```excel
=TODAY()-E3
```

**Explanation**

Excel stores dates as serial numbers, so subtracting the earlier date from `TODAY()` returns the number of elapsed days.

**🏁 Final Answer:** **4,005 days** for the displayed workbook result.

---

## 13. 🔢 Math Functions

### 🔹 Question 53 — Use ROUND, CEILING, and FLOOR for financial calculations

#### ROUND

**Formula**

```excel
=ROUND(H3,2)
```

**Explanation**

Rounds the percentage in `H3` to **two decimal places**.

**🏁 Final Answer:** **74.67**

#### FLOOR

**Formula**

```excel
=FLOOR(H8,5)
```

**Explanation**

Rounds the value in `H8` **down** to the nearest multiple of 5.

**🏁 Final Answer:** **65**

#### CEILING

**Formula**

```excel
=CEILING(H10,5)
```

**Explanation**

Rounds the value in `H10` **up** to the nearest multiple of 5.

**🏁 Final Answer:** **60**

---

## 14. 🧹 Return Multiple Values using FILTER Function

### 🔹 Question 56 — Extract top-performing students based on scores

**Formula**

```excel
=FILTER(A2:C6,C2:C6>80,"No Records")
```

**Explanation**

`FILTER` returns all rows from `A2:C6` where the Math score in `C2:C6` is greater than 80.

If no matching record exists, the formula returns **No Records**.

**🏁 Final Answer:** The filtered result includes the qualifying student record, including **Student 3 (Math = 97)** in the displayed output.

---

# 🛠️ Excel Skills Demonstrated

<div align="center">

| Category | Skills |
|:--|:--|
| **Cell References** | Relative references, Absolute references |
| **Logical Functions** | `IF`, Nested `IF`, `AND`, `OR` |
| **Conditional Aggregation** | `COUNTIF`, `SUMIFS`, `AVERAGEIF` |
| **Lookup Functions** | `VLOOKUP`, `INDEX`, `MATCH`, `XLOOKUP`, `XMATCH` |
| **Text Functions** | `LEFT`, `FIND`, `UPPER` |
| **Dynamic References** | `INDIRECT`, `OFFSET` |
| **Date & Time** | `DATEDIF`, `TODAY` |
| **Math & Rounding** | `ROUND`, `FLOOR`, `CEILING` |
| **Dynamic Arrays** | `FILTER` |
| **Data Handling** | Formatting, structured datasets, formula-based analysis |

</div>

---

# 📈 Key Results Demonstrated

| Result | Value |
|:--|--:|
| Students scoring above 50 in Math | **19** |
| East-region Keyboard sales | **80,349** |
| Average Math score above 60 | **82.26** |
| Student ID 1003 lookup | **Student 3** |
| Person 13 sales in June 2023 | **15,560** |
| Region lookup using INDEX/MATCH | **2001** |
| Employee 2 salary lookup | **31,943** |
| Mouse XMATCH position | **3** |
| Keyboard XMATCH position | **1** |
| Total sales using INDIRECT | **507,725** |
| Dynamic OFFSET total | **48,148** |
| Example year difference | **11 years** |
| Example date difference | **4,005 days** |

---

# 🎓 Learning Outcomes

By completing this project, the following practical Excel skills were developed:

✅ Writing and copying formulas using **relative and absolute references**  
✅ Building decision-making logic with **IF, AND, OR, and nested IF**  
✅ Performing conditional calculations using **SUMIFS, COUNTIF, and AVERAGEIF**  
✅ Retrieving information with **VLOOKUP, XLOOKUP, INDEX/MATCH, and XMATCH**  
✅ Cleaning and transforming text with Excel **TEXT functions**  
✅ Working with **dates, year differences, and elapsed days**  
✅ Using **INDIRECT and OFFSET** for dynamic references  
✅ Applying **rounding functions** for numeric and financial calculations  
✅ Returning filtered datasets with the **FILTER** dynamic-array function  
✅ Structuring real-world data for practical spreadsheet analysis

---

# 📁 Project Files

```text
Excel PR. 1 Fundamental Booster/
│
├── PR. 1 Fundamental Booster task.xlsx   # Original task / instruction workbook
├── PR 1.xlsx                       # Completed practical workbook
└── README.md                        # Project documentation
```

---

# 🚀 How to Use This Project

1. Open **`PR. 1 Fundamental Booster task.xlsx`** to review the original questions and essential tasks.
2. Open **`PR 1.xlsx`** to view the completed formulas, calculations, and results.
3. Use the **question numbers in brackets** inside `PR 1.xlsx` to map each answer back to the original task.
4. Read this README from top to bottom to understand the purpose of each formula and the result it produces.

---

# 👩‍💻 Project Status

<div align="center">

### ✅ ESSENTIAL TASKS COMPLETED

**PR. 1 Fundamental Booster — Excel Practical Project**

</div>

---

<div align="center">

### ⭐ Built to Practice • Learn • Analyze • Excel

</div>
