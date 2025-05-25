
# 📘 Excel Formulas & Functions Study Project – Mastering with Employee Dataset

## 👋 Introduction

Welcome to your personal Excel study playground! 🎉

This project is about **learning, revising, and practicing the most useful Excel formulas and functions** using a realistic HR dataset. The goal is to **build solid Excel skills** that can help in **job tasks, data analysis, and dashboards**, and most importantly, **make you confident using Excel like a pro**.

Whether you're preparing for interviews, building dashboards, or just flexing your spreadsheet muscles — this repo is your go-to Excel function guide 🚀.

---

## 🎯 What This Project Is

✅ A personal study notebook for Excel functions

✅ Hands-on with real formulas, filters, and lookups

✅ Focused on learning *how, when,* and *why* to use different functions

✅ Simple, clear, and practical explanations

✅ Includes bonus challenges and real-life examples


---

## 📁 File Contents

* `EmployeeData.xlsx`:
  Includes:

  * Raw employee data
  * Calculations using different Excel formulas
  * Filters, sorts, dashboards
  * Lookup examples
  * Department reports, salary summaries, and more!

---

## ✅ Business Questions Practiced

Here’s a list of **10 Excel-based challenges** I’ve completed using this dataset:

| #  | Business Question                                                            | Function(s) You'll Learn                                      | Challenge for You                                     |
| -- | ---------------------------------------------------------------------------- | ------------------------------------------------------------- | ----------------------------------------------------- |
| 1  | Total Salary and headcount by department                                     | `SUMIFS`, `COUNTIFS`                                          | Identify only permanent headcount                     |
| 2  | All employees with more than \$100k salary                                   | `FILTER`                                                      |                                                       |
| 3  | All female employees with more than \$100k salary                            | `FILTER`, `*`                                                 | Filter by joining year (2020 or later)                |
| 4  | Lowest, highest and top 5 salary values                                      | `MIN`, `MAX`, `LARGE`, `SORT` + `TAKE`                        |                                                       |
| 5  | List of all departments                                                      | `UNIQUE`, `COUNTA`, `#`                                       | Show all departments in a single comma-separated cell |
| 6  | Employee details lookup                                                      | `VLOOKUP`, `INDEX` + `MATCH`                                  |                                                       |
| 7  | Employee details lookup                                                      | `XLOOKUP`, `IFERROR`                                          | Lookup all employees earning exactly \$120,000        |
| 8  | Complex formula: Highest salary person                                       | `XLOOKUP` + `MAX`                                             |                                                       |
| 9  | Complex formula: All employees joined in March                               | `FILTER` + `MONTH`                                            | Also filter for females who started on a Monday       |
| 10 | Complex formula: Department report of headcounts, salaries & % diff from avg | `UNIQUE`, `SUMIFS`, `COUNTIFS`, `#`, `CONDITIONAL FORMATTING` | Calculate **median salary** and **female ratio**      |

### 📘 Explanations:

* **Permanent Headcount (Q1):** Use `COUNTIFS` with employment type to count only full-time or permanent staff per department.
* **FILTER for \$100k+ (Q2, Q3):** Master conditional row filtering using multiple columns (`Salary`, `Gender`, `Join Year`) using dynamic array functions.
* **Top Salaries (Q4):** Combine `LARGE`, `SORT`, `TAKE` to extract top salary insights — useful for dashboards or HR insights.
* **Unique Departments (Q5):** Use `TEXTJOIN(", ", TRUE, UNIQUE(...))` to list all unique departments in a single neat cell.
* **Employee Lookup (Q6–Q7):** Learn classic vs modern lookup: `VLOOKUP`, `INDEX`+`MATCH`, `XLOOKUP`, and error handling with `IFERROR`.
* **Highest Paid Employee (Q8):** Use `MAX(Salary)` to get top salary and `XLOOKUP` to return corresponding name/role.
* **Joined in March (Q9):** Filter employees by `MONTH(JoinDate)=3`, and use `TEXT(WEEKDAY(Date),"dddd")="Monday"` to find those who joined on a Monday.
* **Department Report (Q10):** Build a mini dashboard using `UNIQUE` for department list, `SUMIFS/COUNTIFS` for metrics, and highlight unusual values using `CONDITIONAL FORMATTING`.

---

## 🧠 Master Excel Formulas – By Category

### 🔍 Lookup Functions

| Function        | Use Case                   | Simple Explanation                                |
| --------------- | -------------------------- | ------------------------------------------------- |
| `VLOOKUP`       | Find employee’s department | Looks **vertically** to match a value in a column |
| `INDEX + MATCH` | More flexible lookup       | `MATCH` finds row number, `INDEX` gets value      |
| `XLOOKUP` ✅     | Modern lookup              | Searches both ways, includes error handling       |

### 🧮 Aggregation Functions

| Function               | Use Case                  | Simple Explanation                       |
| ---------------------- | ------------------------- | ---------------------------------------- |
| `SUMIFS`, `COUNTIFS` ✅ | Aggregate with conditions | Adds or counts based on multiple filters |

### 🔠 Text Functions

| Function   | Use Case              | Simple Explanation                     |
| ---------- | --------------------- | -------------------------------------- |
| `TEXTJOIN` | Comma-separate values | Joins multiple values with a delimiter |

### 📅 Date Functions

| Function           | Use Case            | Simple Explanation                          |
| ------------------ | ------------------- | ------------------------------------------- |
| `MONTH`, `WEEKDAY` | Filter by month/day | Used in filters like March or Monday starts |
| `YEAR`             | Filter by join year | Useful for “joined after 2020” condition    |

### 🔢 Logical & Conditional Functions

| Function        | Use Case           | Simple Explanation                         |
| --------------- | ------------------ | ------------------------------------------ |
| `IF`, `IFERROR` | Conditional output | Show something only when condition is true |

### 📊 Dynamic Array Functions

| Function                | Use Case                | Simple Explanation             |
| ----------------------- | ----------------------- | ------------------------------ |
| `FILTER` ✅              | Show only relevant data | Like SQL WHERE clause          |
| `SORT`, `TAKE`, `LARGE` | Ranking and limits      | Grab top N values easily       |
| `UNIQUE` ✅              | Remove duplicates       | Get list of unique departments |

### 📌 Extras:

| Tool                   | Purpose                              |
| ---------------------- | ------------------------------------ |
| Conditional Formatting | Highlight outliers like top salaries |
| Data Validation        | Dropdowns for department or gender   |
| Named Ranges           | Cleaner and readable formulas        |
| Flash Fill             | Auto-predict based on data patterns  |

---

## 📈 What You’ve Practiced

* Advanced filtering with `FILTER`, `UNIQUE`, `SORT` and formulas such as `MEDIAN`
* Salary and headcount analysis using `SUMIFS`, `COUNTIFS`, `IFERROR`
* Modern lookups using `XLOOKUP`, `VLOOKUP` and `INDEX + MATCH`
* Conditional and date logic with  `MONTH`, `WEEKDAY`
* Text formatting using `TEXTJOIN`, `PROPER`, `TRIM`
* Dashboard basics using Excel tables and formatting


## ✍️ Final Thoughts

This is your **Excel formula battle station** 🧠💥
Keep building on it by solving more HR, finance, or operations scenarios.
This practice set is great for interviews, dashboards, and day-to-day job analysis. The more you play, the more confident you'll become!

---

> 🧮 “Excel doesn’t judge your formulas... but it will definitely return an error if you’re wrong.” 😄

