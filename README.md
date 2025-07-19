# 📊 AI Prompts for Excel & Google Sheets Formulas

Unlock the full power of Excel and Google Sheets using AI. This repository offers a curated collection of effective prompts designed to help you generate, explain, and fix spreadsheet formulas using AI tools like ChatGPT, Claude, Gemini, or Copilot.

---

## 📌 Overview

Working with spreadsheets is powerful—but complex. From basic calculations to advanced lookups and nested logic, formulas can be confusing or time-consuming. With the right **AI prompt engineering**, you can automate the creation, explanation, and optimization of any formula—instantly.

This project provides a structured collection of **AI prompts for Excel and Google Sheets**, enabling users to:
- Generate accurate formulas from plain English instructions
- Translate complex formulas into human-readable explanations
- Debug errors or fix broken spreadsheet logic
- Learn by example and reduce repetitive tasks

Whether you're a student, accountant, analyst, business manager, or just spreadsheet-curious — this repo will transform how you work with data.

---

## ✅ How to Use

1. Choose a prompt that matches your use case below.
2. Paste it into your favorite AI tool (e.g., ChatGPT).
3. Replace placeholders like `[your description]`, `[cell range]`, or `[column names]`.
4. Copy the AI-generated output directly into your spreadsheet.

---

## 🔟 10 AI Prompt Examples with Outputs

### 1. ➕ Generate a formula to sum only if condition is met

**Prompt:**  
```

Write a formula to sum values in column B only if column A says "Sales".

````

**Output:**  
```excel
=SUMIF(A:A, "Sales", B:B)
````

---

### 2. 🧠 Explain a complex formula

**Prompt:**

```
Explain this formula in plain English: =IF(AND(B2>500,C2="Yes"),"Qualified","Not Qualified")
```

**Output:**

> If the value in B2 is greater than 500 and the value in C2 is "Yes", return "Qualified". Otherwise, return "Not Qualified".

---

### 3. 🛠 Fix a broken VLOOKUP

**Prompt:**

```
My formula =VLOKUP(E2, A2:B20, 2, FALSE) gives an error. Can you fix it?
```

**Output:**

```excel
=VLOOKUP(E2, A2:B20, 2, FALSE)
```

> The function name was misspelled — "VLOKUP" should be "VLOOKUP".

---

### 4. 📅 Extract month name from date

**Prompt:**

```
Write a formula to get the full month name from a date in cell A1.
```

**Output:**

```excel
=TEXT(A1, "mmmm")
```

---

### 5. 🎯 Rank values in descending order

**Prompt:**

```
How can I rank numbers in column D from highest to lowest?
```

**Output:**

```excel
=RANK.EQ(D2, D$2:D$100, 0)
```

---

### 6. 📈 Calculate percentage change between two columns

**Prompt:**

```
Create a formula to calculate percentage increase from column A to column B.
```

**Output:**

```excel
=((B2 - A2) / A2) * 100
```

---

### 7. 🔗 Create a hyperlink based on cell value

**Prompt:**

```
Make a formula to create a clickable link using text in cell A2 and URL in cell B2.
```

**Output:**

```excel
=HYPERLINK(B2, A2)
```

---

### 8. 🎯 Conditional formatting: highlight duplicates

**Prompt:**

```
What formula can I use in conditional formatting to highlight duplicate values in column C?
```

**Output:**

```excel
=COUNTIF(C:C, C1) > 1
```

---

### 9. 🧾 Running total of expenses

**Prompt:**

```
Generate a formula to create a running total in column B starting from B2.
```

**Output:**

```excel
=SUM($B$2:B2)
```

---

### 10. 🧮 Count unique values in a column

**Prompt:**

```
Write a formula to count unique values in column A.
```

**Output:**

```excel
=COUNTA(UNIQUE(A:A))
```

---

## 🚀 Benefits of Using AI for Spreadsheet Formulas

* ⏱ Save time on repetitive formula writing
* ❌ Eliminate syntax errors and guesswork
* 💡 Learn how each formula works, step by step
* 🧰 Debug and optimize legacy spreadsheets
* 📚 Build reusable prompt libraries for your team

Whether you're handling sales reports, payroll sheets, dashboards, or school projects, AI-enhanced spreadsheet workflows can save hours of manual effort.

---

## 🔍 Bonus Prompt Ideas

* "Convert a nested IF to SWITCH formula"
* "Extract first and last names from full name column"
* "Split data in column B by comma"
* "Create formula to calculate tax based on tiered brackets"
* "Generate QUERY formula to filter by date range in Google Sheets"
* "Make a formula to flag overdue tasks from a deadline column"

---

## 🌐 Discover More at [Promptshub.net](https://promptshub.net)

Looking for **more professional AI prompt templates** like these?

### ✨ Visit 👉 [https://promptshub.net](https://promptshub.net)

**Promptshub** offers:

* 💼 Prompt packs for Excel, Google Sheets, PowerPoint, Word, Notion & Airtable
* 🔍 AI tools for automation, data entry, reports & analysis
* 🧠 Ready-to-use templates for finance, marketing, HR, and operations
* 🛠 Prompt bundles for ChatGPT, Claude, Gemini, Copilot & more

🔗 Start automating smarter with prompts: **[https://promptshub.net](https://promptshub.net)**

---

## 🤝 Contributing

Have an awesome prompt idea for spreadsheets?
Feel free to open a pull request or start a discussion!

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
Use it freely and share it widely.

