# 📌 CC Expected Assigner (Excel | SQL | Power Query | VBA)

**Date:** January 2025  
**Author:** Fazil Ali  
**Tools Used:** Excel, SQL Server, Power Query, VBA, Macros  

---

## 🧠 Project Brief

The **CC Expected Assigner** automates the task assignment process for Data Analysts (DAs) by integrating data from SQL Server into Excel. It leverages Power Query for dynamic data processing and VBA for automated task distribution. This system eliminates manual intervention, enhances accuracy, and boosts overall productivity.

---

## ✅ Key Benefits

- 🔄 **100% automation** of DA task assignments for Estimate & Forecasted Events.
- ⏱ **90% reduction** in processing time.
- 🙌 **80% less manual effort** required.
- 🎯 **95% improvement** in assignment accuracy — reducing data errors and inconsistencies.

---

## 🔑 Core Concepts

| Field           | Description                                 |
|------------------|---------------------------------------------|
| `CompanyID`       | Unique identifier for each company           |
| `EventID`         | Multiple events under a single CompanyID     |
| `BeginningDate`   | Date of the scheduled event                  |
| `Status`          | Event status — *Confirmed* / *Estimate*      |

---

## 🗂 Priority Segregation Logic

Events are automatically classified into the following categories based on `BeginningDate`:

| Priority Level | Schedule                         |
|----------------|----------------------------------|
| **P1**         | Current day                      |
| **P2**         | Current day + 2 days             |
| **P3**         | Current day + 4 days             |
| **Top P**      | High priority (e.g., FF1–FF3)     |

Each category is assigned to employees using a pre-defined mapping sheet.

---

## ❌ Previous Workflow (Manual)

- Data was pulled manually from the Corporate Communication (CC) system for events scheduled within the next 7 days.
- Events were categorized by priority using VLOOKUP:
  - **Top Priority:** FF1, FF2, FF3  
  - **Expected:** P1, P2, P3
- Tasks were manually assigned to DAs, taking **~30 minutes daily**.

---

## ⚙️ What I Built (Automated Workflow)

- 🔌 **SQL Integration:** Pulls current date + 7 days of event data directly from SQL Server into Excel.
- 📊 **Power Query:** Automatically transforms and filters the data based on rules.
- 🤖 **VBA + Macros:** Assigns events to DAs across categories with a single click.
- 📥 **Output:** Segregated sheets (`Top P`, `P1`, `P2`, `P3`) are generated with assignments, ready to use.

---

## 🔒 Confidentiality Note

Due to confidential company data, actual Excel files and SQL scripts are **not included** in this repository.
Instead, screenshots are shared below to demonstrate the project’s structure, functionality, and output.

---

## 📸 Project Preview

https://github.com/FazilAli-11/Expected-Assigner/blob/main/DA%20Task%20Assigning%20(VBA%2BMacro).png
https://github.com/FazilAli-11/Expected-Assigner/blob/main/Expected%20(P1%2CP2%2CP3).png
https://github.com/FazilAli-11/Expected-Assigner/blob/main/Segreation%20by%20PowerQuery.png
https://github.com/FazilAli-11/Expected-Assigner/blob/main/Top%20Priority%20(FF1%2C%20FF2%2CFF3).png

---

## 👨‍💻 Author

**Fazil Ali**  
[GitHub](https://github.com/FazilAli-11) | [LinkedIn](https://www.linkedin.com/in/fazilali11)
