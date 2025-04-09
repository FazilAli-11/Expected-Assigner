# 📌 CC Expected Assigner (Excel | SQL | Power Query | VBA)

**Date:** January 2025  
**Author:** Fazil Ali  
**Tools Used:** Excel, SQL Server, Power Query, VBA, Macros  

---

## 🧠 Project Brief

The **CC Expected Assigner** automates the task assignment process for Data Analysts (DAs) by integrating data from SQL Server into Excel. Using Power Query for dynamic filtering and VBA/Macros for logic-driven task assignment, the solution eliminates manual effort, minimizes errors, and enhances team productivity.

---

## ✅ Key Benefits

- 🔄 **100% automation** of DA task assignments for Estimate & Forecasted events.
- ⏱ **90% reduction** in daily processing time.
- 🙌 **80% reduction** in manual effort.
- 🎯 **95% improvement** in accuracy — significantly reducing data discrepancies.

---

## 🔑 Core Concepts

| Field            | Description                                 |
|------------------|---------------------------------------------|
| `CompanyID`      | Unique identifier for each company           |
| `EventID`        | Multiple events under a single CompanyID     |
| `BeginningDate`  | Date of the scheduled event                  |
| `Status`         | Event status — *Confirmed* / *Estimate*      |

---

## 🗂 Priority Segregation Logic

Events are automatically classified based on their `BeginningDate`:

| Priority Level | Schedule                          |
|----------------|-----------------------------------|
| **P1**         | Current day                       |
| **P2**         | Current day + 2 days              |
| **P3**         | Current day + 4 days              |
| **Top P**      | High priority (FF1, FF2, FF3)     |

Assignment of each category is handled via a structured DA mapping sheet.

---

## ❌ Previous Manual Workflow

- Manually pulled data from the Corporate Communication (CC) system for events scheduled within the next 7 days.
- Events were segregated into priorities using VLOOKUP.
- Tasks were then assigned manually — this took **~30 minutes daily** and was prone to human error.

---

## ⚙️ Automated Workflow (What I Built)

- 🔌 **SQL Integration:** Automatically pulls real-time data for current date + 7 days.
- 📊 **Power Query:** Applies filters, prioritization, and segregation rules directly within Excel.
- 🤖 **VBA + Macros:** Performs dynamic DA task assignments with just one click.
- 📁 **Output Sheets:** Final categorized sheets (`Top P`, `P1`, `P2`, `P3`) are ready-to-use with zero manual edits required.

---

## 🔒 Confidentiality Note

Due to the sensitivity of company data, the actual Excel files and SQL scripts are **not included** in this repository.  
However, **screenshots** are provided below to demonstrate project structure, workflows, and output.

---

## 📸 Project Preview

| Task | Screenshot |
|------|------------|
| DA Task Assigning via Macro | ![DA Task Assigning](https://github.com/FazilAli-11/Expected-Assigner/blob/main/DA%20Task%20Assigning%20(VBA%2BMacro).png) |
| P1, P2, P3 Output Sheets | ![Expected P1 P2 P3](https://github.com/FazilAli-11/Expected-Assigner/blob/main/Expected%20(P1%2CP2%2CP3).png) |
| Power Query Segregation | ![PowerQuery Logic](https://github.com/FazilAli-11/Expected-Assigner/blob/main/Segreation%20by%20PowerQuery.png) |
| Top Priority Segregation | ![Top Priority](https://github.com/FazilAli-11/Expected-Assigner/blob/main/Top%20Priority%20(FF1%2C%20FF2%2CFF3).png) |

---

## 👨‍💻 Author

**Fazil Ali**  
📍 [GitHub](https://github.com/FazilAli-11) | 🔗 [LinkedIn](https://www.linkedin.com/in/fazilali11)

---

