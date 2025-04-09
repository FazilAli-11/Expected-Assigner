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

| Field         | Description                               |
|---------------|-------------------------------------------|
| `CompanyID`   | Unique identifier for each company         |
| `EventID`     | Multiple events under a single CompanyID   |
| `BeginningDate` | Date of the scheduled event               |
| `Status`      | Event status — *Confirmed* / *Estimate*    |

---

## ❌ Previous Workflow (Manual)

- Data pulled manually from the Corporate Communication (CC) system for events scheduled within the next 7 days.
- Events were categorized by priority using VLOOKUP:
  - **Top Priority:** FF1, FF2, FF3  
  - **Priority Buckets:** P1, P2, P3
- Task assignment to DAs was done manually, taking **~30 minutes daily**.

---

## ⚙️ What I Built (Automated Workflow)

- 🔌 **SQL Integration:** Pulls current date + 7 days data directly from SQL Server into Excel.
- 📊 **Power Query:** Filters, segregates, and transforms data based on business rules.
- 🤖 **VBA + Macros:** Automatically assigns tasks to DAs based on pre-defined logic and mapping.
- 📥 **Output Sheets:** Separate categorized sheets (`Top P`, `P1`, `P2`, `P3`) — fully ready for action with no manual editing.

---



