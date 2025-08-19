## Project Cost Calcuation Templates

Templates for calculating costs in research projects, tailored for DFG requests and reporting. Small, single-purpose Python scripts generate a Excel workbooks to budget (e.g., for one field trip). The workbook includes sheets for **Inputs & Rates**, **Staff & Participants** (incl. Hiwis and graduating students), **Hours Log**, **Travel & Vehicles**, **Materials & Other**, and an aggregated **Summary**.

---

## Features

* **Per-diem calculation** with full/partial day allowances (no meal deductions).
* **Hiwi wages** via hours x hourly rate x (1 + on-cost %).
* **Graduating students** (Bachelor's/Master's): hours tracked, wage cost fixed at 0.
* **Travel**: tickets/day-rates, rental-car per-km, private-car per-km (trip-level cap applied in the **Summary**).
* **Materials & Other**: any additional line items.
* **Named rates** and clear formulas directly in the workbook for transparency.

> The script writes an Excel file (e.g., named `fieldtrip-cost-template.xlsx`) in the `output` directory.

---

## Quick start

```bash
# (optional) create & activate a virtual environment
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install openpyxl
python scripts/generate_template.py  # or your chosen filename
```

After running, open `fieldtrip-cost-template.xlsx` and adjust values on **Inputs & Rates**.

---

## Workbook overview

* **Inputs & Rates** -- Edit per-diem full/partial amounts, private-car rate & cap, rental-car per-km, overnight default, **Hiwi hourly rate** and **on-cost %**. 
* **Staff & Participants** -- One row per person. Choose role: **Employee**, **Hiwi (student assistant)**, or **Grad. Student (unpaid)**. The sheet auto-pulls **Hours** per name from **Hours Log** and computes per-diems, overnights, wages (if applicable), and a participant subtotal.
* **Hours Log** -- Date, task, first/last name, hours. Keeps time tracking in one place; `SUMIFS` aggregates back to **Staff & Participants**.
* **Travel & Vehicles** -- Tickets/day-rates, rental-car variable (km x per-km), private-car (km x per-km). The **Summary** applies the private-car trip cap.
* **Materials & Other** -- Consumables, equipment rentals, shipping, permits, etc.
* **Summary** -- Category subtotals and grand total.

---

## Usage

* Fill **Inputs & Rates** first; all other sheets reference these values.
* In **Staff & Participants**, set the **Role**:

  * **Employee** -- per-diem/overnight only (no wage cost).
  * **Hiwi (student assistant)** -- per-diem/overnight **plus** wages from hours x rate x (1 + on-cost).
  * **Student (unpaid)** -- per-diem/overnight only; hours tracked for effort reporting, wage cost = 0.
* Use **Hours Log** as your single source of truth for time tracking; names must match the **Staff & Participants** sheet.

