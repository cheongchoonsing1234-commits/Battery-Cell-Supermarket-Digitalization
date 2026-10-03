# 🔋 Battery Cell Supermarket Digitalization

> A real-time Excel VBA dashboard and transaction system that replaced manual
> observation and handwritten records in a battery cell supermarket on a
> production line.

![Dashboard](docs/images/01-dashboard.png)

---

## 📌 Table of Contents
1. [Background & Problem](#background--problem)
2. [Objective](#objective)
3. [Solution Overview](#solution-overview)
4. [Process Flow: Before vs After](#process-flow-before-vs-after)
5. [System Modules](#system-modules)
6. [Business Impact](#business-impact)
7. [Technical Highlights](#technical-highlights)
8. [Hardware Setup](#hardware-setup)
9. [Repository Structure](#repository-structure)
10. [How to Run the Demo](#how-to-run-the-demo)
11. [My Role](#my-role)
12. [Lessons Learned & Future Improvements](#lessons-learned--future-improvements)

---

## Background & Problem

In the battery cell supermarket, material is supplied to the production line
and replenished across 100 PGPCELL locations. The process was fully manual:

- Staff checked the stock level **by eye** to decide when to replenish.
- Withdrawals, returns, and replenishments were **keyed in or written by hand**.
- There was **no central, real-time view** of how many cells were left.
- Manual key-in created risks: **typing mistakes, missing records, wrong
  quantities, and delayed replenishment**.
- Tracing who took which material, and when, was difficult.

## Objective

Digitalize the supermarket so that:
1. Stock level can be understood **at a glance**.
2. Data is captured through **scanner and card reader**, not typing.
3. Every action is **traceable** to an employee and a date.
4. **FIFO** (first in, first out) is enforced for battery cells.
5. Replenishment is triggered **earlier and more accurately**.

## Solution Overview

A desktop application built in **Excel VBA** using UserForms. It runs on one
station beside the supermarket and has a sidebar with six functions:
**Chart, Withdraw, Return, Replenish, Edit, Records**.

Operators scan the material barcode and tap their employee card. The system
validates the input, updates the database, and refreshes the dashboard
immediately.

## Process Flow: Before vs After

| Before (manual) | After (digital) |
|---|---|
| ![Before](docs/images/process-flow-before.png) | ![After](docs/images/process-flow-after.png) |

---

## System Modules

### 1️⃣ Dashboard (Chart)

![Dashboard](docs/images/01-dashboard.png)

**Purpose:** Give anyone a clear view of the stock situation in seconds.

**How it works:**
- Each **bar represents one battery cell part number**.
- The bar height shows the current quantity compared to its maximum.
- Bars are **colour-coded** according to the minimum and maximum
  thresholds stored for each part:
  - 🟢 **Green:** sufficient stock
  - 🟠 **Orange / Yellow:** low stock, attention needed
  - 🔴 **Red:** critical level, urgent replenishment
- When a part turns orange or red, the **battery cell loader creates a
  Transfer Order (TO)** with a mobile scanner. The production supervisor
  confirms the TO when replenishment is done (same as the original process).
- The chart refreshes automatically after every withdrawal, return, or
  replenishment.

**Why it matters:** The decision "do we need to replenish?" no longer depends
on someone walking around and judging by eye.

---

### 2️⃣ Withdrawal

![Withdrawal](docs/images/02-withdrawal.png)

**Purpose:** Record every time battery cells are taken from the supermarket.

**Process:**
1. **Employee identification:** the person taps their employee card
   (or types the number). This links the withdrawal to a real person.
2. **Part number input:** scanned or entered using the format
   `x.xxx.xxx.xxx`. The form checks the length and format and rejects
   incomplete numbers.
3. **Carton quantity:** the operator enters how many full cartons were taken.
4. **Automatic conversion:** the system multiplies by the carton size stored
   for that part (e.g. 200 cells per carton) and converts to **units**.
5. **Loose quantity:** if the withdrawal is not a full carton, the operator
   enters the extra quantity in **individual cells**.
6. **Confirmation:** the system checks that all fields are filled and the
   quantity is positive, then:
   - deducts the stock in the database,
   - writes a record (employee, part number, cartons, cells, date)
     to the Withdrawal sheet,
   - updates the dashboard.

**Error prevention built in:** blank fields are blocked, negative or
non-numeric quantities are rejected, and unknown part numbers return a
warning instead of creating bad data.

---

### 3️⃣ Replenishment

![Replenishment](docs/images/03-replenishment.png)

**Purpose:** Record incoming battery cells and control FIFO.

**Process:**
1. **Employee identification** by the battery cell loader.
2. **TO-based input:** the loader enters the part number and the quantity
   (in cells) according to the Transfer Order.
3. **Physical verification:** the loader enters the **rank** and **Lot ID**
   from the stock actually received. This confirms that the system
   matches the physical goods.
4. **System validation and guidance:** the system looks into the database
   for the existing rank and Lot ID, then **recommends the next sticker
   colour** from a fixed sequence
   (Green → Orange → Dark Blue → Pink → Yellow, then repeats).
   The suggested colour is based on the **last colour used**.
5. **Execution:** the loader sticks the system-generated colour on the
   cartons and loads the cells into the correct rack location.
6. The record (employee, part, lot, date, rank, colour, quantity) is
   saved, and the row in the history sheet is coloured to match.

**Why this matters (FIFO):** Colour sequencing lets operators see the
oldest lot at a glance. The older colour is taken first, so no battery cells
sit too long in storage. This is especially important for batteries, where
age and rank matter.

---

### 4️⃣ Return

![Return](docs/images/04-return.png)

**Purpose:** Put unused battery cells back into stock correctly.

**Process:**
1. **Employee identification** to keep returns traceable.
2. **Return input:** part number and quantity in cells.
3. **System update:** the system validates the data (format, positive
   quantity, part exists) and **adds the quantity back** to available stock.
4. The return is logged and the dashboard updates.

**Why it matters:** Without a return function, unused cells are either
forgotten or the stock becomes inaccurate. This keeps the count honest.

---

### 5️⃣ Editor (Add / Edit / Delete)

![Editor](docs/images/05-editor.png)

**Purpose:** Maintain the master data and correct stock when needed.

The Editor is **password-protected** so only authorised staff can change data.

**Three functions:**
- **Add new part number:** define the initial parameters, including
  current quantity, **maximum quantity**, **minimum quantity**,
  **carton quantity**, and pallet/racking quantity. The system blocks
  duplicates (if the part already exists, it warns the user).
- **Edit existing data:** used for the **end-of-shift stock count**. The
  operator counts what is physically there and updates the quantity, so
  the dashboard stays accurate.
- **Delete part number:** removes obsolete parts after a
  confirmation prompt ("Are you sure you want to delete this material?"),
  to prevent accidental deletion.

**Benefit:** flexibility. When new battery models arrive or old ones are
discontinued, no programmer is needed to update the system.

---

### 6️⃣ History (Records)

![History](docs/images/06-history.png)

**Purpose:** Full traceability of all inventory movements.

- Shows the **withdrawal, replenishment, and return records** in one place.
- Every line includes **who, which part, how many, and when**.
- Supports investigation, for example when a quantity doesn't match,
  by tracing back through the records.
- Records are kept for a set period and then **automatically cleared after
  90 days**, so the file stays fast and the database doesn't grow forever.

---

### ➕ Additional Automation

- **Automatic email notification:** the workbook includes a module that can
  send an email (via Outlook) when a delivery update or low-stock alert
  needs to be communicated. **[Confirm how you use this]**
- **90-day auto-reset timer:** on opening, the workbook checks how long it
  has been since the last reset and clears the record sheets when the limit
  is reached. A counter ("X days left before reset") is visible.

---

## Business Impact

- ✅ **Reduced manual key-in errors** through scanner and card-reader input
  and form validation
- ✅ **Real-time stock visibility** that anyone can read at a glance
- ✅ **Earlier replenishment:** colour alerts trigger action before shortage
- ✅ **Full traceability:** every transaction is tied to an employee and a date
- ✅ **FIFO enforced** with a colour-sticker system
- ✅ **Low-cost solution:** built with tools already available (Excel) plus
  basic hardware
- 📊 **[Add your result, e.g. "reduced stock-check time from X to Y minutes"]**

## Technical Highlights

| Area | What was done |
|---|---|
| **Language / Platform** | Excel VBA, UserForms, worksheet database |
| **Custom UI** | Borderless, modern-looking form with draggable title bar, built using Windows API calls (`user32`) |
| **Reusable components** | Class modules (`cSideButton`, `cDragSurface`) for sidebar buttons with hover and press effects |
| **Dynamic chart** | Bar chart drawn from database values, with colours decided by min and max thresholds |
| **Input validation** | Format masks (`x.xxx.xxx.xxx`), required-field checks, positive-number checks, duplicate detection |
| **Colour system** | Separate module that manages the FIFO colour sequence |
| **Data protection** | Password-protected editor, delete confirmation |
| **Maintenance** | Automatic 90-day record reset |
| **Integration** | Barcode scanner and card reader (acting as keyboard input), Outlook email |

## Hardware Setup

- 1 × desktop station (stock update, receiving, issuing)
- 1 × barcode scanner (material scanning)
- 1 × employee card reader (tap to capture employee ID)
- 1 × touch screen monitor

## Repository Structure
(paste the tree from Part 4)

## How to Run the Demo
1. Download `demo/BatteryCell_Demo.xlsm`.
2. Open it and click **Enable Content / Enable Macros**.
3. The dashboard opens automatically.
4. All data is **fictional**. The demo edit password is `[demo password]`.

## My Role
[Write 3 to 5 lines. Example:]
- Collected requirements from operators, loaders, and supervisors
- Designed the process flow and user interface
- Developed the full VBA application (6 modules, validation, chart)
- Tested with users and fixed issues
- Prepared documentation and trained the team

## Lessons Learned & Future Improvements
- Move from Excel to a proper database (SQL) for multi-user access
- Add automatic low-stock alerts to a phone or Teams
- Build a web version so stock can be seen from anywhere
- Add daily/weekly usage analytics for demand forecasting

---
👤 **Author:** [Your Name] · [LinkedIn] · [Email]
