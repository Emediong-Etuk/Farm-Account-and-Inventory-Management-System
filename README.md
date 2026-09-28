# Farm Account and Inventory Management System

An Excel workbook for farm record-keeping, inventory and accounting: **`Farm_Accounting_Inventory_System.xlsx`**.

Its rule is: **enter each transaction once; everything else is calculated.**

```
Purchase → Inventory (store) → Issue / Usage → Pen & Batch → Production → Sale → Profitability
```

A purchase is recorded once and goes into the store. It becomes a pen or batch cost only when feed is
**issued** or a drug is **used**, so one purchase (for example 20 bags × 25 kg) can be shared by any number of pens.

## Hierarchy

`Farm → Pen → Batch → Transactions`

* **Pen**: a permanent physical house (`P001`). Its ID never changes.
* **Batch**: one production cycle in a pen (`BROILER-SEP-001`). When a batch is closed, its history stays and the pen is free for the next batch.

## Sheets

| Sheet | What it is | You type? |
|---|---|---|
| README | Instructions inside the workbook | – |
| DASHBOARD | Farm status, financial, production and mortality cards, charts, and a pen comparison. Filters: period (Today / This Week / This Month / This Year / All Time / Custom from–to), pen, batch, animal type | Filters only |
| PEN_REGISTER | Pens. Current batch, start date, stock and status are calculated | ID, name, location, type |
| BATCH_REGISTER | Batches. Current quantity = initial + transfers in − deaths − animals sold − transfers out | ID, pen, start, qty, status, closing date |
| FEED_PURCHASES | Feed bought (bags × bag size = kg; bags × price = cost) | ✔ |
| FEED_ISSUES | Feed sent to a pen/batch. Cost per kg is filled in automatically | ✔ |
| DRUG_PURCHASES | Drugs and vaccines bought, with expiry date, lot number and expiry warnings | ✔ |
| DRUG_USAGE | Each dose given, per pen/batch, with dosage. Cost is automatic | ✔ |
| MORTALITY_LOG | Deaths, recorded only on days they happen | ✔ |
| PRODUCTION_LOG | Weighings and output such as eggs, recorded only when they happen | ✔ |
| SALES | Sales linked to pen and batch (`GENERAL` for items like manure) | ✔ |
| OTHER_EXPENSES | Chicks, labour, electricity, repairs and so on, charged to a pen, a batch, or `GENERAL` | ✔ |
| CASHBOOK | Actual money in and out, with a running balance | ✔ |
| INVENTORY | Feed and drug stock, value, REORDER and expiry alerts | – |
| BATCH_PROFITABILITY | Full-cycle results for every batch ever recorded, active or closed | – |
| STOCK_MOVEMENTS | Transfers between batches, stock added, count corrections | ✔ |
| REPORTS | Feed by type, drugs by drug, mortality by cause, sales by product and customer, expenses by category | – |
| TRENDS | Monthly and 14-day tables with charts | First month only |
| CHECKS | Data-quality counts and a reconciliation of the cashbook against the logs | – |
| LISTS | Master lists for every dropdown, reorder levels, drug units and settings | ✔ |

Green column headers mark cells you type in. Grey headers mark formulas; new rows fill these in automatically.
Every log ends with a **Check** column that shows `OK` or a plain-English warning. The warnings cover:

* missing pen or batch
* a batch that is not in the pen
* a date outside the batch's life
* an issue that exceeds stock
* use of an expired lot
* purchases entered out of date order
* negative stock

## Accounting rules

* **Costing method: moving (perpetual) weighted average.** After each purchase the new average cost is
  `(stock on hand × old average + purchase cost) / (stock on hand + quantity bought)`.
  Each issue is costed at the average in force on its date. The calculation steps are visible in the grey
  columns on FEED_PURCHASES and DRUG_PURCHASES. Purchases always equal cost consumed plus stock still in the store.
* **Purchase vs consumption.** Buying ₦500,000 of feed is spending, not cost. It becomes cost only when the feed is issued.
* **Cash vs profit.** The CASHBOOK records money movement. Profit uses consumption. The CHECKS sheet reconciles the two
  and shows outstanding credit owed to suppliers and by customers.
* **Overheads.** `GENERAL` expenses count in farm totals only. Batch profitability includes only costs tagged
  with that batch's ID.

## Scale

All logs are Excel Tables and all formulas use structured references, so new rows are included automatically.
You never need to edit a range, whether the farm has 5 pens or 50, or tens of thousands of rows.

The summary sheets have fixed capacities:

| Sheet | Shows up to |
|---|---|
| DASHBOARD | 60 pens |
| BATCH_PROFITABILITY | 300 batches |
| INVENTORY | 20 feed types and 50 drugs |

To raise these limits, change the `*_SLOTS` constants in the generator script and rebuild the workbook.

## Sample data

Every example row is marked `SAMPLE DATA` in its Notes column. The sample includes:

* pens P001–P006
* one closed batch in P006, which shows how history is kept
* a 500 kg Starter purchase shared across P001–P004 (40 / 60 / 35 / 50 kg)
* one multivitamin used in three pens at different dosages

Delete the sample rows before you start real records.

## Regenerating the workbook

The workbook is built by `build_farm_workbook.py` (requires `openpyxl`):

```bash
pip install openpyxl
python3 build_farm_workbook.py            # writes Farm_Accounting_Inventory_System.xlsx
```

The file is saved without cached values and is set to recalculate fully on open. Excel, LibreOffice and
Google Sheets all show the numbers when the file is opened. File previewers that do not calculate formulas
will show blank cells.
