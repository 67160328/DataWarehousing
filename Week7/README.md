# Python Data Pipeline Engineering - Retail DW Lab (Week 7)

## 📌 Overview
This repository contains the complete ETL Data Pipeline solution for omnichannel retail order transactions. The pipeline extracts raw data from Excel batch files, normalizes & validates fields, quarantines invalid records with specific reason codes, deduplicates records, and loads clean records into a SQLite Star Schema Data Warehouse (`retail_dw.db`).

---

## 🛠️ Architecture & Star Schema Design

### Star Schema Structure
The Data Warehouse is built on a **Star Schema** architecture consisting of 3 Dimension tables and 1 Fact table:

- **`dim_customer`**: Customer dimension (`customer_key` PK, `customer_id` natural key, `customer_name`, `province`, `segment`, `signup_date`).
- **`dim_product`**: Product dimension (`product_key` PK, `product_id` natural key, `product_name`, `category`, `unit_price`, `active_flag`).
- **`dim_date`**: Date dimension (`date_key` PK YYYYMMDD, `full_date`, `day`, `month`, `quarter`, `year`).
- **`fact_sales`**: Fact sales table (`order_id` PK, `date_key` FK, `customer_key` FK, `product_key` FK, `quantity`, `unit_price`, `discount_pct`, `gross_amount`, `net_amount`, `payment_method`, `sales_channel`, `updated_at`, `source_batch`).

### Fact Table Grain
> **Grain Definition:** One verified line item purchase per unique `order_id`.

---

## 🔍 Data Quality Rules & Quarantine Handling
Data issues are detected during the transformation phase and rejected into the `quarantine` table (and `quarantine.csv`) with clear `reason_code` tags:

1. `MISSING_ORDER_ID`: Blank or null transaction order key.
2. `INVALID_DATETIME`: Unparseable order timestamp.
3. `MISSING_CUSTOMER_ID` / `CUSTOMER_NOT_FOUND`: Missing customer key or referential integrity failure against `dim_customer`.
4. `MISSING_PRODUCT_ID` / `PRODUCT_NOT_FOUND`: Missing product key or referential integrity failure against `dim_product`.
5. `INVALID_QUANTITY`: Non-numeric, `<= 0`, or `> 20`.
6. `INVALID_UNIT_PRICE`: Non-numeric or `<= 0`.
7. `INVALID_DISCOUNT_PCT`: Non-numeric, `< 0`, or `> 100`.

### Data Cleaning & Normalization:
- **`payment_method`**: Case-folded and standardized (`Credit Card`, `PromptPay`, `Cash`, `Bank Transfer`).
- **`sales_channel`**: Mapped `E-Commerce` / `ecommerce` to `Online`, and standardized channel labels.
- **Calculated Metrics**:
  - `gross_amount = quantity * unit_price`
  - `net_amount = gross_amount * (1 - discount_pct / 100.0)`

---

## 🚀 How to Run the Pipeline

### Prerequisites
- Python 3.8+
- `pandas`, `openpyxl`, `sqlite3`

```bash
pip install pandas openpyxl
```

### Execution
Run `pipeline.py` from the `Week7` directory:

```bash
python pipeline.py
```

---

## 📊 Pipeline Run Results & Metrics

Below is the execution audit log summary across 4 sequential passes:

| Pass | Batch Name | Read Rows | Valid Rows | Rejected Rows | Duplicated Rows | Loaded Rows | Status |
|:---:|:---|:---:|:---:|:---:|:---:|:---:|:---:|
| **1** | `orders_batch_1` | 420 | 386 | 34 | 0 | 386 | **SUCCESS** |
| **2** | `orders_batch_1` *(Re-run)* | 420 | 386 | 34 | 386 | 0 | **SUCCESS** |
| **3** | `orders_batch_2` | 424 | 384 | 40 | 1 | 383 | **SUCCESS** |
| **4** | `orders_batch_3` | 424 | 386 | 38 | 4 | 382 | **SUCCESS** |

### Key Summary Metrics:
- **Total Clean Fact Records Loaded in DW:** `1,150`
- **Total Quarantined Invalid Records:** `146`
- **Total Cumulative Net Sales Revenue:** `$2,812,461.57`
- **Idempotency Verification:** Pass 2 (re-running batch 1) loaded `0` new rows, preventing data duplication.

---

## 💡 Reflection: Availability vs. Strictness in Production Pipelines

In production data engineering, system availability and continuous data flow are almost always prioritized over strict, all-or-nothing pipeline halts. If a pipeline is designed with extreme strictness—failing the entire batch upon encountering a single invalid timestamp or corrupted row—business dashboards, real-time analytics, and downstream operational tools will be completely starved of data. Isolating erroneous records into a dedicated quarantine store while allowing clean data to flow guarantees that operational decisions rely on fresh, available information. Crucially, late-arriving or corrected records in quarantine can easily be reprocessed later without causing downtime or disrupting downstream analytics.

---

## 📁 Deliverable Files in `Week7/`
- [`pipeline.py`](file:///c:/ClassRoomBuraphaUniversity/DataWarehouseingConceptsAndDesign/Week7/pipeline.py) - Complete Python ETL Script
- [`retail_dw.db`](file:///c:/ClassRoomBuraphaUniversity/DataWarehouseingConceptsAndDesign/Week7/retail_dw.db) - SQLite Data Warehouse
- [`quarantine.csv`](file:///c:/ClassRoomBuraphaUniversity/DataWarehouseingConceptsAndDesign/Week7/quarantine.csv) - Rejected Records with `reason_code`
- [`pipeline_run_log.csv`](file:///c:/ClassRoomBuraphaUniversity/DataWarehouseingConceptsAndDesign/Week7/pipeline_run_log.csv) - Execution Audit Log
- [`README.md`](file:///c:/ClassRoomBuraphaUniversity/DataWarehouseingConceptsAndDesign/Week7/README.md) - Pipeline Documentation & Reflection
