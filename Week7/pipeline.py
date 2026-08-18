"""
Week 7 - Python Data Pipeline Engineering
Retail DW ETL Pipeline Script

Author: Student
Course: Data Warehousing Concepts and Design (Week 7)
"""

import os
import sys
import sqlite3
import logging
from dataclasses import dataclass, field
from datetime import datetime
from typing import List, Tuple, Dict, Any, Optional
import pandas as pd
import numpy as np

# Configure logging to stdout
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger("PipelineEngine")


@dataclass
class PipelineConfig:
    input_path: str = "Python_Data_Pipeline_Lab_Dataset.xlsx"
    output_db: str = "retail_dw.db"
    quarantine_file: str = "quarantine.csv"
    log_file: str = "pipeline_run_log.csv"
    batch_list: List[str] = field(default_factory=lambda: [
        "orders_batch_1",
        "orders_batch_2",
        "orders_batch_3"
    ])
    error_mode: str = "quarantine"


class DataWarehouseLoader:
    def __init__(self, db_path: str):
        self.db_path = db_path
        self._init_db()

    def get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_path)
        conn.execute("PRAGMA foreign_keys = ON;")
        return conn

    def _init_db(self):
        with self.get_connection() as conn:
            cursor = conn.cursor()
            # Customer Dimension
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS dim_customer (
                    customer_key INTEGER PRIMARY KEY AUTOINCREMENT,
                    customer_id TEXT UNIQUE NOT NULL,
                    customer_name TEXT,
                    province TEXT,
                    segment TEXT,
                    signup_date TEXT
                );
            """)

            # Product Dimension
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS dim_product (
                    product_key INTEGER PRIMARY KEY AUTOINCREMENT,
                    product_id TEXT UNIQUE NOT NULL,
                    product_name TEXT,
                    category TEXT,
                    unit_price REAL,
                    active_flag TEXT
                );
            """)

            # Date Dimension
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS dim_date (
                    date_key INTEGER PRIMARY KEY,
                    full_date TEXT UNIQUE NOT NULL,
                    day INTEGER,
                    month INTEGER,
                    quarter INTEGER,
                    year INTEGER
                );
            """)

            # Fact Sales Table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS fact_sales (
                    order_id TEXT PRIMARY KEY,
                    date_key INTEGER REFERENCES dim_date(date_key),
                    customer_key INTEGER REFERENCES dim_customer(customer_key),
                    product_key INTEGER REFERENCES dim_product(product_key),
                    quantity INTEGER NOT NULL,
                    unit_price REAL NOT NULL,
                    discount_pct REAL NOT NULL,
                    gross_amount REAL NOT NULL,
                    net_amount REAL NOT NULL,
                    payment_method TEXT,
                    sales_channel TEXT,
                    updated_at TEXT NOT NULL,
                    source_batch INTEGER NOT NULL
                );
            """)

            # Quarantine Table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS quarantine (
                    quarantine_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    order_id TEXT,
                    order_datetime TEXT,
                    customer_id TEXT,
                    product_id TEXT,
                    quantity TEXT,
                    unit_price TEXT,
                    discount_pct TEXT,
                    payment_method TEXT,
                    sales_channel TEXT,
                    updated_at TEXT,
                    source_batch INTEGER,
                    reason_code TEXT,
                    quarantined_at TEXT
                );
            """)

            # Pipeline Run Log Table
            cursor.execute("""
                CREATE TABLE IF NOT EXISTS pipeline_run_log (
                    run_id INTEGER PRIMARY KEY AUTOINCREMENT,
                    batch_name TEXT NOT NULL,
                    started_at TEXT NOT NULL,
                    ended_at TEXT NOT NULL,
                    rows_read INTEGER NOT NULL,
                    rows_valid INTEGER NOT NULL,
                    rows_rejected INTEGER NOT NULL,
                    rows_duplicated INTEGER NOT NULL,
                    rows_loaded INTEGER NOT NULL,
                    status TEXT NOT NULL
                );
            """)
            conn.commit()

    def sync_dimensions(self, excel_path: str):
        """Loads and syncs dim_customer and dim_product from Excel."""
        cust_df = pd.read_excel(excel_path, sheet_name="customers")
        prod_df = pd.read_excel(excel_path, sheet_name="products")

        with self.get_connection() as conn:
            cursor = conn.cursor()
            for _, row in cust_df.iterrows():
                cursor.execute("""
                    INSERT INTO dim_customer (customer_id, customer_name, province, segment, signup_date)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(customer_id) DO UPDATE SET
                        customer_name = excluded.customer_name,
                        province = excluded.province,
                        segment = excluded.segment,
                        signup_date = excluded.signup_date;
                """, (
                    str(row['customer_id']).strip(),
                    str(row['customer_name']).strip(),
                    str(row['province']).strip(),
                    str(row['segment']).strip(),
                    str(row['signup_date'])[:10] if pd.notna(row['signup_date']) else None
                ))

            for _, row in prod_df.iterrows():
                cursor.execute("""
                    INSERT INTO dim_product (product_id, product_name, category, unit_price, active_flag)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(product_id) DO UPDATE SET
                        product_name = excluded.product_name,
                        category = excluded.category,
                        unit_price = excluded.unit_price,
                        active_flag = excluded.active_flag;
                """, (
                    str(row['product_id']).strip(),
                    str(row['product_name']).strip(),
                    str(row['category']).strip(),
                    float(row['unit_price']),
                    str(row['active_flag']).strip()
                ))
            conn.commit()

    def get_dimension_lookups(self) -> Tuple[Dict[str, int], Dict[str, int]]:
        """Returns customer_id -> customer_key and product_id -> product_key mapping."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT customer_id, customer_key FROM dim_customer")
            cust_map = {row[0]: row[1] for row in cursor.fetchall()}

            cursor.execute("SELECT product_id, product_key FROM dim_product")
            prod_map = {row[0]: row[1] for row in cursor.fetchall()}

        return cust_map, prod_map

    def ensure_date_key(self, dt: datetime, cursor: sqlite3.Cursor) -> int:
        """Generates YYYYMMDD date_key and inserts into dim_date if not present."""
        date_key = int(dt.strftime("%Y%m%d"))
        full_date = dt.strftime("%Y-%m-%d")
        quarter = (dt.month - 1) // 3 + 1

        cursor.execute("""
            INSERT OR IGNORE INTO dim_date (date_key, full_date, day, month, quarter, year)
            VALUES (?, ?, ?, ?, ?, ?);
        """, (date_key, full_date, dt.day, dt.month, quarter, dt.year))
        return date_key


class PipelineEngine:
    def __init__(self, config: PipelineConfig):
        self.config = config
        self.loader = DataWarehouseLoader(config.output_db)

    def extract_batch(self, batch_name: str) -> pd.DataFrame:
        logger.info(f"Extracting batch '{batch_name}' from {self.config.input_path}...")
        df = pd.read_excel(self.config.input_path, sheet_name=batch_name)
        logger.info(f"Extracted {len(df)} raw rows from '{batch_name}'.")
        return df

    @staticmethod
    def normalize_payment_method(val: Any) -> str:
        if pd.isna(val):
            return "Unknown"
        s = str(val).strip()
        lower = s.lower()
        if "credit" in lower:
            return "Credit Card"
        elif "prompt" in lower:
            return "PromptPay"
        elif "cash" in lower:
            return "Cash"
        elif "bank" in lower or "transfer" in lower:
            return "Bank Transfer"
        return s.title()

    @staticmethod
    def normalize_sales_channel(val: Any) -> str:
        if pd.isna(val):
            return "Unknown"
        s = str(val).strip()
        lower = s.lower()
        if "e-commerce" in lower or "ecommerce" in lower or "online" in lower:
            return "Online"
        elif "store" in lower:
            return "Store"
        elif "market" in lower:
            return "Marketplace"
        return s.title()

    def transform_and_validate(self, df: pd.DataFrame, batch_name: str) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        cust_map, prod_map = self.loader.get_dimension_lookups()
        
        valid_records: List[Dict[str, Any]] = []
        quarantine_records: List[Dict[str, Any]] = []

        for idx, row in df.iterrows():
            reasons = []
            
            # 1. Order ID
            raw_order_id = row.get("order_id")
            if pd.isna(raw_order_id) or str(raw_order_id).strip() == "":
                reasons.append("MISSING_ORDER_ID")
                order_id = f"UNKNOWN_{idx}"
            else:
                order_id = str(raw_order_id).strip()

            # 2. Order Datetime
            raw_dt = row.get("order_datetime")
            dt_obj = None
            if pd.isna(raw_dt):
                reasons.append("MISSING_DATETIME")
            else:
                try:
                    dt_obj = pd.to_datetime(raw_dt)
                    if pd.isna(dt_obj):
                        reasons.append("INVALID_DATETIME")
                except Exception:
                    reasons.append("INVALID_DATETIME")

            # 3. Customer ID & Foreign Key
            raw_cust_id = row.get("customer_id")
            cust_key = None
            if pd.isna(raw_cust_id) or str(raw_cust_id).strip() == "":
                reasons.append("MISSING_CUSTOMER_ID")
            else:
                cust_id_str = str(raw_cust_id).strip()
                if cust_id_str not in cust_map:
                    reasons.append("CUSTOMER_NOT_FOUND")
                else:
                    cust_key = cust_map[cust_id_str]

            # 4. Product ID & Foreign Key
            raw_prod_id = row.get("product_id")
            prod_key = None
            if pd.isna(raw_prod_id) or str(raw_prod_id).strip() == "":
                reasons.append("MISSING_PRODUCT_ID")
            else:
                prod_id_str = str(raw_prod_id).strip()
                if prod_id_str not in prod_map:
                    reasons.append("PRODUCT_NOT_FOUND")
                else:
                    prod_key = prod_map[prod_id_str]

            # 5. Quantity (1 to 20)
            raw_qty = row.get("quantity")
            qty_val = None
            try:
                qty_val = int(raw_qty)
                if qty_val < 1 or qty_val > 20:
                    reasons.append("INVALID_QUANTITY")
            except (ValueError, TypeError):
                reasons.append("INVALID_QUANTITY")

            # 6. Unit Price (> 0)
            raw_price = row.get("unit_price")
            price_val = None
            try:
                price_val = float(raw_price)
                if pd.isna(price_val) or price_val <= 0:
                    reasons.append("INVALID_UNIT_PRICE")
            except (ValueError, TypeError):
                reasons.append("INVALID_UNIT_PRICE")

            # 7. Discount Pct (0 to 100)
            raw_disc = row.get("discount_pct")
            disc_val = None
            try:
                disc_val = float(raw_disc)
                if pd.isna(disc_val) or disc_val < 0 or disc_val > 100:
                    reasons.append("INVALID_DISCOUNT_PCT")
            except (ValueError, TypeError):
                reasons.append("INVALID_DISCOUNT_PCT")

            # 8. Updated At
            raw_updated_at = row.get("updated_at")
            updated_at_str = str(raw_updated_at) if pd.notna(raw_updated_at) else datetime.now().isoformat()

            # Payment method & sales channel normalization
            payment_method = self.normalize_payment_method(row.get("payment_method"))
            sales_channel = self.normalize_sales_channel(row.get("sales_channel"))

            source_batch = int(row.get("source_batch", 0))

            if len(reasons) > 0:
                # Quarantine record
                quarantine_records.append({
                    "order_id": str(raw_order_id) if pd.notna(raw_order_id) else "",
                    "order_datetime": str(raw_dt) if pd.notna(raw_dt) else "",
                    "customer_id": str(raw_cust_id) if pd.notna(raw_cust_id) else "",
                    "product_id": str(raw_prod_id) if pd.notna(raw_prod_id) else "",
                    "quantity": str(raw_qty) if pd.notna(raw_qty) else "",
                    "unit_price": str(raw_price) if pd.notna(raw_price) else "",
                    "discount_pct": str(raw_disc) if pd.notna(raw_disc) else "",
                    "payment_method": str(row.get("payment_method", "")),
                    "sales_channel": str(row.get("sales_channel", "")),
                    "updated_at": updated_at_str,
                    "source_batch": source_batch,
                    "reason_code": ",".join(reasons),
                    "quarantined_at": datetime.now().isoformat()
                })
            else:
                # Derived Calculations
                gross_amount = round(qty_val * price_val, 2)
                net_amount = round(gross_amount * (1.0 - disc_val / 100.0), 2)

                valid_records.append({
                    "order_id": order_id,
                    "order_datetime": dt_obj,
                    "customer_key": cust_key,
                    "product_key": prod_key,
                    "quantity": qty_val,
                    "unit_price": price_val,
                    "discount_pct": disc_val,
                    "gross_amount": gross_amount,
                    "net_amount": net_amount,
                    "payment_method": payment_method,
                    "sales_channel": sales_channel,
                    "updated_at": updated_at_str,
                    "source_batch": source_batch
                })

        return valid_records, quarantine_records

    def load_to_dw(
        self,
        batch_name: str,
        valid_records: List[Dict[str, Any]],
        quarantine_records: List[Dict[str, Any]],
        rows_read: int
    ) -> Dict[str, Any]:
        started_at = datetime.now().isoformat()
        
        # Intra-batch deduplication: keep latest updated_at per order_id
        dedup_dict: Dict[str, Dict[str, Any]] = {}
        dup_count = 0

        for rec in valid_records:
            oid = rec["order_id"]
            if oid in dedup_dict:
                dup_count += 1
                # Compare updated_at
                existing_updated = dedup_dict[oid]["updated_at"]
                if rec["updated_at"] > existing_updated:
                    dedup_dict[oid] = rec
            else:
                dedup_dict[oid] = rec

        clean_records = list(dedup_dict.values())
        rows_loaded = 0

        try:
            with self.loader.get_connection() as conn:
                cursor = conn.cursor()

                # Load quarantine records into database table
                for q in quarantine_records:
                    cursor.execute("""
                        INSERT INTO quarantine (
                            order_id, order_datetime, customer_id, product_id,
                            quantity, unit_price, discount_pct, payment_method,
                            sales_channel, updated_at, source_batch, reason_code, quarantined_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                    """, (
                        q["order_id"], q["order_datetime"], q["customer_id"], q["product_id"],
                        q["quantity"], q["unit_price"], q["discount_pct"], q["payment_method"],
                        q["sales_channel"], q["updated_at"], q["source_batch"], q["reason_code"],
                        q["quarantined_at"]
                    ))

                # Load clean records into fact_sales with Idempotent UPSERT
                for r in clean_records:
                    dt_obj = r["order_datetime"]
                    date_key = self.loader.ensure_date_key(dt_obj, cursor)

                    # Check if order_id exists
                    cursor.execute("SELECT updated_at FROM fact_sales WHERE order_id = ?", (r["order_id"],))
                    existing = cursor.fetchone()

                    if existing is None:
                        # New record -> INSERT
                        cursor.execute("""
                            INSERT INTO fact_sales (
                                order_id, date_key, customer_key, product_key,
                                quantity, unit_price, discount_pct, gross_amount,
                                net_amount, payment_method, sales_channel, updated_at, source_batch
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                        """, (
                            r["order_id"], date_key, r["customer_key"], r["product_key"],
                            r["quantity"], r["unit_price"], r["discount_pct"], r["gross_amount"],
                            r["net_amount"], r["payment_method"], r["sales_channel"],
                            r["updated_at"], r["source_batch"]
                        ))
                        rows_loaded += 1
                    else:
                        # Existing record -> check updated_at for UPSERT
                        existing_updated = existing[0]
                        if r["updated_at"] > existing_updated:
                            cursor.execute("""
                                UPDATE fact_sales SET
                                    date_key = ?, customer_key = ?, product_key = ?,
                                    quantity = ?, unit_price = ?, discount_pct = ?,
                                    gross_amount = ?, net_amount = ?, payment_method = ?,
                                    sales_channel = ?, updated_at = ?, source_batch = ?
                                WHERE order_id = ?;
                            """, (
                                date_key, r["customer_key"], r["product_key"],
                                r["quantity"], r["unit_price"], r["discount_pct"],
                                r["gross_amount"], r["net_amount"], r["payment_method"],
                                r["sales_channel"], r["updated_at"], r["source_batch"],
                                r["order_id"]
                            ))
                            rows_loaded += 1
                        else:
                            dup_count += 1

                ended_at = datetime.now().isoformat()
                status = "SUCCESS"

                # Log pipeline execution
                cursor.execute("""
                    INSERT INTO pipeline_run_log (
                        batch_name, started_at, ended_at, rows_read, rows_valid,
                        rows_rejected, rows_duplicated, rows_loaded, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
                """, (
                    batch_name, started_at, ended_at, rows_read, len(valid_records),
                    len(quarantine_records), dup_count, rows_loaded, status
                ))

                conn.commit()

        except Exception as e:
            logger.error(f"Error loading batch '{batch_name}': {e}", exc_info=True)
            ended_at = datetime.now().isoformat()
            status = "FAILED"
            with self.loader.get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    INSERT INTO pipeline_run_log (
                        batch_name, started_at, ended_at, rows_read, rows_valid,
                        rows_rejected, rows_duplicated, rows_loaded, status
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);
                """, (
                    batch_name, started_at, ended_at, rows_read, len(valid_records),
                    len(quarantine_records), dup_count, 0, status
                ))
                conn.commit()
            raise e

        # Export Quarantine and Run Log to CSV
        self.export_csv_logs()

        return {
            "batch_name": batch_name,
            "started_at": started_at,
            "ended_at": ended_at,
            "rows_read": rows_read,
            "rows_valid": len(valid_records),
            "rows_rejected": len(quarantine_records),
            "rows_duplicated": dup_count,
            "rows_loaded": rows_loaded,
            "status": status
        }

    def export_csv_logs(self):
        with self.loader.get_connection() as conn:
            q_df = pd.read_sql_query("SELECT * FROM quarantine", conn)
            q_df.to_csv(self.config.quarantine_file, index=False)

            log_df = pd.read_sql_query("SELECT * FROM pipeline_run_log", conn)
            log_df.to_csv(self.config.log_file, index=False)

    def run_single_batch(self, batch_name: str) -> Dict[str, Any]:
        logger.info(f"========== Starting Pipeline Run for '{batch_name}' ==========")
        df_raw = self.extract_batch(batch_name)
        valid_recs, quar_recs = self.transform_and_validate(df_raw, batch_name)
        stats = self.load_to_dw(batch_name, valid_recs, quar_recs, len(df_raw))
        logger.info(f"Completed '{batch_name}': Read={stats['rows_read']}, Valid={stats['rows_valid']}, "
                    f"Rejected={stats['rows_rejected']}, Duplicated={stats['rows_duplicated']}, "
                    f"Loaded={stats['rows_loaded']}, Status={stats['status']}\n")
        return stats


def run_pipeline(config: PipelineConfig):
    engine = PipelineEngine(config)
    
    # 1. Sync Dimensions
    logger.info("Syncing Dimension Tables (customers, products)...")
    engine.loader.sync_dimensions(config.input_path)

    # 2. Run Required Execution Sequence (4 passes for Lab Acceptance Test)
    # Pass 1: orders_batch_1
    # Pass 2: orders_batch_1 (re-run to test Idempotency)
    # Pass 3: orders_batch_2
    # Pass 4: orders_batch_3
    run_sequence = [
        "orders_batch_1",
        "orders_batch_1",  # Re-run test
        "orders_batch_2",
        "orders_batch_3"
    ]

    summary_results = []
    for i, batch in enumerate(run_sequence, start=1):
        res = engine.run_single_batch(batch)
        res["pass_number"] = i
        summary_results.append(res)

    # Final DB Metrics Summary
    with engine.loader.get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM fact_sales;")
        total_fact_rows = cursor.fetchone()[0]

        cursor.execute("SELECT SUM(net_amount) FROM fact_sales;")
        total_net_sales = cursor.fetchone()[0] or 0.0

        cursor.execute("SELECT COUNT(*) FROM quarantine;")
        total_quarantine_rows = cursor.fetchone()[0]

    print("\n=======================================================")
    print("             DATA PIPELINE EXECUTION SUMMARY            ")
    print("=======================================================")
    print(f"Total Fact Sales Records Loaded: {total_fact_rows}")
    print(f"Total Quarantine Records Logged: {total_quarantine_rows}")
    print(f"Total Net Sales Revenue: ${total_net_sales:,.2f}")
    print("-------------------------------------------------------")
    print(f"{'Pass':<6}{'Batch Name':<20}{'Read':<8}{'Valid':<8}{'Rej':<8}{'Dup':<8}{'Loaded':<8}{'Status'}")
    print("-------------------------------------------------------")
    for r in summary_results:
        print(f"{r['pass_number']:<6}{r['batch_name']:<20}{r['rows_read']:<8}{r['rows_valid']:<8}"
              f"{r['rows_rejected']:<8}{r['rows_duplicated']:<8}{r['rows_loaded']:<8}{r['status']}")
    print("=======================================================\n")


if __name__ == "__main__":
    # Ensure working directory is set to script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)

    cfg = PipelineConfig(
        input_path="Python_Data_Pipeline_Lab_Dataset.xlsx",
        output_db="retail_dw.db",
        quarantine_file="quarantine.csv",
        log_file="pipeline_run_log.csv"
    )

    run_pipeline(cfg)
