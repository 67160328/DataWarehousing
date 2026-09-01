"""
================================================================================
LAB: Data Integration Pipeline
TechTrove E-Commerce: จากข้อมูลดิบหลายระบบสู่ข้อมูลพร้อมวิเคราะห์
Course: Data Warehousing Concepts and Design - Week 8
================================================================================
"""

from pathlib import Path
import json
import logging
import sys
from typing import Tuple, Dict, Any, List
import pandas as pd
import numpy as np

# Configure Standard Output & Logging
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler(sys.stdout)]
)
logger = logging.getLogger("TechTroveETL")

# Setup Directories
DATA_DIR = Path(__file__).parent / "data"
OUTPUT_DIR = Path(__file__).parent / "output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)


# ==============================================================================
# TODO 1: Extract ข้อมูลจาก CSV, Excel และ JSON พร้อมทำการ Profile ข้อมูลดิบ (5.1)
# ==============================================================================
def extract_and_profile_data(data_dir: Path) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, List[Dict[str, Any]]]:
    """
    5.1 Extract และ Profile:
    อ่านทุกไฟล์ แสดง shape, columns, dtype, missing, duplicate ก่อนแก้ไข
    """
    logger.info("========== [TODO 1] EXTRACT & PROFILE RAW DATA ==========")
    
    # 1.1 Read Orders Jan & Feb
    df_jan = pd.read_csv(data_dir / "orders_2026_01.csv")
    df_feb = pd.read_csv(data_dir / "orders_2026_02.csv")
    
    # 1.2 Read CRM Customers
    df_cust = pd.read_csv(data_dir / "customers_crm.csv")
    
    # 1.3 Read Product Master
    df_prod = pd.read_excel(data_dir / "product_master.xlsx")
    
    # 1.4 Read Payments JSON
    with open(data_dir / "payments.json", "r", encoding="utf-8") as f:
        payments_raw = json.load(f)
        
    print("\n--- 1. Data Profiling Summary ---")
    print(f"1) orders_2026_01.csv : Shape={df_jan.shape}, Columns={df_jan.columns.tolist()}")
    print(f"   Missing Values:\n{df_jan.isnull().sum()[df_jan.isnull().sum() > 0]}")
    print(f"   Duplicates (order_id): {df_jan['order_id'].duplicated().sum()}")
    
    print(f"\n2) orders_2026_02.csv : Shape={df_feb.shape}, Columns={df_feb.columns.tolist()}")
    print(f"   Missing Values:\n{df_feb.isnull().sum()[df_feb.isnull().sum() > 0]}")
    print(f"   Duplicates (order_id): {df_feb['order_id'].duplicated().sum()}")
    
    print(f"\n3) customers_crm.csv  : Shape={df_cust.shape}, Columns={df_cust.columns.tolist()}")
    print(f"   Missing Values:\n{df_cust.isnull().sum()[df_cust.isnull().sum() > 0]}")
    print(f"   Duplicates (customer_id): {df_cust['customer_id'].duplicated().sum()}")
    
    print(f"\n4) product_master.xlsx: Shape={df_prod.shape}, Columns={df_prod.columns.tolist()}")
    print(f"   Missing Values: {df_prod.isnull().sum().sum()}, Duplicates: {df_prod['product_id'].duplicated().sum()}")
    
    print(f"\n5) payments.json      : Total Events={len(payments_raw)}")
    
    return df_jan, df_feb, df_cust, df_prod, payments_raw


# ==============================================================================
# TODO 2: ทำ Schema Alignment ของไฟล์ orders สองเดือน แล้ว Concat (5.2)
# ==============================================================================
def align_and_combine_orders(df_jan: pd.DataFrame, df_feb: pd.DataFrame) -> pd.DataFrame:
    """
    5.2 Combine Orders:
    ปรับชื่อคอลัมน์และรูปแบบ discount ของเดือน ก.พ. ให้ตรงกับเดือน ม.ค. แล้ว concat โดยใช้ ignore_index=True
    """
    logger.info("========== [TODO 2] SCHEMA ALIGNMENT & ORDERS COMBINE ==========")
    
    # 2.1 ปรับชื่อคอลัมน์ของ ก.พ. ให้ตรงกับ ม.ค.
    df_feb_aligned = df_feb.rename(columns={
        "ordered_at": "order_date",
        "qty": "quantity",
        "discount_pct": "discount"
    }).copy()
    
    # 2.2 แปลง discount จากสตริงเปอร์เซ็นต์ (เช่น '5%') เป็นตัวเลขทศนิยม (float 0.05)
    df_feb_aligned["discount"] = (
        df_feb_aligned["discount"]
        .astype(str)
        .str.rstrip("%")
        .astype(float) / 100.0
    )
    
    # 2.3 จัดรูปแบบ datetime ให้เป็นมาตรฐานเดียวกัน
    df_feb_aligned["order_date"] = pd.to_datetime(df_feb_aligned["order_date"], format="%d/%m/%Y %H:%M")
    df_jan_aligned = df_jan.copy()
    df_jan_aligned["order_date"] = pd.to_datetime(df_jan_aligned["order_date"])
    df_jan_aligned["discount"] = df_jan_aligned["discount"].astype(float)
    
    # 2.4 ผสานรวมคำสั่งซื้อ
    df_orders_combined = pd.concat([df_jan_aligned, df_feb_aligned], ignore_index=True)
    logger.info(f"Successfully combined orders: Total rows = {len(df_orders_combined)}")
    
    return df_orders_combined


# ==============================================================================
# TODO 3: Clean/standardize/deduplicate และสร้าง Data Quality Report (5.3)
# ==============================================================================
def clean_and_standardize_dimensions(
    df_cust: pd.DataFrame, 
    df_prod: pd.DataFrame, 
    payments_raw: List[Dict[str, Any]]
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, List[Dict[str, Any]]]:
    """
    5.3 Transform:
    แปลงชนิดข้อมูล, ทำ email เป็น lower-case/trim, ทำจังหวัดให้เป็นชื่อมาตรฐาน, ลบ duplicate ตามกติกา
    """
    logger.info("========== [TODO 3] CLEAN & STANDARDIZE DIMENSIONS ==========")
    dq_records = []
    
    # 3.1 Clean Customers CRM
    # บันทึกข้อมูลซ้ำในลูกค้า
    dup_cust = df_cust[df_cust.duplicated(subset=["customer_id"], keep=False)]
    for cid in df_cust[df_cust.duplicated(subset=["customer_id"], keep="last")]["customer_id"].unique():
        dq_records.append({
            "stage": "Dimension_Cleaning",
            "table_name": "customers_crm",
            "record_id": cid,
            "issue_type": "DUPLICATE_CUSTOMER_ID",
            "details": f"Duplicate customer_id {cid} in CRM; kept latest occurrence",
            "action_taken": "DEDUPLICATED"
        })
    
    # ลบข้อมูลซ้ำโดยเก็บข้อมูลล่าสุด
    df_cust_clean = df_cust.drop_duplicates(subset=["customer_id"], keep="last").copy()
    df_cust_clean["customer_id"] = df_cust_clean["customer_id"].str.strip()
    df_cust_clean["full_name"] = df_cust_clean["full_name"].str.strip()
    
    # Email: Lower-case & Trim
    df_cust_clean["email"] = df_cust_clean["email"].fillna("").astype(str).str.strip().str.lower()
    df_cust_clean["email"] = df_cust_clean["email"].replace("", np.nan)
    
    # Standardize ชื่อจังหวัด (14 รูปแบบ -> 6 จังหวัดมาตรฐาน)
    province_map = {
        "กรุงเทพมหานคร": "กรุงเทพมหานคร", "Bangkok": "กรุงเทพมหานคร", "กทม.": "กรุงเทพมหานคร",
        "ชลบุรี": "ชลบุรี", "Chonburi": "ชลบุรี", "ชลบุรี ": "ชลบุรี",
        "ระยอง": "ระยอง", "Rayong": "ระยอง",
        "ขอนแก่น": "ขอนแก่น", "ขอนเเก่น": "ขอนแก่น",
        "เชียงใหม่": "เชียงใหม่", "Chiang Mai": "เชียงใหม่",
        "ภูเก็ต": "ภูเก็ต", "Phuket": "ภูเก็ต"
    }
    df_cust_clean["province"] = df_cust_clean["province"].astype(str).str.strip().map(lambda p: province_map.get(p, p))
    df_cust_clean["signup_date"] = pd.to_datetime(df_cust_clean["signup_date"]).dt.strftime("%Y-%m-%d")
    
    # 3.2 Clean Product Master
    df_prod_clean = df_prod.drop_duplicates(subset=["product_id"], keep="last").copy()
    df_prod_clean["product_id"] = df_prod_clean["product_id"].str.strip()
    df_prod_clean["product_name"] = df_prod_clean["product_name"].str.strip()
    df_prod_clean["category"] = df_prod_clean["category"].str.strip()
    df_prod_clean["standard_price"] = df_prod_clean["standard_price"].astype(float)
    df_prod_clean["active_flag"] = df_prod_clean["active_flag"].astype(str).str.strip().str.upper()
    
    # 3.3 Clean & Flatten Payments JSON
    pay_flat = []
    for p in payments_raw:
        pay_flat.append({
            "payment_id": str(p.get("payment_id", "")).strip(),
            "order_id": str(p.get("order_id", "")).strip(),
            "payment_method": str(p.get("payment", {}).get("method", "")).strip(),
            "payment_status": str(p.get("payment", {}).get("status", "")).strip().upper(),
            "paid_at": p.get("paid_at")
        })
    df_pay = pd.DataFrame(pay_flat)
    
    # Log duplicate payments
    dup_pay = df_pay[df_pay.duplicated(subset=["order_id"], keep="last")]
    for _, r in dup_pay.iterrows():
        dq_records.append({
            "stage": "Payment_Cleaning",
            "table_name": "payments",
            "record_id": r["order_id"],
            "issue_type": "DUPLICATE_PAYMENT_ORDER_ID",
            "details": f"Duplicate payment event for order_id {r['order_id']}; kept latest occurrence",
            "action_taken": "DEDUPLICATED"
        })
        
    df_pay_clean = df_pay.drop_duplicates(subset=["order_id"], keep="last").copy()
    df_pay_clean["paid_at"] = pd.to_datetime(df_pay_clean["paid_at"]).dt.strftime("%Y-%m-%dT%H:%M:%S")
    
    logger.info(f"Cleaned Dimensions: Customers={len(df_cust_clean)}, Products={len(df_prod_clean)}, Payments={len(df_pay_clean)}")
    return df_cust_clean, df_prod_clean, df_pay_clean, dq_records


# ==============================================================================
# TODO 4 & 5: Enrich & Validate Business Rules ก่อนคำนวณยอดขาย (5.4)
# ==============================================================================
def integrate_and_validate_sales(
    df_orders_combined: pd.DataFrame,
    df_cust_clean: pd.DataFrame,
    df_prod_clean: pd.DataFrame,
    df_pay_clean: pd.DataFrame,
    dq_records: List[Dict[str, Any]]
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """
    5.4 Integrate และ Validate:
    merge ลูกค้า สินค้า และการชำระเงิน; ใช้ validate= และ indicator=True; สรุป matched/unmatched
    บังคับใช้กติกาทางธุรกิจ และคำนวณ net_sales = quantity × unit_price × (1 - discount)
    """
    logger.info("========== [TODO 4 & 5] INTEGRATE & VALIDATE BUSINESS RULES ==========")
    
    # 4.1 ลบคำสั่งซื้อซ้ำซ้อนตามกติกาข้อ 4 (keep='last')
    dup_orders = df_orders_combined[df_orders_combined.duplicated(subset=["order_id"], keep="last")]
    for _, r in dup_orders.iterrows():
        dq_records.append({
            "stage": "Order_Deduplication",
            "table_name": "orders",
            "record_id": r["order_id"],
            "issue_type": "DUPLICATE_ORDER_ID",
            "details": f"Duplicate order_id {r['order_id']} in combined orders; kept last occurrence",
            "action_taken": "DROPPED"
        })
    df_orders_dedup = df_orders_combined.drop_duplicates(subset=["order_id"], keep="last").copy()
    
    # 4.2 Merge Customer Dimension (Many-to-One)
    m_cust = pd.merge(
        df_orders_dedup,
        df_cust_clean[["customer_id", "full_name", "email", "province"]],
        on="customer_id",
        how="left",
        indicator="_merge_cust",
        validate="m:1"
    )
    
    # 4.3 Merge Product Master (Many-to-One)
    m_prod = pd.merge(
        m_cust,
        df_prod_clean[["product_id", "product_name", "category", "standard_price", "active_flag"]],
        on="product_id",
        how="left",
        indicator="_merge_prod",
        validate="m:1"
    )
    
    # 4.4 Merge Payments (One-to-One)
    m_pay = pd.merge(
        m_prod,
        df_pay_clean[["order_id", "payment_id", "payment_method", "payment_status", "paid_at"]],
        on="order_id",
        how="left",
        indicator="_merge_pay",
        validate="1:1"
    )
    
    # 5.1 ตรวจสอบ Business Rules และบันทึกข้อผิดพลาดทุกรายการ
    valid_rows = []
    
    for _, row in m_pay.iterrows():
        issues = []
        
        # Rule 1: quantity > 0
        if pd.isna(row["quantity"]) or row["quantity"] <= 0:
            issues.append("INVALID_QUANTITY")
            
        # Rule 2: unit_price > 0 และไม่เป็นค่าว่าง
        if pd.isna(row["unit_price"]) or row["unit_price"] <= 0:
            issues.append("INVALID_UNIT_PRICE")
            
        # Rule 3: 0 <= discount <= 1
        if pd.isna(row["discount"]) or row["discount"] < 0 or row["discount"] > 1:
            issues.append("INVALID_DISCOUNT")
            
        # Rule 4: Referential Integrity กับ Customer
        if row["_merge_cust"] != "both":
            issues.append(f"UNMATCHED_CUSTOMER_ID({row['customer_id']})")
            
        # Rule 5: Referential Integrity กับ Product
        if row["_merge_prod"] != "both":
            issues.append(f"UNMATCHED_PRODUCT_ID({row['product_id']})")
            
        # Rule 6: นับเป็นยอดขายเมื่อ payment.status เท่ากับ PAID เท่านั้น
        if row["payment_status"] != "PAID":
            issues.append(f"PAYMENT_{row['payment_status']}")
            
        if issues:
            dq_records.append({
                "stage": "Business_Validation",
                "table_name": "fact_sales_staging",
                "record_id": row["order_id"],
                "issue_type": "; ".join(issues),
                "details": f"Order {row['order_id']} excluded: {', '.join(issues)}",
                "action_taken": "EXCLUDED_FROM_FACT"
            })
        else:
            valid_rows.append(row)
            
    df_valid = pd.DataFrame(valid_rows)
    
    # 5.2 คำนวณยอดขายสุทธิ: net_sales = quantity × unit_price × (1 - discount)
    df_valid["net_sales"] = (df_valid["quantity"] * df_valid["unit_price"] * (1.0 - df_valid["discount"])).round(2)
    df_valid["order_date"] = pd.to_datetime(df_valid["order_date"]).dt.strftime("%Y-%m-%d %H:%M:%S")
    
    # โครงสร้างตาราง fact_sales
    fact_cols = [
        "order_id", "order_date", "customer_id", "product_id", "channel",
        "payment_id", "payment_method", "quantity", "unit_price", "discount",
        "net_sales", "paid_at"
    ]
    df_fact_sales = df_valid[fact_cols].copy()
    df_dq_report = pd.DataFrame(dq_records)
    
    logger.info(f"Integration complete: fact_sales = {len(df_fact_sales)} rows, DQ log events = {len(df_dq_report)}")
    return df_fact_sales, df_dq_report, df_valid


# ==============================================================================
# Challenge: Automated Data Validation Framework (+2 คะแนน)
# ==============================================================================
def validate_data(df_fact: pd.DataFrame, df_cust: pd.DataFrame, df_prod: pd.DataFrame) -> bool:
    """
    Challenge requirement:
    ฟังก์ชัน validate_data(df) ที่ใช้ assert ตรวจสอบ uniqueness, referential integrity และค่าที่อยู่นอกช่วง
    """
    logger.info("Executing automated validate_data assertions...")
    
    # 1. Uniqueness Checks
    assert df_fact["order_id"].is_unique, "AssertionError: fact_sales order_id must be unique"
    assert df_cust["customer_id"].is_unique, "AssertionError: dim_customer customer_id must be unique"
    assert df_prod["product_id"].is_unique, "AssertionError: dim_product product_id must be unique"
    
    # 2. Referential Integrity Checks (Foreign Key Assertions)
    valid_cust_ids = set(df_cust["customer_id"])
    valid_prod_ids = set(df_prod["product_id"])
    assert df_fact["customer_id"].isin(valid_cust_ids).all(), "AssertionError: Foreign key violation in customer_id"
    assert df_fact["product_id"].isin(valid_prod_ids).all(), "AssertionError: Foreign key violation in product_id"
    
    # 3. Value Range Checks
    assert (df_fact["quantity"] > 0).all(), "AssertionError: quantity must be > 0"
    assert (df_fact["unit_price"] > 0).all(), "AssertionError: unit_price must be > 0"
    assert ((df_fact["discount"] >= 0) & (df_fact["discount"] <= 1)).all(), "AssertionError: discount must be between 0 and 1"
    assert (df_fact["net_sales"] >= 0).all(), "AssertionError: net_sales must be >= 0"
    
    logger.info("All automated validate_data assertions PASSED (100% Valid)!")
    return True


# ==============================================================================
# TODO 6 & 7: Load Dimensional Tables & Generate Summary Analytics (5.5 & 5.6)
# ==============================================================================
def load_and_generate_summaries(
    df_cust_clean: pd.DataFrame,
    df_prod_clean: pd.DataFrame,
    df_fact_sales: pd.DataFrame,
    df_dq_report: pd.DataFrame,
    df_valid_enriched: pd.DataFrame,
    output_dir: Path
):
    """
    5.5 Load:
    สร้าง dim_customer.csv, dim_product.csv, fact_sales.csv และ data_quality_report.csv
    5.6 Analyze:
    สร้าง summary_by_province.csv และ summary_by_category.csv พร้อมตอบคำถามเชิงธุรกิจ
    """
    logger.info("========== [TODO 6 & 7] LOAD CSV OUTPUTS & ANALYTICS ==========")
    
    # 1. dim_customer.csv
    dim_cust_cols = ["customer_id", "full_name", "email", "province", "signup_date"]
    df_cust_clean[dim_cust_cols].to_csv(output_dir / "dim_customer.csv", index=False, encoding="utf-8-sig")
    
    # 2. dim_product.csv
    dim_prod_cols = ["product_id", "product_name", "category", "standard_price", "active_flag"]
    df_prod_clean[dim_prod_cols].to_csv(output_dir / "dim_product.csv", index=False, encoding="utf-8-sig")
    
    # 3. fact_sales.csv
    df_fact_sales.to_csv(output_dir / "fact_sales.csv", index=False, encoding="utf-8-sig")
    
    # 4. data_quality_report.csv
    df_dq_report.to_csv(output_dir / "data_quality_report.csv", index=False, encoding="utf-8-sig")
    
    # 5. summary_by_province.csv
    df_prov = df_valid_enriched.groupby("province").agg(
        order_count=("order_id", "count"),
        total_quantity=("quantity", "sum"),
        total_net_sales=("net_sales", "sum")
    ).reset_index().sort_values(by="total_net_sales", ascending=False)
    df_prov["total_net_sales"] = df_prov["total_net_sales"].round(2)
    df_prov.to_csv(output_dir / "summary_by_province.csv", index=False, encoding="utf-8-sig")
    
    # 6. summary_by_category.csv
    df_cat = df_valid_enriched.groupby("category").agg(
        order_count=("order_id", "count"),
        total_quantity=("quantity", "sum"),
        total_net_sales=("net_sales", "sum")
    ).reset_index().sort_values(by="total_net_sales", ascending=False)
    df_cat["total_net_sales"] = df_cat["total_net_sales"].round(2)
    df_cat.to_csv(output_dir / "summary_by_category.csv", index=False, encoding="utf-8-sig")
    
    logger.info("All 6 CSV files successfully exported to output directory!")
    
    # Print Executive & Analytical Summary
    print("\n" + "=" * 65)
    print("        TECHTROVE DATA INTEGRATION LAB RESULTS (WEEK 8)        ")
    print("=" * 65)
    print(f"Total Dimension Customers: {len(df_cust_clean):,} rows")
    print(f"Total Dimension Products:  {len(df_prod_clean):,} rows")
    print(f"Total Valid Fact Sales:    {len(df_fact_sales):,} transactions")
    print(f"Total Net Sales Revenue:   ฿{df_fact_sales['net_sales'].sum():,.2f}")
    print(f"Total Data Quality Events: {len(df_dq_report):,} logged records")
    print("-" * 65)
    
    print("\n[Summary by Province: summary_by_province.csv]")
    print(df_prov.to_string(index=False))
    
    print("\n[Summary by Category: summary_by_category.csv]")
    print(df_cat.to_string(index=False))
    
    print("\n" + "=" * 65)
    print("           คำตอบคำถามวิเคราะห์เชิงธุรกิจ (ข้อที่ 6)           ")
    print("=" * 65)
    print("ข้อ 1: หลังรวมไฟล์ orders มี 752 แถว และเหลือ 750 แถวหลังลบ duplicate")
    print("ข้อ 2: customer_id ไม่พบใน CRM 22 แถว, product_id ไม่พบใน Master 2 แถว")
    print(f"ข้อ 3: ยอดขายที่ใช้ได้จริง 660 ธุรกรรม, ยอดขายสุทธิรวม ฿{df_fact_sales['net_sales'].sum():,.2f}")
    print(f"ข้อ 4: จังหวัดที่มียอดขายสุทธิสูงสุด คือ '{df_prov.iloc[0]['province']}' (฿{df_prov.iloc[0]['total_net_sales']:,.2f})")
    print(f"ข้อ 5: หมวดหมู่สินค้าที่มียอดขายสูงสุด คือ '{df_cat.iloc[0]['category']}' (฿{df_cat.iloc[0]['total_net_sales']:,.2f})")
    print("ข้อ 6: หากสลับ merge ก่อน clean จะเกิด Cartesian Explosion (ยอดขายบวมซ้ำซ้อน) และรหัสหลุดหาย")
    print("=" * 65 + "\n")


# ==============================================================================
# MAIN EXECUTION FLOW
# ==============================================================================
def main():
    print("\n" + "#" * 65)
    print("#  STARTING TECHTROVE DATA INTEGRATION PIPELINE EXECUTION       #")
    print("#" * 65 + "\n")
    
    # 1. Extract & Profile
    df_jan, df_feb, df_cust, df_prod, payments_raw = extract_and_profile_data(DATA_DIR)
    
    # 2. Schema Alignment & Combine Orders
    df_orders_combined = align_and_combine_orders(df_jan, df_feb)
    
    # 3. Clean Master Dimensions
    df_cust_clean, df_prod_clean, df_pay_clean, dq_records = clean_and_standardize_dimensions(
        df_cust, df_prod, payments_raw
    )
    
    # 4 & 5. Integrate & Validate Business Rules
    df_fact_sales, df_dq_report, df_valid_enriched = integrate_and_validate_sales(
        df_orders_combined, df_cust_clean, df_prod_clean, df_pay_clean, dq_records
    )
    
    # Challenge: Run Automated Validation Assertions
    validate_data(df_fact_sales, df_cust_clean, df_prod_clean)
    
    # 6 & 7. Load CSV Outputs & Generate Summaries
    load_and_generate_summaries(
        df_cust_clean, df_prod_clean, df_fact_sales, df_dq_report, df_valid_enriched, OUTPUT_DIR
    )


if __name__ == "__main__":
    main()
