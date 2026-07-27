from __future__ import annotations

from pathlib import Path
import re
import sqlite3
import pandas as pd

PROJECT_DIR = Path(__file__).resolve().parent
DATA_PATH = PROJECT_DIR / "retail_logs.csv"
OUTPUT_DIR = PROJECT_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)
DB_PATH = OUTPUT_DIR / "retail_warehouse.db"

PROVINCE_TO_REGION = {
    "Chiang Mai": "North",
    "Bangkok": "Central",
    "Rayong": "East",
    "Chonburi": "East",
    "Phuket": "South",
    "Khon Kaen": "Northeast"
}

def clean_text(value: object, default: str = "Unknown") -> str:
    if pd.isna(value) or str(value).strip() == "":
        return default
    return " ".join(str(value).strip().split()).title()

def clean_store_code(value: object) -> str:
    if pd.isna(value) or str(value).strip() == "":
        return "Unknown"
    return str(value).strip().upper()

def parse_mixed_date(value: object) -> pd.Timestamp:
    text = "" if pd.isna(value) else str(value).strip()
    for fmt in ["%Y-%m-%d", "%d-%b-%Y", "%d/%m/%Y"]:
        parsed = pd.to_datetime(text, format=fmt, errors="coerce")
        if not pd.isna(parsed):
            return parsed
    return pd.NaT

def extract() -> pd.DataFrame:
    df = pd.read_csv(DATA_PATH, dtype=str, encoding="utf-8-sig")
    df.columns = [c.strip().lower() for c in df.columns]
    print(f"[Extract] raw rows={len(df):,}, duplicate sale_id={df.duplicated('sale_id').sum():,}")
    return df

def transform(df: pd.DataFrame):
    # 1. Clean data rows
    df = df.drop_duplicates(subset=["sale_id"], keep="first").copy()
    
    # Transform fields
    df["store_code"] = df["store_code"].apply(clean_store_code)
    df["branch"] = df["branch"].apply(clean_text)
    df["province"] = df["province"].apply(clean_text)
    
    # Handle empty regions by mapping from cleaned province
    df["region"] = df["region"].apply(lambda x: clean_text(x, default=""))
    mask_empty_region = df["region"].eq("") | df["region"].eq("Unknown")
    df.loc[mask_empty_region, "region"] = df.loc[mask_empty_region, "province"].map(PROVINCE_TO_REGION).fillna("Unknown")
    
    # Product cleaning
    df["product_name"] = df["product_name"].apply(clean_text)
    df["category"] = df["category"].apply(clean_text)
    
    # Date parsing
    df["sale_date"] = df["sale_date"].apply(parse_mixed_date)
    
    # Numeric values
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce").fillna(0).astype(int)
    df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce").fillna(0.0)
    df["discount_percent"] = pd.to_numeric(df["discount_percent"], errors="coerce").fillna(0.0)
    
    # Calculate Total Amount: (quantity * unit_price) * (1 - discount_percent / 100)
    df["total_amount"] = (df["quantity"] * df["unit_price"]) * (1 - df["discount_percent"] / 100.0)
    df["total_amount"] = df["total_amount"].round(2)
    
    # Drop rows that are missing critical business data
    df = df.dropna(subset=["sale_date"])
    df = df[(df["quantity"] > 0) & (df["unit_price"] >= 0)].copy()

    # 2. Build Dimension Tables
    # dim_location
    location_candidates = df[["store_code", "branch", "province", "region"]].copy()
    dim_location = (location_candidates.sort_values(["store_code", "branch"])
                    .drop_duplicates("store_code")
                    .reset_index(drop=True))
    dim_location.insert(0, "location_id", range(1, len(dim_location) + 1))

    # dim_product
    product_candidates = df[["product_name", "category"]].copy()
    dim_product = (product_candidates.sort_values(["product_name", "category"])
                   .drop_duplicates("product_name")
                   .reset_index(drop=True))
    dim_product.insert(0, "product_id", range(1, len(dim_product) + 1))

    # dim_date
    dim_date = df[["sale_date"]].drop_duplicates().sort_values("sale_date").reset_index(drop=True).rename(columns={"sale_date": "full_date"})
    dim_date["date_id"] = dim_date["full_date"].dt.strftime("%Y%m%d").astype(int)
    dim_date["day"] = dim_date["full_date"].dt.day
    dim_date["month"] = dim_date["full_date"].dt.month
    dim_date["month_name"] = dim_date["full_date"].dt.month_name()
    dim_date["quarter"] = "Q" + dim_date["full_date"].dt.quarter.astype(str)
    dim_date["year"] = dim_date["full_date"].dt.year
    dim_date["full_date"] = dim_date["full_date"].dt.strftime("%Y-%m-%d")
    dim_date = dim_date[["date_id", "full_date", "day", "month", "month_name", "quarter", "year"]]

    # 3. Build Fact Table by merging keys
    mapped = df.merge(dim_location[["location_id", "store_code"]], on="store_code", how="left")
    mapped = mapped.merge(dim_product[["product_id", "product_name"]], on="product_name", how="left")
    
    # Need to match sale_date back to full_date (which was string-formatted or timestamp-matched)
    mapped["sale_date_str"] = mapped["sale_date"].dt.strftime("%Y-%m-%d")
    mapped = mapped.merge(dim_date[["date_id", "full_date"]], left_on="sale_date_str", right_on="full_date", how="left")

    fact_sales = mapped[["sale_id", "location_id", "product_id", "date_id", "quantity", "unit_price", "discount_percent", "total_amount"]]
    fact_sales = fact_sales.rename(columns={"sale_id": "transaction_id"})
    
    fact_sales[["location_id", "product_id", "date_id"]] = fact_sales[["location_id", "product_id", "date_id"]].astype(int)

    print(f"[Transform] locations={len(dim_location)}, products={len(dim_product)}, dates={len(dim_date)}, facts={len(fact_sales)}")
    return dim_location, dim_product, dim_date, fact_sales

def load(dim_location, dim_product, dim_date, fact_sales):
    with sqlite3.connect(DB_PATH) as conn:
        conn.execute("PRAGMA foreign_keys = ON;")
        conn.executescript("""
        DROP TABLE IF EXISTS fact_sales;
        DROP TABLE IF EXISTS dim_location;
        DROP TABLE IF EXISTS dim_product;
        DROP TABLE IF EXISTS dim_date;
        
        CREATE TABLE dim_location (
            location_id INTEGER PRIMARY KEY,
            store_code TEXT NOT NULL UNIQUE,
            branch TEXT NOT NULL,
            province TEXT NOT NULL,
            region TEXT NOT NULL
        );
        
        CREATE TABLE dim_product (
            product_id INTEGER PRIMARY KEY,
            product_name TEXT NOT NULL UNIQUE,
            category TEXT NOT NULL
        );
        
        CREATE TABLE dim_date (
            date_id INTEGER PRIMARY KEY,
            full_date TEXT NOT NULL UNIQUE,
            day INTEGER NOT NULL,
            month INTEGER NOT NULL,
            month_name TEXT NOT NULL,
            quarter TEXT NOT NULL,
            year INTEGER NOT NULL
        );
        
        CREATE TABLE fact_sales (
            transaction_id TEXT PRIMARY KEY,
            location_id INTEGER NOT NULL REFERENCES dim_location(location_id),
            product_id INTEGER NOT NULL REFERENCES dim_product(product_id),
            date_id INTEGER NOT NULL REFERENCES dim_date(date_id),
            quantity INTEGER NOT NULL CHECK(quantity > 0),
            unit_price REAL NOT NULL CHECK(unit_price >= 0),
            discount_percent REAL NOT NULL CHECK(discount_percent >= 0),
            total_amount REAL NOT NULL CHECK(total_amount >= 0)
        );
        """)
        dim_location.to_sql("dim_location", conn, if_exists="append", index=False)
        dim_product.to_sql("dim_product", conn, if_exists="append", index=False)
        dim_date.to_sql("dim_date", conn, if_exists="append", index=False)
        fact_sales.to_sql("fact_sales", conn, if_exists="append", index=False)
        conn.commit()
    print(f"[Load] database={DB_PATH}")

def verify():
    # Verify by calculating Total Revenue per Region and Branch
    sql = """
    SELECT l.region, l.branch, ROUND(SUM(f.total_amount), 2) AS total_revenue
    FROM fact_sales f
    JOIN dim_location l ON f.location_id = l.location_id
    GROUP BY l.region, l.branch
    ORDER BY total_revenue DESC;
    """
    with sqlite3.connect(DB_PATH) as conn:
        result = pd.read_sql_query(sql, conn)
    print("\n[Verify] Revenue Summary by Store Branch:")
    print(result.to_string(index=False))

if __name__ == "__main__":
    raw = extract()
    dim_loc, dim_prod, dim_dt, fact = transform(raw)
    load(dim_loc, dim_prod, dim_dt, fact)
    verify()
    print("\nETL Pipeline run successfully for homework!")
