"""
SmartMart Retail Analytics
Standalone desktop GUI for chunk-based retail CSV processing.

Requirements:
    pip install pandas matplotlib openpyxl

Run:
    python smartmart_analytics.py

Notes:
- The supplied assignment dataset is already CSV, so it can be selected directly.
- Excel (.xlsx/.xls) inputs are also supported and converted to CSV in chunks where possible.
- CSV/log readers are separated from processing logic.
- Invalid rows are written to a dead-letter CSV with a rejection reason.
"""

from __future__ import annotations
import csv
import logging
import os
import platform
import sys
import queue
import threading
import time
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    import matplotlib
    matplotlib.use("TkAgg")
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    from matplotlib.figure import Figure
    CHARTS_AVAILABLE = True
except Exception:
    CHARTS_AVAILABLE = False


APP_NAME = "SmartMart Retail Analytics"
EXPECTED_COLUMNS = [
    "transaction_id", "transaction_datetime", "store_id", "store_city",
    "region", "warehouse_id", "customer_id", "customer_age",
    "customer_gender", "customer_city", "product_category", "product_name",
    "quantity", "unit_price", "discount_percent", "promotion_name",
    "payment_method", "channel", "total_amount", "stock_on_hand",
    "reorder_level", "transaction_type",
]
NUMERIC_COLUMNS = [
    "customer_age", "quantity", "unit_price", "discount_percent",
    "total_amount", "stock_on_hand", "reorder_level",
]
REQUIRED_COLUMNS = [
    "transaction_id", "transaction_datetime", "store_id", "region",
    "product_category", "product_name", "quantity", "unit_price",
    "discount_percent", "total_amount",
]


@dataclass
class ProcessingSummary:
    source_file: str
    total_rows: int = 0
    accepted_rows: int = 0
    rejected_rows: int = 0
    duplicate_rows: int = 0
    sales_rows: int = 0
    return_rows: int = 0
    revenue: float = 0.0
    quantity_sold: float = 0.0
    elapsed_seconds: float = 0.0
    output_dir: str = ""
    errors: int = 0


class DataSourceFactory:
    """Creates the appropriate reader based on the selected file extension."""

    @staticmethod
    def create_reader(path: str, chunk_size: int):
        suffix = Path(path).suffix.lower()
        if suffix == ".csv":
            return CSVReader(path, chunk_size)
        if suffix in (".xlsx", ".xls"):
            return ExcelReader(path, chunk_size)
        if suffix in (".log", ".txt"):
            return LogReader(path, chunk_size)
        raise ValueError("Unsupported file type. Select CSV, Excel (.xlsx/.xls), LOG or TXT.")


class CSVReader:
    """Reads CSV input incrementally to avoid loading the whole file into memory."""

    def __init__(self, path: str, chunk_size: int):
        self.path = path
        self.chunk_size = chunk_size

    def read_chunks(self):
        # utf-8-sig handles common CSV files exported with a BOM.
        yield from pd.read_csv(
            self.path,
            chunksize=self.chunk_size,
            encoding="utf-8-sig",
            low_memory=False,
            on_bad_lines="warn",
        )


class ExcelReader:
    """Reads an Excel workbook and yields DataFrame chunks.

    Pandas/openpyxl may load workbook structures before yielding rows, so CSV is
    recommended for the 50MB+ assignment dataset and best memory efficiency.
    """

    def __init__(self, path: str, chunk_size: int):
        self.path = path
        self.chunk_size = chunk_size

    def read_chunks(self):
        try:
            frame = pd.read_excel(self.path, engine="openpyxl")
        except ImportError as exc:
            raise RuntimeError("Excel support requires openpyxl. Run: pip install openpyxl") from exc
        for start in range(0, len(frame), self.chunk_size):
            yield frame.iloc[start:start + self.chunk_size].copy()


class LogReader:
    """Reads simple delimited log files; each non-empty line becomes a record."""

    def __init__(self, path: str, chunk_size: int):
        self.path = path
        self.chunk_size = chunk_size

    def read_chunks(self):
        batch = []
        with open(self.path, "r", encoding="utf-8", errors="replace") as handle:
            for line_number, line in enumerate(handle, start=1):
                line = line.strip()
                if not line:
                    continue
                # Supports key=value fields separated by commas; otherwise stores raw line.
                record = {"line_number": line_number, "raw_record": line}
                if "=" in line:
                    for part in line.split(","):
                        if "=" in part:
                            key, value = part.split("=", 1)
                            record[key.strip()] = value.strip()
                batch.append(record)
                if len(batch) >= self.chunk_size:
                    yield pd.DataFrame(batch)
                    batch.clear()
        if batch:
            yield pd.DataFrame(batch)


class ErrorLogger:
    """Central application logger."""

    def __init__(self, output_dir: Path):
        self.path = output_dir / "smartmart_processing.log"
        self.logger = logging.getLogger("SmartMart")
        self.logger.setLevel(logging.INFO)
        for old_handler in list(self.logger.handlers):
            self.logger.removeHandler(old_handler)
            try:
                old_handler.close()
            except Exception:
                pass
        handler = logging.FileHandler(self.path, encoding="utf-8")
        handler.setFormatter(logging.Formatter("%(asctime)s | %(levelname)s | %(message)s"))
        self.logger.addHandler(handler)

    def info(self, message: str):
        self.logger.info(message)

    def error(self, message: str):
        self.logger.error(message)


class DeadLetterQueue:
    """Stores rejected records and the reason each record was rejected."""

    def __init__(self, output_path: Path):
        self.output_path = output_path
        self.first_write = True

    def write(self, rejected: pd.DataFrame):
        if rejected.empty:
            return
        rejected.to_csv(
            self.output_path,
            mode="w" if self.first_write else "a",
            header=self.first_write,
            index=False,
            encoding="utf-8-sig",
        )
        self.first_write = False


class DataValidator:
    """Validates required fields and numeric business rules."""

    @staticmethod
    def validate(frame: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
        data = frame.copy()
        data.columns = [str(c).strip().lower() for c in data.columns]

        # This application is designed for the SmartMart transaction schema.
        missing = [column for column in REQUIRED_COLUMNS if column not in data.columns]
        if missing:
            raise ValueError(
                "The selected dataset does not contain the required SmartMart columns: "
                + ", ".join(missing)
            )

        for column in NUMERIC_COLUMNS:
            if column in data.columns:
                data[column] = pd.to_numeric(data[column], errors="coerce")

        if "transaction_datetime" in data.columns:
            data["transaction_datetime"] = pd.to_datetime(
                data["transaction_datetime"], errors="coerce", dayfirst=False
            )

        reasons = pd.Series("", index=data.index, dtype="object")

        def add_reason(mask: pd.Series, text: str):
            nonlocal reasons
            mask = mask.fillna(False)
            empty_reason = mask & reasons.eq("")
            existing_reason = mask & reasons.ne("")
            reasons.loc[empty_reason] = text
            reasons.loc[existing_reason] = reasons.loc[existing_reason] + "; " + text

        for column in ["transaction_id", "store_id", "region", "product_category", "product_name"]:
            add_reason(data[column].isna() | data[column].astype(str).str.strip().eq(""), f"Missing {column}")

        add_reason(data["transaction_datetime"].isna(), "Invalid transaction_datetime")
        add_reason(data["quantity"].isna(), "Invalid quantity")
        add_reason(data["unit_price"].isna(), "Invalid unit_price")
        add_reason(data["discount_percent"].isna(), "Invalid discount_percent")
        add_reason(data["total_amount"].isna(), "Invalid total_amount")
        is_return_record = data.get("transaction_type", pd.Series("Sale", index=data.index)).astype(str).str.lower().str.contains("return|refund", regex=True)
        add_reason(data["quantity"].fillna(0) == 0, "Quantity is zero")
        add_reason((data["quantity"].fillna(0) < 0) & ~is_return_record, "Negative quantity on non-return record")
        add_reason(data["unit_price"].fillna(-1) < 0, "Negative unit_price")
        add_reason(
            (data["discount_percent"].fillna(-1) < 0) | (data["discount_percent"].fillna(101) > 100),
            "discount_percent outside 0-100",
        )
        add_reason((data["total_amount"].fillna(-1) < 0) & ~is_return_record, "Negative total_amount on non-return record")

        data["rejection_reason"] = reasons
        rejected = data.loc[reasons.ne("")].copy()
        accepted = data.loc[reasons.eq("")].copy()
        rejected["rejection_reason"] = rejected["rejection_reason"].astype(str)
        accepted.drop(columns=["rejection_reason"], inplace=True, errors="ignore")
        return accepted, rejected


class TransformationPipeline:
    """Applies reusable cleaning and enrichment operations to accepted records."""

    @staticmethod
    def apply(frame: pd.DataFrame) -> pd.DataFrame:
        data = frame.copy()
        data["transaction_date"] = data["transaction_datetime"].dt.date.astype(str)
        data["transaction_hour"] = data["transaction_datetime"].dt.hour
        data["transaction_day"] = data["transaction_datetime"].dt.day_name()
        data["transaction_month"] = data["transaction_datetime"].dt.month
        data["calculated_amount"] = (
            data["quantity"] * data["unit_price"] *
            (1 - data["discount_percent"] / 100)
        ).round(2)
        if "region" in data.columns:
            data["region"] = data["region"].astype(str).str.strip().str.title()
        if "product_category" in data.columns:
            data["product_category"] = data["product_category"].astype(str).str.strip().str.title()
        return data


class AnalyticsService:
    """Calculates key retail analytics from each accepted chunk."""

    @staticmethod
    def analyse(frame: pd.DataFrame, accumulators: dict):
        if frame.empty:
            return
        work = frame.copy()
        type_series = work.get("transaction_type", pd.Series("Sale", index=work.index))
        is_return = type_series.astype(str).str.lower().str.contains("return|refund", regex=True)
        is_sale = ~is_return

        accumulators["sales_rows"] += int(is_sale.sum())
        accumulators["return_rows"] += int(is_return.sum())
        accumulators["revenue"] += float(work.loc[is_sale, "total_amount"].fillna(0).sum())
        accumulators["quantity_sold"] += float(work.loc[is_sale, "quantity"].fillna(0).sum())

        by_hour = work.loc[is_sale].groupby("transaction_hour")["total_amount"].sum()
        for key, value in by_hour.items():
            accumulators["hour_sales"][int(key)] = accumulators["hour_sales"].get(int(key), 0.0) + float(value)

        by_day = work.loc[is_sale].groupby("transaction_day")["total_amount"].sum()
        for key, value in by_day.items():
            accumulators["day_sales"][str(key)] = accumulators["day_sales"].get(str(key), 0.0) + float(value)

        if "transaction_month" in work.columns:
            by_month = work.loc[is_sale].groupby("transaction_month")["total_amount"].sum()
            for key, value in by_month.items():
                accumulators["month_sales"][int(key)] = accumulators["month_sales"].get(int(key), 0.0) + float(value)

        if "promotion_name" in work.columns:
            promo_work = work.loc[is_sale].copy()
            promo_work["promotion_name"] = promo_work["promotion_name"].fillna("No promotion").astype(str).str.strip()
            promo_work.loc[promo_work["promotion_name"].eq(""), "promotion_name"] = "No promotion"
            by_promotion = promo_work.groupby("promotion_name")["total_amount"].sum()
            for key, value in by_promotion.items():
                accumulators["promotion_sales"][str(key)] = accumulators["promotion_sales"].get(str(key), 0.0) + float(value)

        by_store = work.loc[is_sale].groupby(["store_id", "store_city", "region"], dropna=False)["total_amount"].sum()
        for key, value in by_store.items():
            key_string = " | ".join(str(part) for part in key)
            accumulators["store_sales"][key_string] = accumulators["store_sales"].get(key_string, 0.0) + float(value)

        by_category = work.loc[is_sale].groupby("product_category")["total_amount"].sum()
        for key, value in by_category.items():
            accumulators["category_sales"][str(key)] = accumulators["category_sales"].get(str(key), 0.0) + float(value)

        region_work = work.loc[is_sale].copy()
        region_work["region"] = region_work["region"].fillna("Unknown").astype(str).str.strip().str.title()
        region_work.loc[region_work["region"].eq(""), "region"] = "Unknown"
        by_region = region_work.groupby("region")["total_amount"].sum()
        for key, value in by_region.items():
            accumulators["region_sales"][str(key)] = accumulators["region_sales"].get(str(key), 0.0) + float(value)

        if "customer_id" in work.columns:
            customer_rows = work.loc[is_sale & work["customer_id"].notna()]
            grouped = customer_rows.groupby("customer_id").agg(
                transactions=("transaction_id", "count"),
                spending=("total_amount", "sum"),
            )
            for customer_id, row in grouped.iterrows():
                item = accumulators["customers"].setdefault(str(customer_id), [0, 0.0])
                item[0] += int(row["transactions"])
                item[1] += float(row["spending"])

        # Daily totals support sales-drop detection after all chunks are processed.
        if "transaction_date" in work.columns:
            daily = work.loc[is_sale].groupby("transaction_date")["total_amount"].sum()
            for key, value in daily.items():
                accumulators["date_sales"][str(key)] = accumulators["date_sales"].get(str(key), 0.0) + float(value)

        # Time-based heatmap data: weekday by hour.
        if "transaction_day" in work.columns and "transaction_hour" in work.columns:
            heat = work.loc[is_sale].groupby(["transaction_day", "transaction_hour"])["total_amount"].sum()
            for (day, hour), value in heat.items():
                key = f"{day}|{int(hour)}"
                accumulators["weekday_hour_sales"][key] = accumulators["weekday_hour_sales"].get(key, 0.0) + float(value)

        # Demographic and geographic breakdowns for customer segmentation.
        for column, target in (("customer_age", "age_sales"), ("customer_gender", "gender_sales"), ("customer_city", "customer_city_sales"), ("store_city", "store_city_sales")):
            if column in work.columns:
                subset = work.loc[is_sale & work[column].notna()]
                if column == "customer_age":
                    bins = [0, 17, 24, 34, 44, 54, 64, 120]
                    labels = ["0-17", "18-24", "25-34", "35-44", "45-54", "55-64", "65+"]
                    categories = pd.cut(subset[column], bins=bins, labels=labels, include_lowest=True).astype(str)
                else:
                    categories = subset[column].astype(str).str.strip().replace("", "Unknown")
                grouped = subset.assign(_category=categories).groupby("_category", observed=False)["total_amount"].sum()
                for key, value in grouped.items():
                    accumulators[target][str(key)] = accumulators[target].get(str(key), 0.0) + float(value)

        # Count transaction and return activity by store for return-rate checks.
        if "store_id" in work.columns:
            for store, count in work.groupby("store_id").size().items():
                accumulators["store_transaction_rows"][str(store)] = accumulators["store_transaction_rows"].get(str(store), 0) + int(count)
            for store, count in work.loc[is_return].groupby("store_id").size().items():
                accumulators["store_return_rows"][str(store)] = accumulators["store_return_rows"].get(str(store), 0) + int(count)

        # Configurable rule-based anomaly flags.
        if "discount_percent" in work.columns:
            discount_mask = work["discount_percent"] >= 70
            for _, row in work.loc[discount_mask].iterrows():
                accumulators["anomalies"].append({
                    "anomaly_type": "High discount",
                    "transaction_id": row.get("transaction_id", ""),
                    "store_id": row.get("store_id", ""),
                    "value": row.get("discount_percent", ""),
                    "details": "Discount is 70% or higher",
                })
        if "stock_on_hand" in work.columns:
            stock_mask = work["stock_on_hand"] < 0
            for _, row in work.loc[stock_mask].iterrows():
                accumulators["anomalies"].append({
                    "anomaly_type": "Negative stock",
                    "transaction_id": row.get("transaction_id", ""),
                    "store_id": row.get("store_id", ""),
                    "value": row.get("stock_on_hand", ""),
                    "details": "Stock on hand is negative",
                })


def empty_accumulators():
    """Create all aggregate values used by AnalyticsService and report generation."""
    return {
        "sales_rows": 0,
        "return_rows": 0,
        "revenue": 0.0,
        "quantity_sold": 0.0,
        "hour_sales": {},
        "day_sales": {},
        "month_sales": {},
        "promotion_sales": {},
        "store_sales": {},
        "category_sales": {},
        "region_sales": {},
        "customers": {},
        "date_sales": {},
        "weekday_hour_sales": {},
        "age_sales": {},
        "gender_sales": {},
        "customer_city_sales": {},
        "store_city_sales": {},
        "store_transaction_rows": {},
        "store_return_rows": {},
        "anomalies": [],
    }


class SmartMartProcessor:
    """Runs the complete chunk-based processing workflow."""

    def __init__(self, source_path: str, output_dir: str, chunk_size: int,
                 progress_callback: Callable[[int, str], None],
                 log_callback: Callable[[str], None],
                 cancel_event: threading.Event):
        self.source_path = source_path
        self.output_dir = Path(output_dir)
        self.chunk_size = chunk_size
        self.progress_callback = progress_callback
        self.log_callback = log_callback
        self.cancel_event = cancel_event

    def run(self) -> tuple[ProcessingSummary, dict]:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        logger = ErrorLogger(self.output_dir)
        dlq = DeadLetterQueue(self.output_dir / "dead_letter_queue.csv")
        summary = ProcessingSummary(source_file=self.source_path, output_dir=str(self.output_dir))
        accum = empty_accumulators()
        cleaned_path = self.output_dir / "cleaned_transactions.csv"
        duplicate_ids = set()
        first_cleaned_write = True
        started = time.time()

        # Remove previous outputs to avoid appending to an old job.
        for filename in [
            "cleaned_transactions.csv", "dead_letter_queue.csv", "anomalies.csv",
            "peak_sales_by_hour.csv", "sales_by_day.csv", "sales_by_month.csv", "sales_by_promotion.csv", "sales_by_store.csv",
            "sales_by_category.csv", "sales_by_region.csv", "customer_segments.csv",
            "sales_by_weekday_hour_heatmap.csv", "sales_by_age_group.csv", "sales_by_gender.csv",
            "sales_by_customer_city.csv", "sales_by_store_city.csv", "sales_by_date.csv",
            "processing_summary.txt",
        ]:
            target = self.output_dir / filename
            if target.exists():
                target.unlink()

        reader = DataSourceFactory.create_reader(self.source_path, self.chunk_size)
        self.log_callback(f"Input selected: {self.source_path}")
        self.log_callback(f"Chunk size: {self.chunk_size:,}")
        logger.info(f"Started processing {self.source_path}")

        try:
            for chunk_number, raw_chunk in enumerate(reader.read_chunks(), start=1):
                if self.cancel_event.is_set():
                    self.log_callback("Cancellation requested. Stopping after the current chunk.")
                    break

                summary.total_rows += len(raw_chunk)
                raw_chunk.columns = [str(c).strip().lower() for c in raw_chunk.columns]

                accepted, rejected = DataValidator.validate(raw_chunk)
                if not accepted.empty:
                    # Remove duplicates across chunks by transaction_id.
                    transaction_ids = accepted["transaction_id"].astype(str)
                    duplicate_mask = transaction_ids.isin(duplicate_ids) | transaction_ids.duplicated(keep="first")
                    new_ids = transaction_ids.loc[~duplicate_mask]
                    duplicate_ids.update(new_ids.tolist())
                    duplicate_rows = accepted.loc[duplicate_mask].copy()
                    if not duplicate_rows.empty:
                        duplicate_rows["rejection_reason"] = "Duplicate transaction_id"
                        rejected = pd.concat([rejected, duplicate_rows], ignore_index=True)
                        summary.duplicate_rows += len(duplicate_rows)
                    accepted = accepted.loc[~duplicate_mask].copy()

                if not rejected.empty:
                    dlq.write(rejected)
                    summary.rejected_rows += len(rejected)
                    logger.info(f"Chunk {chunk_number}: rejected {len(rejected)} records")

                transformed = TransformationPipeline.apply(accepted) if not accepted.empty else accepted
                if not transformed.empty:
                    transformed.to_csv(
                        cleaned_path,
                        mode="w" if first_cleaned_write else "a",
                        header=first_cleaned_write,
                        index=False,
                        encoding="utf-8-sig",
                    )
                    first_cleaned_write = False
                    summary.accepted_rows += len(transformed)
                    AnalyticsService.analyse(transformed, accum)

                self.progress_callback(0, f"Processed chunk {chunk_number:,} — {summary.total_rows:,} rows read")
                self.log_callback(
                    f"Chunk {chunk_number}: read {len(raw_chunk):,}, accepted {len(transformed):,}, "
                    f"rejected {len(rejected):,}"
                )
                logger.info(
                    f"Chunk {chunk_number}: read={len(raw_chunk)}, accepted={len(transformed)}, rejected={len(rejected)}"
                )

        except Exception as exc:
            summary.errors += 1
            logger.error(f"Processing failed: {exc}\n{traceback.format_exc()}")
            self.log_callback(f"ERROR: {exc}")
            raise
        finally:
            summary.elapsed_seconds = time.time() - started

        # Detect large day-over-day sales drops (20% or more) after totals are complete.
        ordered_dates = sorted(accum["date_sales"].items())
        for index in range(1, len(ordered_dates)):
            previous_date, previous_value = ordered_dates[index - 1]
            current_date, current_value = ordered_dates[index]
            if previous_value > 0:
                drop_percent = (previous_value - current_value) / previous_value * 100
                if drop_percent >= 20:
                    accum["anomalies"].append({
                        "anomaly_type": "Sudden sales drop", "transaction_id": "",
                        "store_id": "All stores", "value": round(drop_percent, 2),
                        "details": f"Revenue dropped {drop_percent:.1f}% from {previous_date} ({previous_value:.2f}) to {current_date} ({current_value:.2f})"
                    })

        # Flag store-level return rates above 10% (only where there is activity).
        for store, total_count in accum["store_transaction_rows"].items():
            return_count = accum["store_return_rows"].get(store, 0)
            rate = (return_count / total_count * 100) if total_count else 0
            if return_count and rate >= 10:
                accum["anomalies"].append({
                    "anomaly_type": "Abnormal return rate", "transaction_id": "",
                    "store_id": store, "value": round(rate, 2),
                    "details": f"{return_count} return/refund rows out of {total_count} transactions ({rate:.1f}%)"
                })

        self._write_outputs(accum, summary)
        self.progress_callback(100, "Processing finished")
        self.log_callback(f"Completed in {summary.elapsed_seconds:.1f} seconds.")
        logger.info("Processing finished")
        return summary, accum

    def _write_outputs(self, accum: dict, summary: ProcessingSummary):
        out = self.output_dir
        pd.DataFrame(
            [{"hour": k, "revenue": v} for k, v in sorted(accum["hour_sales"].items())]
        ).to_csv(out / "peak_sales_by_hour.csv", index=False, encoding="utf-8-sig")
        pd.DataFrame(
            [{"day": k, "revenue": v} for k, v in sorted(accum["day_sales"].items())]
        ).to_csv(out / "sales_by_day.csv", index=False, encoding="utf-8-sig")
        month_names = {1:"January", 2:"February", 3:"March", 4:"April", 5:"May", 6:"June", 7:"July", 8:"August", 9:"September", 10:"October", 11:"November", 12:"December"}
        pd.DataFrame(
            [{"month_number": k, "month": month_names.get(k, str(k)), "revenue": v} for k, v in sorted(accum["month_sales"].items())]
        ).to_csv(out / "sales_by_month.csv", index=False, encoding="utf-8-sig")
        pd.DataFrame(
            [{"promotion": k, "revenue": v} for k, v in sorted(accum["promotion_sales"].items(), key=lambda x: x[1], reverse=True)]
        ).to_csv(out / "sales_by_promotion.csv", index=False, encoding="utf-8-sig")
        pd.DataFrame(
            [{"store": k, "revenue": v} for k, v in sorted(accum["store_sales"].items(), key=lambda x: x[1], reverse=True)]
        ).to_csv(out / "sales_by_store.csv", index=False, encoding="utf-8-sig")
        pd.DataFrame(
            [{"category": k, "revenue": v} for k, v in sorted(accum["category_sales"].items(), key=lambda x: x[1], reverse=True)]
        ).to_csv(out / "sales_by_category.csv", index=False, encoding="utf-8-sig")
        pd.DataFrame(
            [{"region": k, "revenue": v} for k, v in sorted(accum["region_sales"].items(), key=lambda x: x[1], reverse=True)]
        ).to_csv(out / "sales_by_region.csv", index=False, encoding="utf-8-sig")
        segments = []
        for customer_id, (transactions, spending) in accum["customers"].items():
            segment = "One-time" if transactions == 1 else ("Occasional" if transactions <= 5 else "Frequent")
            segments.append({
                "customer_id": customer_id, "transaction_count": transactions,
                "total_spending": round(spending, 2), "segment": segment,
            })
        pd.DataFrame(segments).to_csv(out / "customer_segments.csv", index=False, encoding="utf-8-sig")
        pd.DataFrame(accum["anomalies"], columns=["anomaly_type", "transaction_id", "store_id", "value", "details"]).to_csv(
            out / "anomalies.csv", index=False, encoding="utf-8-sig"
        )
        weekday_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
        heat_rows = []
        for key, value in accum["weekday_hour_sales"].items():
            day, hour = key.rsplit("|", 1)
            heat_rows.append({"weekday": day, "hour": int(hour), "revenue": value})
        heat_df = pd.DataFrame(heat_rows, columns=["weekday", "hour", "revenue"])
        if not heat_df.empty:
            heat_df["_day_order"] = heat_df["weekday"].map({day: i for i, day in enumerate(weekday_order)})
            heat_df.sort_values(["_day_order", "hour"]).drop(columns="_day_order").to_csv(out / "sales_by_weekday_hour_heatmap.csv", index=False, encoding="utf-8-sig")
        else:
            heat_df.to_csv(out / "sales_by_weekday_hour_heatmap.csv", index=False, encoding="utf-8-sig")
        for accumulator_key, filename, label in [
            ("age_sales", "sales_by_age_group.csv", "age_group"),
            ("gender_sales", "sales_by_gender.csv", "gender"),
            ("customer_city_sales", "sales_by_customer_city.csv", "customer_city"),
            ("store_city_sales", "sales_by_store_city.csv", "store_city"),
            ("date_sales", "sales_by_date.csv", "date"),
        ]:
            pd.DataFrame([{label: key, "revenue": value} for key, value in sorted(accum[accumulator_key].items(), key=lambda item: item[1], reverse=(accumulator_key != "date_sales"))]).to_csv(out / filename, index=False, encoding="utf-8-sig")

        summary_text = (
            f"{APP_NAME} — Processing Summary\n"
            f"{'=' * 42}\n"
            f"Input file: {summary.source_file}\n"
            f"Total rows read: {summary.total_rows:,}\n"
            f"Accepted rows: {summary.accepted_rows:,}\n"
            f"Rejected rows: {summary.rejected_rows:,}\n"
            f"Duplicate rows: {summary.duplicate_rows:,}\n"
            f"Sale rows: {accum['sales_rows']:,}\n"
            f"Return/refund rows: {accum['return_rows']:,}\n"
            f"Calculated sales revenue (source total_amount): {accum['revenue']:,.2f}\n"
            f"Quantity sold: {accum['quantity_sold']:,.0f}\n"
            f"Elapsed seconds: {summary.elapsed_seconds:.1f}\n"
            f"Output folder: {summary.output_dir}\n"
        )
        (out / "processing_summary.txt").write_text(summary_text, encoding="utf-8")


class SmartMartGUI(tk.Tk):
    """Tkinter GUI for selecting, processing and exploring a retail dataset."""

    def __init__(self):
        super().__init__()
        self.title(APP_NAME)
        self.geometry("1360x850")
        self.minsize(1050, 680)
        self.configure(bg="#f5f2fb")
        self.events = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker = None
        self.last_summary = None
        self.last_accum = None
        self.source_var = tk.StringVar()
        self.output_var = tk.StringVar(value=str(Path.cwd() / "smartmart_output"))
        self.chunk_var = tk.StringVar(value="20000")
        self.status_var = tk.StringVar(value="Ready — select a dataset to begin.")
        self._style()
        self._build_ui()
        self.after(100, self._poll_events)

    def _style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TFrame", background="#f5f2fb")
        style.configure("TLabel", background="#f5f2fb", foreground="#244653", font=("Segoe UI", 10))
        style.configure("TButton", font=("Segoe UI", 10), padding=8)
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), foreground="#ffffff", background="#7046a8")
        style.map("Accent.TButton", background=[("active", "#583486"), ("disabled", "#c9bfd6")])
        style.configure("TNotebook", background="#f5f2fb", borderwidth=0)
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=(13, 8))
        style.configure("Treeview", rowheight=29, font=("Segoe UI", 9), background="#ffffff", fieldbackground="#ffffff")
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

    def _build_ui(self):
        # Main shell: teal navigation sidebar + light dashboard workspace.
        shell = tk.Frame(self, bg="#f5f2fb")
        shell.pack(fill="both", expand=True)

        sidebar = tk.Frame(shell, bg="#45266f", width=190)
        sidebar.pack(side="left", fill="y")
        sidebar.pack_propagate(False)
        tk.Label(sidebar, text="SMARTMART", font=("Segoe UI", 17, "bold"),
                 bg="#45266f", fg="white").pack(anchor="w", padx=18, pady=(24, 2))
        tk.Label(sidebar, text="RETAIL ANALYTICS", font=("Segoe UI", 8, "bold"),
                 bg="#45266f", fg="#d9c8f2").pack(anchor="w", padx=19, pady=(0, 28))

        nav_items = [
            ("▦   Dashboard", "Dashboard"),
            ("◉   Data Processing", "Data Processing"),
            ("▤   Analytics Reports", "Analytics Reports"),
            ("≡   Processing log", "Processing log"),
            ("⚙   System", "System"),
        ]
        self.nav_buttons = {}
        for label, tab_name in nav_items:
            btn = tk.Button(sidebar, text=label, anchor="w", relief="flat", bd=0,
                            bg="#5b348a", fg="white", activebackground="#7046a8",
                            activeforeground="white", font=("Segoe UI", 10, "bold"),
                            padx=18, pady=12, cursor="hand2",
                            command=lambda name=tab_name: self._select_tab(name))
            btn.pack(fill="x", padx=10, pady=4)
            self.nav_buttons[tab_name] = btn
        tk.Frame(sidebar, bg="#45266f").pack(fill="both", expand=True)
        tk.Label(sidebar, text="SmartMart Retail Group\nData processing workspace",
                 justify="left", bg="#45266f", fg="#ded4ef",
                 font=("Segoe UI", 8)).pack(anchor="w", padx=18, pady=18)

        workspace = tk.Frame(shell, bg="#f5f2fb")
        workspace.pack(side="left", fill="both", expand=True)

        topbar = tk.Frame(workspace, bg="#ffffff", padx=24, pady=16,
                          highlightbackground="#e5dcf1", highlightthickness=1)
        topbar.pack(fill="x")
        top_text = tk.Frame(topbar, bg="#ffffff")
        top_text.pack(side="left", fill="x", expand=True)
        tk.Label(top_text, text="SmartMart Retail Analytics", bg="#ffffff", fg="#45266f",
                 font=("Segoe UI", 20, "bold")).pack(anchor="w")
        tk.Label(top_text, text="Retail intelligence  •  Data quality  •  Daily insights",
                 bg="#ffffff", fg="#81728f", font=("Segoe UI", 10)).pack(anchor="w", pady=(3, 0))
        tk.Label(topbar, text="●  SYSTEM READY", bg="#eee5fa", fg="#684095",
                 font=("Segoe UI", 9, "bold"), padx=12, pady=8).pack(side="right", padx=(12, 0))

        controls = tk.Frame(workspace, bg="#f5f2fb", padx=20, pady=12)
        controls.pack(fill="x")
        tk.Label(controls, text="Input dataset", bg="#f5f2fb", fg="#59466f",
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 8), pady=5)
        self.source_entry = ttk.Entry(controls, textvariable=self.source_var)
        self.source_entry.grid(row=0, column=1, sticky="ew", pady=5)
        ttk.Button(controls, text="Browse…", command=self._browse_source).grid(row=0, column=2, padx=6, pady=5)

        tk.Label(controls, text="Output folder", bg="#f5f2fb", fg="#59466f",
                 font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", padx=(0, 8), pady=5)
        ttk.Entry(controls, textvariable=self.output_var).grid(row=1, column=1, sticky="ew", pady=5)
        ttk.Button(controls, text="Choose…", command=self._browse_output).grid(row=1, column=2, padx=6, pady=5)

        tk.Label(controls, text="Chunk size", bg="#f5f2fb", fg="#59466f",
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=3, padx=(12, 6), pady=5)
        ttk.Combobox(controls, textvariable=self.chunk_var, values=["5000", "10000", "20000", "50000"],
                     width=9, state="readonly").grid(row=0, column=4, pady=5)
        self.start_button = ttk.Button(controls, text="▶  Process dataset", style="Accent.TButton", command=self._start)
        self.start_button.grid(row=1, column=3, padx=(12, 6), pady=5, sticky="ew")
        self.cancel_button = ttk.Button(controls, text="Cancel", command=self._cancel, state="disabled")
        self.cancel_button.grid(row=1, column=4, pady=5, sticky="ew")
        controls.columnconfigure(1, weight=1)

        progress_frame = tk.Frame(workspace, bg="#f5f2fb", padx=20)
        progress_frame.pack(fill="x", pady=(0, 8))
        self.progress = ttk.Progressbar(progress_frame, mode="indeterminate", maximum=100)
        self.progress.pack(fill="x")
        tk.Label(progress_frame, textvariable=self.status_var, bg="#f5f2fb", fg="#756581",
                 font=("Segoe UI", 9)).pack(anchor="w", pady=(4, 0))

        self.notebook = ttk.Notebook(workspace)
        self.notebook.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        self.dashboard_tab = ttk.Frame(self.notebook, padding=0)
        self.processing_tab = ttk.Frame(self.notebook, padding=12)
        self.results_tab = ttk.Frame(self.notebook, padding=12)
        self.logs_tab = ttk.Frame(self.notebook, padding=12)
        self.system_tab = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(self.dashboard_tab, text="Dashboard")
        self.notebook.add(self.processing_tab, text="Data Processing")
        self.notebook.add(self.results_tab, text="Analytics Reports")
        self.notebook.add(self.logs_tab, text="Processing log")
        self.notebook.add(self.system_tab, text="System")
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self._build_dashboard_scroll_area()
        self._build_dashboard()
        self._build_processing_page()
        self._build_results()
        self._build_logs()
        self._build_system_page()
        self._select_tab("Dashboard")

    def _select_tab(self, tab_name):
        tab_map = {
            "Dashboard": self.dashboard_tab,
            "Data Processing": self.processing_tab,
            "Analytics Reports": self.results_tab,
            "Processing log": self.logs_tab,
            "System": self.system_tab,
        }
        if tab_name in tab_map:
            self.notebook.select(tab_map[tab_name])
        self._on_tab_changed(None)

    def _on_tab_changed(self, _event):
        try:
            selected = self.notebook.tab(self.notebook.select(), "text")
            for name, button in self.nav_buttons.items():
                active = name == selected
                button.configure(bg="#7046a8" if active else "#5b348a")
        except (tk.TclError, AttributeError):
            pass

    def _build_dashboard_scroll_area(self):
        """Make the complete dashboard vertically scrollable on smaller screens."""
        self.dashboard_canvas = tk.Canvas(self.dashboard_tab, bg="#f5f2fb", highlightthickness=0)
        self.dashboard_scrollbar = ttk.Scrollbar(
            self.dashboard_tab, orient="vertical", command=self.dashboard_canvas.yview
        )
        self.dashboard_canvas.configure(yscrollcommand=self.dashboard_scrollbar.set)
        self.dashboard_scrollbar.pack(side="right", fill="y")
        self.dashboard_canvas.pack(side="left", fill="both", expand=True)
        self.dashboard_content = ttk.Frame(self.dashboard_canvas, padding=12)
        self.dashboard_window = self.dashboard_canvas.create_window(
            (0, 0), window=self.dashboard_content, anchor="nw"
        )
        self.dashboard_content.bind(
            "<Configure>",
            lambda _event: self.after_idle(self._refresh_dashboard_scrollregion),
        )
        self.dashboard_canvas.bind(
            "<Configure>",
            self._resize_dashboard_content,
        )

        # Bind at the application level so the wheel also works when the pointer
        # is over a chart, label, text box, or other child widget.
        self.bind_all("<MouseWheel>", self._dashboard_mousewheel, add="+")
        self.bind_all("<Button-4>", self._dashboard_mousewheel_linux, add="+")
        self.bind_all("<Button-5>", self._dashboard_mousewheel_linux, add="+")

    def _refresh_dashboard_scrollregion(self):
        """Recalculate the scrollable height after dashboard widgets resize."""
        try:
            self.dashboard_canvas.configure(
                scrollregion=self.dashboard_canvas.bbox("all") or (0, 0, 0, 0)
            )
        except tk.TclError:
            pass

    def _resize_dashboard_content(self, event):
        try:
            self.dashboard_canvas.itemconfigure(
                self.dashboard_window, width=event.width
            )
            self.after_idle(self._refresh_dashboard_scrollregion)
        except tk.TclError:
            pass

    def _dashboard_mousewheel(self, event):
        """Scroll smoothly on Windows/macOS, including small touchpad deltas."""
        try:
            if self.notebook.select() != str(self.dashboard_tab):
                return
            delta = getattr(event, "delta", 0)
            if not delta:
                return
            units = max(1, round(abs(delta) / 120))
            self.dashboard_canvas.yview_scroll(-units if delta > 0 else units, "units")
            return "break"
        except (AttributeError, tk.TclError):
            return

    def _dashboard_mousewheel_linux(self, event):
        try:
            if self.notebook.select() != str(self.dashboard_tab):
                return
            self.dashboard_canvas.yview_scroll(-1 if event.num == 4 else 1, "units")
            return "break"
        except (AttributeError, tk.TclError):
            return

    def _build_dashboard(self):
        self.kpi_frame = ttk.Frame(self.dashboard_content)
        self.kpi_frame.pack(fill="x")
        self.kpi_labels = {}
        kpis = [
            ("rows", "Total records", "#245bff"),
            ("accepted", "Valid records", "#16864a"),
            ("rejected", "Invalid records", "#d62828"),
            ("duplicates", "Duplicates", "#c45a00"),
            ("revenue", "Sales revenue", "#7046a8"),
        ]
        for i, (key, title, value_color) in enumerate(kpis):
            card = tk.Frame(self.kpi_frame, bg="#ffffff", highlightbackground="#e5dcf1", highlightthickness=1, padx=12, pady=13)
            card.grid(row=0, column=i, sticky="nsew", padx=4, pady=5)
            tk.Label(card, text=title.upper(), bg="white", fg="#81728f", font=("Segoe UI", 8, "bold")).pack(anchor="w")
            value = tk.Label(card, text="—", bg="white", fg=value_color, font=("Segoe UI", 17, "bold"))
            value.pack(anchor="w", pady=(7, 0))
            self.kpi_labels[key] = value
            self.kpi_frame.columnconfigure(i, weight=1)

        self.summary_box = tk.Text(self.dashboard_content, height=8, wrap="word", font=("Segoe UI", 10),
                                   bg="#ffffff", fg="#40334f", relief="flat", padx=16, pady=14,
                                   highlightbackground="#e5dcf1", highlightthickness=1)
        self.summary_box.pack(fill="x", padx=5, pady=10)
        self.summary_box.insert("1.0", "Select a CSV dataset and click “Process dataset”.\n"
                                      "Recommended for the supplied 60MB dataset: CSV input and chunk size 20,000.\n"
                                      "Excel input is supported, but converting large Excel workbooks to CSV first is recommended.")
        self.summary_box.configure(state="disabled")

        if CHARTS_AVAILABLE:
            chart_controls = ttk.Frame(self.dashboard_content)
            chart_controls.pack(fill="x", padx=5, pady=(8, 0))
            ttk.Label(chart_controls, text="Visualisation:").pack(side="left")
            self.chart_var = tk.StringVar(value="Peak sales by hour")
            self.chart_select = ttk.Combobox(chart_controls, textvariable=self.chart_var, state="readonly", width=32,
                values=["Peak sales by hour", "Sales by day", "Monthly sales trend", "Product performance", "Regional performance", "Sales by promotion", "Sales by age group", "Sales by gender", "Sales by customer city", "Sales by store city", "Weekday-hour heatmap"])
            self.chart_select.pack(side="left", padx=8)
            ttk.Button(chart_controls, text="Open Charts Window", command=self._open_charts_window).pack(side="right", padx=6)
            self.chart_select.bind("<<ComboboxSelected>>", lambda _event: self._draw_selected_chart())
            chart_frame = ttk.Frame(self.dashboard_content)
            chart_frame.pack(fill="both", expand=True, padx=5, pady=5)
            self.figure = Figure(figsize=(9, 3.2), dpi=90, facecolor="#ffffff")
            self.ax = self.figure.add_subplot(111)
            self.ax.set_facecolor("#ffffff")
            self.ax.set_title("Peak sales by transaction hour", color="#45266f", pad=14, fontweight="bold")
            self.ax.set_xlabel("Hour of day")
            self.ax.set_ylabel("Revenue")
            self.canvas = FigureCanvasTkAgg(self.figure, master=chart_frame)
            self.canvas.get_tk_widget().configure(height=300)
            self.canvas.get_tk_widget().pack(fill="x", expand=False)
            self._build_dashboard_results_table()
        else:
            ttk.Label(self.dashboard_content, text="Install matplotlib to display embedded charts: pip install matplotlib").pack(anchor="w")

    def _build_results(self):
        """Build the searchable Analytics Reports workspace."""
        heading = ttk.Frame(self.results_tab)
        heading.pack(fill="x", pady=(0, 3))
        ttk.Label(
            heading, text="Analytics Reports",
            font=("Segoe UI", 17, "bold")
        ).pack(side="left")
        ttk.Label(
            heading, text="SEARCH AND REVIEW GENERATED REPORTS"
        ).pack(side="right")

        controls = ttk.Frame(self.results_tab)
        controls.pack(fill="x", pady=(10, 10))

        ttk.Label(controls, text="Report:").grid(row=0, column=0, padx=(0, 6), sticky="w")
        self.table_var = tk.StringVar(value="Peak sales by hour")
        self.table_select = ttk.Combobox(
            controls, textvariable=self.table_var, state="readonly", width=25,
            values=[
                "Peak sales by hour", "Sales by day", "Sales by month",
                "Sales by promotion", "Sales by store", "Sales by category", "Product performance",
                "Sales by region", "Regional performance", "Customer segments", "Weekday-hour heatmap",
                "Sales by age group", "Sales by gender", "Sales by customer city",
                "Sales by store city", "Sales by date", "Anomalies",
                "Dead-letter records"
            ]
        )
        self.table_select.grid(row=0, column=1, padx=(0, 12), sticky="w")
        self.table_select.bind("<<ComboboxSelected>>", lambda _event: self._load_table())

        ttk.Label(controls, text="Search:").grid(row=0, column=2, padx=(0, 6), sticky="w")
        self.report_search_var = tk.StringVar()
        self.report_search_entry = ttk.Entry(
            controls, textvariable=self.report_search_var, width=20
        )
        self.report_search_entry.grid(row=0, column=3, padx=(0, 8), sticky="ew")
        self.report_search_entry.bind("<Return>", lambda _event: self._search_table())

        self.report_count_var = tk.StringVar(value="0 of 0 rows")
        ttk.Label(
            controls, textvariable=self.report_count_var, width=17,
            anchor="center"
        ).grid(row=0, column=4, padx=(0, 8))

        ttk.Button(controls, text="Search", command=self._search_table).grid(row=0, column=5, padx=3)
        ttk.Button(controls, text="Clear", command=self._clear_table_search).grid(row=0, column=6, padx=3)
        ttk.Button(controls, text="View Charts", command=self._view_report_charts).grid(row=0, column=7, padx=3)
        ttk.Button(controls, text="Open Outputs", command=self._open_output).grid(row=0, column=8, padx=(3, 0))
        controls.columnconfigure(3, weight=1)

        table_frame = ttk.Frame(self.results_tab)
        table_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(table_frame, show="headings")
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        self.current_report_df = pd.DataFrame()
        self.filtered_report_df = pd.DataFrame()
        self._load_table()

    def _build_logs(self):
        log_frame = ttk.Frame(self.logs_tab)
        log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, wrap="word", bg="#2b1b3f", fg="#f1e9fb",
                                insertbackground="white", font=("Consolas", 9), relief="flat")
        log_scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=log_scroll.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")
        self.log_text.insert("end", "SmartMart application log is ready.\n")
        self.log_text.configure(state="disabled")

    def _browse_source(self):
        path = filedialog.askopenfilename(
            title="Select retail dataset",
            filetypes=[
                ("Supported datasets", "*.csv *.xlsx *.xls *.log *.txt"),
                ("CSV files", "*.csv"), ("Excel files", "*.xlsx *.xls"),
                ("Log/text files", "*.log *.txt"), ("All files", "*.*"),
            ],
        )
        if path:
            self.source_var.set(path)
            if not self.output_var.get().strip():
                self.output_var.set(str(Path(path).parent / "smartmart_output"))
            self._refresh_system_page()

    def _browse_output(self):
        path = filedialog.askdirectory(title="Choose output folder")
        if path:
            self.output_var.set(path)
            self._refresh_system_page()

    def _start(self):
        source = self.source_var.get().strip()
        output = self.output_var.get().strip()
        if not source or not Path(source).is_file():
            messagebox.showerror("Input required", "Please select an existing dataset file.")
            return
        if not output:
            messagebox.showerror("Output required", "Please choose an output folder.")
            return
        try:
            chunk_size = int(self.chunk_var.get())
            if chunk_size < 1:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid chunk size", "Choose a positive chunk size.")
            return
        self.cancel_event.clear()
        self.start_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.progress.configure(mode="indeterminate")
        self.progress.start(10)
        self.status_var.set("Processing dataset…")
        self._append_log("Starting processing job.")
        processor = SmartMartProcessor(
            source, output, chunk_size,
            lambda percent, text: self.events.put(("progress", percent, text)),
            lambda text: self.events.put(("log", text)),
            self.cancel_event,
        )
        self.worker = threading.Thread(target=self._run_worker, args=(processor,), daemon=True)
        self.worker.start()

    def _run_worker(self, processor):
        try:
            summary, accum = processor.run()
            self.events.put(("done", summary, accum))
        except Exception as exc:
            self.events.put(("failed", str(exc)))

    def _cancel(self):
        self.cancel_event.set()
        self.status_var.set("Cancellation requested; waiting for current chunk to finish…")
        self._append_log("Cancellation requested by user.")

    def _poll_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "log":
                    self._append_log(event[1])
                elif event[0] == "progress":
                    self.status_var.set(event[2])
                elif event[0] == "done":
                    self.progress.stop()
                    self.progress.configure(mode="determinate", value=100)
                    self.start_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    self.last_summary, self.last_accum = event[1], event[2]
                    self._show_summary(self.last_summary, self.last_accum)
                    self._load_table()
                    self.status_var.set("Processing completed. Outputs are saved in the selected folder.")
                    messagebox.showinfo("Processing complete", f"Dataset processing completed.\n\nOutput folder:\n{self.last_summary.output_dir}")
                elif event[0] == "failed":
                    self.progress.stop()
                    self.progress.configure(mode="determinate", value=0)
                    self.start_button.configure(state="normal")
                    self.cancel_button.configure(state="disabled")
                    self.status_var.set("Processing failed.")
                    self._append_log("FAILED: " + event[1])
                    messagebox.showerror("Processing failed", event[1])
        except queue.Empty:
            pass
        self.after(100, self._poll_events)

    def _summary_text(self, summary: ProcessingSummary, accum: dict) -> str:
        return (
            f"Run status: Completed\n"
            f"Rows: {summary.total_rows:,} total, {summary.accepted_rows:,} valid, "
            f"{summary.rejected_rows:,} invalid, {summary.duplicate_rows:,} duplicates\n"
            f"Revenue: {accum['revenue']:,.2f}    Returns/refunds: {accum['return_rows']:,}\n"
            f"Validation notes: check dead_letter_queue.csv for rejected rows and reasons.\n"
            f"Processing time: {summary.elapsed_seconds:.1f} seconds\n"
            f"Output folder: {summary.output_dir}"
        )

    def _build_dashboard_results_table(self):
        frame = ttk.LabelFrame(self.dashboard_content, text="Analytics Results", padding=8)
        frame.pack(fill="x", padx=5, pady=(8, 12))
        columns = ("analysis", "rows", "status", "file")
        self.dashboard_results_tree = ttk.Treeview(frame, columns=columns, show="headings", height=5)
        for key, title, width in [("analysis", "Analysis", 190), ("rows", "Result Rows", 100), ("status", "Status", 110), ("file", "Output File", 250)]:
            self.dashboard_results_tree.heading(key, text=title)
            self.dashboard_results_tree.column(key, width=width, anchor="w")
        self.dashboard_results_tree.pack(fill="x", expand=False)
        self.dashboard_results_tree.insert("", "end", values=("No processing run yet", "—", "Waiting", "—"))

    def _refresh_dashboard_results_table(self, summary):
        if not hasattr(self, "dashboard_results_tree"):
            return
        self.dashboard_results_tree.delete(*self.dashboard_results_tree.get_children())
        output = Path(summary.output_dir)
        for label, filename in [
            ("Peak Sales", "peak_sales_by_hour.csv"),
            ("Customer Segmentation", "customer_segments.csv"),
            ("Seasonal Promotion", "sales_by_month.csv"),
            ("Product Performance", "sales_by_category.csv"),
            ("Regional Performance", "sales_by_region.csv"),
            ("Anomaly Detection", "anomalies.csv"),
        ]:
            path = output / filename
            try:
                rows = sum(1 for _ in path.open("r", encoding="utf-8-sig")) - 1 if path.exists() else 0
            except OSError:
                rows = 0
            self.dashboard_results_tree.insert("", "end", values=(label, f"{max(0, rows):,}", "Completed" if path.exists() else "Unavailable", filename))

    def _build_processing_page(self):
        ttk.Label(self.processing_tab, text="Data Processing", font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(0, 10))
        ttk.Label(self.processing_tab, text="Select a dataset, run validation and review the latest processing status.").pack(anchor="w", pady=(0, 12))
        dataset_frame = ttk.LabelFrame(self.processing_tab, text="Selected Dataset", padding=12)
        dataset_frame.pack(fill="x", pady=(0, 10))
        self.processing_dataset_label = ttk.Label(dataset_frame, text="No dataset selected", wraplength=950)
        self.processing_dataset_label.pack(anchor="w", fill="x")
        buttons = ttk.Frame(self.processing_tab)
        buttons.pack(fill="x", pady=(0, 10))
        ttk.Button(buttons, text="Browse CSV", command=self._browse_source).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="Process Dataset", style="Accent.TButton", command=self._start).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="Open Outputs", command=self._open_output).pack(side="left", padx=(0, 8))
        ttk.Button(buttons, text="View Charts", command=self._open_charts_window).pack(side="left", padx=(0, 8))
        ttk.Label(self.processing_tab, text="LATEST RUN", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 4))
        self.processing_latest = tk.Text(self.processing_tab, height=11, wrap="word", font=("Segoe UI", 10), bg="#ffffff", fg="#40334f", relief="flat", padx=12, pady=10)
        self.processing_latest.pack(fill="both", expand=True)
        self.processing_latest.insert("1.0", "No processing run yet. Select a dataset above and click Process Dataset.")
        self.processing_latest.configure(state="disabled")
        self.source_var.trace_add("write", lambda *_: self._update_dataset_label())
        self._update_dataset_label()

    def _update_dataset_label(self):
        if hasattr(self, "processing_dataset_label"):
            self.processing_dataset_label.configure(text=self.source_var.get().strip() or "No dataset selected")

    def _build_system_page(self):
        ttk.Label(self.system_tab, text="System Information", font=("Segoe UI", 18, "bold")).pack(anchor="w", pady=(0, 12))
        self.system_info = tk.Text(self.system_tab, height=14, wrap="word", font=("Consolas", 10), bg="#ffffff", fg="#40334f", relief="flat", padx=14, pady=12)
        self.system_info.pack(fill="both", expand=True)
        ttk.Button(self.system_tab, text="Refresh System Information", command=self._refresh_system_page).pack(anchor="e", pady=(10, 0))
        self._refresh_system_page()

    def _refresh_system_page(self):
        if not hasattr(self, "system_info"):
            return
        output = Path(self.output_var.get().strip() or ".")
        lines = [
            f"Application: {APP_NAME}",
            f"Python version: {sys.version.split()[0]}",
            f"Operating system: {platform.platform()}",
            f"Pandas version: {pd.__version__}",
            f"Matplotlib charts available: {'Yes' if CHARTS_AVAILABLE else 'No'}",
            f"Selected dataset: {self.source_var.get().strip() or 'Not selected'}",
            f"Output directory: {output}",
            f"Output directory exists: {'Yes' if output.exists() else 'No'}",
            f"Latest processing status: {'Completed' if self.last_summary else 'Not run yet'}",
        ]
        self.system_info.configure(state="normal")
        self.system_info.delete("1.0", "end")
        self.system_info.insert("1.0", "\n".join(lines))
        self.system_info.configure(state="disabled")

    def _open_charts_window(self):
        if not CHARTS_AVAILABLE:
            messagebox.showwarning("Charts unavailable", "Install matplotlib with: pip install matplotlib")
            return
        if not self.last_accum:
            messagebox.showinfo("No analytics yet", "Process a dataset before opening charts.")
            return
        window = tk.Toplevel(self)
        window.title("SmartMart Retail Analytics | Charts")
        window.geometry("1050x650")
        window.minsize(760, 480)
        toolbar = ttk.Frame(window, padding=8)
        toolbar.pack(fill="x")
        ttk.Label(toolbar, text="Select Chart:").pack(side="left")
        choice = tk.StringVar(value=self.chart_var.get())
        options = ["Peak sales by hour", "Sales by day", "Monthly sales trend", "Product performance", "Regional performance", "Sales by promotion", "Sales by age group", "Sales by gender", "Sales by customer city", "Sales by store city", "Weekday-hour heatmap"]
        select = ttk.Combobox(toolbar, textvariable=choice, values=options, state="readonly", width=30)
        select.pack(side="left", padx=8)
        fig = Figure(figsize=(10, 5.5), dpi=100, facecolor="white")
        ax = fig.add_subplot(111)
        canvas = FigureCanvasTkAgg(fig, master=window)
        canvas.get_tk_widget().pack(fill="both", expand=True, padx=8, pady=(0, 8))
        def draw(_event=None):
            self._draw_chart_on_axes(ax, choice.get())
            fig.tight_layout()
            canvas.draw_idle()
        select.bind("<<ComboboxSelected>>", draw)
        draw()

    def _draw_chart_on_axes(self, ax, chart_name):
        accum = self.last_accum or {}
        ax.clear()
        ax.set_facecolor("white")
        ax.tick_params(axis="x", labelrotation=25, labelsize=8)
        if chart_name == "Peak sales by hour":
            items = sorted(accum.get("hour_sales", {}).items())
            ax.bar([str(k) for k, _ in items], [v for _, v in items], color="#8054b8")
            ax.set_title("Peak Sales by Hour", fontweight="bold", color="#45266f")
            ax.set_xlabel("Hour of day"); ax.set_ylabel("Revenue")
        elif chart_name == "Sales by day":
            order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            items = [(k, accum.get("day_sales", {}).get(k, 0)) for k in order if k in accum.get("day_sales", {})]
            ax.bar([k[:3] for k, _ in items], [v for _, v in items], color="#8054b8")
            ax.set_title("Sales by Day", fontweight="bold", color="#45266f"); ax.set_ylabel("Revenue")
        elif chart_name == "Monthly sales trend":
            items = sorted(accum.get("month_sales", {}).items())
            ax.plot([k for k, _ in items], [v for _, v in items], marker="o", color="#8054b8")
            ax.set_title("Monthly Sales Trend", fontweight="bold", color="#45266f"); ax.set_xlabel("Month"); ax.set_ylabel("Revenue")
        elif chart_name == "Product performance":
            items = sorted(accum.get("category_sales", {}).items(), key=lambda x: x[1], reverse=True)[:12]
            ax.barh([str(k) for k, _ in reversed(items)], [v for _, v in reversed(items)], color="#8054b8")
            ax.set_title("Top Product Categories by Revenue", fontweight="bold", color="#45266f"); ax.set_xlabel("Revenue")
        elif chart_name == "Regional performance":
            combined = {}
            for region, value in accum.get("region_sales", {}).items():
                label = str(region).strip().title() or "Unknown"
                combined[label] = combined.get(label, 0) + float(value)
            items = sorted(combined.items(), key=lambda x: x[1], reverse=True)
            ax.barh([k for k, _ in reversed(items)], [v for _, v in reversed(items)], color="#8054b8")
            ax.set_title("Revenue by Region", fontweight="bold", color="#45266f"); ax.set_xlabel("Revenue")
        elif chart_name == "Sales by promotion":
            items = sorted(accum.get("promotion_sales", {}).items(), key=lambda x: x[1], reverse=True)[:12]
            ax.barh([str(k) for k, _ in reversed(items)], [v for _, v in reversed(items)], color="#8054b8")
            ax.set_title("Sales by Promotion", fontweight="bold", color="#45266f"); ax.set_xlabel("Revenue")
        elif chart_name in ("Sales by age group", "Sales by gender", "Sales by customer city", "Sales by store city"):
            key_map = {"Sales by age group":"age_sales", "Sales by gender":"gender_sales", "Sales by customer city":"customer_city_sales", "Sales by store city":"store_city_sales"}
            items = sorted(accum.get(key_map[chart_name], {}).items(), key=lambda x: x[1], reverse=True)[:12]
            ax.barh([str(k) for k, _ in reversed(items)], [v for _, v in reversed(items)], color="#8054b8")
            ax.set_title(chart_name, fontweight="bold", color="#45266f"); ax.set_xlabel("Revenue")
        elif chart_name == "Weekday-hour heatmap":
            days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            matrix = [[accum.get("weekday_hour_sales", {}).get(f"{day}|{hour}", 0.0) for hour in range(24)] for day in days]
            ax.imshow(matrix, aspect="auto", cmap="Purples")
            ax.set_yticks(range(len(days)), days); ax.set_xticks(range(0, 24, 2))
            ax.set_title("Sales Heatmap: Weekday vs Hour", fontweight="bold", color="#45266f"); ax.set_xlabel("Hour of day")

    def _show_summary(self, summary: ProcessingSummary, accum: dict):
        self.kpi_labels["rows"].configure(text=f"{summary.total_rows:,}")
        self.kpi_labels["accepted"].configure(text=f"{summary.accepted_rows:,}")
        self.kpi_labels["rejected"].configure(text=f"{summary.rejected_rows:,}")
        self.kpi_labels["duplicates"].configure(text=f"{summary.duplicate_rows:,}")
        self.kpi_labels["revenue"].configure(text=f"{accum['revenue']:,.0f}")
        if hasattr(self, "processing_latest"):
            self.processing_latest.configure(state="normal")
            self.processing_latest.delete("1.0", "end")
            self.processing_latest.insert("1.0", self._summary_text(summary, accum))
            self.processing_latest.configure(state="disabled")
        if hasattr(self, "system_info"):
            self._refresh_system_page()
        text = (
            f"Input: {summary.source_file}\n"
            f"Total rows read: {summary.total_rows:,}    Accepted: {summary.accepted_rows:,}    "
            f"Rejected: {summary.rejected_rows:,}    Duplicates: {summary.duplicate_rows:,}\n"
            f"Sale records: {accum['sales_rows']:,}    Returns/refunds: {accum['return_rows']:,}\n"
            f"Revenue from source total_amount (sale records): {accum['revenue']:,.2f}\n"
            f"Quantity sold: {accum['quantity_sold']:,.0f}    Processing time: {summary.elapsed_seconds:.1f}s\n"
            f"Output folder: {summary.output_dir}\n\n"
            "Generated files: cleaned_transactions.csv, dead_letter_queue.csv, anomalies.csv, "
            "peak_sales_by_hour.csv, sales_by_day.csv, sales_by_month.csv, sales_by_promotion.csv, sales_by_store.csv, sales_by_category.csv, "
            "sales_by_region.csv, customer_segments.csv, sales_by_weekday_hour_heatmap.csv, sales_by_age_group.csv, sales_by_gender.csv, sales_by_customer_city.csv, sales_by_store_city.csv, sales_by_date.csv, processing_summary.txt, smartmart_processing.log"
        )
        self.summary_box.configure(state="normal")
        self.summary_box.delete("1.0", "end")
        self.summary_box.insert("1.0", text)
        self.summary_box.configure(state="disabled")
        self._refresh_dashboard_results_table(summary)
        if CHARTS_AVAILABLE:
            self._draw_selected_chart()

    def _draw_selected_chart(self):
        """Render one of the available analytics visualisations in the dashboard."""
        if not CHARTS_AVAILABLE or not self.last_accum:
            return
        accum = self.last_accum
        chart_name = self.chart_var.get() if hasattr(self, "chart_var") else "Peak sales by hour"
        self.ax.clear()
        self.ax.set_facecolor("#ffffff")
        self.ax.set_ylabel("Revenue")
        self.ax.tick_params(axis="x", labelrotation=25, labelsize=8)
        if chart_name == "Peak sales by hour":
            items = sorted(accum["hour_sales"].items())
            self.ax.bar([str(k) for k, _ in items], [v for _, v in items], color="#8054b8")
            self.ax.set_title("Peak sales by transaction hour", color="#45266f", pad=14, fontweight="bold")
            self.ax.set_xlabel("Hour of day")
        elif chart_name == "Sales by day":
            order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            items = [(k, accum["day_sales"].get(k, 0)) for k in order if k in accum["day_sales"]]
            self.ax.bar([k[:3] for k, _ in items], [v for _, v in items], color="#a58bd1")
            self.ax.set_title("Sales by day of week", color="#45266f", pad=14, fontweight="bold")
        elif chart_name == "Monthly sales trend":
            items = sorted(accum["month_sales"].items())
            self.ax.plot([k for k, _ in items], [v for _, v in items], marker="o", linewidth=2, color="#7046a8")
            self.ax.set_xticks(range(1, 13))
            self.ax.set_title("Monthly sales trend", color="#45266f", pad=14, fontweight="bold")
            self.ax.set_xlabel("Month number")
        elif chart_name == "Product performance":
            items = sorted(accum["category_sales"].items(), key=lambda item: item[1], reverse=True)[:12]
            self.ax.barh([k for k, _ in reversed(items)], [v for _, v in reversed(items)], color="#8054b8")
            self.ax.set_title("Top Product Categories by Revenue", color="#45266f", pad=14, fontweight="bold")
            self.ax.set_xlabel("Revenue")
            self.ax.set_ylabel("Product category")
        elif chart_name == "Regional performance":
            # Combine case variants so East and EAST do not appear as separate bars.
            region_totals = {}
            for region, value in accum["region_sales"].items():
                label = str(region).strip().title() or "Unknown"
                region_totals[label] = region_totals.get(label, 0.0) + float(value)
            items = sorted(region_totals.items(), key=lambda item: item[1], reverse=True)
            self.ax.barh([k for k, _ in reversed(items)], [v for _, v in reversed(items)], color="#8054b8")
            self.ax.set_title("Revenue by Region", color="#45266f", pad=14, fontweight="bold")
            self.ax.set_xlabel("Revenue")
            self.ax.set_ylabel("Region")
        elif chart_name == "Sales by promotion":
            items = sorted(accum["promotion_sales"].items(), key=lambda item: item[1], reverse=True)[:12]
            self.ax.barh([k for k, _ in reversed(items)], [v for _, v in reversed(items)], color="#8054b8")
            self.ax.set_title("Sales by Promotion", color="#45266f", pad=14, fontweight="bold")
            self.ax.set_xlabel("Revenue")
            self.ax.set_ylabel("Promotion")
        elif chart_name in ("Sales by age group", "Sales by gender", "Sales by customer city", "Sales by store city"):
            key_map = {"Sales by age group":"age_sales", "Sales by gender":"gender_sales", "Sales by customer city":"customer_city_sales", "Sales by store city":"store_city_sales"}
            items = sorted(accum[key_map[chart_name]].items(), key=lambda item: item[1], reverse=True)[:12]
            self.ax.barh([k for k, _ in reversed(items)], [v for _, v in reversed(items)], color="#8054b8")
            self.ax.set_title(chart_name, color="#45266f", pad=14, fontweight="bold")
            self.ax.set_xlabel("Revenue")
            self.ax.set_ylabel("")
        elif chart_name == "Weekday-hour heatmap":
            days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
            matrix = [[accum["weekday_hour_sales"].get(f"{day}|{hour}", 0.0) for hour in range(24)] for day in days]
            self.ax.imshow(matrix, aspect="auto", cmap="Purples")
            self.ax.set_yticks(range(len(days)), days)
            self.ax.set_xticks(range(0, 24, 2))
            self.ax.set_title("Sales heatmap: weekday vs hour", color="#45266f", pad=14, fontweight="bold")
            self.ax.set_xlabel("Hour of day")
            self.ax.set_ylabel("Weekday")
        self.figure.tight_layout()
        self.canvas.draw_idle()
        # Matplotlib can change the embedded widget's requested height; update
        # the scrollable region after Tk has processed the geometry change.
        self.after_idle(self._refresh_dashboard_scrollregion)

    def _report_file_map(self):
        return {
            "Peak sales by hour": "peak_sales_by_hour.csv",
            "Sales by day": "sales_by_day.csv",
            "Sales by month": "sales_by_month.csv",
            "Sales by promotion": "sales_by_promotion.csv",
            "Sales by store": "sales_by_store.csv",
            "Sales by category": "sales_by_category.csv",
            "Product performance": "sales_by_category.csv",
            "Sales by region": "sales_by_region.csv",
            "Regional performance": "sales_by_region.csv",
            "Customer segments": "customer_segments.csv",
            "Weekday-hour heatmap": "sales_by_weekday_hour_heatmap.csv",
            "Sales by age group": "sales_by_age_group.csv",
            "Sales by gender": "sales_by_gender.csv",
            "Sales by customer city": "sales_by_customer_city.csv",
            "Sales by store city": "sales_by_store_city.csv",
            "Sales by date": "sales_by_date.csv",
            "Anomalies": "anomalies.csv",
            "Dead-letter records": "dead_letter_queue.csv",
        }

    def _load_table(self):
        """Load the selected generated report and display all its rows."""
        output = Path(self.output_var.get().strip() or ".")
        filename = self._report_file_map().get(self.table_var.get())
        path = output / filename if filename else None

        if path is None or not path.exists():
            self.current_report_df = pd.DataFrame()
            self.filtered_report_df = pd.DataFrame()
            self._populate_tree([], [])
            self.report_count_var.set("0 of 0 rows")
            return

        try:
            self.current_report_df = pd.read_csv(path, encoding="utf-8-sig", low_memory=False)
            self._search_table()
        except Exception as exc:
            self.current_report_df = pd.DataFrame()
            self.filtered_report_df = pd.DataFrame()
            self._populate_tree([], [])
            self.report_count_var.set("Unable to load report")
            self._append_log(f"Could not display table {path.name}: {exc}")

    def _search_table(self):
        """Filter report rows by a case-insensitive search across all columns."""
        if self.current_report_df is None or self.current_report_df.empty:
            self.filtered_report_df = pd.DataFrame()
            self._populate_tree([], [])
            self.report_count_var.set("0 of 0 rows")
            return

        query = self.report_search_var.get().strip() if hasattr(self, "report_search_var") else ""
        if query:
            searchable = self.current_report_df.fillna("").astype(str)
            mask = searchable.apply(
                lambda column: column.str.contains(query, case=False, regex=False), axis=0
            ).any(axis=1)
            self.filtered_report_df = self.current_report_df.loc[mask].copy()
        else:
            self.filtered_report_df = self.current_report_df.copy()

        shown = self.filtered_report_df.fillna("").astype(str)
        self._populate_tree(list(shown.columns), shown.values.tolist())
        self.report_count_var.set(
            f"{len(self.filtered_report_df):,} of {len(self.current_report_df):,} rows"
        )

    def _clear_table_search(self):
        self.report_search_var.set("")
        self._search_table()

    def _view_report_charts(self):
        """Open the dashboard and select a matching chart where available."""
        chart_map = {
            "Peak sales by hour": "Peak sales by hour",
            "Sales by day": "Sales by day",
            "Sales by month": "Monthly sales trend",
            "Sales by category": "Product performance",
            "Product performance": "Product performance",
            "Sales by region": "Regional performance",
            "Regional performance": "Regional performance",
            "Sales by promotion": "Sales by promotion",
            "Sales by age group": "Sales by age group",
            "Sales by gender": "Sales by gender",
            "Sales by customer city": "Sales by customer city",
            "Sales by store city": "Sales by store city",
            "Weekday-hour heatmap": "Weekday-hour heatmap",
        }
        chart_name = chart_map.get(self.table_var.get(), "Peak sales by hour")
        if CHARTS_AVAILABLE and hasattr(self, "chart_var"):
            self.chart_var.set(chart_name)
            self._draw_selected_chart()
        self.notebook.select(self.dashboard_tab)
        self._on_tab_changed(None)

    def _populate_tree(self, columns, rows):
        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col.replace("_", " ").title())
            self.tree.column(col, width=max(120, min(240, len(col) * 12)), anchor="w")
        for row in rows:
            self.tree.insert("", "end", values=row)

    def _append_log(self, message: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", f"[{time.strftime('%H:%M:%S')}] {message}\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _open_output(self):
        path = Path(self.output_var.get().strip() or ".")
        path.mkdir(parents=True, exist_ok=True)
        try:
            if os.name == "nt":
                os.startfile(path)  # Windows
            elif os.name == "posix":
                import subprocess
                subprocess.Popen(["xdg-open", str(path)])
        except Exception as exc:
            messagebox.showerror("Cannot open folder", str(exc))


if __name__ == "__main__":
    app = SmartMartGUI()
    app.mainloop()

import json
from pathlib import Path
from datetime import datetime

import pandas as pd


class OutputManager:
    """Save SmartMart reports, rejected records, summaries and charts."""

    def __init__(self, output_dir: str):
        self.output_dir = Path(output_dir)
        self.charts_dir = self.output_dir / "charts"

        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.charts_dir.mkdir(parents=True, exist_ok=True)

    def save_report(
        self,
        data: pd.DataFrame,
        report_name: str
    ) -> Path:
        """Save an analytical report as an Excel file."""

        safe_name = self._safe_filename(report_name)
        file_path = self.output_dir / f"{safe_name}.xlsx"

        data.to_excel(file_path, index=False, engine="openpyxl")
        return file_path

    def save_csv(
        self,
        data: pd.DataFrame,
        report_name: str
    ) -> Path:
        """Save a report as a CSV file."""

        safe_name = self._safe_filename(report_name)
        file_path = self.output_dir / f"{safe_name}.csv"

        data.to_csv(
            file_path,
            index=False,
            encoding="utf-8-sig"
        )
        return file_path

    def save_rejected_records(
        self,
        data: pd.DataFrame
    ) -> Path:
        """Save rejected records with a timestamp."""

        timestamp = datetime.now().strftime("%Y%m%dT%H%M%S")
        file_path = (
            self.output_dir / f"invalid_records_{timestamp}.xlsx"
        )

        data.to_excel(file_path, index=False, engine="openpyxl")
        return file_path

    def save_dead_letter_queue(
        self,
        data: pd.DataFrame
    ) -> Path:
        """Save rejected records for later investigation."""

        file_path = self.output_dir / "dead_letter_queue.csv"

        data.to_csv(
            file_path,
            index=False,
            encoding="utf-8-sig"
        )
        return file_path

    def save_summary(
        self,
        summary: dict,
        summary_name: str = "processing_summary"
    ) -> Path:
        """Save processing statistics in JSON format."""

        timestamp = datetime.now().strftime("%Y%m%dT%H%M%S")
        safe_name = self._safe_filename(summary_name)

        file_path = (
            self.output_dir / f"{safe_name}_{timestamp}.json"
        )

        with file_path.open("w", encoding="utf-8") as file:
            json.dump(
                summary,
                file,
                indent=4,
                ensure_ascii=False,
                default=str
            )

        return file_path

    def save_chart(self, figure, chart_name: str) -> Path:
        """Save a Matplotlib chart as a PNG image."""

        safe_name = self._safe_filename(chart_name)
        file_path = self.charts_dir / f"{safe_name}.png"

        figure.savefig(
            file_path,
            dpi=300,
            bbox_inches="tight"
        )

        return file_path

    @staticmethod
    def _safe_filename(name: str) -> str:
        """Remove characters unsuitable for filenames."""

        safe_name = "".join(
            char if char.isalnum() or char in "_-" else "_"
            for char in name.strip()
        )

        return safe_name or "report"
