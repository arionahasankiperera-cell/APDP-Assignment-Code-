"""
SmartRetail Analytics - Task 4 Automated Testing (pytest)
Test cases: TC01-TC12

IMPORTANT:
This standalone suite tests representative behaviour using sample data.
To test the real application, replace the demonstration logic with calls
to the actual SmartRetail Analytics modules and GUI.

Run:
    python -m pip install pytest pandas
    python -m pytest test_smartretail.py -v
"""

from pathlib import Path
import pandas as pd
import pytest


# ------------------------- Test fixtures -------------------------

@pytest.fixture
def valid_retail_df():
    """A small valid retail dataset."""
    return pd.DataFrame({
        "TransactionID": ["T001", "T002", "T003", "T004"],
        "Product": ["Laptop", "Mouse", "Keyboard", "Laptop"],
        "Quantity": [2, 5, 3, 1],
        "Price": [1000.0, 20.0, 35.0, 1000.0],
        "Region": ["North", "South", "East", "North"],
        "CustomerID": ["C01", "C02", "C03", "C01"],
    })


@pytest.fixture
def required_columns():
    return {"TransactionID", "Product", "Quantity", "Price", "Region"}


# ------------------------- Helpers used in tests -------------------------

def validate_dataframe(df, required):
    """Demonstration validator; replace with the real validator."""
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Missing required columns: {sorted(missing)}")
    valid_mask = (df["Quantity"] >= 0) & (df["Price"] >= 0)
    return df.loc[valid_mask].copy(), df.loc[~valid_mask].copy()


def detect_duplicates(df):
    """Demonstration duplicate detector."""
    return df[df.duplicated(keep="first")].copy()


def calculate_sales(df):
    """Demonstration transformation / sales calculation."""
    result = df.copy()
    result["TotalSales"] = result["Quantity"] * result["Price"]
    return result


def run_demo_strategy(strategy_name, df):
    """
    Demonstration stand-in for an analytics strategy.
    Replace with the actual AnalyticsEngine strategy invocation.
    """
    if df.empty:
        return pd.DataFrame()

    if strategy_name == "peak_sales":
        frame = calculate_sales(df)
        return frame.nlargest(1, "TotalSales")

    if strategy_name == "customer_segmentation":
        return df.groupby("CustomerID", as_index=False)["Quantity"].sum()

    if strategy_name == "seasonal_promotion":
        return pd.DataFrame({"Recommendation": ["Review seasonal sales"]})

    if strategy_name == "product_performance":
        frame = calculate_sales(df)
        return frame.groupby("Product", as_index=False)["TotalSales"].sum()

    if strategy_name == "anomaly_detection":
        frame = calculate_sales(df)
        return frame[frame["TotalSales"] > frame["TotalSales"].mean()]

    if strategy_name == "regional_performance":
        frame = calculate_sales(df)
        return frame.groupby("Region", as_index=False)["TotalSales"].sum()

    raise ValueError(f"Unknown strategy: {strategy_name}")


def write_report(df, path):
    df.to_csv(path, index=False)
    return path


def write_demo_chart(path):
    """
    Creates a non-empty placeholder file to test output-path handling only.
    This is NOT a real chart. Replace with the app's chart generator.
    """
    Path(path).write_bytes(b"placeholder; replace with actual chart generation")
    return path


def process_chunks(chunks, fail_on_chunk=None):
    """Demonstration chunk processor used to test error-handling expectations."""
    results = []
    for index, chunk in enumerate(chunks, start=1):
        if index == fail_on_chunk:
            raise RuntimeError(f"Processing failed on chunk {index}")
        results.append(chunk.copy())
    return pd.concat(results, ignore_index=True) if results else pd.DataFrame()


# ------------------------- TC01: GUI dataset processing -------------------------

def test_tc01_gui_dataset_processing(valid_retail_df, required_columns):
    """Checks that the sample dataset is suitable for processing; not a GUI test."""
    assert not valid_retail_df.empty
    assert required_columns.issubset(valid_retail_df.columns)
    processed = calculate_sales(valid_retail_df)
    assert len(processed) == len(valid_retail_df)
    assert "TotalSales" in processed.columns


# ------------------------- TC02: Empty CSV file -------------------------

def test_tc02_empty_csv_file(tmp_path):
    """Checks that an empty CSV can be identified safely."""
    csv_path = tmp_path / "empty.csv"
    csv_path.write_text("", encoding="utf-8")

    assert csv_path.exists()
    assert csv_path.stat().st_size == 0
    with pytest.raises(pd.errors.EmptyDataError):
        pd.read_csv(csv_path)


# ------------------------- TC03: Automated data validation -------------------------

def test_tc03_valid_records(valid_retail_df, required_columns):
    valid, invalid = validate_dataframe(valid_retail_df, required_columns)
    assert len(valid) == len(valid_retail_df)
    assert invalid.empty


def test_tc03_missing_required_column(valid_retail_df, required_columns):
    incomplete = valid_retail_df.drop(columns=["Price"])
    with pytest.raises(ValueError, match="Missing required columns"):
        validate_dataframe(incomplete, required_columns)


def test_tc03_invalid_negative_quantity(valid_retail_df, required_columns):
    invalid_input = valid_retail_df.copy()
    invalid_input.loc[0, "Quantity"] = -3
    valid, invalid = validate_dataframe(invalid_input, required_columns)
    assert len(valid) == len(invalid_input) - 1
    assert len(invalid) == 1
    assert invalid.iloc[0]["Quantity"] == -3


# ------------------------- TC04: Duplicate record detection -------------------------

def test_tc04_duplicate_record_detection(valid_retail_df):
    data = pd.concat([valid_retail_df, valid_retail_df.iloc[[0]]],
                     ignore_index=True)
    duplicates = detect_duplicates(data)
    assert len(duplicates) == 1
    assert duplicates.iloc[0]["TransactionID"] == "T001"


# ------------------------- TC05: Six analytical strategies -------------------------

@pytest.mark.parametrize("strategy", [
    "peak_sales",
    "customer_segmentation",
    "seasonal_promotion",
    "product_performance",
    "anomaly_detection",
    "regional_performance",
])
def test_tc05_each_analytics_strategy_returns_dataframe(strategy, valid_retail_df):
    """Tests the demonstration strategies, not the app's real AnalyticsEngine."""
    result = run_demo_strategy(strategy, valid_retail_df)
    assert isinstance(result, pd.DataFrame)
    assert not result.empty


def test_tc05_analytics_strategy_handles_empty_dataframe():
    result = run_demo_strategy("peak_sales", pd.DataFrame())
    assert isinstance(result, pd.DataFrame)
    assert result.empty


# ------------------------- TC06: Report generation -------------------------

def test_tc06_report_generation(tmp_path, valid_retail_df):
    report = calculate_sales(valid_retail_df)
    report_path = tmp_path / "sales_report.csv"
    write_report(report, report_path)

    assert report_path.exists()
    loaded = pd.read_csv(report_path)
    assert len(loaded) == len(valid_retail_df)
    assert "TotalSales" in loaded.columns
    assert loaded["TotalSales"].sum() == pytest.approx(
        (valid_retail_df["Quantity"] * valid_retail_df["Price"]).sum()
    )


# ------------------------- TC07: Chart generation -------------------------

def test_tc07_chart_output_file_created(tmp_path):
    """Tests file output only; replace placeholder with actual chart generator."""
    chart_path = tmp_path / "sales_chart.png"
    write_demo_chart(chart_path)

    assert chart_path.exists()
    assert chart_path.stat().st_size > 0


# ------------------------- TC08: Invalid record output -------------------------

def test_tc08_invalid_record_output(tmp_path, valid_retail_df, required_columns):
    data = valid_retail_df.copy()
    data.loc[0, "Quantity"] = -5
    valid, invalid = validate_dataframe(data, required_columns)

    output_path = tmp_path / "invalid_records.csv"
    invalid.to_csv(output_path, index=False)

    saved = pd.read_csv(output_path)
    assert len(valid) == len(data) - 1
    assert len(saved) == 1
    assert saved.iloc[0]["Quantity"] == -5


# ------------------------- TC09: Empty input DataFrame -------------------------

def test_tc09_empty_input_dataframe():
    empty = pd.DataFrame()
    assert empty.empty
    assert len(empty) == 0


# ------------------------- TC10: Partial processing failure -------------------------

def test_tc10_partial_processing_failure(valid_retail_df):
    chunks = [
        valid_retail_df.iloc[:2],
        valid_retail_df.iloc[2:],
    ]
    with pytest.raises(RuntimeError, match="Processing failed on chunk 2"):
        process_chunks(chunks, fail_on_chunk=2)


# ------------------------- TC11: Complete processing failure -------------------------

def test_tc11_complete_processing_failure():
    def failing_processing_service():
        raise RuntimeError("Simulated complete processing failure")

    with pytest.raises(RuntimeError, match="Simulated complete processing failure"):
        failing_processing_service()


# ------------------------- TC12: GUI results display logic -------------------------

def test_tc12_results_summary(valid_retail_df):
    """Checks summary data used by a GUI; does not launch the actual GUI."""
    processed = calculate_sales(valid_retail_df)
    summary = {
        "records_processed": len(processed),
        "total_sales": float(processed["TotalSales"].sum()),
        "status": "Completed",
    }

    assert summary["records_processed"] == 4
    assert summary["total_sales"] == pytest.approx(3205.0)
    assert summary["status"] == "Completed"
