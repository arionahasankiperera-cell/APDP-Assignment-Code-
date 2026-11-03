
import pytest
import pandas as pd


# 1. Unit testing: sales calculation
def calculate_sales(quantity, price):
    return quantity * price


def test_normal_sales_calculation():
    assert calculate_sales(2, 100) == 200


def test_zero_quantity():
    assert calculate_sales(0, 100) == 0


def test_large_quantity():
    assert calculate_sales(100, 25) == 2500


# 2. Parameterized boundary-value testing
@pytest.mark.parametrize(
    "quantity, price, expected",
    [
        (0, 100, 0),
        (1, 100, 100),
        (100, 25, 2500),
    ]
)
def test_sales_boundaries(quantity, price, expected):
    assert calculate_sales(quantity, price) == pytest.approx(expected)


# 3. Negative testing and exception handling
def calculate_valid_sales(quantity, price):
    if quantity < 0:
        raise ValueError("Quantity cannot be negative")
    if price < 0:
        raise ValueError("Price cannot be negative")
    return quantity * price


def test_negative_quantity():
    with pytest.raises(ValueError):
        calculate_valid_sales(-1, 100)


def test_negative_price():
    with pytest.raises(ValueError):
        calculate_valid_sales(2, -10)


# 4. Binary search algorithm testing
def binary_search(items, target):
    low = 0
    high = len(items) - 1

    while low <= high:
        middle = (low + high) // 2

        if items[middle] == target:
            return middle
        elif items[middle] < target:
            low = middle + 1
        else:
            high = middle - 1

    return -1


def test_binary_search_middle():
    assert binary_search([10, 20, 30, 40, 50], 30) == 2


def test_binary_search_start():
    assert binary_search([10, 20, 30, 40, 50], 10) == 0


def test_binary_search_end():
    assert binary_search([10, 20, 30, 40, 50], 50) == 4


def test_binary_search_not_found():
    assert binary_search([10, 20, 30, 40, 50], 35) == -1


def test_binary_search_empty_list():
    assert binary_search([], 10) == -1


# 5. CSV file handling and data checks
def test_csv_file_reading(tmp_path):
    csv_file = tmp_path / "retail_sales.csv"

    expected = pd.DataFrame({
        "Product": ["Laptop", "Mouse"],
        "Quantity": [2, 5],
        "Price": [1000, 20]
    })

    expected.to_csv(csv_file, index=False)
    actual = pd.read_csv(csv_file)

    assert len(actual) == 2
    assert list(actual.columns) == [
        "Product", "Quantity", "Price"
    ]
    assert actual["Quantity"].sum() == 7
