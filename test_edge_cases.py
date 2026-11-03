import pytest

@pytest.mark.parametrize(
    "quantity, price, expected",
    [
        (0, 100, 0),
        (1, 100, 100),
        (100, 25, 2500),
    ]
)
def test_sales_boundaries(quantity, price, expected):
    assert quantity * price == pytest.approx(expected)
