def binary_search(items, target):
    low = 0
    high = len(items) - 1
    while low <= high:
        middle = (low + high) // 2
        if items[middle] == target:
            return middle
        if items[middle] < target:
            low = middle + 1
        else:
            high = middle - 1
    return -1
def test_target_in_middle():
    assert binary_search([10, 20, 30, 40, 50], 30) == 2
def test_target_at_start():
    assert binary_search([10, 20, 30, 40, 50], 10) == 0
def test_target_at_end():
    assert binary_search([10, 20, 30, 40, 50], 50) == 4
def test_target_not_found():
    assert binary_search([10, 20, 30, 40, 50], 35) == -1
def test_empty_list():
    assert binary_search([], 10) == -1
