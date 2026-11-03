{
 "cells": [
  {
   "cell_type": "code",
   "execution_count": 3,
   "id": "44a2186e-be45-428b-8ffb-0b571b2f9438",
   "metadata": {},
   "outputs": [],
   "source": [
    "def calculate_sales(quantity, price):\n",
    "    return quantity * price\n",
    "def test_sales_calculation():\n",
    "    assert calculate_sales(2, 100) == 200\n",
    "def test_zero_quantity():\n",
    "    assert calculate_sales(0, 100) == 0\n",
    "def test_large_quantity():\n",
    "    assert calculate_sales(100, 25) == 2500\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 4,
   "id": "b6c4a609-3b0e-4d5a-b97b-d37fede5f919",
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "Overwriting test_retail.py\n"
     ]
    }
   ],
   "source": [
    "%%writefile test_retail.py\n",
    "\n",
    "def calculate_sales(quantity, price):\n",
    "    return quantity * price\n",
    "\n",
    "\n",
    "def test_sales_calculation():\n",
    "    assert calculate_sales(2, 100) == 200\n",
    "\n",
    "\n",
    "def test_zero_quantity():\n",
    "    assert calculate_sales(0, 100) == 0\n",
    "\n",
    "\n",
    "def test_large_quantity():\n",
    "    assert calculate_sales(100, 25) == 2500"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 5,
   "id": "db4aac84-436d-4c14-a974-e62fc9dd7303",
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "C:\\Users\\USER'\\OneDrive\\Documents\\APDP\n"
     ]
    }
   ],
   "source": [
    "import os\n",
    "print(os.getcwd())"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 1,
   "id": "50113ed5-6d15-4a99-802d-44570fe96e7d",
   "metadata": {},
   "outputs": [],
   "source": [
    "import pytest\n",
    "@pytest.mark.parametrize(\n",
    "    \"quantity, price, expected\",\n",
    "    [\n",
    "        (0, 100, 0),\n",
    "        (1, 100, 100),\n",
    "        (100, 25, 2500),\n",
    "    ]\n",
    ")\n",
    "def test_sales_boundaries(quantity, price, expected):\n",
    "    assert quantity * price == pytest.approx(expected)\n"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 2,
   "id": "887f585e-82b9-495b-96bc-3a636e1d24a6",
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "Writing test_edge_cases.py\n"
     ]
    }
   ],
   "source": [
    "%%writefile test_edge_cases.py\n",
    "import pytest\n",
    "\n",
    "@pytest.mark.parametrize(\n",
    "    \"quantity, price, expected\",\n",
    "    [\n",
    "        (0, 100, 0),\n",
    "        (1, 100, 100),\n",
    "        (100, 25, 2500),\n",
    "    ]\n",
    ")\n",
    "def test_sales_boundaries(quantity, price, expected):\n",
    "    assert quantity * price == pytest.approx(expected)"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 3,
   "id": "2d30edc7-bdc6-43fb-8b7d-54792ff8d62f",
   "metadata": {},
   "outputs": [
    {
     "name": "stdout",
     "output_type": "stream",
     "text": [
      "Notebook folder: C:\\Users\\USER'\\OneDrive\\Documents\\APDP\n",
      "Test file exists: True\n"
     ]
    }
   ],
   "source": [
    "import os\n",
    "\n",
    "print(\"Notebook folder:\", os.getcwd())\n",
    "print(\"Test file exists:\", os.path.exists(\"test_edge_cases.py\"))"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "id": "8d3285de-569b-4c86-8ef5-5406b130f35d",
   "metadata": {},
   "outputs": [],
   "source": []
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3 (ipykernel)",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.14.1"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 5
}
