# Data Quality Report

## Summary
- Input: C:\XL\Project\07_ExcelSkills\data\store_ops_testdata_50rows.xlsx
- Sheet: (first)
- Rows: 50
- Columns: 18
- Duplicate rows: 0 (0.00%)
- Cleaned rows: 45

## Missing Rate by Column
| Column | Missing | Missing % |
| --- | --- | --- |
| order_id | 0 | 0.00% |
| order_date | 0 | 0.00% |
| payment_date | 1 | 2.00% |
| store_id | 0 | 0.00% |
| store_name | 0 | 0.00% |
| product_id | 0 | 0.00% |
| product_name | 0 | 0.00% |
| category | 0 | 0.00% |
| unit_price | 1 | 2.00% |
| quantity | 0 | 0.00% |
| order_amount | 1 | 2.00% |
| customer_id | 0 | 0.00% |
| customer_type | 0 | 0.00% |
| city | 1 | 2.00% |
| country | 0 | 0.00% |
| order_status | 0 | 0.00% |
| refund_amount | 0 | 0.00% |
| shipping_fee | 0 | 0.00% |

## Duplicate Values by Column
| Column | Duplicate Values | Duplicate % |
| --- | --- | --- |
| order_id | 1 | 2.00% |
| order_date | 17 | 34.00% |
| payment_date | 14 | 28.57% |
| store_id | 47 | 94.00% |
| store_name | 47 | 94.00% |
| product_id | 40 | 80.00% |
| product_name | 40 | 80.00% |
| category | 44 | 88.00% |
| unit_price | 37 | 75.51% |
| quantity | 45 | 90.00% |
| order_amount | 19 | 38.78% |
| customer_id | 0 | 0.00% |
| customer_type | 47 | 94.00% |
| city | 38 | 77.55% |
| country | 48 | 96.00% |
| order_status | 45 | 90.00% |
| refund_amount | 42 | 84.00% |
| shipping_fee | 44 | 88.00% |

## Type Corrections
| Column | From | To | Parse Rate |
| --- | --- | --- | --- |
| order_date | object | datetime64 | 98.00% |
| payment_date | object | datetime64 | 100.00% |
| unit_price | object | numeric | 97.96% |
| quantity | int64 | numeric | 100.00% |
| order_amount | object | numeric | 100.00% |
| refund_amount | float64 | numeric | 100.00% |
| shipping_fee | float64 | numeric | 100.00% |

## Outliers (IQR)
| Column | Count | % of Non-missing | Lower | Upper |
| --- | --- | --- | --- | --- |
| unit_price | 7 | 15.56% | -13.45 | 38.15 |
| quantity | 3 | 6.67% | -0.5 | 3.5 |
| order_amount | 2 | 4.44% | -38.9 | 83.5 |
| refund_amount | 7 | 15.56% | 0 | 0 |
| shipping_fee | 0 | 0.00% | -4.5 | 7.5 |