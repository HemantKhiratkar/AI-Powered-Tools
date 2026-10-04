from pathlib import Path

import pandas as pd


PROJECT_DIR = Path(__file__).resolve().parent
SUPPLIERS_FILE = PROJECT_DIR / "suppliers.xlsx"
PURCHASE_ORDERS_FILE = PROJECT_DIR / "Purchase Orders.xlsx"
OUTPUT_FILE = PROJECT_DIR / "baseline_supplier_metrics.csv"

SUPPLIER_COLUMNS = {"supplier_id"}
PURCHASE_ORDER_COLUMNS = {
    "po_id",
    "supplier_id",
    "days_late",
    "qty_ordered",
    "qty_rejected",
    "price_change_vs_last_po",
    "responsiveness_rating_1to5",
}
NUMERIC_PO_COLUMNS = {
    "days_late",
    "qty_ordered",
    "qty_rejected",
    "price_change_vs_last_po",
    "responsiveness_rating_1to5",
}


def validate_schema(
    dataframe: pd.DataFrame, required_columns: set[str], source_name: str
) -> None:
    missing_columns = required_columns.difference(dataframe.columns)
    if missing_columns:
        missing = ", ".join(sorted(missing_columns))
        raise ValueError(f"{source_name} is missing required columns: {missing}")


def calculate_baseline_metrics(
    suppliers: pd.DataFrame, purchase_orders: pd.DataFrame
) -> pd.DataFrame:
    validate_schema(suppliers, SUPPLIER_COLUMNS, "Suppliers data")
    validate_schema(purchase_orders, PURCHASE_ORDER_COLUMNS, "Purchase Orders data")

    if suppliers["supplier_id"].isna().any():
        raise ValueError("Suppliers data contains blank supplier_id values.")
    if purchase_orders["supplier_id"].isna().any():
        raise ValueError("Purchase Orders data contains blank supplier_id values.")
    if purchase_orders["po_id"].isna().any():
        raise ValueError("Purchase Orders data contains blank po_id values.")
    if suppliers["supplier_id"].duplicated().any():
        raise ValueError("Suppliers data must contain one row per supplier_id.")

    purchase_orders = purchase_orders.copy()
    for column in NUMERIC_PO_COLUMNS:
        purchase_orders[column] = pd.to_numeric(
            purchase_orders[column], errors="raise"
        )
        if purchase_orders[column].isna().any():
            raise ValueError(
                f"Purchase Orders data contains blank values in {column}."
            )

    metrics = purchase_orders.groupby("supplier_id").agg(
        order_count=("po_id", "count"),
        on_time_delivery_percentage=(
            "days_late",
            lambda days_late: days_late.eq(0).mean() * 100,
        ),
        total_quantity_ordered=("qty_ordered", "sum"),
        total_quantity_rejected=("qty_rejected", "sum"),
        average_price_change=("price_change_vs_last_po", "mean"),
        average_responsiveness=("responsiveness_rating_1to5", "mean"),
    )
    metrics["overall_rejection_rate"] = metrics[
        "total_quantity_rejected"
    ].div(metrics["total_quantity_ordered"].where(
        metrics["total_quantity_ordered"] != 0
    ))

    return suppliers.merge(metrics.reset_index(), on="supplier_id", how="left")


def main() -> None:
    for input_file in (SUPPLIERS_FILE, PURCHASE_ORDERS_FILE):
        if not input_file.is_file():
            raise FileNotFoundError(
                f"Required input file not found: {input_file}. "
                "Place both Excel files next to this script."
            )

    suppliers = pd.read_excel(SUPPLIERS_FILE)
    purchase_orders = pd.read_excel(PURCHASE_ORDERS_FILE)
    baseline_metrics = calculate_baseline_metrics(suppliers, purchase_orders)
    baseline_metrics.to_csv(OUTPUT_FILE, index=False)
    print(f"Baseline metrics saved to {OUTPUT_FILE}")
    print(baseline_metrics.head().to_string(index=False))


if __name__ == "__main__":
    main()
