from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

# Resolve files from this script so the app works regardless of the launch directory.
PROJECT_DIR = Path(__file__).resolve().parent
SUBMISSION_WORKBOOK = PROJECT_DIR / "T18_Supplier_scorecard_app_completed.xlsx"
SOURCE_WORKBOOK = PROJECT_DIR / "T18_Supplier_scorecard_app_AI.xlsx"
TIER_ORDER = ["Strategic", "Maintain", "Develop", "Exit"]

st.set_page_config(page_title="Supplier Scorecard", layout="wide")
st.markdown(
    """
    <style>
    :root { --page: #0e1117; --surface: #1f2430; --text: #f4f6fb; --muted: #c4ccd8; --accent: #ff4b4b; }
    .stApp { background: var(--page); color: var(--text); color-scheme: dark; }
    [data-testid="stAppViewContainer"], [data-testid="stMain"],
    [data-testid="stSidebar"], [data-testid="stHeader"] { background: var(--page) !important; color: var(--text); }
    .stApp h1, .stApp h2, .stApp h3, .stApp p, .stApp label,
    .stApp [data-testid="stWidgetLabel"] { color: var(--text); }
    [data-testid="stMetric"] { background: var(--surface); color: var(--text); border-left: 3px solid var(--accent); padding: 12px 16px; }
    [data-testid="stMetric"] * { color: var(--text); }
    .stApp input, .stApp textarea,
    [data-testid="stTextAreaRootElement"] { background: var(--surface) !important; color: var(--text) !important; }
    .stApp input::placeholder, .stApp textarea::placeholder { color: var(--muted); }
    [data-testid="stExpander"] summary,
    [data-testid="stExpander"] [data-testid="stTextAreaRootElement"] { background: var(--surface) !important; color: var(--text) !important; }
    [data-testid="stExpander"] summary * { color: var(--text) !important; }
    [data-testid="stExpander"] textarea { background: var(--surface) !important; color: var(--text) !important; }
    [data-testid="stDownloadButton"] button { background: #ff4b4b !important; border-color: #ff4b4b !important; color: #fff !important; }
    [data-testid="stDownloadButton"] button * { color: #fff !important; }
    [data-testid="stDownloadButton"] button:hover { background: #e23d3d !important; border-color: #e23d3d !important; }
    div[data-testid="stDataFrame"] { border-top: 2px solid var(--accent); }
    </style>
    """,
    unsafe_allow_html=True,
)


@st.cache_data
def load_scorecard() -> pd.DataFrame:
    suppliers = pd.read_excel(SOURCE_WORKBOOK, sheet_name="Suppliers")
    purchase_orders = pd.read_excel(SOURCE_WORKBOOK, sheet_name="Purchase Orders")

    po_summary = purchase_orders.groupby("supplier_id").agg(
        po_count=("po_id", "count"),
        on_time_delivery_rate=("days_late", lambda values: values.eq(0).mean() * 100),
        total_qty_rejected=("qty_rejected", "sum"),
        total_qty_ordered=("qty_ordered", "sum"),
        average_price_change=("price_change_vs_last_po", "mean"),
        mean_responsiveness_rating=("responsiveness_rating_1to5", "mean"),
    )
    po_summary["rejection_rate"] = (
        po_summary["total_qty_rejected"] / po_summary["total_qty_ordered"]
    )
    po_summary = po_summary.drop(
        columns=["total_qty_rejected", "total_qty_ordered"]
    ).reset_index()

    scorecard = suppliers.merge(po_summary, on="supplier_id", how="left")

    for metric in ("on_time_delivery_rate", "mean_responsiveness_rating"):
        values = scorecard[metric]
        minimum = values.min()
        maximum = values.max()
        score_column = f"{metric}_score"
        if pd.isna(minimum) or pd.isna(maximum):
            scorecard[score_column] = values
        elif minimum == maximum:
            scorecard[score_column] = values.where(values.isna(), 50.0)
        else:
            scorecard[score_column] = (values - minimum) / (maximum - minimum) * 100

    scorecard["rejection_rate_score"] = (
        100 - scorecard["rejection_rate"] * 100
    ).clip(lower=0)

    price_change = scorecard["average_price_change"]
    min_price_change = price_change.min()
    max_price_change = price_change.max()
    if pd.isna(min_price_change) or pd.isna(max_price_change):
        scorecard["average_price_change_score"] = price_change
    elif min_price_change == max_price_change:
        scorecard["average_price_change_score"] = price_change.where(
            price_change.isna(), 50.0
        )
    else:
        scorecard["average_price_change_score"] = (
            (max_price_change - price_change)
            / (max_price_change - min_price_change)
            * 100
        )

    return scorecard


def write_baseline_sheet(workbook_path: Path, scorecard_df: pd.DataFrame) -> None:
    """Replace or add the auditable baseline sheet while preserving other workbook sheets."""
    if not workbook_path.exists():
        # If completed workbook doesn't exist yet, copy template or create from SOURCE_WORKBOOK
        import shutil
        shutil.copy2(SOURCE_WORKBOOK, workbook_path)

    suppliers = pd.read_excel(SOURCE_WORKBOOK, sheet_name="Suppliers")
    purchase_orders = pd.read_excel(SOURCE_WORKBOOK, sheet_name="Purchase Orders")
    
    grouped = purchase_orders.groupby("supplier_id", as_index=False).agg(
        po_count=("po_id", "count"),
        on_time_delivery_rate_pct=("days_late", lambda values: values.eq(0).mean() * 100),
        qty_ordered_total=("qty_ordered", "sum"),
        qty_rejected_total=("qty_rejected", "sum"),
        average_price_change=("price_change_vs_last_po", "mean"),
        average_responsiveness=("responsiveness_rating_1to5", "mean"),
    )
    grouped["rejection_rate_pct"] = (
        grouped["qty_rejected_total"] / grouped["qty_ordered_total"] * 100
    )
    baseline = suppliers.merge(grouped, on="supplier_id", how="left")

    export_columns = [
        "supplier_id",
        "supplier_name",
        "category",
        "po_count",
        "on_time_delivery_rate_pct",
        "qty_ordered_total",
        "qty_rejected_total",
        "rejection_rate_pct",
        "average_price_change",
        "average_responsiveness",
    ]
    with pd.ExcelWriter(
        workbook_path,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace",
    ) as writer:
        baseline[export_columns].to_excel(writer, sheet_name="Baseline Metrics", index=False)
        worksheet = writer.book["Baseline Metrics"]
        worksheet.freeze_panes = "A2"
        worksheet.auto_filter.ref = worksheet.dimensions
        worksheet.row_dimensions[1].height = 30
        for cell in worksheet[1]:
            cell.font = Font(color="FFFFFF", bold=True)
            cell.fill = PatternFill("solid", fgColor="173C36")
        widths = {
            "A": 14, "B": 26, "C": 18, "D": 12, "E": 25,
            "F": 22, "G": 22, "H": 18, "I": 22, "J": 25,
        }
        for column, width in widths.items():
            worksheet.column_dimensions[column].width = width
        for row_number in range(2, worksheet.max_row + 1):
            worksheet.cell(row_number, 5).number_format = "0.00"
            worksheet.cell(row_number, 8).number_format = "0.00"
            worksheet.cell(row_number, 9).number_format = "0.00%"
            worksheet.cell(row_number, 10).number_format = "0.00"
        table = Table(displayName="SupplierBaseline", ref=worksheet.dimensions)
        table.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium4",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        worksheet.add_table(table)


st.title("Supplier Scorecard")
st.caption("Procurement performance across delivery, quality, price movement, and responsiveness")

st.sidebar.header("Scoring model")
st.sidebar.caption("Each KPI is normalized to 0-100. Weights are automatically normalized to total 100%.")

weight_labels = {
    "On-Time Delivery": "on_time_delivery_weight",
    "Rejection Rate": "rejection_rate_weight",
    "Price Change": "price_change_weight",
    "Responsiveness": "responsiveness_weight",
}

def normalize_slider_weights(changed_key: str) -> None:
    other_keys = [key for key in weight_labels.values() if key != changed_key]
    remaining = 100 - st.session_state[changed_key]
    other_values = [st.session_state[key] for key in other_keys]
    other_total = sum(other_values)

    if other_total == 0:
        exact_shares = [remaining / len(other_keys)] * len(other_keys)
    else:
        exact_shares = [
            remaining * value / other_total for value in other_values
        ]

    rounded_shares = [int(share) for share in exact_shares]
    remainder = remaining - sum(rounded_shares)
    order = sorted(
        range(len(other_keys)),
        key=lambda index: exact_shares[index] - rounded_shares[index],
        reverse=True,
    )
    for index in order[:remainder]:
        rounded_shares[index] += 1

    for key, value in zip(other_keys, rounded_shares):
        st.session_state[key] = value

defaults = {"on_time_delivery_weight": 25, "rejection_rate_weight": 25, "price_change_weight": 10, "responsiveness_weight": 40}
for label, key in weight_labels.items():
    if key not in st.session_state:
        st.session_state[key] = defaults[key]
    st.sidebar.slider(
        label,
        min_value=0,
        max_value=100,
        step=5,
        key=key,
        on_change=normalize_slider_weights,
        args=(key,),
    )

weights = {}
for label, key in weight_labels.items():
    weights[label] = st.session_state[key] / 100

last_label = next(reversed(weights))
weights[last_label] = 1.0 - sum(
    weight for label, weight in weights.items() if label != last_label
)

effective_str = " | ".join(f"{label} {weight:.0%}" for label, weight in weights.items())
st.sidebar.caption(f"Effective weights: {effective_str}")

st.sidebar.subheader("Tier cutoffs")
strategic_cutoff = st.sidebar.slider("Strategic at or above", 1, 100, 80)
maintain_cutoff = st.sidebar.slider("Maintain at or above", 0, 99, 65)
develop_cutoff = st.sidebar.slider("Develop at or above", 0, 98, 50)

if not 0 <= develop_cutoff < maintain_cutoff < strategic_cutoff <= 100:
    st.error("Set tier cutoffs in ascending order: Develop < Maintain < Strategic.")
    st.stop()

scorecard = load_scorecard()
score_columns = {
    "On-Time Delivery": "on_time_delivery_rate_score",
    "Rejection Rate": "rejection_rate_score",
    "Price Change": "average_price_change_score",
    "Responsiveness": "mean_responsiveness_rating_score",
}

scorecard["total_score"] = sum(
    scorecard[column] * weights[label]
    for label, column in score_columns.items()
)

scorecard["strategic_tier"] = pd.cut(
    scorecard["total_score"],
    bins=[float("-inf"), develop_cutoff, maintain_cutoff, strategic_cutoff, float("inf")],
    labels=["Exit", "Develop", "Maintain", "Strategic"],
    right=False,
).astype(str)

tier_counts = scorecard["strategic_tier"].value_counts().reindex(TIER_ORDER, fill_value=0)

# Top summary metric cards
metric_cols = st.columns(4)
metric_cols[0].metric("Suppliers", f"{len(scorecard)}")
metric_cols[1].metric("Purchase orders", f"{int(scorecard['po_count'].sum()):,}")
metric_cols[2].metric("Mean supplier on-time", f"{scorecard['on_time_delivery_rate'].mean():.1f}%")
metric_cols[3].metric("Mean supplier rejection", f"{scorecard['rejection_rate'].mean() * 100:.2f}%")

# Tier distribution chart & score interpretation side-by-side
summary_col, interp_col = st.columns([1.25, 1])
with summary_col:
    st.subheader("Tier distribution")
    chart_data = tier_counts.rename_axis("Tier").to_frame("Suppliers")
    st.bar_chart(chart_data, color="#dc684d", height=250)
with interp_col:
    st.subheader("Score interpretation")
    st.markdown(
        "- On-time delivery: observed rate, higher is better.\n"
        "- Rejection: `100 - rejection rate (%)`, bounded to 0-100.\n"
        "- Price: lower average price change scores higher, scaled across suppliers in this dataset.\n"
        "- Responsiveness: rating from 1-5 mapped linearly to 0-100.\n\n"
        "Price normalization is relative to this dataset and should be reviewed before commercial use."
    )

# Supplier register filters
st.subheader("Supplier register")
filter_col, category_col, tier_col = st.columns([1.5, 1, 1])
with filter_col:
    query = st.text_input("Search supplier or ID", placeholder="e.g. Nova or SUP001")
with category_col:
    categories = sorted(scorecard["category"].dropna().unique().tolist())
    selected_categories = st.multiselect("Category", categories, default=categories)
with tier_col:
    selected_tiers = st.multiselect("Tier", TIER_ORDER, default=TIER_ORDER)

filtered = scorecard[
    scorecard["category"].isin(selected_categories) & scorecard["strategic_tier"].isin(selected_tiers)
]
if query.strip():
    needle = query.strip().casefold()
    filtered = filtered[
        filtered["supplier_name"].str.casefold().str.contains(needle)
        | filtered["supplier_id"].str.casefold().str.contains(needle)
    ]

visible_columns = [
    "supplier_id",
    "supplier_name",
    "category",
    "po_count",
    "on_time_delivery_rate",
    "rejection_rate",
    "average_price_change",
    "mean_responsiveness_rating",
    "total_score",
    "strategic_tier",
]
display_table = filtered[visible_columns].sort_values("total_score", ascending=False).copy()
display_table["rejection_rate"] *= 100
display_table["average_price_change"] *= 100

st.dataframe(
    display_table,
    use_container_width=True,
    hide_index=True,
    column_config={
        "on_time_delivery_rate": st.column_config.NumberColumn("On-time %", format="%.1f%%"),
        "rejection_rate": st.column_config.NumberColumn("Rejection %", format="%.2f%%"),
        "average_price_change": st.column_config.NumberColumn("Avg price change", format="%.2f%%"),
        "mean_responsiveness_rating": st.column_config.NumberColumn("Responsiveness / 5", format="%.2f"),
        "total_score": st.column_config.ProgressColumn("Weighted score", min_value=0, max_value=100, format="%.1f"),
        "strategic_tier": "Tier",
    },
)
st.caption("Average price change is shown in percent; baseline stores it as a decimal ratio.")

# Corrective feedback letters for underperformers
st.subheader("Performance feedback drafts")
action_suppliers = scorecard[scorecard["strategic_tier"].isin(["Develop", "Exit"])].nsmallest(2, "total_score")

if action_suppliers.empty:
    st.info("No supplier is currently in Develop or Exit. Lower the cutoffs to review the lowest scorers.")
else:
    for _, supplier in action_suppliers.iterrows():
        with st.expander(f"{supplier['supplier_name']} · {supplier['strategic_tier']} · {supplier['total_score']:.1f}/100"):
            on_time = supplier["on_time_delivery_rate"]
            rejection = supplier["rejection_rate"] * 100
            price = supplier["average_price_change"] * 100
            resp = supplier["mean_responsiveness_rating"]
            
            letter = (
                f"Subject: Supplier performance improvement request — {supplier['supplier_id']}\n\n"
                f"Dear {supplier['supplier_name']} team,\n\n"
                f"Your current supplier score is {supplier['total_score']:.1f}/100 and your assigned tier is {supplier['strategic_tier']}. "
                f"Please review the following performance items:\n"
                f"- On-time delivery ({on_time:.1f}%): Submit a recovery plan with milestones.\n"
                f"- Rejection rate ({rejection:.2f}%): Provide a root-cause quality review.\n"
                f"- Average price change ({price:+.2f}%): Provide cost breakdown for increases.\n"
                f"- Responsiveness ({resp:.2f}/5): Commit to response-time targets.\n\n"
                f"Regards,\nProcurement Team"
            )
            st.text_area("Draft letter", letter, height=300, key=f"letter_{supplier['supplier_id']}")
            st.download_button(
                "Download feedback letter",
                data=letter,
                file_name=f"corrective_action_{supplier['supplier_id']}.txt",
                mime="text/plain",
                key=f"download_{supplier['supplier_id']}",
            )

# Baseline workbook update button
st.subheader("Baseline workbook")
st.write(f"Submission file: `{SUBMISSION_WORKBOOK.name}`")
if st.button("Update Baseline Metrics sheet in submission workbook", type="primary"):
    try:
        write_baseline_sheet(SUBMISSION_WORKBOOK, scorecard)
        st.success("Baseline Metrics was written to the submission workbook.")
    except (FileNotFoundError, PermissionError, OSError, ValueError) as error:
        st.error(f"Could not update the workbook: {error}. Close it in Excel and try again.")