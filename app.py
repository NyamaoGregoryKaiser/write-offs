import re
from pathlib import Path

import pandas as pd
import streamlit as st


APP_DIR = Path(__file__).parent
REPAYMENTS_PATH = APP_DIR / "Repayments.xlsx"
WRITEOFFS_PATH = APP_DIR / "writeoffs.xlsx"


def extract_last_nine_digits(value: object) -> str:
	"""Return the last 9 numeric digits from a value; empty string if none.

	This function converts the input to string, strips non-digits, and returns
	the last 9 characters. If there are fewer than 9 digits, it returns the
	entire string of digits (which may be empty). Using string return type keeps
	leading zeros if present in the last 9 digits.
	"""
	if pd.isna(value):
		return ""
	digits = re.sub(r"\D", "", str(value))
	return digits[-9:]


@st.cache_data(show_spinner=False)
def load_data(repayments_path: Path, writeoffs_path: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
	"""Load Excel files into DataFrames with expected columns."""
	repayments = pd.read_excel(repayments_path, dtype=str)
	writeoffs = pd.read_excel(writeoffs_path, dtype=str)
	return repayments, writeoffs


def compute_phone_aggregates(repayments: pd.DataFrame) -> tuple[pd.Series, pd.Series]:
	"""Compute per-phone last-9-digit aggregates: counts and sum of Total Repaid:."""
	if "Phone Number" not in repayments.columns:
		raise KeyError("Expected 'Phone Number' column (column S) in Repayments.xlsx.")
	if "Total Repaid:" not in repayments.columns:
		raise KeyError("Expected 'Total Repaid:' column in Repayments.xlsx.")

	# Optional: warn if not at expected column S (index 18 in 0-based)
	try:
		col_idx = list(repayments.columns).index("Phone Number")
		if col_idx != 18:
			st.warning("'Phone Number' is not at expected column S; proceeding with the named column.")
	except Exception:
		pass

	repayments_last9 = repayments["Phone Number"].map(extract_last_nine_digits)
	amounts = pd.to_numeric(repayments["Total Repaid:"], errors="coerce").fillna(0.0)

	# Counts per last-9 phone
	counts = repayments_last9.value_counts(dropna=False)
	counts.index = counts.index.astype(str)

	# Sums per last-9 phone
	sums = pd.DataFrame({"k": repayments_last9, "v": amounts}).groupby("k", dropna=False)["v"].sum()
	sums.index = sums.index.astype(str)

	return counts, sums


def main() -> None:
	st.set_page_config(page_title="Write-offs vs Repayments Analysis", layout="wide")
	st.title("Write-offs vs Repayments Analysis")


	missing_files = [p.name for p in [REPAYMENTS_PATH, WRITEOFFS_PATH] if not p.exists()]
	if missing_files:
		st.error(f"Missing required file(s): {', '.join(missing_files)}. Place them next to app.py and reload.")
		return

	with st.spinner("Loading data..."):
		repayments_df, writeoffs_df = load_data(REPAYMENTS_PATH, WRITEOFFS_PATH)

	# Normalize and aggregate phone numbers in repayments
	try:
		phone_last9_counts, phone_last9_amount_sums = compute_phone_aggregates(repayments_df)
	except Exception as exc:
		st.error(str(exc))
		return

	# Normalize writeoffs mobile column using exact name
	if "mobile" not in writeoffs_df.columns:
		st.error("Expected 'mobile' column (column D) in writeoffs.xlsx.")
		return

	# Optional: warn if not at expected column D (index 3 in 0-based)
	try:
		w_col_idx = list(writeoffs_df.columns).index("mobile")
		if w_col_idx != 3:
			st.warning("'mobile' is not at expected column D; proceeding with the named column.")
	except Exception:
		pass

	writeoffs_df = writeoffs_df.copy()
	writeoffs_df["last9_mobile"] = writeoffs_df["mobile"].map(extract_last_nine_digits)
	writeoffs_df["Repayment Phone Matches"] = writeoffs_df["last9_mobile"].map(lambda x: int(phone_last9_counts.get(x, 0)))
	writeoffs_df["amount repayed"] = writeoffs_df["last9_mobile"].map(lambda x: float(phone_last9_amount_sums.get(x, 0.0)))
	writeoffs_df.drop(columns=["last9_mobile"], inplace=True)

	# --- KPI Cards ---
	num_writeoffs = len(writeoffs_df)
	try:
		total_written_off_series = pd.to_numeric(writeoffs_df["Total Writtenoff Derived"], errors="coerce").fillna(0.0)
	except KeyError:
		st.error("Expected 'Total Writtenoff Derived' column in writeoffs.xlsx.")
		return
	total_written_off = float(total_written_off_series.sum())
	total_repaid_for_writeoffs = float(pd.to_numeric(writeoffs_df["amount repayed"], errors="coerce").fillna(0.0).sum())
	percent_repaid = (total_repaid_for_writeoffs / total_written_off * 100.0) if total_written_off > 0 else 0.0

	c1, c2, c3, c4 = st.columns(4)
	with c1:
		st.metric(label="Number of write-offs", value=f"{num_writeoffs:,}")
	with c2:
		st.metric(label="Total amount written off", value=f"{total_written_off:,.2f}")
	with c3:
		st.metric(label="Total repaid write-offs", value=f"{total_repaid_for_writeoffs:,.2f}")
	with c4:
		st.metric(label="% of repaid write-offs", value=f"{percent_repaid:,.2f}%")

	st.subheader("Write-offs with Repayment Phone Match Counts")
	st.dataframe(writeoffs_df, use_container_width=True)

	st.download_button(
		label="Download result as CSV",
		data=writeoffs_df.to_csv(index=False).encode("utf-8"),
		file_name="writeoffs_with_match_counts.csv",
		mime="text/csv",
	)


if __name__ == "__main__":
	main()


