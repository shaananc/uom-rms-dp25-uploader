import pandas as pd
import panflute as pf

from create_docs import create_doc_latex, create_doc_markdown


def extract_and_convert_to_markdown_with_hdrs(data, positions, ci_rows, hdr_rows, total_row, output_path):
    # Extract costs for each year for each CI, HDR, and directly use the total row values
    costs = {}
    totals = {}
    for year_idx, (row, start_col, end_col) in enumerate(positions):
        year_key = f"Year {year_idx + 1}"
        costs[year_key] = []
        # Fetch costs for CIs, using specific description
        for ci_row in ci_rows:
            name = data.iloc[ci_row, 0]
            name = str(name) if not pd.isna(name) else "Unknown CI Name"
            cost = data.iloc[ci_row, end_col]
            #print(name,cost, ci_row, start_col)
            cost = 0 if pd.isna(cost) else cost
            if cost:
                costs[year_key].append((name, "Level at 0.2 FTE, including on-costs.", cost))
        # Fetch costs for HDRs, using specific description
        for hdr_row in hdr_rows:
            name = data.iloc[hdr_row, 0]
            name = str(name) if not pd.isna(name) else "Unknown HDR Name"
            cost = data.iloc[hdr_row, start_col]
            cost = 0 if pd.isna(cost) else cost
            if cost:
                costs[year_key].append((name, "Level at 0.2 FTE, including on-costs.", cost))
        # Directly fetch the total from the specified row
        total_cost = data.iloc[total_row, start_col] + data.iloc[total_row, end_col]
        total_cost = 0 if pd.isna(total_cost) else total_cost
        totals[year_key] = total_cost
    
    print(costs)
    create_doc_latex(costs, totals, "budget-uom.tex")
    create_doc_markdown(costs, totals, output_path)

    return output_path


# Load your data
data = pd.read_excel('DP25-Budget.xlsx', sheet_name='2. DP25 Budget Tool')


# Assuming the Excel data is already loaded
output_md_path = '/mnt/data/Final_Budget_Justification_with_HDRs.md'
year_positions = [(13, 3, 4), (13, 10, 11), (13, 17, 18)]  # Row 14 in Excel, Columns D-E, K-L, R-S
ci_rows = [13, 14, 15, 16]  # Rows for individual CIs are 14, 15, 16, 17 in Excel
hdr_rows = [17,18]
total_row = 12  # Row for total personnel is 13 in Excel
output_md_path = './Budget_Justification_UOM.md'
final_md_path_with_hdrs = extract_and_convert_to_markdown_with_hdrs(data, year_positions, ci_rows, hdr_rows, total_row, output_md_path)
print(final_md_path_with_hdrs)
