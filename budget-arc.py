from collections import defaultdict
import pandas as pd
import yaml
from RMSBudgetUploader import BudgetCategory, RMSBudgetBuilder
import pprint

def fetch_data(data, row, col):
    """Fetch and validate cost data from a spreadsheet."""
    cost = data.iloc[row, col] if not pd.isna(data.iloc[row, col]) else 0
    return cost

def sum_costs(data, rows, col):
    """Sum the costs for specified rows."""
    return sum(data.iloc[row, col] for row in rows if not pd.isna(data.iloc[row, col]))

def get_all_costs(data, year_columns, row_configs):
    """Extract and convert budget data for multiple years."""
    for row_config in row_configs:
        for row in row_config['rows']:
            name = fetch_data(data, row, 0)
            row_config['names'].append(name)
            row_config[name] = defaultdict(dict)
            row_config[name]['arc_cash'] = []
            row_config[name]['admin_cash'] = []
            row_config[name]['admin_inkind'] = []
            
            for year_index, _ in enumerate(year_columns):
                start_column = year_columns[year_index]
                row_config[name]['arc_cash'].append(fetch_data(data, row, start_column))
                row_config[name]['admin_cash'].append(fetch_data(data, row, start_column+1))
                row_config[name]['admin_inkind'].append(fetch_data(data, row, start_column+2))
    return row_configs

def upload_to_rms(row_configs, year_columns):
    """Upload the budget data to the RMS."""
    builder = RMSBudgetBuilder()
    builder.setup()
    builder.login()
    builder.goto_budget()
    builder.apply_teaching_relief()
    for year, _ in enumerate(year_columns):
        builder.goto_budget_year(year+1)
        for row_config in row_configs:
            for name in row_config['names']:
                arc_cash = row_config[name]['arc_cash'][year]
                admin_cash = row_config[name]['admin_cash'][year]
                admin_inkind = row_config[name]['admin_inkind'][year]
                if row_config['category'] != BudgetCategory.TOTAL and row_config['category'] != BudgetCategory.TEACHING_RELIEF:
                    builder.input_category(year+1, row_config['category'], name, arc_cash, admin_cash, admin_inkind)
                print(year, row_config['category'], name, arc_cash, admin_cash, admin_inkind)
    builder.save_budget()

# Function to convert Excel column name to zero-based index
def excel_to_index(column_name):
    index = 0
    for c in column_name:
        index = index * 26 + (ord(c) - ord('A') + 1)
    return index - 1


def main():
    # Load the YAML data from a file
    config_data = None
    with open('config.yml', 'r') as file:
        config_data = yaml.safe_load(file)

    # Convert column names to indices
    config_data['year_columns'] = [excel_to_index(col) for col in config_data['year_columns']]

    for config in config_data['row_configs']:
        # Adjust the row start and end to zero-based index and store them as a range list
        start = config['start_row'] - 2
        end = config['end_row'] - 2
        config['rows'] = range(start, end + 1)
        config['names'] = []  # Add an empty names list

    excel_data = pd.read_excel(config_data['xlsx_path'], config_data['sheet_name'])
    config_data['row_configs'] = get_all_costs(excel_data, config_data['year_columns'], config_data['row_configs'])
    
    pprint.pprint(config_data)
    upload_to_rms(config_data['row_configs'], config_data['year_columns'])


if __name__ == '__main__':
    main()