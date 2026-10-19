import pandas as pd

orders, tax = pd.read_csv('orders_raw.csv'), pd.read_csv('tax_ids_reference.csv')
orders['Customer ID'] = orders['Customer ID'].astype(str).str.strip()
merged = orders.merge(tax[['Customer ID', 'Tax ID']], on='Customer ID', how='left').fillna({'Tax ID': 'Pending Verification'})
merged.to_csv('orders_enriched.csv', index=False)
