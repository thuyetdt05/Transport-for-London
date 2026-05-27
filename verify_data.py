# -*- coding: utf-8 -*-
import sys
import pandas as pd
import numpy as np
sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_csv('outputs/london_tfl_cleaned.csv')
print('=== SO GA PHAN TICH ===')
print(f'Tong so ga: {len(df)}')

# COVID impact
total_2019 = df['passengers_2019'].sum()
total_2020 = df['passengers_2020'].sum()
total_2021 = df['passengers_2021'].sum()
covid_pct = (total_2020 - total_2019)/total_2019*100
recovery_pct = (total_2021 - total_2020)/total_2020*100

print('\n=== COVID ANALYSIS ===')
print(f'Tong 2019: {total_2019:,.0f}')
print(f'Tong 2020: {total_2020:,.0f}')
print(f'Tong 2021: {total_2021:,.0f}')
print(f'COVID impact 2019->2020: {covid_pct:.4f}%')
print(f'Recovery 2020->2021: {recovery_pct:.4f}%')

print('\n=== CLUSTERING ===')
cluster_col = 'cluster_name'
id_col = 'cluster_id'
print(f'So cum: {df[cluster_col].nunique()}')
grp = df.groupby([id_col, cluster_col])['station'].count().reset_index()
grp = grp.sort_values(id_col)
print(grp.to_string())

print('\n=== TOP 10 GA DONG KHACH NHAT 2021 ===')
top10 = df.nlargest(10, 'passengers_2021')[['station','passengers_2021', cluster_col]].reset_index(drop=True)
for i, row in top10.iterrows():
    print(f"{i+1}. {row['station']} - {row['passengers_2021']:,.0f} ({row[cluster_col]})")

print('\n=== XU HUONG ===')
trend_col = 'trend_category'
print(df[trend_col].value_counts().to_string())

print('\n=== CLUSTER SUMMARY ===')
cs = df.groupby([id_col, cluster_col]).agg(
    so_ga=('station','count'),
    hk_tb=('passengers_2021','mean'),
    tong_hk=('passengers_2021','sum'),
    covid_tb=('covid_impact_pct','mean'),
    recovery_tb=('recovery_rate_pct','mean')
).reset_index().sort_values('hk_tb', ascending=False)
for _, r in cs.iterrows():
    print(f"  Cum {int(r[id_col])} - {r[cluster_col]}: {int(r['so_ga'])} ga, HK_tb={r['hk_tb']:,.0f}, COVID={r['covid_tb']:.2f}%, Recovery={r['recovery_tb']:.2f}%")

print('\n=== COLUMNS AVAILABLE ===')
print(list(df.columns))

print('\n=== TOP 5 GA 2021 CHINH XAC ===')
top5 = df.nlargest(5, 'passengers_2021')[['station','passengers_2021','cluster_name','num_lines','borough']].reset_index(drop=True)
for i, row in top5.iterrows():
    print(f"{i+1}. {row['station']}: {row['passengers_2021']:,.0f} luot | Cum: {row['cluster_name']} | Tuyen: {row['num_lines']} | Borough: {row['borough']}")
