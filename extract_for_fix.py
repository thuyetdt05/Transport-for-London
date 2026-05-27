# -*- coding: utf-8 -*-
"""Thu thap so lieu chuyen sau de viet fix.docx"""
import sys, pandas as pd, numpy as np
sys.stdout.reconfigure(encoding='utf-8')
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import silhouette_score

df = pd.read_csv('outputs/london_tfl_cleaned.csv')

print("=== 1. THONG KE CO BAN ===")
print(f"Tong ga: {len(df)}")
print(f"So cot: {len(df.columns)}")
print(f"Cot co NaN: {df.isnull().sum()[df.isnull().sum()>0].to_dict()}")

print("\n=== 2. COVID ANALYSIS CHINH XAC ===")
t19 = df['passengers_2019'].sum()
t20 = df['passengers_2020'].sum()
t21 = df['passengers_2021'].sum()
covid = (t20-t19)/t19*100
rec   = (t21-t20)/t20*100
ratio_vs_2019 = t21/t19*100
print(f"2019: {t19:,}")
print(f"2020: {t20:,}")
print(f"2021: {t21:,}")
print(f"COVID 2019->2020: {covid:.4f}%")
print(f"Recovery 2020->2021: {rec:.4f}%")
print(f"2021 vs 2019: {ratio_vs_2019:.2f}%")

print("\n=== 3. ELBOW METHOD - INERTIA ===")
features = ['passengers_2021','num_lines','lat','lon']
X = df[features].dropna().values
Xs = StandardScaler().fit_transform(X)
inertias, silhouettes = [], []
for k in range(2, 11):
    km = KMeans(n_clusters=k, random_state=42, n_init=10)
    labels = km.fit_predict(Xs)
    inertias.append((k, km.inertia_))
    sil = silhouette_score(Xs, labels)
    silhouettes.append((k, sil))
    print(f"  K={k}: Inertia={km.inertia_:.1f}, Silhouette={sil:.4f}")

print("\n=== 4. CLUSTER SUMMARY CHINH XAC ===")
cs = df.groupby(['cluster_id','cluster_name']).agg(
    so_ga=('station','count'),
    hk_tb_2021=('passengers_2021','mean'),
    tong_hk_2021=('passengers_2021','sum'),
    hk_tb_2019=('passengers_2019','mean'),
    covid_tb=('covid_impact_pct','mean'),
    recovery_tb=('recovery_rate_pct','mean'),
    num_lines_tb=('num_lines','mean'),
    trend_giam_manh=('trend_category', lambda x: (x=='Giảm mạnh').sum()),
).reset_index().sort_values('hk_tb_2021', ascending=False)
for _, r in cs.iterrows():
    print(f"  [{int(r['cluster_id'])}] {r['cluster_name']}: {int(r['so_ga'])} ga | "
          f"HK_tb={r['hk_tb_2021']:,.0f} | COVID={r['covid_tb']:.2f}% | "
          f"Rec={r['recovery_tb']:.2f}% | Lines_tb={r['num_lines_tb']:.2f} | "
          f"GiamManh={int(r['trend_giam_manh'])}")

print("\n=== 5. TOP 10 GA ===")
top10 = df.nlargest(10,'passengers_2021')[['station','passengers_2021','cluster_name','num_lines','borough','covid_impact_pct','recovery_rate_pct']]
for i,(_, r) in enumerate(top10.iterrows(),1):
    print(f"  {i}. {r['station']}: {r['passengers_2021']:,.0f} | {r['cluster_name']} | "
          f"Lines={r['num_lines']} | COVID={r['covid_impact_pct']:.1f}% | Rec={r['recovery_rate_pct']:.1f}%")

print("\n=== 6. XU HUONG ===")
print(df['trend_category'].value_counts().to_string())

print("\n=== 7. FEATURES STATS ===")
print(df[features].describe().round(2).to_string())

print("\n=== 8. CLUSTER NAME LOGIC ===")
# Verify naming: cluster_means sorted descending -> CLUSTER_NAMES[rank]
CLUSTER_NAMES = ["Siêu trung tâm","Ga lớn","Ga trung bình","Ga nhỏ","Ga ít khách","Ga rất ít khách"]
cm = df.groupby('cluster_id')['passengers_2021'].mean().sort_values(ascending=False)
print("Thu tu cum theo luu luong giam dan:")
for rank, (cid, mean) in enumerate(cm.items()):
    print(f"  Rank {rank} -> cluster_id={cid} -> Ten={CLUSTER_NAMES[rank]} (mean={mean:,.0f})")

print("\n=== 9. NORMALIZE STATION TEST ===")
# Test normalize function
import re
LINE_SUFFIXES = [
    " underground station"," overground station"," rail station",
    " dlr station"," elizabeth line station"," tram stop"," station",
    " lu"," lo"," nr"," dlr"," el"," tfl"," (dis)"," (bak)"," (h&c)",
    " (d & p)"," h&c"," dis"," bak"," d & p"," bakerloo"," circle",
    " central"," hammersmith & city"," hammersmith and city",
]
def normalize(name):
    if pd.isna(name): return ""
    text = str(name).strip().lower()
    text = text.replace("&"," and").replace("st.","st")
    text = re.sub(r"\s+"," ",text)
    changed=True
    while changed:
        changed=False
        for s in LINE_SUFFIXES:
            if text.endswith(s):
                text=text[:-len(s)].strip(); changed=True; break
    text=re.sub(r"[^\w\s']"," ",text)
    text=re.sub(r"\s+"," ",text).strip()
    if text in ["bank","monument"]: return "bank and monument"
    if text=="crossharbour and london arena": return "crossharbour"
    if text in ["heathrow terminals 123","heathrow terminals 2 and 3","heathrow terminals 1 2 3"]: return "heathrow terminals 1 2 3"
    if text in ["hammersmith d and p","hammersmith (d and p)","hammersmith"]: return "hammersmith"
    return text

test_cases=[
    ("Bank Underground Station","bank and monument"),
    ("King's Cross St. Pancras Underground Station","king's cross st pancras"),
    ("Liverpool Street NR","liverpool street"),
    ("Heathrow Terminals 1 2 3 Underground Station","heathrow terminals 1 2 3"),
    ("Hammersmith (D & P) Station","hammersmith"),
    ("Bank","bank and monument"),
    ("Oxford Circus Underground Station","oxford circus"),
]
print("Test normalize_station_name:")
all_ok=True
for inp, expected in test_cases:
    result=normalize(inp)
    ok = "OK" if result==expected else "FAIL"
    if ok=="FAIL": all_ok=False
    print(f"  [{ok}] '{inp}' -> '{result}' (expected: '{expected}')")
print(f"All tests passed: {all_ok}")

print("\n=== 10. UNMATCHED STATS ===")
try:
    with open('outputs/unmatched_stations_log.txt','r',encoding='utf-8') as f:
        lines = f.readlines()
    for l in lines[:10]:
        print(' ', l.rstrip())
except: print("  (khong doc duoc log)")
