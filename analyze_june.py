import pandas as pd

file_path = r"c:\Users\user\Desktop\Antigravity\データ分析自動化 - スタジアム\スタジアム_データ.xlsx"
df = pd.read_excel(file_path, sheet_name='【データ】蓄積用')
df['日付'] = pd.to_datetime(df['日付'], errors='coerce')
df_june = df[df['日付'].dt.month == 6].copy()

# Phase 1: Total diff by date
df_june['差枚'] = pd.to_numeric(df_june['差枚'], errors='coerce').fillna(0)
daily_diff = df_june.groupby(df_june['日付'].dt.date)['差枚'].sum()
tokubi_diff = daily_diff[pd.to_datetime(daily_diff.index).day.astype(str).str.contains('5|6|7')]
print("--- Phase 1: June Daily Diff (Tokubi) ---")
print(tokubi_diff)
print(f"Total June Tokubi Diff: {tokubi_diff.sum()}")
print(f"Total June Diff: {daily_diff.sum()}")

# Phase 2: Juggler REG prob & VvV/Karakuri
df_june['REG確率分母'] = pd.to_numeric(df_june['REG確率分母'], errors='coerce')
df_june['G数'] = pd.to_numeric(df_june['G数'], errors='coerce').fillna(0)
df_jug = df_june[df_june['機種名'].str.contains('ジャグラー', na=False)]
avg_reg = df_jug['REG確率分母'].mean()
print("\n--- Phase 2: Juggler & Smart Slot ---")
print(f"Juggler Avg REG Denom: 1/{avg_reg:.1f}")

df_sm = df_june[df_june['機種名'].str.contains('ヴァルヴレイヴ|からくり', na=False)]
print(df_sm.groupby('機種名')[['G数', '差枚']].mean())

# Phase 3: New Machines (appeared in June, not in May)
df_may = df[df['日付'].dt.month == 5]
may_models = set(df_may['機種名'].unique())
june_models = set(df_june['機種名'].unique())
new_models = list(june_models - may_models)
print("\n--- Phase 3: New Models ---")
print(f"New Models in June: {new_models}")
if new_models:
    df_new = df_june[df_june['機種名'].isin(new_models)]
    early_new = df_new[df_new['日付'].dt.day <= 10].groupby('機種名')[['G数', '差枚']].mean()
    late_new = df_new[df_new['日付'].dt.day >= 20].groupby('機種名')[['G数', '差枚']].mean()
    print("Early June Averages:")
    print(early_new)
    print("Late June Averages:")
    print(late_new)

# Phase 4: Patterns on specific days (ending numbers, etc.)
# Let's check average G_count for top winning machines early vs late
df_june['末尾'] = df_june['台番号'].astype(str).str[-1]
early_tokubi = df_june[(df_june['日付'].dt.day <= 16) & (df_june['日付'].dt.day.astype(str).str.endswith(('5','6','7')))]
late_tokubi = df_june[(df_june['日付'].dt.day > 16) & (df_june['日付'].dt.day.astype(str).str.endswith(('5','6','7')))]

print("\n--- Phase 4: Pattern Analysis (End numbers on Tokubi) ---")
print("Early Tokubi (Top 3 End Numbers by Avg G):")
print(early_tokubi.groupby('末尾')['G数'].mean().sort_values(ascending=False).head(3))
print("Late Tokubi (Top 3 End Numbers by Avg G):")
print(late_tokubi.groupby('末尾')['G数'].mean().sort_values(ascending=False).head(3))
print("Early Tokubi (Top 3 End Numbers by Total Diff):")
print(early_tokubi.groupby('末尾')['差枚'].sum().sort_values(ascending=False).head(3))
print("Late Tokubi (Top 3 End Numbers by Total Diff):")
print(late_tokubi.groupby('末尾')['差枚'].sum().sort_values(ascending=False).head(3))
