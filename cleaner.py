import pandas as pd
import numpy as np

# 1. Load data
df = pd.read_excel("[FILE PATH]")

# 2. Make a copy (safety best practice)
df_clean = df.copy()

# 3. Remove exact duplicate rows (safe)
df_clean = df_clean.drop_duplicates()

# 4. Handle missing NUMERIC values
numeric_cols = df_clean.select_dtypes(include=["int64", "float64"]).columns

for col in numeric_cols:
    df_clean[col] = df_clean[col].fillna(df_clean[col].median())

# 5. Handle missing CATEGORICAL values
categorical_cols = df_clean.select_dtypes(include=["object"]).columns

for col in categorical_cols:
    df_clean[col] = df_clean[col].fillna("Unknown")

# 6. Handle date columns (if any)
date_cols = df_clean.select_dtypes(include=["datetime64"]).columns

for col in date_cols:
    df_clean[col] = pd.to_datetime(df_clean[col], errors="coerce")

# 7. Outlier handling (optional but recommended)
for col in numeric_cols:
    Q1 = df_clean[col].quantile(0.25)
    Q3 = df_clean[col].quantile(0.75)
    IQR = Q3 - Q1
    lower = Q1 - 1.5 * IQR
    upper = Q3 + 1.5 * IQR
    df_clean[col] = np.clip(df_clean[col], lower, upper)

# 8. Business-friendly summary
summary_numeric = df_clean.describe()
summary_full = df_clean.describe(include="all")

# 9. Save outputs
df_clean.to_excel("cleaned_company_data.xlsx", index=False)
summary_numeric.to_excel("numeric_summary.xlsx")
summary_full.to_excel("full_summary.xlsx")

# 10. Quick validation output
print("Original rows:", df.shape[0])
print("Cleaned rows:", df_clean.shape[0])
print("Missing values after cleaning:\n", df_clean.isna().sum())
