import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
from scipy.stats import pearsonr

# Load dataset
df = pd.read_excel("Documents/Mar24_Mar25_Cleansed.xlsx", sheet_name="Main")

# --- Prepare relevant columns ---
# Convert Gender to numeric
df["Gender_num"] = df["Gender"].map({"Male": 1, "Female": 0})

# Use existing Age column or calculate from DOB
if "RajiNewColumn-Age" in df.columns:
    df["Age"] = pd.to_numeric(df["RajiNewColumn-Age"], errors="coerce")
elif "DOB" in df.columns:
    df["DOB"] = pd.to_datetime(df["DOB"], errors="coerce")
    df["Age"] = (pd.Timestamp("today") - df["DOB"]).dt.days // 365

# Keep only relevant numeric columns
cols = ["Gender", "RajiNewColumn-Age", "IMD rank"]
df_corr = df[cols].dropna()

# --- Pearson correlation matrix ---
corr = df_corr.corr(method="pearson")
print("📊 Pearson Correlation Matrix:")
print(corr, "\n")

# --- Individual correlation values ---
print("Pairwise Correlations:")
for i, c1 in enumerate(cols):
    for c2 in cols[i+1:]:
        r, p = pearsonr(df_corr[c1], df_corr[c2])
        print(f"{c1} vs {c2}: r={r:.3f}, p={p:.3f}")

# --- Heatmap ---
plt.figure(figsize=(6, 4))
sns.heatmap(corr, annot=True, cmap="coolwarm", fmt=".2f")
plt.title("Correlation: Gender, Age, and IMD Rank")
plt.tight_layout()
plt.show()
