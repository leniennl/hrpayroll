import pandas as pd

df_payrate = pd.read_csv(
    r"H:\PAY RATE.CSV",
    usecols=["Employee code", "Employee Name", "Trans Rate"],
    index_col=0,
    dtype={"Employee code": str},
)
df_payrate.rename(columns={"Trans Rate": "Pay Rate"}, inplace=True)
# print(df_payrate)

df_employment = pd.read_csv(
    r"H:\EMPLOYMENT TAX.CSV",
    usecols=[
        "Employee code",
        "Emp category",
        "Description",
        "Class code",
        "Classification Description",
    ],
    keep_default_na=False,
    index_col=0,
    dtype={"Employee code": str},
)
df_employment.rename(columns={"Description": "Hours/FN"}, inplace=True)
df_employment = df_employment.drop(index="", errors="ignore")
# print(df_employment)

# from numpy import int64, string_
# df_payrate=df_payrate.index.astype(int64)
# df_employment.index=df_employment.index.astype(int64)

df1 = pd.merge(df_payrate, df_employment, on="Employee code")
# print(df1)

df_costcenter = pd.read_csv(
    r"H:\COST CENTER.CSV",
    usecols=["Employee code", "Cost Centre", "Description"],
    keep_default_na=False,
    index_col=0,
    dtype={"Employee code": str},
)
df_costcenter = df_costcenter.drop(index="", errors="ignore")
df_costcenter.rename(columns={"Description": "CC Description"}, inplace=True)
# print(df_costcenter)

"""
df_costcenter.index=df_costcenter.index.astype(int64)
"""

df2 = pd.merge(df1, df_costcenter, on="Employee code")
# print(df2)

df_biographical = pd.read_csv(
    r"H:\BIOGRAPHICAL.CSV",
    usecols=["Employee code", "State", "Date hired", "Term Date", "Sex"],
    keep_default_na=False,
    index_col=0,
    dtype={"Employee code": str},
)
df_biographical = df_biographical.drop(index="", errors="ignore")
df_biographical = df_biographical.loc[df_biographical["Term Date"] == ""]
# print(df_biographical)
"""
df_biographical.index=df_biographical.index.astype(int64)
"""
df3 = pd.merge(df2, df_biographical, on="Employee code", how="left")
# df=df.drop(columns=['Term Date'])
# print(df3)

df4 = df3.drop(columns=["Term Date"])
# print(df4)

df4.to_csv(r"h:\4.Employee Listing.csv")

print("File created.")
