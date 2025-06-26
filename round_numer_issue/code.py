## 🧮✨ 11. Smart Integer Adjustment for KPI Breakdown Totals

def adjust_kpi_parts(df, kpi_mapping):
    """
    🧮 Adjust multiple KPI part columns so that their integer-rounded sum matches the integer-rounded total KPI.

    Parameters:
        df (pd.DataFrame): 📋 The input DataFrame containing KPI totals and parts.
        kpi_mapping (dict): 🗺️ A dictionary where:
            🔑 Keys = total KPI column names (e.g., 'VoLTE Traffic (Erl)')
            📌 Values = list of sub-component KPI column names (e.g., ['Volte L900', 'Volte L1800', ...])

    Returns:
        pd.DataFrame: ✅ A new DataFrame with adjusted parts so that:
                      ➕ Rounded total = sum of rounded parts
    """
    df_adj = df.copy()

    for total_col, part_cols in kpi_mapping.items():
        # 1️⃣ Round total KPI to nearest integer
        total_rounded = df_adj[total_col].round().astype("Int64")

        # 2️⃣ Floor each part (sub-KPI) and calculate remaining decimals (remainders)
        parts_float = df_adj[part_cols]
        parts_floor = parts_float.apply(np.floor)  # ⬇️ Floor values
        remainders = parts_float - parts_floor     # 🔁 Remainder to guide rounding

        # 3️⃣ Calculate how many units need to be added to parts to match the rounded total
        units_needed = (total_rounded - parts_floor.sum(axis=1)).astype(int)

        # 4️⃣ Distribute missing units to the parts with the largest remainders
        adjusted_parts = parts_floor.copy()
        for i in range(len(df_adj)):
            n = units_needed[i]
            if n > 0:
                top_indices = remainders.iloc[i].nlargest(n).index  # 🥇 Get top remainders
                adjusted_parts.loc[i, top_indices] += 1             # ➕ Add 1 to the top parts

        # 5️⃣ Assign adjusted parts and final rounded total to output DataFrame
        df_adj[part_cols] = adjusted_parts.astype("Int64")  # 🔢 Ensure final integer format
        df_adj[total_col] = total_rounded                   # 🧮 Final total

    return df_adj


## 🧩✨ 12. Apply Integer Consistency for KPI Breakdown Components
# 🗺️ Define mapping of total KPIs and their corresponding part columns
kpi_mapping = {
    '2G + 3G + VoLTE Voice Traffic (Erl)': [         # 📞 Total voice
        "2G Voice Traffic (Erl)",
        "3G Voice Traffic (Erl)",
        "VoLTE Traffic (Erl)"
    ],
    "3G + 4G Data Traffic (GB)": [                   # 🌐 Total data
        "3G Data Traffic (GB)",
        "4G Data Traffic (GB)"
    ],
    "VoLTE Traffic (Erl)": [                         # 📶 VoLTE by LTE layers
        "4G Volte Traffic L900",
        "4G Volte Traffic L2100",
        "4G Volte Traffic L1800"
    ],
    "4G Data Traffic (GB)": [                        # 💾 4G data by layers
        "4G Data Volume (GB) L1800",
        "4G Data Volume (GB) L2100",
        "4G Data Volume (GB) L900"
    ]
}

# 🧮 Apply adjustment so integer-rounded sub-KPIs sum up to the rounded total KPIs
merged_df = adjust_kpi_parts(merged_df, kpi_mapping)
