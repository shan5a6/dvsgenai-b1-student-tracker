# ================================
# IMPORTS
# ================================
import os
import json
import gspread
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
from google.oauth2.service_account import Credentials

sns.set_theme(style="whitegrid")

# ================================
# CONFIG
# ================================
SHEET_ID            = os.environ["SHEET_ID"]
WORKSHEET_INDEX     = int(os.environ.get("WORKSHEET_INDEX", "3"))
CREDENTIALS_JSON    = os.environ["GOOGLE_CREDENTIALS_JSON"]

# ================================
# 1. LOAD & CLEAN DATA
# ================================
def load_and_prepare_data():
    creds_dict = json.loads(CREDENTIALS_JSON)
    creds = Credentials.from_service_account_info(
        creds_dict,
        scopes=[
            "https://spreadsheets.google.com/feeds",
            "https://www.googleapis.com/auth/drive"
        ]
    )
    gc = gspread.authorize(creds)
    sh = gc.open_by_key(SHEET_ID)
    df = pd.DataFrame(sh.get_worksheet(WORKSHEET_INDEX).get_all_records())

    raw_date_cols = df.columns[2:]
    parsed_dates = {}
    for col in raw_date_cols:
        try:
            parsed_dates[col] = pd.to_datetime(col)
        except:
            parsed_dates[col] = None

    working_date_cols = [
        col for col, dt in parsed_dates.items()
        if dt is not None and dt.weekday() != 6
    ]

    df["NAME"] = df["NAME"].astype(str).str.strip()
    df = df[df["NAME"].str.len() > 0]
    df = df[df["NAME"].str.lower() != "nan"]

    for col in working_date_cols:
        df[col] = df[col].astype(str).str.strip().str.lower()

    done_aliases = ["done", "completed", "yes", "y"]
    df["Total_Done"] = df[working_date_cols].isin(done_aliases).sum(axis=1)
    df["Total_Tasks"] = len(working_date_cols)
    df["Completion_Pct"] = (df["Total_Done"] / df["Total_Tasks"] * 100).round(2)

    return df, working_date_cols


# ================================
# 2. PLOT DASHBOARD
# ================================
def plot_dashboard(df, working_date_cols):
    total_enrolled = len(df)
    total_days     = len(working_date_cols)

    active_df   = df[df["Total_Done"] >= (total_days * 0.5)]
    passive_df  = df[(df["Total_Done"] > 0) & (df["Total_Done"] < (total_days * 0.5))]
    inactive_df = df[df["Total_Done"] == 0]

    num_active              = len(active_df)
    num_passive             = len(passive_df)
    num_total_participating = num_active + num_passive

    total_tasks_possible = num_total_participating * total_days
    total_tasks_done     = df["Total_Done"].sum()
    completion_rate      = (total_tasks_done / total_tasks_possible * 100) if total_tasks_possible > 0 else 0

    fig = plt.figure(figsize=(10.8, 20.5))
    gs  = fig.add_gridspec(3, 3, height_ratios=[1.0, 2.8, 2.8])

    kpis = [
        (f"{total_enrolled}",              f"Total Enrolled\n({num_total_participating} Active)"),
        (f"{num_active}A | {num_passive}P", "Active vs Passive"),
        (f"{completion_rate:.1f}%",         "Batch Completion")
    ]

    for i in range(3):
        ax = fig.add_subplot(gs[0, i])
        ax.axis("off")
        ax.text(0.5, 0.6, kpis[i][0], fontsize=32, ha='center', weight='bold')
        ax.text(0.5, 0.3, kpis[i][1], fontsize=12, ha='center', fontweight='bold')
        ax.add_patch(plt.Rectangle((0.05, 0.05), 0.9, 0.9, fill=False, lw=1, transform=ax.transAxes))

    ax_bar  = fig.add_subplot(gs[1, :])
    plot_df = df.sort_values("Total_Done", ascending=False).head(20)
    colors  = ["green" if x >= (total_days * 0.5) else "orange" for x in plot_df["Total_Done"]]
    ax_bar.bar(plot_df["NAME"], plot_df["Total_Done"], color=colors)
    ax_bar.set_title(f"Top 20 Student Progress (Goal: {total_days})", fontsize=16, fontweight='bold')
    ax_bar.set_ylabel("Tasks Completed", fontweight='bold')
    ax_bar.set_xlabel("Students",        fontweight='bold')
    plt.setp(ax_bar.get_xticklabels(), rotation=35, ha="right", fontweight='bold')
    plt.setp(ax_bar.get_yticklabels(), fontweight='bold')

    def create_table(ax, data, title, color):
        ax.axis("off")
        ax.set_facecolor("#f5f7fa")
        ax.set_title(title, fontsize=16, fontweight='bold', color=color, pad=10)
        table = ax.table(
            cellText=data.values,
            colLabels=["Name", "Done"],
            loc="center", cellLoc="center",
            bbox=[0, 0, 1, 1]
        )
        table.auto_set_font_size(False)
        table.set_fontsize(11)
        for (row, col), cell in table.get_celld().items():
            if row == 0:
                cell.set_text_props(weight='bold', color='white')
                cell.set_facecolor("#2c3e50")
            if col == 0:
                cell.set_width(0.75)
                cell.set_text_props(weight='bold', ha='left', fontsize=12)
            if col == 1:
                cell.set_width(0.25)
                cell.set_text_props(ha='center', fontsize=12)

    TOP_N = 30
    ax_t1 = fig.add_subplot(gs[2, 0])
    ax_t2 = fig.add_subplot(gs[2, 1])
    ax_t3 = fig.add_subplot(gs[2, 2])

    create_table(ax_t1, active_df.sort_values("Total_Done", ascending=False)[["NAME","Total_Done"]].head(TOP_N),  "Active Students",   "green")
    create_table(ax_t2, passive_df.sort_values("Total_Done", ascending=False)[["NAME","Total_Done"]].head(TOP_N), "Passive Students",  "orange")
    create_table(ax_t3, inactive_df[["NAME","Total_Done"]].head(TOP_N),                                           "Inactive Students", "red")

    plt.tight_layout()

    now         = datetime.utcnow()
    date_folder = now.strftime("%Y-%m-%d")
    time_stamp  = now.strftime("%H-%M")
    output_dir  = f"dashboards/{date_folder}"
    os.makedirs(output_dir, exist_ok=True)

    output_file = f"{output_dir}/dashboard_{time_stamp}.jpg"
    plt.savefig(output_file, dpi=120, bbox_inches="tight", format="jpg")
    plt.close()

    print(f"✅ Dashboard saved: {output_file}")
    return output_file


# ================================
# 3. MAIN
# ================================
if __name__ == "__main__":
    print("🚀 Starting AI Student Tracker...")
    df, working_date_cols = load_and_prepare_data()
    plot_dashboard(df, working_date_cols)