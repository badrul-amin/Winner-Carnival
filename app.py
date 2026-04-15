import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import io

# ======================================================
# 🎉 PAGE CONFIG
# ======================================================
st.set_page_config(
    page_title="AEON Carnival Winner Selector",
    layout="wide",
    page_icon="🎉"
)

st.title("🎉 AEON Carnival Winner Selector")
st.markdown("---")

# ======================================================
# 📂 UPLOAD FILE
# ======================================================
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:

    # ======================================================
    # 1️⃣ READ FIRST SHEET
    # ======================================================
    xl = pd.ExcelFile(uploaded_file)
    first_sheet = xl.sheet_names[0]
    st.info(f"📂 Reading from sheet: {first_sheet}")

    df = pd.read_excel(uploaded_file, sheet_name=first_sheet, dtype=str)

    # --- HEADER MAPPING ---
    COLUMN_MAP = {
        "Member Card": "Member ID",
        "Store": "Store Code"
    }

    df.columns = df.columns.astype(str).str.strip()
    df.rename(columns=COLUMN_MAP, inplace=True)

    df["Store Code"] = df["Store Code"].astype(str).str.strip()
    df["Member ID"] = df["Member ID"].astype(str).str.strip()
    df["Original Store Code"] = df["Store Code"]

    STORE_ALIAS = {"5603": "1015"}
    df["Store Code"] = df["Store Code"].replace(STORE_ALIAS)

    st.success("✅ File processed successfully!")

    # ======================================================
    # 2️⃣ PARAMETERS
    # ======================================================
    ALLOW_MULTIPLE_WINS = False

    overrides = {
        "3009": 5, "3015": 5, "3019": 5, "3020": 5, "3011": 5,
        "3021": 5, "3027": 5, "1046": 5, "1037": 5,
        "1005": 10, "1009": 10, "1010": 10, "1027": 10, "1028": 10,
        "1029": 10, "1030": 10, "1038": 10, "1002": 10,
        "1003": 10, "1006": 10, "1008": 10, "1011": 10, "1012": 10,
        "1014": 10, "1015": 10, "1020": 10, "1022": 10, "1025": 10,
        "1026": 10, "1032": 10, "1035": 10, "1039": 10, "1040": 10,
        "1043": 10, "1044": 10,
        "1001": 28, "1004": 28, "1007": 28, "1013": 28,
        "1016": 28, "1018": 27, "1021": 28
    }

    BACKUPS_PER_STORE = 5

    # ======================================================
    # 🚀 RUN BUTTON
    # ======================================================
    if st.button("🚀 Run Winner Selection"):

        store_results = []
        selected_members = set()

        # ======================================================
        # 3️⃣ STORE SPECIFIC SELECTION
        # ======================================================
        for store, main_count in overrides.items():

            pool = df[df["Store Code"] == store].copy()

            if not ALLOW_MULTIPLE_WINS:
                pool = pool[~pool["Member ID"].isin(selected_members)]

            if pool.empty:
                continue

            pool = pool.sample(frac=1, random_state=42).reset_index(drop=True)

            mains = pool.head(main_count).copy()
            mains["Winner Type"] = "Main"

            if not ALLOW_MULTIPLE_WINS:
                selected_members.update(mains["Member ID"])

            backups = pool.iloc[main_count: main_count + BACKUPS_PER_STORE].copy()
            backups["Winner Type"] = "Backup"

            if not ALLOW_MULTIPLE_WINS:
                selected_members.update(backups["Member ID"])

            store_results.append(mains)
            store_results.append(backups)

        # ======================================================
        # 4️⃣ ANY STORE POOL
        # ======================================================
        any_store_pool = df[~df["Member ID"].isin(selected_members)].copy()
        any_store_pool = any_store_pool.sample(frac=1, random_state=99).reset_index(drop=True)

        any_mains = any_store_pool.head(10).copy()
        any_mains["Winner Type"] = "Main"

        any_backups = any_store_pool.iloc[10:20].copy()
        any_backups["Winner Type"] = "Backup"

        # ======================================================
        # 5️⃣ SAVE TO EXCEL (IN MEMORY)
        # ======================================================
        output = io.BytesIO()

        df_sheet1 = pd.concat(store_results, ignore_index=True) if store_results else pd.DataFrame()
        df_sheet2 = pd.concat([any_mains, any_backups], ignore_index=True)

        REVERSE_MAP = {v: k for k, v in COLUMN_MAP.items()}

        for temp_df in [df_sheet1, df_sheet2]:
            if not temp_df.empty:
                temp_df["Store Code"] = temp_df["Original Store Code"]
                temp_df.drop(columns=["Original Store Code"], inplace=True)
                temp_df.rename(columns=REVERSE_MAP, inplace=True)

        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_sheet1.to_excel(writer, index=False, sheet_name="StoreSpecificWinners")
            df_sheet2.to_excel(writer, index=False, sheet_name="AnyStoreWinners")

        # ======================================================
        # 🎨 COLOR HIGHLIGHTING
        # ======================================================
        output.seek(0)
        wb = load_workbook(output)

        fill_main = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        fill_backup = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            header = [cell.value for cell in ws[1]]

            if "Winner Type" in header:
                type_col = header.index("Winner Type") + 1

                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    w_type = row[type_col - 1].value
                    color = fill_main if w_type == "Main" else fill_backup

                    for cell in row:
                        cell.fill = color

        final_output = io.BytesIO()
        wb.save(final_output)

        # ======================================================
        # 🎉 SUCCESS + CONFETTI
        # ======================================================
        st.success("🎉 Selection Complete!")

        st.balloons()  # 🎈 CONFETTI EFFECT

        # ======================================================
        # ⬇️ DOWNLOAD
        # ======================================================
        st.download_button(
            label="⬇️ Download Official Winners File",
            data=final_output.getvalue(),
            file_name="Official_Winners_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
