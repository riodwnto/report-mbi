import streamlit as st
import pandas as pd

st.set_page_config(page_title="Report Dashboard", layout="wide")
st.title("📈 Dashboard Report")

tab1, tab2 = st.tabs(["Weekly", "Monthly"])

# -------------------------- TAB WEEKLY --------------------------
with tab1:
    st.header("📊 Weekly Report")

    uploaded_files_weekly = st.file_uploader(
        "Upload file Excel (Advanced Report) - Weekly",
        type=["xlsx"], accept_multiple_files=True, key="weekly_upload"
    )

    final_df = None
    structured_df = None

    if uploaded_files_weekly:
        st.subheader("📁 Daftar File Advanced Report yang Diupload:")
        for uploaded_file in uploaded_files_weekly:
            st.write(f"- {uploaded_file.name}")

        combined_data = []

        for uploaded_file in uploaded_files_weekly:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Advanced Report", skiprows=8)

                if {'Device Name', 'Total Utilization(%)', 'Interface Name'}.issubset(df.columns):
                    df_filtered = df[['Device Name', 'Total Utilization(%)', 'Interface Name']].copy()
                    df_filtered.dropna(subset=['Device Name'], inplace=True)
                    df_filtered = df_filtered[df_filtered['Device Name'].astype(str).str.startswith("RTR")]
                    df_filtered["Source File"] = uploaded_file.name
                    combined_data.append(df_filtered)
                else:
                    st.warning(f"Kolom tidak lengkap di file: {uploaded_file.name}")
            except Exception as e:
                st.error(f"❌ Gagal memproses file {uploaded_file.name}: {str(e)}")

        if combined_data:
            final_df = pd.concat(combined_data, ignore_index=True)
            st.subheader("📄 Data Gabungan (Device diawali 'RTR'):")
            st.dataframe(final_df, use_container_width=True)

            st.subheader("📊 Pivot Table (Average)")

            pivot_df = final_df.pivot_table(
                index=['Device Name', 'Interface Name'],
                values='Total Utilization(%)',
                aggfunc='mean'
            )

            pivot_df_sorted = pivot_df.sort_values(by='Total Utilization(%)', ascending=False)

            st.write("📌 Pivot Table Struktur Device → Interface:")

            pivot_df_reset = pivot_df_sorted.reset_index()
            structured_data = []

            for device in pivot_df_reset['Device Name'].unique():
                sub_df = pivot_df_reset[pivot_df_reset['Device Name'] == device]

                structured_data.append({
                    "Device Name": device,
                    "Interface Name": "",
                    "Total Utilization(%)": round(sub_df['Total Utilization(%)'].max(), 2)
                })

                for _, row in sub_df.iterrows():
                    structured_data.append({
                        "Device Name": "",
                        "Interface Name": row['Interface Name'],
                        "Total Utilization(%)": round(row['Total Utilization(%)'], 2)
                    })

            structured_df = pd.DataFrame(structured_data)
            st.dataframe(structured_df, use_container_width=True)

        else:
            st.warning("Tidak ada data yang berhasil digabungkan.")

    # ------------------ Upload File Traffic Analytic ------------------
st.divider()
st.subheader("📥 Upload File Traffic Analytic")

file_traffic = st.file_uploader(
    "Upload file Excel untuk Traffic Analytic (Weekly)",
    type=["xlsx"], key="traffic_upload"
)

if file_traffic:
    try:
        traffic_raw_df = pd.read_excel(file_traffic, sheet_name="Traffic Analytic", header=None)

        tanggal = None
        extracted_data = []

        for index, row in traffic_raw_df.iterrows():
            row_data = row.dropna().tolist()

            if any(isinstance(cell, str) and "-" in cell for cell in row_data):
                tanggal = row_data[0]

            if len(row_data) >= 5 and isinstance(row_data[2], str) and row_data[2].startswith('RTR'):
                entry = {
                    'Tanggal': traffic_raw_df.iloc[index, 2],
                    'Cabang': traffic_raw_df.iloc[index, 3],
                }
                extracted_data.append(entry)

        if extracted_data:
            selected_traffic = pd.DataFrame(extracted_data)
            selected_traffic['Tanggal'] = pd.to_datetime(selected_traffic['Tanggal'], dayfirst=True, errors='coerce')

            # ------------------ Tambahan Filter Range Tanggal ------------------
            st.subheader("📅 Pilih Rentang Tanggal")
            min_date = selected_traffic['Tanggal'].min()
            max_date = selected_traffic['Tanggal'].max()

            start_date, end_date = st.date_input(
                "Pilih rentang tanggal:",
                value=(min_date, max_date),
                min_value=min_date,
                max_value=max_date
            )

            mask = (selected_traffic['Tanggal'] >= pd.to_datetime(start_date)) & (selected_traffic['Tanggal'] <= pd.to_datetime(end_date))
            filtered_traffic = selected_traffic.loc[mask]

            st.subheader("📄 Data Traffic Analytic (Kolom: Tanggal dan Cabang - Sesuai Range Tanggal)")
            st.dataframe(filtered_traffic, use_container_width=True)

            # ------------------ Tambahan: Rekap Cabang dan Jumlah Kemunculannya ------------------
            if not filtered_traffic.empty:
                st.subheader("🏢 Gabungan Daftar Cabang, Jumlah Kemunculan, dan Average Utilization")

                # Hitung jumlah munculnya cabang
                cabang_count = filtered_traffic['Cabang'].value_counts().reset_index()
                cabang_count.columns = ['Cabang', 'Jumlah Muncul']

                # Gabungkan dengan pivot_df (average utilization)
                if structured_df is not None:
                    # Sesuaikan nama kolom supaya bisa merge
                    structured_df.rename(columns={"Device Name": "Cabang"}, inplace=True)
                    merged_df = pd.merge(cabang_count, structured_df, on='Cabang', how='left')

                    st.dataframe(merged_df, use_container_width=True)
                else:
                    st.warning("❗ Data Advanced Report belum tersedia untuk digabungkan.")
            else:
                st.info("Tidak ada data dalam rentang tanggal yang dipilih.")
        else:
            st.warning("Tidak ada data yang bisa diekstrak dari file Traffic Analytic.")
    except Exception as e:
        st.error(f"❌ Gagal memproses file Traffic Analytic: {str(e)}")

# -------------------------- TAB MONTHLY --------------------------
with tab2:
    st.header("📊 Monthly Report")

    uploaded_files_monthly = st.file_uploader(
        "Upload file Excel (Advanced Report) - Monthly",
        type=["xlsx"], accept_multiple_files=True, key="monthly_upload"
    )

    if uploaded_files_monthly:
        st.subheader("📁 Daftar File Advanced Report yang Diupload:")
        for uploaded_file in uploaded_files_monthly:
            st.write(f"- {uploaded_file.name}")

        combined_data = []

        for uploaded_file in uploaded_files_monthly:
            try:
                df = pd.read_excel(uploaded_file, sheet_name="Advanced Report", skiprows=8)

                if {'Device Name', 'Total Utilization(%)', 'Interface Name'}.issubset(df.columns):
                    df_filtered = df[['Device Name', 'Total Utilization(%)', 'Interface Name']].copy()
                    df_filtered.dropna(subset=['Device Name'], inplace=True)
                    df_filtered = df_filtered[df_filtered['Device Name'].astype(str).str.startswith("RTR")]
                    df_filtered["Source File"] = uploaded_file.name
                    combined_data.append(df_filtered)
                else:
                    st.warning(f"Kolom tidak lengkap di file: {uploaded_file.name}")
            except Exception as e:
                st.error(f"❌ Gagal memproses file {uploaded_file.name}: {str(e)}")

        if combined_data:
            final_df = pd.concat(combined_data, ignore_index=True)
            st.subheader("📄 Data Gabungan (Device diawali 'RTR'):")
            st.dataframe(final_df, use_container_width=True)

            st.subheader("📊 Pivot Table (Average)")

            pivot_df = final_df.pivot_table(
                index=['Device Name', 'Interface Name'],
                values='Total Utilization(%)',
                aggfunc='mean'
            )

            pivot_df_sorted = pivot_df.sort_values(by='Total Utilization(%)', ascending=False)

            st.write("📌 Pivot Table Struktur Device → Interface:")

            pivot_df_reset = pivot_df_sorted.reset_index()
            structured_data = []

            for device in pivot_df_reset['Device Name'].unique():
                sub_df = pivot_df_reset[pivot_df_reset['Device Name'] == device]

                structured_data.append({
                    "Device Name": device,
                    "Interface Name": "",
                    "Total Utilization(%)": round(sub_df['Total Utilization(%)'].max(), 2)
                })

                for _, row in sub_df.iterrows():
                    structured_data.append({
                        "Device Name": "",
                        "Interface Name": row['Interface Name'],
                        "Total Utilization(%)": round(row['Total Utilization(%)'], 2)
                    })

            structured_df = pd.DataFrame(structured_data)
            st.dataframe(structured_df, use_container_width=True)

        else:
            st.warning("Tidak ada data yang berhasil digabungkan.")
