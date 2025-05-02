import streamlit as st
import pandas as pd
import io

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
                    st.subheader("🏢 Gabungan Daftar Cabang, Jumlah Traffic Tinggi, dan Average Utilization per Provider")

                    # Hitung jumlah munculnya cabang
                    cabang_count = filtered_traffic['Cabang'].value_counts().reset_index()
                    cabang_count.columns = ['Cabang', 'Jumlah Traffic Tinggi']

                    if 'pivot_df_reset' in locals():
                        # Filter berdasarkan nama interface masing-masing provider
                        telkom_df = pivot_df_reset[pivot_df_reset['Interface Name'] == "GigabitEthernet0/0/0-= WAN MPLS TELKOM ="]
                        lintasarta_df = pivot_df_reset[pivot_df_reset['Interface Name'] == "GigabitEthernet0/0/1-= WAN INTERNET LA ="]

                        # Ubah nama kolom Device Name menjadi Cabang
                        telkom_df = telkom_df.rename(columns={"Device Name": "Cabang", "Total Utilization(%)": "Util % 1"})
                        lintasarta_df = lintasarta_df.rename(columns={"Device Name": "Cabang", "Total Utilization(%)": "Util % 2"})

                        # Ambil hanya kolom yang diperlukan
                        telkom_df = telkom_df[["Cabang", "Util % 1"]]
                        lintasarta_df = lintasarta_df[["Cabang", "Util % 2"]]

                        # Bulatkan ke dua angka di belakang koma
                        telkom_df["Util % 1"] = telkom_df["Util % 1"].round(2)
                        lintasarta_df["Util % 2"] = lintasarta_df["Util % 2"].round(2)

                        # Gabungkan semua data
                        merged_df = cabang_count.merge(telkom_df, on='Cabang', how='left')
                        merged_df = merged_df.merge(lintasarta_df, on='Cabang', how='left')

                        # Tambahkan nama provider
                        merged_df["Provider 1"] = "Telkom"
                        merged_df["Provider 2"] = "Lintasarta"

                        # Susun ulang kolom
                        merged_df = merged_df[["Cabang", "Jumlah Traffic Tinggi", "Provider 1", "Util % 1", "Provider 2", "Util % 2"]]

                        st.dataframe(merged_df, use_container_width=True)

                        # Export ke file Excel
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            merged_df.to_excel(writer, index=False, sheet_name='Rekap Traffic')
                            writer.save()
                            processed_data = output.getvalue()

                        # Tombol download
                        st.download_button(
                            label="📥 Download Rekap Cabang (.xlsx)",
                            data=processed_data,
                            file_name='rekap_cabang_traffic.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
                        )
                    else:
                        st.warning("❗ Data Pivot Table belum tersedia untuk digabungkan.")
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

            # Export structured_df ke file Excel
            output_structured = io.BytesIO()
            with pd.ExcelWriter(output_structured, engine='xlsxwriter') as writer:
                structured_df.to_excel(writer, index=False, sheet_name='Structured Pivot')
                writer.save()
                processed_structured_data = output_structured.getvalue()

            # Tombol download
            st.download_button(
                label="📥 Download Pivot Table (Device → Interface) (.xlsx)",
                data=processed_structured_data,
                file_name='pivot_table_device_interface.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

        else:
            st.warning("Tidak ada data yang berhasil digabungkan.")