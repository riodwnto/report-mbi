import streamlit as st
import pandas as pd

st.set_page_config(page_title="Weekly & Monthly Report", layout="wide")
st.title("📊 Report Weekly & Monthly")

# Upload file
uploaded_files = st.file_uploader("Upload file Excel (bisa lebih dari satu)", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    st.subheader("📁 Daftar File yang Diupload:")
    for uploaded_file in uploaded_files:
        st.write(f"- {uploaded_file.name}")

    combined_data = []

    for uploaded_file in uploaded_files:
        try:
            # Baca sheet dan skip sampai baris ke-9
            df = pd.read_excel(uploaded_file, sheet_name="Advanced Report", skiprows=8)

            # Ambil kolom yang dibutuhkan
            if {'Device Name', 'Total Utilization(%)', 'Interface Name'}.issubset(df.columns):
                df_filtered = df[['Device Name', 'Total Utilization(%)', 'Interface Name']].copy()

                # Filter hanya Device yang diawali 'RTR'
                df_filtered.dropna(subset=['Device Name'], inplace=True)
                df_filtered = df_filtered[df_filtered['Device Name'].astype(str).str.startswith("RTR")]

                df_filtered["Source File"] = uploaded_file.name
                combined_data.append(df_filtered)
            else:
                st.warning(f"Kolom tidak lengkap di file: {uploaded_file.name}")
        except Exception as e:
            st.error(f"❌ Gagal memproses file {uploaded_file.name}: {str(e)}")

    # Gabungkan dan tampilkan data
    if combined_data:
        final_df = pd.concat(combined_data, ignore_index=True)
        st.subheader("📄 Data Gabungan (Device diawali 'RTR'):")
        st.dataframe(final_df, use_container_width=True)

        # Pilih metode agregasi
        st.subheader("📊 Pivot Table")
        agg_func = st.radio("Pilih metode agregasi untuk 'Total Utilization(%)':", ['max', 'average'])

        # Pivot table
        if agg_func == 'max':
            pivot_df = final_df.pivot_table(
                index=['Device Name', 'Interface Name'],
                values='Total Utilization(%)',
                aggfunc='max'
            )
        else:
            pivot_df = final_df.pivot_table(
                index=['Device Name', 'Interface Name'],
                values='Total Utilization(%)',
                aggfunc='mean'
            )

        # Urutkan dari yang tertinggi
        pivot_df_sorted = pivot_df.sort_values(by='Total Utilization(%)', ascending=False)

        # Struktur seperti parent-child
        st.write("📌 Pivot Table Struktur Device → Interface:")

        pivot_df_reset = pivot_df_sorted.reset_index()
        structured_data = []

        for device in pivot_df_reset['Device Name'].unique():
            sub_df = pivot_df_reset[pivot_df_reset['Device Name'] == device]

            # Baris utama
            structured_data.append({
                "Device Name": device,
                "Interface Name": "",
                "Total Utilization(%)": round(sub_df['Total Utilization(%)'].max(), 2)
            })

            # Baris anak (interface)
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