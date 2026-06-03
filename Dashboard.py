import streamlit as st
import pandas as pd
import plotly.express as px
import io
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from streamlit_option_menu import option_menu

# PENGATURAN HALAMAN & TEMA
st.set_page_config(page_title="Dashboard TA | Kasus Kekerasan di Indonesia", layout="wide", initial_sidebar_state="expanded")

# CSS KUSTOM UNTUK MENGATUR TAMPILAN
st.markdown("""
    <style>
    /* Mengatur jarak halaman utama (Kanan) */
    .block-container {
        padding-top: 1.5rem !important; 
        padding-bottom: 1rem !important;
        padding-left: 2rem !important;
        padding-right: 2rem !important;
    }
    header { visibility: hidden !important; }
    
    /* Mengatur logo */
    img { object-fit: contain !important; }
    
    /* Sidebar Supaya Padat */
    [data-testid="stSidebar"] [data-testid="stSidebarUserContent"] {
        padding-top: 0rem !important;
        padding-bottom: 0rem !important;
        gap: 0rem !important; 
    }
    
    /* Membuat Teks Menjadi Rata Tengah */
    h1, h2, h3, h4, h5, h6, .stInfo { 
        text-align: center !important; 
        justify-content: center !important;
    }
    div.stAlert { text-align: center !important; display: flex; justify-content: center;}
    
    /* Warna Utama Halaman dan Sidebar */
    .stApp { background-color: #041C32; color: #E0E0E0; }
    [data-testid="stSidebar"] { background-color: #04293A !important; border-right: 2px solid #ECB365; }
    
    /* Kotak Metrik (Total Provinsi, Kasus, dll) */
    .stMetric { 
        background-color: #04293A; 
        padding: 10px; 
        border-radius: 8px; 
        border: 1px solid #064663; 
        text-align: center !important; 
        box-shadow: 1px 1px 5px rgba(0,0,0,0.3);
    }
    [data-testid="stMetricLabel"] {
        display: flex !important;
        justify-content: center !important;
        align-items: center !important;
        width: 100% !important;
    }
    [data-testid="stMetricLabel"] > div {
        display: flex !important;
        justify-content: center !important;
        text-align: center !important;
        width: 100% !important;
    }
    [data-testid="stMetricValue"] { 
        color: #ECB365 !important; 
        font-weight: bold; 
        font-size: 1.6rem !important; 
        justify-content: center !important;
        display: flex !important;
        width: 100% !important;
    }
    
    h1, h2, h3, h4, h5 { color: #ECB365 !important; margin-top: 0px !important; margin-bottom: 5px !important;}
    
    /* Margin Checkbox Sidebar */
    div[data-testid="stCheckbox"] { min-height: 1rem !important; padding: 0 !important; justify-content: flex-start;}
    label[data-baseweb="checkbox"] { margin-bottom: -15px !important;}
    .stSelectbox, .stMultiSelect { margin-bottom: -10px !important; }
    </style>
    """, unsafe_allow_html=True)

# MEMORI UNTUK MENYIMPAN JUMLAH CLUSTER
if 'jumlah_cluster' not in st.session_state:
    st.session_state.jumlah_cluster = 3

# SIDEBAR (LOGO, NAVIGASI & FILTER)
with st.sidebar:
    col_img1, col_img2, col_img3 = st.columns([1, 6, 1])
    with col_img2:
        try:
            st.image("Logo USU-Vertikal.png", use_container_width=True) 
        except:
            st.markdown("<p>*(Logo USU)*</p>", unsafe_allow_html=True)
            
    st.markdown("<h3 style='font-size: 20px; margin-bottom: 0px;'>TUGAS AKHIR</h3>", unsafe_allow_html=True)
    
    # Menu Navigasi
    menu = option_menu(
        menu_title=None, 
        options=["Beranda", "Analisis Cluster", "Dataset"],
        icons=["house", "diagram-3", "database"],
        default_index=0,
        styles={
            "container": {"padding": "10px !important", "background-color": "#04293A", "margin-bottom": "5px", "margin-top": "10px", "border": "none"},
            "icon": {"color": "#ECB365", "font-size": "14px"},
            "nav-link": {"font-size": "13px", "text-align": "left", "margin": "5px 0px", "padding": "10px", "border-radius": "8px", "background-color": "#041C32", "border": "1px solid #064663", "color": "#E0E0E0"},
            "nav-link-selected": {"background-color": "#3EDBF0", "font-weight": "bold", "color": "#041C32", "border": "1px solid #3EDBF0"},
        }
    )
    st.markdown('<p style="font-size: 14px; font-weight: bold; margin-top: 15px; margin-bottom: 0px; text-align: center;">Filter Data</p>', unsafe_allow_html=True)
    
    # Membaca daftar tahun (nama sheet) secara live langsung dari file Excel asli
    try:
        xl = pd.ExcelFile("Data Kekerasan di Indonesia.xlsx")
        daftar_tahun_live = xl.sheet_names
        # Pastikan pilihan "Integrasi" tetap berada di paling bawah
        if "Integrasi" in daftar_tahun_live:
            daftar_tahun_live.remove("Integrasi")
            daftar_tahun_live.append("Integrasi")
        else:
            daftar_tahun_live.append("Integrasi")
    except:
        # Pilihan darurat jika file excel belum terbaca
        daftar_tahun_live = ["2021", "2022", "2023", "2024", "2025", "Integrasi"]

    # Kotak pilihan tahun sekarang otomatis mengikuti isi Excel terbaru
    pilihan_tahun = st.selectbox("Tahun:", daftar_tahun_live, label_visibility="collapsed")
    st.markdown('<p style="font-size: 14px; font-weight: bold; margin-top: 15px; margin-bottom: 0px; text-align: center;">Cari Provinsi:</p>', unsafe_allow_html=True)
    
# LOAD DATA & LOGIKA FILTER
@st.cache_data
def get_data(sheet):
    try:
        return pd.read_excel("Data Kekerasan di Indonesia.xlsx", sheet_name=sheet)
    except Exception as e:
        return None

df_raw = get_data(pilihan_tahun)

if df_raw is not None:
    expected_cols = ["Fisik", "Psikis", "Seksual", "Eksploitasi", "TPPO", "Penelantaran", "Lainnya"]
    jenis_kekerasan_all = [col for col in expected_cols if col in df_raw.columns]
    
    with st.sidebar:
        list_Provinsi = df_raw['Provinsi'].unique().tolist()
        Provinsi_terpilih = st.multiselect("Provinsi", list_Provinsi, default=[], label_visibility="collapsed", placeholder="Ketik provinsi...")
        if not Provinsi_terpilih:
            Provinsi_terpilih = list_Provinsi
            
        st.markdown('<p style="font-size: 14px; font-weight: bold; margin-top: 15px; margin-bottom: 0px; text-align: center;">Jenis Kekerasan:</p>', unsafe_allow_html=True)
        kolom_terpilih = []
        for col in jenis_kekerasan_all:
            if st.checkbox(col, value=True):
                kolom_terpilih.append(col)

    df = df_raw[df_raw['Provinsi'].isin(Provinsi_terpilih)].copy()

    # JUDUL UTAMA DINAMIS SESUAI HALAMAN
    if menu == "Beranda":
        judul_halaman = "Dashboard Kasus Kekerasan terhadap Anak dan Perempuan di Indonesia"
    elif menu == "Analisis Cluster":
        judul_halaman = "Analisis Cluster Kasus Kekerasan terhadap Anak dan Perempuan di Indonesia"
    elif menu == "Dataset":
        judul_halaman = "Dataset Kasus Kekerasan terhadap Anak dan Perempuan di Indonesia"

    st.markdown(f"""
        <div style='background-color: #04293A; border: 1px solid #ECB365; border-radius: 8px; padding: 15px; margin-bottom: 15px; text-align: center;'>
            <h4 style='margin: 0; color: #ECB365; font-weight: bold; font-size: 30px;'>{judul_halaman}</h4>
        </div>
    """, unsafe_allow_html=True)

    if not df.empty and len(kolom_terpilih) > 0:
        df['Total Kasus'] = df[kolom_terpilih].sum(axis=1)
        jenis_dominan = df[kolom_terpilih].sum().idxmax()
        prov_tertinggi = df.loc[df['Total Kasus'].idxmax()]['Provinsi']
        
        # PROSES K-MEANS
        if len(df) >= 5:
            X = df[kolom_terpilih]
            scaler = StandardScaler()
            X_scaled = scaler.fit_transform(X)
            
            kmeans = KMeans(n_clusters=st.session_state.jumlah_cluster, random_state=42, n_init=10)
            df['Cluster_ID'] = kmeans.fit_predict(X_scaled)
            
            # Mengurutkan Cluster berdasarkan rata-rata (Cluster 1 = Tertinggi)
            cluster_means_sorted = df.groupby('Cluster_ID')['Total Kasus'].mean().sort_values(ascending=False).index.tolist()
            
            def assign_cluster_name(c_id):
                rank = cluster_means_sorted.index(c_id) + 1
                return f"Cluster {rank}"
                
            df['Nama Cluster'] = df['Cluster_ID'].apply(assign_cluster_name)
            
           # Map Warna berdasarkan urutan
            warna_cluster_map = {}
            warna_tengah = ["#FFE235", "#FFA500", "#3EDBF0", "#B28DFF"]
            
            for i in range(1, st.session_state.jumlah_cluster + 1):
                nama_c = f"Cluster {i}"
                if i == 1:
                    warna_cluster_map[nama_c] = "#FF4B4B" # Selalu Merah (Tinggi)
                elif i == st.session_state.jumlah_cluster:
                    warna_cluster_map[nama_c] = "#00FF00" # Selalu Hijau (Rendah)
                else:
                    # Berikan warna berbeda dari palet untuk tiap cluster di tengah
                    warna_cluster_map[nama_c] = warna_tengah[i - 2]

        # 1. HALAMAN BERANDA
        if menu == "Beranda":
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Total Provinsi", f"{len(df)}")
            c2.metric("Total Kasus", f"{df['Total Kasus'].sum():,}")
            c3.metric("Rata-rata Kasus", f"{df['Total Kasus'].mean():.0f}")
            c4.metric(f"Kasus Terbanyak ({jenis_dominan})", f"{df[jenis_dominan].sum():,}")
            
            # Ringkasan Data
            st.error(f"PERHATIAN: Jumlah kekerasan tertinggi didominasi oleh Kekerasan {jenis_dominan} dengan provinsi tertinggi yaitu {prov_tertinggi}.")
            
            # Grafik Rata-rata Jenis Kekerasan Per Cluster
            if len(df) >= 5:
                df_mean_jenis = df.groupby('Nama Cluster')[kolom_terpilih].mean().reset_index()
                df_mean_jenis = df_mean_jenis.sort_values('Nama Cluster')
                
                df_mean_melt = df_mean_jenis.melt(id_vars='Nama Cluster', value_vars=kolom_terpilih, var_name='Jenis Kekerasan', value_name='Rata-rata Kasus')
                fig_cluster_avg = px.bar(df_mean_melt, x='Nama Cluster', y='Rata-rata Kasus', color='Jenis Kekerasan', barmode='group',
                                         color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_cluster_avg.update_layout(title={'text': "Rata-rata Jenis Kekerasan pada Tiap Cluster", 'x': 0.5, 'xanchor': 'center'},
                                              template="plotly_dark", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=40, b=10))
                st.plotly_chart(fig_cluster_avg, use_container_width=True)
                
            # Diagram Garis
            df_melt = df.melt(id_vars=['Provinsi'], value_vars=kolom_terpilih, var_name='Jenis Kekerasan', value_name='Jumlah Kasus')
            fig_line = px.line(df_melt, x='Provinsi', y='Jumlah Kasus', color='Jenis Kekerasan', markers=True, 
                               color_discrete_sequence=px.colors.qualitative.Pastel)
            fig_line.update_layout(title={'text': "Jumlah Kasus Kekerasan tiap Provinsi", 'x': 0.5, 'xanchor': 'center'},
                                   xaxis_tickangle=-45, template="plotly_dark", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=40, b=10))
            st.plotly_chart(fig_line, use_container_width=True)
            
            # Proporsi & Top 10
            col_1a, col_1b = st.columns([1, 1.2])
            with col_1a:
                df_pie = df[kolom_terpilih].sum().reset_index()
                df_pie.columns = ['Jenis', 'Jumlah']
                fig_pie = px.pie(df_pie, values='Jumlah', names='Jenis', hole=0.5, 
                                 color_discrete_sequence=px.colors.sequential.Oryel)
                fig_pie.update_layout(title={'text': "Total kasus berdasarkan jenis kekerasan", 'x': 0.5, 'xanchor': 'center'},
                                      template="plotly_dark", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=40, b=10, l=0, r=0))
                st.plotly_chart(fig_pie, use_container_width=True)
            with col_1b:
                df_top = df.nlargest(10, 'Total Kasus').sort_values('Total Kasus', ascending=True)
                fig_bar_top = px.bar(df_top, x='Total Kasus', y='Provinsi', orientation='h', 
                                     color='Total Kasus', color_continuous_scale='Oryel')
                fig_bar_top.update_layout(title={'text': "10 Provinsi dengan Kasus Tertinggi", 'x': 0.5, 'xanchor': 'center'},
                                          template="plotly_dark", plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=40, b=10, l=0, r=0))
                st.plotly_chart(fig_bar_top, use_container_width=True)
                
            # Korelasi & Outlier
            # Korelasi & Outlier
            col_4a, col_4b = st.columns(2)
            
            with col_4a:
                fig_box = px.box(df_melt, x='Jenis Kekerasan', y='Jumlah Kasus', color='Jenis Kekerasan',
                                 template="plotly_dark", color_discrete_sequence=px.colors.qualitative.Pastel)
                fig_box.update_layout(title={'text': "Pencilan Data (Boxplot)", 'x': 0.5, 'xanchor': 'center'},
                                      plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', showlegend=False, margin=dict(t=40, b=10))
                st.plotly_chart(fig_box, use_container_width=True)
                
                # --- TAMBAHAN KETERANGAN BOXPLOT ---
                with st.expander("💡 Cara Membaca Visualisasi Boxplot"):
                    st.write("""
                    **Boxplot** digunakan untuk melihat sebaran jumlah kasus dan menemukan **provinsi dengan kasus ekstrem**.
                    * **Area Kotak :** Mewakili 50% data provinsi yang berada di tingkat normal. Garis di tengah kotak adalah nilai tengah (median).
                    * **Garis Batas atas dan bawah :** Menunjukkan rentang sebaran umum dari jumlah kasus kekerasan di berbagai Provinsi.
                    * **Titik-titik di luar garis:** Ini adalah **Pencilan (Outlier)**. Artinya, ada provinsi-provinsi tertentu yang jumlah kasusnya **sangat tinggi dan tidak wajar** jika dibandingkan dengan rata-rata provinsi lain di Indonesia.
                    """)

            with col_4b:
                korelasi = df[kolom_terpilih].corr()
                fig_corr = px.imshow(korelasi, text_auto=".2f", aspect="auto",
                                     color_continuous_scale='RdBu_r', template="plotly_dark")
                fig_corr.update_layout(title={'text': "Korelasi Antar Jenis Kekerasan", 'x': 0.5, 'xanchor': 'center'},
                                       plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=40, b=10))
                st.plotly_chart(fig_corr, use_container_width=True)
                
                # --- TAMBAHAN KETERANGAN KORELASI ---
                with st.expander("💡 Cara Membaca Matriks Korelasi"):
                    st.write("""
                    **Korelasi** mengukur seberapa kuat hubungan antar dua jenis kekerasan (Nilai dari 0 hingga 1).
                    * **Mendekati Angka 1 (Warna Merah):** Hubungannya **sangat kuat**. Artinya, jika suatu kasus kekerasan tinggi di suatu provinsi, maka sangat besar kemungkinannya satu kasus kekerasan lainnya juga ikut tinggi (sering terjadi bersamaan).
                    * **Mendekati Angka 0 (Warna Biru):** Hubungannya **lemah**. Artinya, kedua jenis kekerasan tersebut tidak saling mempengaruhi atau jarang terjadi secara bersamaan di wilayah yang sama.
                    """)

        # 2. HALAMAN ANALISIS CLUSTER
        elif menu == "Analisis Cluster":
            if len(df) >= 5: 
                st.markdown('<p style="font-size: 16px; font-weight: bold; text-align: left;">Atur Jumlah Cluster (k):</p>', unsafe_allow_html=True)
                st.number_input("Jumlah Cluster", min_value=2, max_value=6, step=1, key='jumlah_cluster', label_visibility="collapsed")
                st.write("---")
                
                pca = PCA(n_components=2)
                X = df[kolom_terpilih]
                scaler = StandardScaler()
                X_scaled = scaler.fit_transform(X)
                pca_result = pca.fit_transform(X_scaled)
                df['PC1'] = pca_result[:, 0]
                df['PC2'] = pca_result[:, 1]
                
                col_pca, col_desc = st.columns([2, 1])
                with col_pca:
                    st.markdown("**Visualisasi K-Means Cluster (PCA)**")
                    
                    # Membuat daftar urutan cluster yang benar (1 sampai k)
                    urutan_cluster = [f"Cluster {i}" for i in range(1, st.session_state.jumlah_cluster + 1)]
                    
                    fig_scatter = px.scatter(
                        df, x='PC1', y='PC2', color='Nama Cluster', hover_data=['Provinsi'],
                        template="plotly_dark", color_discrete_map=warna_cluster_map, size='Total Kasus',
                        category_orders={"Nama Cluster": urutan_cluster} # <-- INI KODE TAMBAHANNYA
                    )
                    
                    fig_scatter.update_layout(title_x=0.5, plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=0, b=0))
                    st.plotly_chart(fig_scatter, use_container_width=True)
                    
                with col_desc:
                    st.markdown("**Karakteristik Cluster**")
                    
                    for i in range(1, st.session_state.jumlah_cluster + 1):
                        nama_c = f"Cluster {i}"
                        jml_prov = len(df[df['Nama Cluster'] == nama_c])
                        
                        # Menentukan teks karakteristik
                        if i == 1:
                            teks_kat = "Kategori: TINGGI"
                        elif i == st.session_state.jumlah_cluster:
                            teks_kat = "Kategori: RENDAH"
                        else:
                            teks_kat = "Kategori: SEDANG"
                            
                        st.info(f"**{nama_c}**\n\n{teks_kat}\n\nTotal: {jml_prov} Provinsi", icon=None)

                st.write("---")
                st.markdown("**Daftar Anggota Tiap Cluster:**")
                cols_anggota = st.columns(st.session_state.jumlah_cluster)
                
                for i in range(1, st.session_state.jumlah_cluster + 1):
                    nama_c = f"Cluster {i}"
                    anggota = df[df['Nama Cluster'] == nama_c]['Provinsi'].tolist()
                    with cols_anggota[i-1]:
                        with st.expander(f"{nama_c} ({len(anggota)})"):
                            # Menggunakan list komprehensi dan newline (\n) murni, tanpa tag <br>
                            daftar_urut = "\n\n".join([f"{idx+1}. {prov}" for idx, prov in enumerate(anggota)])
                            st.markdown(daftar_urut)
            else:
                st.warning("Data tidak cukup untuk dilakukan clustering.")

       # 3. HALAMAN DATASET
        elif menu == "Dataset":
            # --- BAGIAN ATAS: Tabel Input & Edit Data (Lebar Penuh & Besar) ---
            st.markdown("#### Input & Edit Data")
            st.caption("Ketik langsung di dalam tabel untuk mengedit angka atau menambah data baru.")
            
            kolom_tampil = ['Provinsi'] + kolom_terpilih
            edited_df = st.data_editor(df[kolom_tampil], num_rows="dynamic", use_container_width=True, height=520)
            
            st.write("---") # Garis pembatas horizontal tengah yang rapi
            
            # BAGIAN BAWAH: Menu Unduh & Unggah
            col_unduh, col_unggah = st.columns(2)
            
            # KOTAK SEBELAH KIRI: UNDUH DATA
            with col_unduh:
                st.markdown("#### \u2193 Unduh Data")
                format_unduh = st.selectbox("Pilih format file untuk diunduh:", [".csv", ".xlsx"], key="download_box_fmt")
                
                if format_unduh == ".csv":
                    csv_data = edited_df.to_csv(index=False).encode('utf-8')
                    st.download_button(
                        label="Unduh File (.csv)",
                        data=csv_data, 
                        file_name=f'Data_Diedit_{pilihan_tahun}.csv', 
                        mime='text/csv',
                        use_container_width=True
                    )
                elif format_unduh == ".xlsx":
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                        edited_df.to_excel(writer, index=False)
                    excel_data = buffer.getvalue()
                    st.download_button(
                        label="Unduh File (.xlsx)",
                        data=excel_data,
                        file_name=f'Data_Diedit_{pilihan_tahun}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        use_container_width=True
                    )
            
            # KOTAK SEBELAH KANAN: UNGGAH DATA (FORM KANAN-KIRI & SIMPAN OTOMATIS)
            with col_unggah:
                st.markdown("#### \u2191 Unggah Data Baru")
                
                # Membungkus input ke dalam Form agar menunggu klik tombol submit
                with st.form(key="form_unggah_dataset_baru"):
                    # MEMBUAT INPUT TAHUN DAN UPLOAD FILE SEJAJAR KANAN-KIRI 
                    col_input_thn, col_input_file = st.columns([1, 1.8])
                    
                    with col_input_thn:
                        st.markdown("<p style='font-size:14px; margin-bottom:5px;'>Tahun Data:</p>", unsafe_allow_html=True)
                        tahun_input = st.text_input("Tahun Data", placeholder="Cth: 2026", label_visibility="collapsed")
                        
                    with col_input_file:
                        st.markdown("<p style='font-size:14px; margin-bottom:5px;'>Pilih File Dataset:</p>", unsafe_allow_html=True)
                        file_unggahan = st.file_uploader("Upload File", type=["csv", "xlsx"], label_visibility="collapsed")
                    
                    # Tombol submit di bawahnya pas di tengah form
                    tombol_submit = st.form_submit_button("Unggah Data", use_container_width=True)
                
                # Eksekusi setelah tombol diklik
                if tombol_submit:
                    if file_unggahan is None:
                        st.warning("\u26A0 Silakan pilih file dataset terlebih dahulu!")
                    elif not tahun_input.strip():
                        st.warning("\u26A0 Silakan ketik tahun data pada kotak tersedia!")
                    else:
                        tahun_data = tahun_input.strip()
                        try:
                            if file_unggahan.name.endswith('.csv'):
                                df_baru = pd.read_csv(file_unggahan)
                            elif file_unggahan.name.endswith('.xlsx'):
                                df_baru = pd.read_excel(file_unggahan)
                                
                            kolom_df_baru = df_baru.columns.tolist()
                            format_sesuai = all(kolom in kolom_df_baru for kolom in kolom_tampil)
                            
                            if format_sesuai:
                                # PROSES SIMPAN DATA KE FILE EXCEL ASLI
                                path_file_asli = "Data Kekerasan di Indonesia.xlsx"
                                
                                # Baca seluruh sheet yang sudah ada biar tidak hilang
                                try:
                                    semua_sheet = pd.read_excel(path_file_asli, sheet_name=None)
                                except:
                                    semua_sheet = {}
                                
                                # Masukkan sheet baru / perbarui sheet lama
                                semua_sheet[tahun_data] = df_baru
                                
                                # Tulis ulang semua sheet kembali ke file Excel utama
                                with pd.ExcelWriter(path_file_asli, engine='openpyxl') as writer:
                                    for nama_sheet, data_sheet in semua_sheet.items():
                                        data_sheet.to_excel(writer, sheet_name=str(nama_sheet), index=False)
                                
                                # Hapus cache fungsi pembaca data agar Streamlit tahu ada data baru masuk
                                get_data.clear()
                                
                                st.success(f"\u2713 Sukses! Tahun {tahun_data} berhasil tersimpan")
                                st.rerun() # Muat ulang halaman agar filter di sidebar langsung berubah
                            else:
                                st.error("\u2715 Format kolom tidak sesuai!")
                                st.info(f"Kolom yang wajib ada di file Anda: **{', '.join(kolom_tampil)}**")
                                
                        except Exception as e:
                            st.error(f"Gagal memproses file: {e}")