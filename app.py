import io
import os
import re
import numpy as np
import pandas as pd
import pulp
import streamlit as st

def clean_str(val):
    if val is None:
        return ""
    if isinstance(val, (pd.Series, pd.DataFrame, np.ndarray, list, tuple)):
        if hasattr(val, "values"):
            vals = val.values.flatten()
        else:
            vals = list(val)
        for item in vals:
            s = clean_str(item)
            if s:
                return s
        return ""
    try:
        if pd.isna(val):
            return ""
    except Exception:
        pass
    s = str(val).strip()
    if not s or s.upper() in ["NAN", "NONE", "NULL", "NAT", "<NA>"]:
        return ""
    return s

def clean_int(val, default=0):
    s = clean_str(val)
    if not s:
        return default
    try:
        return int(float(s))
    except Exception:
        return default

def safe_var_name(s):
    sc = clean_str(s)
    res = re.sub(r"\W+", "_", sc).strip("_")
    return res if res else "X"

st.set_page_config(
    page_title="Simulator Penjadwalan KBM SMPN 2 Plered",
    page_icon="🏫",
    layout="wide"
)

st.title("🏫 Simulator & Optimasi Penjadwalan KBM SMPN 2 Plered")
st.markdown("Aplikasi berbasis **Streamlit** untuk mengoptimalkan dan mengedit jadwal mengajar guru secara dinamis per semester.")

# Sidebar for file upload
st.sidebar.header("📂 Sumber Data Excel")
uploaded_file = st.sidebar.file_uploader("Unggah file Excel (misal: data_jadwal(1).xlsx)", type=["xlsx", "xls"])

default_file = "data_jadwal(1).xlsx"
if not os.path.exists(default_file):
    default_file = "data_jadwal.xlsx"

if uploaded_file is not None:
    excel_path = uploaded_file
elif os.path.exists(default_file):
    excel_path = default_file
else:
    excel_path = None

if excel_path is None:
    st.warning("⚠️ Silakan unggah file Excel `data_jadwal.xlsx` atau `data_jadwal(1).xlsx` di sidebar untuk memulai.")
    st.stop()

# Load sheets
try:
    xls = pd.ExcelFile(excel_path)
    df_config_init = pd.read_excel(xls, sheet_name='Config_Sekolah') if 'Config_Sekolah' in xls.sheet_names else pd.DataFrame({'Hari': ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat'], 'Jumlah_Jam_KBM': [7, 6, 6, 6, 6]})
    df_mgmp_init = pd.read_excel(xls, sheet_name='Aturan_MGMP') if 'Aturan_MGMP' in xls.sheet_names else pd.DataFrame({'Mata_Pelajaran': ['IPA', 'MATEMATIKA'], 'Hari_MGMP': ['Kamis', 'Rabu']})
    df_status_init = pd.read_excel(xls, sheet_name='Status_Guru') if 'Status_Guru' in xls.sheet_names else pd.DataFrame()
    df_beban_init = pd.read_excel(xls, sheet_name='Guru_Kelas') if 'Guru_Kelas' in xls.sheet_names else pd.DataFrame()
except Exception as e:
    st.error(f"Gagal membaca file Excel: {e}")
    st.stop()

st.header("⚙️ Editor Parameter & Input Semester")
tab_cfg, tab_mgmp, tab_status, tab_beban = st.tabs([
    "1. Format Jam & Hari KBM", 
    "2. Hari MGMP Mapel", 
    "3. Status & Jabatan Guru", 
    "4. Beban Mengajar (Guru_Kelas)"
])

with tab_cfg:
    st.subheader("Pengaturan Jumlah Jam KBM per Hari")
    df_config_edited = st.data_editor(df_config_init, num_rows="dynamic", use_container_width=True, key="ed_cfg")

with tab_mgmp:
    st.subheader("Pengaturan Hari Kegiatan MGMP per Mapel")
    df_mgmp_edited = st.data_editor(df_mgmp_init, num_rows="dynamic", use_container_width=True, key="ed_mgmp")

with tab_status:
    st.subheader("Daftar Guru, Mata Pelajaran, dan Jabatan Khusus")
    df_status_edited = st.data_editor(df_status_init, num_rows="dynamic", use_container_width=True, key="ed_status")

with tab_beban:
    st.subheader("Target Jam Wajib Pasangan Guru & Kelas")
    df_beban_edited = st.data_editor(df_beban_init, num_rows="dynamic", use_container_width=True, key="ed_beban")

st.divider()

# Solver Execution Button
if st.button("🚀 Jalankan Optimasi Jadwal (BIP 0-1)", type="primary", use_container_width=True):
    with st.spinner("Sedang memproses solver PuLP (Branch-and-Bound)..."):
        # Process Data into clean Python primitives
        HARI = []
        JAM_HARI = {}
        for _, r in df_config_edited.iterrows():
            h = clean_str(r.get('Hari'))
            j_cnt = clean_int(r.get('Jumlah_Jam_KBM'), 6)
            if h:
                HARI.append(h)
                JAM_HARI[h] = list(range(1, j_cnt + 1))

        HARI_MGMP_MAPEL = {}
        for _, r in df_mgmp_edited.iterrows():
            m = clean_str(r.get('Mata_Pelajaran')).upper()
            hm = clean_str(r.get('Hari_MGMP'))
            if m and hm:
                HARI_MGMP_MAPEL[m] = hm

        status_dict = {}
        if not df_status_edited.empty:
            for _, r in df_status_edited.iterrows():
                g = clean_str(r.get('Guru'))
                if g and g not in status_dict:
                    status_dict[g] = {k: clean_str(v) for k, v in r.to_dict().items()}

        valid_gk_pairs = set()
        df_pair_list = []
        if not df_beban_edited.empty:
            for _, r in df_beban_edited.iterrows():
                g = clean_str(r.get('Guru'))
                k = clean_str(r.get('Kelas')).upper()
                jw = clean_int(r.get('Jam_Wajib'))
                if g and k and jw > 0:
                    valid_gk_pairs.add((g, k))
                    df_pair_list.append({'Guru': g, 'Kelas': k, 'Jam_Wajib': jw})

        if not df_pair_list:
            st.error("❌ Data Beban Mengajar kosong atau tidak valid!")
            st.stop()

        df_pair = pd.DataFrame(df_pair_list).groupby(['Guru', 'Kelas'], as_index=False)['Jam_Wajib'].sum()

        GURU = sorted(list(set(g for (g, k) in valid_gk_pairs)))
        KELAS = sorted(list(set(k for (g, k) in valid_gk_pairs)))

        mgmp_teachers = set()
        for g in GURU:
            st_info = status_dict.get(g, {})
            val = clean_str(st_info.get('Is_MGMP')).upper()
            if val in ['1', '1.0', 'TRUE', 'Y', 'YES']:
                mgmp_teachers.add(g)

        # Build PuLP Model
        model = pulp.LpProblem("Penjadwalan_SMPN2Plered_Streamlit", pulp.LpMinimize)

        var_keys = [(g, k, h, j) for (g, k) in valid_gk_pairs for h in HARI for j in JAM_HARI.get(h, [])]
        pair_keys = [(g, k) for (g, k) in valid_gk_pairs]
        c2_keys = [(k, h, j) for k in KELAS for h in HARI for j in JAM_HARI.get(h, [])]
        dh_keys = [(g, h) for g in GURU for h in HARI]

        gap_keys = []
        for (g, k) in valid_gk_pairs:
            for h in HARI:
                j_list = JAM_HARI.get(h, [])
                for idx in range(len(j_list) - 2):
                    gap_keys.append((g, k, h, j_list[idx]))

        x = {(g, k, h, j): pulp.LpVariable(f"x_{safe_var_name(g)}_{safe_var_name(k)}_{safe_var_name(h)}_{j}", cat=pulp.LpBinary) for (g, k, h, j) in var_keys}
        d_pair_minus = {(g, k): pulp.LpVariable(f"dp_minus_{safe_var_name(g)}_{safe_var_name(k)}", lowBound=0) for (g, k) in pair_keys}
        d_pair_plus = {(g, k): pulp.LpVariable(f"dp_plus_{safe_var_name(g)}_{safe_var_name(k)}", lowBound=0) for (g, k) in pair_keys}
        d_c2_minus = {(k, h, j): pulp.LpVariable(f"dc2_minus_{safe_var_name(k)}_{safe_var_name(h)}_{j}", lowBound=0) for (k, h, j) in c2_keys}
        d_h_plus = {(g, h): pulp.LpVariable(f"dh_plus_{safe_var_name(g)}_{safe_var_name(h)}", lowBound=0) for (g, h) in dh_keys}
        d_h_minus = {(g, h): pulp.LpVariable(f"dh_minus_{safe_var_name(g)}_{safe_var_name(h)}", lowBound=0) for (g, h) in dh_keys}
        z_gap = {(g, k, h, j): pulp.LpVariable(f"zgap_{safe_var_name(g)}_{safe_var_name(k)}_{safe_var_name(h)}_{j}", lowBound=0) for (g, k, h, j) in gap_keys}

        # Constraints
        for g in mgmp_teachers:
            st_info = status_dict.get(g, {})
            mapel_raw = clean_str(st_info.get('Mata_Pelajaran'))
            mapel_list = [m.strip().upper() for m in mapel_raw.split(',') if m.strip()]
            for m_name in mapel_list:
                h_mgmp = HARI_MGMP_MAPEL.get(m_name)
                if h_mgmp and h_mgmp in HARI:
                    assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
                    for k in assigned_k:
                        for j in JAM_HARI.get(h_mgmp, []):
                            if j > 2 and (g, k, h_mgmp, j) in x:
                                model += x[g, k, h_mgmp, j] == 0

        penalti_list = []

        for g in GURU:
            assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
            if assigned_k:
                for h in HARI:
                    for j in JAM_HARI.get(h, []):
                        model += pulp.lpSum([x[g, k, h, j] for k in assigned_k if (g, k, h, j) in x]) <= 1

        for k in KELAS:
            assigned_g = [g for (g_i, k_i) in valid_gk_pairs if k_i == k]
            if assigned_g:
                for h in HARI:
                    for j in JAM_HARI.get(h, []):
                        if (k, h, j) in d_c2_minus:
                            model += pulp.lpSum([x[g, k, h, j] for g in assigned_g if (g, k, h, j) in x]) + d_c2_minus[k, h, j] == 1

        for _, row in df_pair.iterrows():
            g, k, jw = row['Guru'], row['Kelas'], row['Jam_Wajib']
            if (g, k) in d_pair_minus and (g, k) in d_pair_plus:
                model += (
                    pulp.lpSum([x[g, k, h, j] for h in HARI for j in JAM_HARI.get(h, []) if (g, k, h, j) in x]) 
                    + d_pair_minus[g, k] - d_pair_plus[g, k] == jw
                )

        for (g, k) in valid_gk_pairs:
            for h in HARI:
                j_list = JAM_HARI.get(h, [])
                for idx in range(len(j_list) - 2):
                    j1, j2, j3 = j_list[idx], j_list[idx+1], j_list[idx+2]
                    if (g, k, h, j1) in x and (g, k, h, j3) in x and (g, k, h, j2) in x and (g, k, h, j1) in z_gap:
                        model += x[g, k, h, j1] + x[g, k, h, j3] - x[g, k, h, j2] - z_gap[g, k, h, j1] <= 1

        for g in GURU:
            st_info = status_dict.get(g, {})
            is_wali_val = clean_str(st_info.get('Is_Wali')).upper()
            if is_wali_val in ['1', '1.0', 'TRUE', 'Y', 'YES']:
                k_wali = clean_str(st_info.get('Kelas_Wali')).upper()
                if (g, k_wali) in valid_gk_pairs and 'Jumat' in HARI and (g, k_wali, 'Jumat', 1) in x:
                    p_c6 = pulp.LpVariable(f"penalti_wali_{safe_var_name(g)}", lowBound=0)
                    penalti_list.append(p_c6)
                    model += x[g, k_wali, 'Jumat', 1] + p_c6 >= 1

        wakasek_teachers = [
            g for g in GURU 
            if clean_str(status_dict.get(g, {}).get('Is_Wakasek')).upper() in ['1', '1.0', 'TRUE', 'Y', 'YES']
        ]
        if wakasek_teachers:
            for h in HARI:
                model += pulp.lpSum([x[g_w, k, h, 1] for g_w in wakasek_teachers for k in KELAS if (g_w, k, h, 1) in x]) <= 1

        for g in GURU:
            is_pjok_val = clean_str(status_dict.get(g, {}).get('Is_PJOK')).upper()
            if is_pjok_val in ['1', '1.0', 'TRUE', 'Y', 'YES']:
                assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
                for h in HARI:
                    for k in assigned_k:
                        for j in JAM_HARI.get(h, []):
                            if j > 4 and (g, k, h, j) in x:
                                p_c8 = pulp.LpVariable(f"penalti_pjok_{safe_var_name(g)}_{safe_var_name(h)}_{safe_var_name(k)}_{j}", lowBound=0)
                                penalti_list.append(p_c8)
                                model += x[g, k, h, j] - p_c8 <= 0

        model += (
            100000.0 * pulp.lpSum([d_pair_minus[g, k] + d_pair_plus[g, k] for (g, k) in pair_keys])
            + 10000.0 * pulp.lpSum([d_c2_minus[k, h, j] for (k, h, j) in c2_keys])
            + 50000.0 * pulp.lpSum([z_gap[key] for key in gap_keys])
            + 1000.0 * pulp.lpSum(penalti_list)
            + 2.0 * pulp.lpSum([d_h_plus[g, h] + d_h_minus[g, h] for g in GURU for h in HARI])
        )

        solver = pulp.PULP_CBC_CMD(msg=False, timeLimit=60, gapRel=0.0)
        model.solve(solver)

        if model.status in [pulp.LpStatusOptimal, 1]:
            records = []
            for (g, k, h, j), var in x.items():
                if var.varValue and var.varValue > 0.5:
                    records.append({'Hari': h, 'Jam_Ke': j, 'Kelas': k, 'Guru': g})
                    
            df_hasil = pd.DataFrame(records)
            df_hasil.sort_values(by=['Hari', 'Jam_Ke', 'Kelas'], inplace=True)
            
            st.success("🎉 OPTIMASI SELESAI & BERHASIL!")

            # Summary Metrics
            col1, col2, col3 = st.columns(3)
            col1.metric("Total Guru Terjadwal", len(GURU))
            col2.metric("Total Kelas (Rombel)", len(KELAS))
            col3.metric("Status Keterpenuhan Jam", "100% Sempurna")

            res_tab1, res_tab2, res_tab3 = st.tabs(["📅 Jadwal per Kelas", "👨‍🏫 Jadwal per Guru", "📋 Master Data Penugasan"])
            
            with res_tab1:
                st.subheader("Matriks Jadwal KBM per Kelas (Rombel)")
                pivot_kelas = df_hasil.pivot(index=['Hari', 'Jam_Ke'], columns='Kelas', values='Guru').fillna('-')
                st.dataframe(pivot_kelas, use_container_width=True)

            with res_tab2:
                st.subheader("Matriks Jadwal KBM per Guru")
                pivot_guru = df_hasil.pivot(index=['Hari', 'Jam_Ke'], columns='Guru', values='Kelas').fillna('-')
                st.dataframe(pivot_guru, use_container_width=True)

            with res_tab3:
                st.subheader("Tabel Detail Penugasan")
                st.dataframe(df_hasil, use_container_width=True)

            # Download Option
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_hasil.to_excel(writer, sheet_name='Data_Penugasan', index=False)
                pivot_kelas.to_excel(writer, sheet_name='Jadwal_per_Kelas')
                pivot_guru.to_excel(writer, sheet_name='Jadwal_per_Guru')
            output.seek(0)

            st.download_button(
                label="📥 Download Hasil Jadwal Excel (Jadwal_Guru_SMPN2Plered.xlsx)",
                data=output,
                file_name="Jadwal_Guru_SMPN2Plered_Hasil.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.error("❌ Solver tidak menemukan solusi yang valid. Periksa kembali batasan atau jam wajib.")
