import io
import os
import re
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pulp
import streamlit as st

st.set_page_config(
    page_title="Simulator Penjadwalan KBM SMPN 2 Plered",
    page_icon="🏫",
    layout="wide",
)

st.title("🏫 Simulator & Optimasi Penjadwalan KBM SMPN 2 Plered")
st.markdown(
    "Aplikasi berbasis **Streamlit** dan **Binary Integer Programming (BIP 0-1)** "
    "untuk mengoptimalkan jadwal mengajar guru dan KBM kelas secara fleksibel."
)


def safe_name(s):
  """Sanitasi nama variabel untuk PuLP."""
  sc = str(s).strip() if s is not None else "X"
  res = re.sub(r"\W+", "_", sc).strip("_")
  return res if res else "X"


def buat_dict_var(nama_prefix, daftar_kunci, lowBound=None, cat=pulp.LpContinuous):
  """Membuat dictionary variabel keputusan PuLP secara aman."""
  hasil = {}
  for k in daftar_kunci:
    if isinstance(k, (tuple, list)):
      s_name = "_".join(safe_name(x) for x in k)
    else:
      s_name = safe_name(k)
    var_name = f"{nama_prefix}_{s_name}"

    try:
      v = pulp.LpVariable(var_name, lowBound=lowBound, cat=cat)
    except Exception:
      try:
        v = pulp.LpVariable(var_name, lowBound=lowBound, upBound=None, cat=cat)
      except Exception:
        v = pulp.LpVariable(var_name)
        if lowBound is not None:
          v.lowBound = lowBound
        v.cat = cat
    hasil[k] = v
  return hasil


# ==========================================
# SIDEBAR CONTROL: SAKELAR ATURAN (TOGGLES)
# ==========================================
st.sidebar.header("📂 1. Sumber Data Excel")
uploaded_file = st.sidebar.file_uploader(
    "Unggah file Excel (misal: data_jadwal(1).xlsx)", type=["xlsx", "xls"]
)

default_file = "data_jadwal(1).xlsx"
if not os.path.exists(default_file):
  default_file = "data_jadwal.xlsx"

if uploaded_file is not None:
  excel_path = uploaded_file
elif os.path.exists(default_file):
  excel_path = default_file
else:
  excel_path = None

st.sidebar.markdown("---")
st.sidebar.header("🎛️ 2. Sakelar Aturan & Batasan")
st.sidebar.caption("Aktifkan atau nonaktifkan aturan optimasi sesuai kebutuhan:")

toggle_c1 = st.sidebar.toggle(
    "1. Maksimal 1 Kelas per Guru per Jam", value=True
)
toggle_c2 = st.sidebar.toggle(
    "2. Tepat 1 Guru per Kelas per Jam (Tidak Ada Jam Kosong)", value=True
)
toggle_c3 = st.sidebar.toggle(
    "3. Pemenuhan Target Wajib Jam Mengajar (Guru-Kelas)", value=True
)
toggle_block2 = st.sidebar.toggle(
    "4. Block Teaching 2 Jam (Mod 2 & Berurutan)", value=True
)
toggle_jam1_guru = st.sidebar.toggle(
    "5. Pemerataan Jam ke-1 per Guru", value=True
)

# Rule 8: MGMP Options
toggle_mgmp = st.sidebar.toggle(
    "6. Aturan Penanggung Jawab MGMP (Maksimal Jam ke-2)", value=True
)
mgmp_mode = "Gunakan Status Guru (Is_MGMP)"
manual_mgmp_teachers = []
if toggle_mgmp:
  mgmp_mode = st.sidebar.radio(
      "Mode Pemilihan Guru MGMP:",
      [
          "Gunakan Status Guru (Is_MGMP)",
          "Pilih Manual Guru MGMP",
          "Otomatis 1 Guru per Mapel MGMP",
      ],
      key="mgmp_mode_radio",
  )

toggle_wali = st.sidebar.toggle(
    "7. Aturan Wali Kelas (Jumat Jam ke-1)", value=True
)
toggle_wakasek = st.sidebar.toggle(
    "8. Aturan Wakasek (Maksimal 1 di Jam ke-1 per Hari)", value=True
)
toggle_pjok = st.sidebar.toggle(
    "9. Aturan PJOK (Maksimal Jam 1 s.d. 4)", value=True
)
toggle_beban_harian = st.sidebar.toggle(
    "10. Pemerataan Jam Mengajar Guru per Hari", value=True
)

# Data Inisialisasi
df_config_init = pd.DataFrame({
    "Hari": ["Senin", "Selasa", "Rabu", "Kamis", "Jumat"],
    "Jumlah_Jam_KBM": (7, 6, 6, 6, 6),
})
df_mgmp_init = pd.DataFrame({
    "Mata_Pelajaran": ["IPA", "MATEMATIKA"],
    "Hari_MGMP": ["Kamis", "Rabu"],
})
df_status_init = pd.DataFrame({
    "Guru": ["Guru A", "Guru B"],
    "Mata_Pelajaran": ["IPA", "MATEMATIKA"],
    "Is_Wali": ["0", "0"],
    "Kelas_Wali": ["", ""],
    "Is_Wakasek": ["0", "0"],
    "Is_PJOK": ["0", "0"],
    "Is_MGMP": ["0", "0"],
})
df_beban_init = pd.DataFrame({
    "Guru": ["Guru A", "Guru B"],
    "Kelas": ["7A", "7B"],
    "Mata_Pelajaran": ["IPA", "MATEMATIKA"],
    "Jam_Wajib": (4, 4),
})

if excel_path is not None:
  try:
    xls = pd.ExcelFile(excel_path)
    sheet_lower = {str(s).lower().strip(): str(s) for s in xls.sheet_names}
    if "config_sekolah" in sheet_lower:
      df_config_init = pd.read_excel(
          xls, sheet_name=sheet_lower["config_sekolah"]
      )
    if "aturan_mgmp" in sheet_lower:
      df_mgmp_init = pd.read_excel(xls, sheet_name=sheet_lower["aturan_mgmp"])
    if "status_guru" in sheet_lower:
      df_status_init = pd.read_excel(
          xls, sheet_name=sheet_lower["status_guru"]
      )
    if "guru_kelas" in sheet_lower:
      df_beban_init = pd.read_excel(xls, sheet_name=sheet_lower["guru_kelas"])
  except Exception as e:
    st.error(f"Gagal membaca file Excel: {e}")

# Inisialisasi Session State
if "solved" not in st.session_state:
  st.session_state["solved"] = False
if "df_hasil" not in st.session_state:
  st.session_state["df_hasil"] = None
if "df_beban_clean" not in st.session_state:
  st.session_state["df_beban_clean"] = None
if "df_config_clean" not in st.session_state:
  st.session_state["df_config_clean"] = None

st.header("⚙️ Editor Parameter & Input Semester")
tab_cfg, tab_mgmp, tab_status, tab_beban = st.tabs([
    "1. Format Jam & Hari KBM",
    "2. Hari MGMP Mapel",
    "3. Status & Jabatan Guru",
    "4. Beban Mengajar (Guru_Kelas)",
])

with tab_cfg:
  st.subheader("Pengaturan Jumlah Jam KBM per Hari")
  df_config = st.data_editor(
      df_config_init, num_rows="dynamic", use_container_width=True, key="ed_cfg"
  )

with tab_mgmp:
  st.subheader("Pengaturan Hari Kegiatan MGMP per Mapel")
  df_mgmp_rule = st.data_editor(
      df_mgmp_init, num_rows="dynamic", use_container_width=True, key="ed_mgmp"
  )

with tab_status:
  st.subheader("Daftar Guru, Mata Pelajaran, dan Jabatan Khusus")
  df_status_raw = st.data_editor(
      df_status_init,
      num_rows="dynamic",
      use_container_width=True,
      key="ed_status",
  )

with tab_beban:
  st.subheader("Target Jam Wajib Pasangan Guru & Kelas")
  df_beban_raw = st.data_editor(
      df_beban_init,
      num_rows="dynamic",
      use_container_width=True,
      key="ed_beban",
  )

# Form Manual MGMP Selection jika mode Manual dipilih
if toggle_mgmp and mgmp_mode == "Pilih Manual Guru MGMP":
  all_teachers_list = sorted(
      list(set(df_status_raw["Guru"].dropna().astype(str).str.strip()))
  )
  manual_mgmp_teachers = st.sidebar.multiselect(
      "Pilih Nama-Nama Guru MGMP:", options=all_teachers_list, default=[]
  )

st.divider()

# Tombol Eksekusi Solver
if st.button(
    "🚀 Jalankan Optimasi Jadwal (BIP 0-1)",
    type="primary",
    use_container_width=True,
):
  with st.spinner("Sedang memproses solver PuLP (Branch-and-Bound)..."):
    INVALID_TEXTS = ["NAN", "NONE", "", "NULL", "GURU", "KELAS", "UNNAMED: 0"]

    col_guru = (
        "Guru" if "Guru" in df_beban_raw.columns else df_beban_raw.columns[0]
    )
    col_kelas = (
        "Kelas" if "Kelas" in df_beban_raw.columns else df_beban_raw.columns[1]
    )
    col_jw = (
        "Jam_Wajib"
        if "Jam_Wajib" in df_beban_raw.columns
        else df_beban_raw.columns[-1]
    )

    df_beban_raw[col_guru] = (
        df_beban_raw[col_guru].replace(r"^\s*$", None, regex=True).ffill()
    )
    df_beban_raw[col_guru] = df_beban_raw[col_guru].astype(str).str.strip()
    df_beban_raw[col_kelas] = (
        df_beban_raw[col_kelas].astype(str).str.strip().str.upper()
    )
    df_beban_raw[col_jw] = (
        pd.to_numeric(df_beban_raw[col_jw], errors="coerce")
        .fillna(0)
        .astype(int)
    )

    df_beban = df_beban_raw[
        (~df_beban_raw[col_guru].str.upper().isin(INVALID_TEXTS))
        & (~df_beban_raw[col_kelas].isin(INVALID_TEXTS))
        & (df_beban_raw[col_jw] > 0)
    ].copy()

    col_status_guru = (
        "Guru" if "Guru" in df_status_raw.columns else df_status_raw.columns[0]
    )
    df_status_raw[col_status_guru] = (
        df_status_raw[col_status_guru].astype(str).str.strip()
    )
    df_status = (
        df_status_raw[
            ~df_status_raw[col_status_guru].str.upper().isin(INVALID_TEXTS)
        ]
        .drop_duplicates(subset=[col_status_guru])
        .copy()
    )

    df_pair = df_beban.groupby(
        [col_guru, col_kelas], as_index=False
    )[col_jw].sum()
    df_pair.rename(
        columns={
            col_guru: "Guru",
            col_kelas: "Kelas",
            col_jw: "Jam_Wajib",
        },
        inplace=True,
    )
    df_pair = df_pair[df_pair["Jam_Wajib"] > 0].copy()

    col_hari = "Hari" if "Hari" in df_config.columns else df_config.columns[0]
    col_jam_count = (
        "Jumlah_Jam_KBM"
        if "Jumlah_Jam_KBM" in df_config.columns
        else df_config.columns[1]
    )

    df_config[col_jam_count] = (
        pd.to_numeric(df_config[col_jam_count], errors="coerce")
        .fillna(6)
        .astype(int)
    )
    df_config[col_hari] = df_config[col_hari].astype(str).str.strip()

    HARI = [
        str(h).strip()
        for h in df_config[col_hari]
        if pd.notna(h)
        and str(h).strip()
        and str(h).strip().upper() not in INVALID_TEXTS
    ]
    if not HARI:
      HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat"]

    JAM_HARI = {}
    for _, r in df_config.iterrows():
      h_name = str(r[col_hari]).strip()
      if h_name and h_name.upper() not in INVALID_TEXTS:
        j_count = int(r[col_jam_count])
        JAM_HARI[h_name] = list(range(1, max(1, j_count) + 1))

    for h in HARI:
      if h not in JAM_HARI:
        JAM_HARI[h] = list(range(1, 7))

    col_mgmp_mapel = (
        "Mata_Pelajaran"
        if "Mata_Pelajaran" in df_mgmp_rule.columns
        else df_mgmp_rule.columns[0]
    )
    col_mgmp_hari = (
        "Hari_MGMP"
        if "Hari_MGMP" in df_mgmp_rule.columns
        else df_mgmp_rule.columns[1]
    )

    HARI_MGMP_MAPEL = dict(
        zip(
            df_mgmp_rule[col_mgmp_mapel].astype(str).str.strip().str.upper(),
            df_mgmp_rule[col_mgmp_hari].astype(str).str.strip(),
        )
    )

    col_mapel_beban = (
        "Mata_Pelajaran"
        if "Mata_Pelajaran" in df_beban_raw.columns
        else "Mapel"
    )
    status_dict = (
        df_status.set_index(col_status_guru).to_dict("index")
        if not df_status.empty
        else {}
    )

    pair_mapel = {}
    for _, row in df_beban.iterrows():
      g_name, k_name = row[col_guru], row[col_kelas]
      m_name = ""
      if col_mapel_beban in row and pd.notna(row[col_mapel_beban]):
        m_name = str(row[col_mapel_beban]).strip()
      if not m_name:
        m_name = str(
            status_dict.get(g_name, {}).get("Mata_Pelajaran", "")
        ).strip()
      pair_mapel[(g_name, k_name)] = m_name

    GURU = sorted(list(df_pair["Guru"].unique()))
    KELAS = sorted(list(df_pair["Kelas"].unique()))
    valid_gk_pairs = set(
        (row["Guru"], row["Kelas"])
        for _, row in df_pair.iterrows()
        if row["Jam_Wajib"] > 0
    )
    pair_jw_dict = dict(
        zip(zip(df_pair["Guru"], df_pair["Kelas"]), df_pair["Jam_Wajib"])
    )

    # Evaluasi Penentuan Guru MGMP
    mgmp_teachers = set()
    if toggle_mgmp:
      if mgmp_mode == "Gunakan Status Guru (Is_MGMP)":
        for g in GURU:
          st_info = status_dict.get(g, {})
          val = st_info.get("Is_MGMP")
          if pd.notna(val) and str(val).strip().upper() in [
              "1",
              "1.0",
              "TRUE",
              "Y",
              "YES",
          ]:
            mgmp_teachers.add(g)
      elif mgmp_mode == "Pilih Manual Guru MGMP":
        mgmp_teachers = set(manual_mgmp_teachers)
      elif mgmp_mode == "Otomatis 1 Guru per Mapel MGMP":
        for m_name in HARI_MGMP_MAPEL.keys():
          for (g, k), m_idx in pair_mapel.items():
            if m_idx.upper() == m_name:
              mgmp_teachers.add(g)
              break

    # PuLP Model
    model = pulp.LpProblem("Penjadwalan_SMPN2Plered_Streamlit", pulp.LpMinimize)

    # 1. Variabel Keputusan Biner x(g, k, h, j)
    var_keys = [
        (g, k, h, j)
        for (g, k) in valid_gk_pairs
        for h in HARI
        for j in JAM_HARI.get(h, list(range(1, 7)))
    ]
    x = {
        (g, k, h, j): pulp.LpVariable(
            f"x_{safe_name(g)}_{safe_name(k)}_{safe_name(h)}_{j}",
            cat=pulp.LpBinary,
        )
        for (g, k, h, j) in var_keys
    }

    # Variabel Deviasi
    pair_keys = list(valid_gk_pairs)
    d_pair_minus = {
        (g, k): pulp.LpVariable(
            f"dp_minus_{safe_name(g)}_{safe_name(k)}", lowBound=0
        )
        for (g, k) in pair_keys
    }
    d_pair_plus = {
        (g, k): pulp.LpVariable(
            f"dp_plus_{safe_name(g)}_{safe_name(k)}", lowBound=0
        )
        for (g, k) in pair_keys
    }

    c2_keys = [
        (k, h, j)
        for k in KELAS
        for h in HARI
        for j in JAM_HARI.get(h, list(range(1, 7)))
    ]
    d_c2_minus = {
        (k, h, j): pulp.LpVariable(
            f"dc2_minus_{safe_name(k)}_{safe_name(h)}_{j}", lowBound=0
        )
        for (k, h, j) in c2_keys
    }

    penalti_list = []

    # RULE 3: Maksimal 1 Kelas per Guru per Jam (Toggle 1)
    if toggle_c1:
      for g in GURU:
        assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
        if assigned_k:
          for h in HARI:
            for j in JAM_HARI.get(h, list(range(1, 7))):
              model += pulp.lpSum([x[g, k, h, j] for k in assigned_k]) <= 1

    # RULE 4: Tepat 1 Guru per Kelas per Jam (Tidak Ada Jam Kelas Kosong) (Toggle 2)
    for k in KELAS:
      assigned_g_list = [g_i for (g_i, k_i) in valid_gk_pairs if k_i == k]
      if assigned_g_list:
        for h in HARI:
          for j in JAM_HARI.get(h, list(range(1, 7))):
            if toggle_c2:
              # Bobot sangat berat untuk mencegah jam kosong
              model += (
                  pulp.lpSum([x[g_i, k, h, j] for g_i in assigned_g_list])
                  + d_c2_minus[k, h, j]
                  == 1
              )
            else:
              model += (
                  pulp.lpSum([x[g_i, k, h, j] for g_i in assigned_g_list]) <= 1
              )

    # RULE 5: Target Jam Wajib Guru-Kelas (Toggle 3)
    for _, row in df_pair.iterrows():
      g, k, jw = row["Guru"], row["Kelas"], row["Jam_Wajib"]
      if toggle_c3:
        model += (
            pulp.lpSum([
                x[g, k, h, j]
                for h in HARI
                for j in JAM_HARI.get(h, list(range(1, 7)))
            ])
            + d_pair_minus[g, k]
            - d_pair_plus[g, k]
            == jw
        )

    # RULE 6: Block Teaching 2 Jam (Mod 2 & Berurutan Tanpa Jeda) (Toggle 4)
    if toggle_block2:
      for (g, k), jw in pair_jw_dict.items():
        max_daily = 3 if jw == 5 else (2 if jw in (2, 3, 4) else 1)
        for h in HARI:
          j_list = JAM_HARI.get(h, list(range(1, 7)))
          # Limit harian per pasangan
          model += pulp.lpSum([x[g, k, h, j] for j in j_list]) <= max_daily
          # No-Gap Constraint
          for idx in range(len(j_list) - 2):
            j1, j2, j3 = j_list[idx], j_list[idx + 1], j_list[idx + 2]
            z_gap = pulp.LpVariable(
                f"z_gap_{safe_name(g)}_{safe_name(k)}_{safe_name(h)}_{j1}",
                lowBound=0,
            )
            penalti_list.append(50000.0 * z_gap)
            model += (
                x[g, k, h, j1] + x[g, k, h, j3] - x[g, k, h, j2] - z_gap <= 1
            )

    # RULE 7: Pemerataan Jam ke-1 per Guru (Toggle 5)
    if toggle_jam1_guru:
      target_jam1 = int(round(len(HARI) * len(KELAS) / max(1, len(GURU))))
      for g in GURU:
        d_j1_plus = pulp.LpVariable(f"dj1_plus_{safe_name(g)}", lowBound=0)
        d_j1_minus = pulp.LpVariable(f"dj1_minus_{safe_name(g)}", lowBound=0)
        penalti_list.append(500.0 * (d_j1_plus + d_j1_minus))
        assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
        model += (
            pulp.lpSum([x[g, k, h, 1] for h in HARI for k in assigned_k])
            + d_j1_minus
            - d_j1_plus
            == target_jam1
        )

    # RULE 8: Aturan MGMP (Toggle 6)
    if toggle_mgmp and mgmp_teachers:
      for g in mgmp_teachers:
        st_info = status_dict.get(g, {})
        mapel_raw = str(st_info.get("Mata_Pelajaran", ""))
        mapel_list = [
            m.strip().upper() for m in mapel_raw.split(",") if m.strip()
        ]
        for m_name in mapel_list:
          h_mgmp = HARI_MGMP_MAPEL.get(m_name)
          if h_mgmp and h_mgmp in HARI:
            assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
            for k in assigned_k:
              for j in JAM_HARI.get(h_mgmp, list(range(1, 7))):
                if j > 2:
                  model += x[g, k, h_mgmp, j] == 0

    # RULE 9: Wali Kelas (Jumat Jam ke-1) (Toggle 7)
    if toggle_wali:
      for g in GURU:
        st_info = status_dict.get(g, {})
        is_wali_val = st_info.get("Is_Wali")
        if pd.notna(is_wali_val) and str(is_wali_val).strip().upper() in [
            "1",
            "1.0",
            "TRUE",
            "Y",
            "YES",
        ]:
          if pd.notna(st_info.get("Kelas_Wali")):
            k_wali = str(st_info.get("Kelas_Wali")).strip().upper()
            if (g, k_wali) in valid_gk_pairs and "Jumat" in HARI:
              p_c6 = pulp.LpVariable(f"penalti_wali_{safe_name(g)}", lowBound=0)
              penalti_list.append(1000.0 * p_c6)
              model += x[g, k_wali, "Jumat", 1] + p_c6 >= 1

    # RULE 10: Wakasek Bergantian Jam ke-1 (Toggle 8)
    if toggle_wakasek:
      wakasek_teachers = [
          g
          for g in GURU
          if pd.notna(status_dict.get(g, {}).get("Is_Wakasek"))
          and str(status_dict.get(g, {}).get("Is_Wakasek")).strip().upper()
          in ["1", "1.0", "TRUE", "Y", "YES"]
      ]
      wakasek_pairs = [
          (g, k) for (g, k) in valid_gk_pairs if g in wakasek_teachers
      ]
      if wakasek_pairs:
        for h in HARI:
          model += (
              pulp.lpSum([x[g, k, h, 1] for (g, k) in wakasek_pairs]) <= 1
          )

    # RULE 11: PJOK Jam 1 s.d. 4 (Toggle 9)
    if toggle_pjok:
      for g in GURU:
        is_pjok_val = status_dict.get(g, {}).get("Is_PJOK")
        if pd.notna(is_pjok_val) and str(is_pjok_val).strip().upper() in [
            "1",
            "1.0",
            "TRUE",
            "Y",
            "YES",
        ]:
          assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
          for h in HARI:
            for k in assigned_k:
              for j in JAM_HARI.get(h, list(range(1, 7))):
                if j > 4:
                  model += x[g, k, h, j] == 0

    # RULE 12: Pemerataan Jam Mengajar Harian Guru (Toggle 10)
    if toggle_beban_harian:
      for g in GURU:
        assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
        total_jw_g = sum(pair_jw_dict.get((g, k), 0) for k in assigned_k)
        target_daily = total_jw_g / len(HARI)
        for h in HARI:
          d_dh_plus = pulp.LpVariable(
              f"ddh_plus_{safe_name(g)}_{safe_name(h)}", lowBound=0
          )
          d_dh_minus = pulp.LpVariable(
              f"ddh_minus_{safe_name(g)}_{safe_name(h)}", lowBound=0
          )
          penalti_list.append(10.0 * (d_dh_plus + d_dh_minus))
          model += (
              pulp.lpSum([
                  x[g, k, h, j]
                  for k in assigned_k
                  for j in JAM_HARI.get(h, list(range(1, 7)))
              ])
              + d_dh_minus
              - d_dh_plus
              == target_daily
          )

    # Objective Function
    model += (
        100000.0
        * pulp.lpSum([
            d_pair_minus[g, k] + d_pair_plus[g, k] for (g, k) in pair_keys
        ])
        + 10000.0
        * pulp.lpSum([d_c2_minus[key] for key in c2_keys if toggle_c2])
        + pulp.lpSum(penalti_list)
    )

    solver = pulp.PULP_CBC_CMD(msg=False, timeLimit=60, gapRel=0.01)
    model.solve(solver)

    records = []
    for (g, k, h, j), var in x.items():
      if var.varValue and var.varValue > 0.5:
        mapel_name = pair_mapel.get((g, k), "")
        records.append({
            "Hari": h,
            "Jam_Ke": j,
            "Kelas": k,
            "Mata_Pelajaran": mapel_name,
            "Guru": g,
            "Tampilan_Kelas": (
                f"{mapel_name}\n({g})" if mapel_name else str(g)
            ),
            "Tampilan_Guru": f"{k}\n({mapel_name})" if mapel_name else str(k),
        })

    if len(records) > 0:
      df_hasil = pd.DataFrame(records)
      df_hasil.sort_values(by=["Hari", "Jam_Ke", "Kelas"], inplace=True)

      st.session_state["df_hasil"] = df_hasil
      st.session_state["df_beban_clean"] = df_beban
      st.session_state["df_config_clean"] = df_config
      st.session_state["solved"] = True
      st.success("🎉 OPTIMASI SELESAI & BERHASIL!")
    else:
      st.session_state["solved"] = False
      st.error(
          "❌ Solver tidak menemukan solusi yang valid. Periksa kembali batasan"
          " atau jam wajib."
      )

# Display Results from Session State
if st.session_state["solved"] and st.session_state["df_hasil"] is not None:
  df_hasil = st.session_state["df_hasil"]
  df_beban = st.session_state["df_beban_clean"]
  df_config = st.session_state["df_config_clean"]

  GURU = sorted(list(df_hasil["Guru"].unique()))
  KELAS = sorted(list(df_hasil["Kelas"].unique()))
  HARI = [
      str(h).strip()
      for h in df_config["Hari"]
      if pd.notna(h)
      and str(h).strip()
      and str(h).strip().upper() not in ["NAN", "NONE", "", "NULL"]
  ]

  st.success("✅ Status Optimasi: Solusi Optimal Ditemukan & Tersimpan.")

  # EVALUASI KETERPENUHAN JAM GURU & KELAS
  total_target_jam_guru = df_beban["Jam_Wajib"].sum()
  total_realisasi_jam_guru = len(df_hasil)
  pct_guru = (
      (total_realisasi_jam_guru / total_target_jam_guru * 100)
      if total_target_jam_guru > 0
      else 100.0
  )

  total_slot_kbm_sekolah = sum(
      int(r["Jumlah_Jam_KBM"]) for _, r in df_config.iterrows()
  ) * len(KELAS)
  jam_kosong_kelas = total_slot_kbm_sekolah - total_realisasi_jam_guru
  pct_kelas = (
      (total_realisasi_jam_guru / total_slot_kbm_sekolah * 100)
      if total_slot_kbm_sekolah > 0
      else 100.0
  )

  col1, col2, col3, col4 = st.columns(4)
  col1.metric("Total Guru Terjadwal", len(GURU))
  col2.metric("Total Kelas (Rombel)", len(KELAS))
  col3.metric(
      "Keterpenuhan Jam Guru",
      f"{pct_guru:.1f}%",
      f"{total_realisasi_jam_guru}/{total_target_jam_guru} Jam",
  )
  col4.metric(
      "Keterpenuhan Jam Kelas",
      f"{pct_kelas:.1f}%",
      f"{total_realisasi_jam_guru}/{total_slot_kbm_sekolah} Slot Terisi",
  )

  st.info(
      "📌 **Rincian Evaluasi Keterpenuhan Jam:**\n"
      f"- **Keterpenuhan Jam Guru**: **{pct_guru:.1f}%** ({total_realisasi_jam_guru}"
      f" dari target {total_target_jam_guru} jam mengajar guru dialokasikan).\n"
      f"- **Keterpenuhan Jam Kelas**: **{pct_kelas:.1f}%** ({total_realisasi_jam_guru}"
      f" dari total {total_slot_kbm_sekolah} slot jam KBM terisi).\n"
      f"- **Keterangan Jam Kosong**: Terdapat **{jam_kosong_kelas} slot jam"
      " kosong** secara agregat karena total beban mengajar di `Guru_Kelas`"
      f" ({total_target_jam_guru} jam) lebih sedikit dibanding total slot KBM"
      f" sekolah ({total_slot_kbm_sekolah} slot)."
  )

  res_tab1, res_tab2, res_tab3, res_tab4, res_tab5 = st.tabs([
      "📅 Jadwal per Kelas",
      "👨‍🏫 Jadwal per Guru",
      "🖼️ Visualisasi Kartu Kelas",
      "🖼️ Visualisasi Kartu Guru",
      "📋 Master Data & Export",
  ])

  pivot_kelas = df_hasil.pivot(
      index=["Hari", "Jam_Ke"], columns="Kelas", values="Tampilan_Kelas"
  ).fillna("-")
  pivot_guru = df_hasil.pivot(
      index=["Hari", "Jam_Ke"], columns="Guru", values="Tampilan_Guru"
  ).fillna("-")
  df_master = df_hasil[["Hari", "Jam_Ke", "Kelas", "Mata_Pelajaran", "Guru"]]

  with res_tab1:
    st.subheader("Matriks Jadwal KBM per Kelas (Mata Pelajaran & Guru)")
    st.dataframe(pivot_kelas, use_container_width=True)

  with res_tab2:
    st.subheader("Matriks Jadwal KBM per Guru (Kelas & Mata Pelajaran)")
    st.dataframe(pivot_guru, use_container_width=True)

  with res_tab3:
    st.subheader("🖼️ Visualisasi Kartu Jadwal per Kelas (Bisa Diunduh)")
    selected_kelas = st.selectbox(
        "Pilih Kelas/Rombel untuk Dilihat & Diunduh:", KELAS, key="sb_kelas"
    )

    df_k = df_hasil[df_hasil["Kelas"] == selected_kelas].copy()
    max_jam = max(int(r["Jumlah_Jam_KBM"]) for _, r in df_config.iterrows())

    grid_k = np.full((max_jam, len(HARI)), "-", dtype=object)
    for _, row in df_k.iterrows():
      h_str = str(row["Hari"]).strip()
      h_idx = HARI.index(h_str) if h_str in HARI else -1
      j_idx = int(row["Jam_Ke"]) - 1
      if h_idx >= 0 and 0 <= j_idx < max_jam:
        lbl = (
            f"{row['Mata_Pelajaran']}\n[{row['Guru']}]"
            if row["Mata_Pelajaran"]
            else str(row["Guru"])
        )
        grid_k[j_idx, h_idx] = lbl

    fig_k, ax_k = plt.subplots(figsize=(10, 6))
    ax_k.axis("off")

    table_k = ax_k.table(
        cellText=grid_k,
        colLabels=HARI,
        rowLabels=[f"Jam {i+1}" for i in range(max_jam)],
        loc="center",
        cellLoc="center",
    )
    table_k.auto_set_font_size(False)
    table_k.set_fontsize(9)
    table_k.scale(1.2, 2.2)

    for (r, c), cell in table_k.get_celld().items():
      if r == 0 or c == -1:
        cell.set_facecolor("#1a365d")
        cell.get_text().set_color("white")
        cell.get_text().set_weight("bold")
      else:
        txt = cell.get_text().get_text()
        if txt != "-":
          cell.set_facecolor("#ebf8ff")
          cell.set_edgecolor("#3182ce")
        else:
          cell.set_facecolor("#f7fafc")
          cell.set_edgecolor("#e2e8f0")

    plt.title(
        f"JADWAL KBM KELAS {selected_kelas} - SMPN 2 PLERED",
        fontsize=14,
        fontweight="bold",
        pad=15,
        color="#1a365d",
    )
    plt.tight_layout()
    st.pyplot(fig_k)

    buf_k = io.BytesIO()
    fig_k.savefig(buf_k, format="png", dpi=250, bbox_inches="tight")
    buf_k.seek(0)
    st.download_button(
        label=f"📥 Download Gambar Jadwal Kelas {selected_kelas} (PNG)",
        data=buf_k,
        file_name=f"Jadwal_Kelas_{selected_kelas}.png",
        mime="image/png",
        use_container_width=True,
    )
    plt.close(fig_k)

  with res_tab4:
    st.subheader("🖼️ Visualisasi Kartu Jadwal per Guru (Bisa Diunduh)")
    selected_guru = st.selectbox(
        "Pilih Guru untuk Dilihat & Diunduh:", GURU, key="sb_guru"
    )

    df_g = df_hasil[df_hasil["Guru"] == selected_guru].copy()
    max_jam = max(int(r["Jumlah_Jam_KBM"]) for _, r in df_config.iterrows())

    grid_g = np.full((max_jam, len(HARI)), "-", dtype=object)
    for _, row in df_g.iterrows():
      h_str = str(row["Hari"]).strip()
      h_idx = HARI.index(h_str) if h_str in HARI else -1
      j_idx = int(row["Jam_Ke"]) - 1
      if h_idx >= 0 and 0 <= j_idx < max_jam:
        lbl = (
            f"Kelas {row['Kelas']}\n[{row['Mata_Pelajaran']}]"
            if row["Mata_Pelajaran"]
            else f"Kelas {row['Kelas']}"
        )
        grid_g[j_idx, h_idx] = lbl

    fig_g, ax_g = plt.subplots(figsize=(10, 6))
    ax_g.axis("off")

    table_g = ax_g.table(
        cellText=grid_g,
        colLabels=HARI,
        rowLabels=[f"Jam {i+1}" for i in range(max_jam)],
        loc="center",
        cellLoc="center",
    )
    table_g.auto_set_font_size(False)
    table_g.set_fontsize(9)
    table_g.scale(1.2, 2.2)

    for (r, c), cell in table_g.get_celld().items():
      if r == 0 or c == -1:
        cell.set_facecolor("#2b6cb0")
        cell.get_text().set_color("white")
        cell.get_text().set_weight("bold")
      else:
        txt = cell.get_text().get_text()
        if txt != "-":
          cell.set_facecolor("#feebc8")
          cell.set_edgecolor("#dd6b20")
        else:
          cell.set_facecolor("#f7fafc")
          cell.set_edgecolor("#e2e8f0")

    plt.title(
        f"JADWAL MENGAJAR GURU: {selected_guru} - SMPN 2 PLERED",
        fontsize=14,
        fontweight="bold",
        pad=15,
        color="#2b6cb0",
    )
    plt.tight_layout()
    st.pyplot(fig_g)

    buf_g = io.BytesIO()
    fig_g.savefig(buf_g, format="png", dpi=250, bbox_inches="tight")
    buf_g.seek(0)
    st.download_button(
        label=f"📥 Download Gambar Jadwal Guru {selected_guru} (PNG)",
        data=buf_g,
        file_name=f"Jadwal_Guru_{safe_name(selected_guru)}.png",
        mime="image/png",
        use_container_width=True,
    )
    plt.close(fig_g)

  with res_tab5:
    st.subheader("📋 Master Data Penugasan, Grafik & Unduh Excel Lengkap")
    st.dataframe(df_master, use_container_width=True)

    st.divider()

    # Grafik Keterpenuhan Jam Guru
    st.markdown("##### 1. Grafik Target vs Realisasi Jam Mengajar Guru")
    df_realisasi = (
        df_hasil.groupby("Guru", as_index=False)["Jam_Ke"]
        .count()
        .rename(columns={"Jam_Ke": "Realisasi_Jam"})
    )
    df_target = (
        df_beban.groupby("Guru", as_index=False)["Jam_Wajib"]
        .sum()
        .rename(columns={"Jam_Wajib": "Target_Jam"})
    )
    df_keterpenuhan = pd.merge(
        df_target, df_realisasi, on="Guru", how="left"
    ).fillna(0)

    fig_g_bar, ax_g_bar = plt.subplots(
        figsize=(10, max(5, len(df_keterpenuhan) * 0.4))
    )
    y_pos = np.arange(len(df_keterpenuhan))
    bars_g_bar = ax_g_bar.barh(
        y_pos,
        df_keterpenuhan["Realisasi_Jam"],
        color="#2b5c8f",
        height=0.55,
        label="Realisasi",
    )
    ax_g_bar.scatter(
        df_keterpenuhan["Target_Jam"],
        y_pos,
        color="#d9534f",
        s=70,
        zorder=5,
        label="Target Wajib",
    )

    for bar_i, target_i in zip(bars_g_bar, df_keterpenuhan["Target_Jam"]):
      w = bar_i.get_width()
      pct = int((w / target_i * 100)) if target_i > 0 else 100
      ax_g_bar.text(
          w + 0.3,
          bar_i.get_y() + bar_i.get_height() / 2,
          f"{int(w)}/{int(target_i)} Jam ({pct}%)",
          ha="left",
          va="center",
          fontsize=9,
          fontweight="bold",
          color="#1a365d",
      )

    ax_g_bar.set_yticks(y_pos)
    ax_g_bar.set_yticklabels(
        df_keterpenuhan["Guru"], fontsize=9.5, fontweight="bold"
    )
    ax_g_bar.invert_yaxis()
    ax_g_bar.set_xlabel(
        "Jumlah Jam Mengajar per Minggu", fontsize=10, fontweight="bold"
    )
    ax_g_bar.set_title(
        "Keterpenuhan Beban Mengajar Guru (Target Wajib vs Realisasi)",
        fontsize=12,
        fontweight="bold",
        pad=12,
    )
    ax_g_bar.legend(loc="lower right", frameon=True)
    plt.tight_layout()
    st.pyplot(fig_g_bar)
    plt.close(fig_g_bar)

    st.divider()
    st.subheader("📥 Export Hasil Jadwal Excel")

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_master.to_excel(writer, sheet_name="Data_Penugasan", index=False)
      pivot_kelas.to_excel(writer, sheet_name="Jadwal_per_Kelas")
      pivot_guru.to_excel(writer, sheet_name="Jadwal_per_Guru")
    output.seek(0)

    st.download_button(
        label="📥 Download Seluruh Paket Jadwal Excel (.xlsx)",
        data=output,
        file_name="Jadwal_Guru_SMPN2Plered_Hasil.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
