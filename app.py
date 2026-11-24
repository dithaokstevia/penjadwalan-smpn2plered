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


def clean_str(s):
  if s is None or pd.isna(s):
    return "X"
  if isinstance(s, (tuple, list)):
    s = "_".join(str(x) for x in s if pd.notna(x))
  res = "".join(c if c.isalnum() else "_" for c in str(s)).strip("_")
  return res if res else "X"


def flatten_columns(df):
  if isinstance(df.columns, pd.MultiIndex):
    df.columns = [
        "_"
        .join(str(e) for e in col if pd.notna(e) and str(e).strip() != "")
        .strip()
        for col in df.columns
    ]
  else:
    df.columns = [
        str(col).strip()
        if not isinstance(col, tuple)
        else "_".join(str(e) for e in col)
        for col in df.columns
    ]
  return df


def find_col(df, candidate_names, default_name=""):
  cols = [str(c).strip() for c in df.columns]
  for name in candidate_names:
    for idx, c in enumerate(cols):
      if c.lower() == name.lower():
        return df.columns[idx]
  for name in candidate_names:
    for idx, c in enumerate(cols):
      if name.lower() in c.lower() or c.lower() in name.lower():
        return df.columns[idx]
  if len(df.columns) > 0:
    return df.columns
  return str(default_name)


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
    "4. Block Teaching Modulo 2 (Berurutan + Sisa 1 Jam Fleksibel)", value=True
)
toggle_jam1_guru = st.sidebar.toggle(
    "5. Pemerataan Jam ke-1 per Guru", value=True
)

toggle_mgmp = st.sidebar.toggle(
    "6. Aturan Penanggung Jawab / Pengurus MGMP (Maksimal Jam ke-2)", value=True
)
mgmp_mode = "Gunakan Pengurus MGMP (Is_MGMP_P) + Bagi Sisa Mapel ke Pengurus"
manual_mgmp_teachers = []
if toggle_mgmp:
  mgmp_mode = st.sidebar.radio(
      "Mode Pemilihan Guru MGMP:",
      [
          "Gunakan Pengurus MGMP (Is_MGMP_P) + Bagi Sisa Mapel ke Pengurus",
          "Gunakan Status Penanggung Jawab (Is_MGMP)",
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

# Data Inisialisasi Default
df_config_init = pd.DataFrame({
    "Hari": ["Senin", "Selasa", "Rabu", "Kamis", "Jumat"],
    "Jumlah_Jam_KBM": [7, 6, 6, 6, 6],
})
df_mgmp_init = pd.DataFrame({
    "Mata_Pelajaran": ["IPA", "MATEMATIKA"],
    "Hari_MGMP": ["Kamis", "Rabu"],
})
df_status_init = pd.DataFrame({
    "Guru": ["Guru A", "Guru B", "Guru C"],
    "Mata_Pelajaran": ["IPA", "MATEMATIKA", "BAHASA INDONESIA"],
    "Is_Wali": ["0", "0", "0"],
    "Kelas_Wali": ["", "", ""],
    "Is_Wakasek": ["0", "0", "0"],
    "Is_PJOK": ["0", "0", "0"],
    "Is_MGMP": ["1", "0", "0"],
    "Is_MGMP_P": ["1", "1", "1"],
})
df_beban_init = pd.DataFrame({
    "Guru": ["Guru A", "Guru B", "Guru C"],
    "Kelas": ["7A", "7B", "7C"],
    "Mata_Pelajaran": ["IPA", "MATEMATIKA", "BAHASA INDONESIA"],
    "Jam_Wajib": [4, 4, 4],
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

# Session State Initializer
if "solved" not in st.session_state:
  st.session_state["solved"] = False
if "df_hasil" not in st.session_state:
  st.session_state["df_hasil"] = None
if "df_beban_clean" not in st.session_state:
  st.session_state["df_beban_clean"] = None
if "df_config_clean" not in st.session_state:
  st.session_state["df_config_clean"] = None
if "mgmp_selected_info" not in st.session_state:
  st.session_state["mgmp_selected_info"] = []

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

    df_beban_raw = flatten_columns(df_beban_raw)
    df_config = flatten_columns(df_config)
    df_status_raw = flatten_columns(df_status_raw)
    df_mgmp_rule = flatten_columns(df_mgmp_rule)

    col_guru = find_col(
        df_beban_raw, ["Guru", "Nama_Guru", "Nama Guru"], "Guru"
    )
    col_kelas = find_col(
        df_beban_raw, ["Kelas", "Nama_Kelas", "Rombel"], "Kelas"
    )
    col_jw = find_col(
        df_beban_raw,
        ["Jam_Wajib", "Jam Wajib", "Target_Jam", "Jam"],
        "Jam_Wajib",
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

    col_status_guru = find_col(
        df_status_raw, ["Guru", "Nama_Guru", "Nama Guru"], "Guru"
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
        columns={col_guru: "Guru", col_kelas: "Kelas", col_jw: "Jam_Wajib"},
        inplace=True,
    )
    df_pair = df_pair[df_pair["Jam_Wajib"] > 0].copy()

    col_hari = find_col(df_config, ["Hari", "Nama_Hari", "Day"], "Hari")
    col_jam_count = find_col(
        df_config,
        ["Jumlah_Jam_KBM", "Jumlah_Jam", "Jam_KBM", "Jam"],
        "Jumlah_Jam_KBM",
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

    col_mgmp_mapel = find_col(
        df_mgmp_rule,
        ["Mata_Pelajaran", "Mapel", "Mata Pelajaran"],
        "Mata_Pelajaran",
    )
    col_mgmp_hari = find_col(
        df_mgmp_rule, ["Hari_MGMP", "Hari MGMP", "Hari"], "Hari_MGMP"
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
      g_name, k_name = str(row[col_guru]), str(row[col_kelas])
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
        (str(row["Guru"]), str(row["Kelas"]))
        for _, row in df_pair.iterrows()
        if row["Jam_Wajib"] > 0
    )
    pair_jw_dict = dict(
        zip(zip(df_pair["Guru"], df_pair["Kelas"]), df_pair["Jam_Wajib"])
    )

    # LOGIKA PEMILIHAN ATURAN GURU MGMP & IS_MGMP_P
    mgmp_teachers = set()
    mgmp_info_list = []
    mgmp_teacher_mapel_restrictions = []

    if toggle_mgmp:
      if (
          mgmp_mode
          == "Gunakan Pengurus MGMP (Is_MGMP_P) + Bagi Sisa Mapel ke Pengurus"
      ):
        mgmp_p_teachers = []
        for g in GURU:
          st_info = status_dict.get(g, {})
          val_p = st_info.get("Is_MGMP_P")
          if pd.notna(val_p) and str(val_p).strip().upper() in [
              "1",
              "1.0",
              "TRUE",
              "Y",
              "YES",
          ]:
            mgmp_p_teachers.append(g)

        if not mgmp_p_teachers:
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
              mgmp_p_teachers.append(g)

        if mgmp_p_teachers:
          teacher_load = {g: 0 for g in mgmp_p_teachers}
          unassigned_mapels = []

          for m_name, h_mgmp in HARI_MGMP_MAPEL.items():
            m_upper = m_name.upper()
            assigned_g = None
            for g in mgmp_p_teachers:
              st_info = status_dict.get(g, {})
              g_mapels = [
                  x.strip().upper()
                  for x in str(st_info.get("Mata_Pelajaran", ""))
                  .upper()
                  .split(",")
                  if x.strip()
              ]
              if m_upper in g_mapels:
                assigned_g = g
                break

            if assigned_g:
              teacher_load[assigned_g] += 1
              mgmp_teachers.add(assigned_g)
              mgmp_teacher_mapel_restrictions.append(
                  (assigned_g, m_upper, h_mgmp)
              )
              mgmp_info_list.append({
                  "Mata_Pelajaran": m_upper,
                  "Hari_MGMP": h_mgmp,
                  "Guru_Yang_Datang_MGMP": assigned_g,
                  "Kategori": "Pengurus & Pengampu Mapel (Is_MGMP_P)",
              })
            else:
              unassigned_mapels.append((m_upper, h_mgmp))

          for m_upper, h_mgmp in unassigned_mapels:
            min_g = min(mgmp_p_teachers, key=lambda g: teacher_load[g])
            teacher_load[min_g] += 1
            mgmp_teachers.add(min_g)
            mgmp_teacher_mapel_restrictions.append((min_g, m_upper, h_mgmp))
            mgmp_info_list.append({
                "Mata_Pelajaran": m_upper,
                "Hari_MGMP": h_mgmp,
                "Guru_Yang_Datang_MGMP": min_g,
                "Kategori": "Pengurus (Penugasan Mapel Tambahan)",
            })

    st.session_state["mgmp_selected_info"] = mgmp_info_list

    # Model PuLP
    model = pulp.LpProblem("Penjadwalan_SMPN2Plered_Streamlit", pulp.LpMinimize)

    # Variabel Keputusan Biner x(g, k, h, j)
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

    # RULE 3: Maksimal 1 Kelas per Guru per Jam
    if toggle_c1:
      for g in GURU:
        assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
        if assigned_k:
          for h in HARI:
            for j in JAM_HARI.get(h, list(range(1, 7))):
              model += (
                  pulp.lpSum([
                      x[g, k, h, j] for k in assigned_k if (g, k, h, j) in x
                  ])
                  <= 1
              )

    # RULE 4: Tepat 1 Guru per Kelas per Jam (Pencegahan Jam Kosong)
    for k in KELAS:
      assigned_g_list = [g_i for (g_i, k_i) in valid_gk_pairs if k_i == k]
      if assigned_g_list:
        for h in HARI:
          for j in JAM_HARI.get(h, list(range(1, 7))):
            if toggle_c2:
              model += (
                  pulp.lpSum([
                      x[g_i, k, h, j]
                      for g_i in assigned_g_list
                      if (g_i, k, h, j) in x
                  ])
                  + d_c2_minus[k, h, j]
                  == 1
              )
            else:
              model += (
                  pulp.lpSum([
                      x[g_i, k, h, j]
                      for g_i in assigned_g_list
                      if (g_i, k, h, j) in x
                  ])
                  <= 1
              )

    # RULE 5: Target Jam Wajib Guru-Kelas
    for _, row in df_pair.iterrows():
      g, k, jw = str(row["Guru"]), str(row["Kelas"]), int(row["Jam_Wajib"])
      if toggle_c3 and (g, k) in d_pair_minus:
        model += (
            pulp.lpSum([
                x[g, k, h, j]
                for h in HARI
                for j in JAM_HARI.get(h, list(range(1, 7)))
                if (g, k, h, j) in x
            ])
            + d_pair_minus[g, k]
            - d_pair_plus[g, k]
            == jw
        )

    # RULE 6: Block Teaching (Modulo 2 + Fleksibilitas Sisa 1 Jam untuk Slot Ganjil)
    v_block2 = {}
    v_single = {}
    for (g, k), jw in pair_jw_dict.items():
      if toggle_block2:
        for h in HARI:
          j_list = JAM_HARI.get(h, list(range(1, 7)))
          for idx in range(len(j_list) - 1):
            j_start = j_list[idx]
            v_block2[(g, k, h, j_start)] = pulp.LpVariable(
                f"vb2_{safe_name(g)}_{safe_name(k)}_{safe_name(h)}_{j_start}",
                cat=pulp.LpBinary,
            )
          for j in j_list:
            v_single[(g, k, h, j)] = pulp.LpVariable(
                f"vsig_{safe_name(g)}_{safe_name(k)}_{safe_name(h)}_{j}",
                cat=pulp.LpBinary,
            )

    if toggle_block2:
      for (g, k), jw in pair_jw_dict.items():
        for h in HARI:
          j_list = JAM_HARI.get(h, list(range(1, 7)))

          for j in j_list:
            active_elements = []
            if (g, k, h, j) in v_block2:
              active_elements.append(v_block2[(g, k, h, j)])
            if j - 1 in j_list and (g, k, h, j - 1) in v_block2:
              active_elements.append(v_block2[(g, k, h, j - 1)])
            if (g, k, h, j) in v_single:
              active_elements.append(v_single[(g, k, h, j)])
              penalti_list.append(5.0 * v_single[(g, k, h, j)])

            if active_elements and (g, k, h, j) in x:
              model += x[g, k, h, j] == pulp.lpSum(active_elements)
            elif (g, k, h, j) in x:
              model += x[g, k, h, j] == 0

    # RULE 7: Pemerataan Jam ke-1 per Guru
    if toggle_jam1_guru:
      target_jam1 = int(round(len(HARI) * len(KELAS) / max(1, len(GURU))))
      for g in GURU:
        d_j1_plus = pulp.LpVariable(f"dj1_plus_{safe_name(g)}", lowBound=0)
        d_j1_minus = pulp.LpVariable(f"dj1_minus_{safe_name(g)}", lowBound=0)
        penalti_list.append(500.0 * (d_j1_plus + d_j1_minus))
        assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
        model += (
            pulp.lpSum([
                x[g, k, h, 1]
                for h in HARI
                for k in assigned_k
                if (g, k, h, 1) in x
            ])
            + d_j1_minus
            - d_j1_plus
            == target_jam1
        )

    # RULE 8: Aturan MGMP (Batas Jam ke-2 pada Hari MGMP)
    if toggle_mgmp and mgmp_teacher_mapel_restrictions:
      for g, m_name, h_mgmp in mgmp_teacher_mapel_restrictions:
        if h_mgmp in HARI:
          assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
          for k in assigned_k:
            for j in JAM_HARI.get(h_mgmp, list(range(1, 7))):
              if j > 2 and (g, k, h_mgmp, j) in x:
                model += x[g, k, h_mgmp, j] == 0

    # RULE 9: Wali Kelas (Jumat Jam ke-1)
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
            if (
                (g, k_wali) in valid_gk_pairs
                and "Jumat" in HARI
                and (g, k_wali, "Jumat", 1) in x
            ):
              p_c6 = pulp.LpVariable(
                  f"penalti_wali_{safe_name(g)}", lowBound=0
              )
              penalti_list.append(1000.0 * p_c6)
              model += x[g, k_wali, "Jumat", 1] + p_c6 >= 1

    # RULE 10: Wakasek Bergantian Jam ke-1
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
              pulp.lpSum([
                  x[g, k, h, 1]
                  for (g, k) in wakasek_pairs
                  if (g, k, h, 1) in x
              ])
              <= 1
          )

    # RULE 11: PJOK Jam 1 s.d. 4
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
                if j > 4 and (g, k, h, j) in x:
                  model += x[g, k, h, j] == 0

    # HIERARKI PENALTI UTAMA:
    # 1. Keterpenuhan Jam Wajib Guru (1.000.000)
    # 2. PENCEGAHAN KELAS KOSONG (500.000) -> Sangat tinggi agar slot terisi 100%!
    model += (
        1000000.0
        * pulp.lpSum([
            d_pair_minus[g, k] + d_pair_plus[g, k] for (g, k) in pair_keys
        ])
        + 500000.0 * pulp.lpSum([d_c2_minus[key] for key in c2_keys if toggle_c2])
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
          "❌ Solver tidak menemukan solusi yang valid. Periksa kembali"
          " batasan atau jam wajib."
      )

# ==========================================
# DISPLAY HASIL & ANALISIS EVALUASI LENGKAP
# ==========================================
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

  total_target_jam_guru = int(df_beban["Jam_Wajib"].sum())
  total_realisasi_jam_guru = int(len(df_hasil))
  pct_guru = (
      (total_realisasi_jam_guru / total_target_jam_guru * 100)
      if total_target_jam_guru > 0
      else 100.0
  )

  total_slot_kbm_sekolah = int(
      sum(int(r["Jumlah_Jam_KBM"]) for _, r in df_config.iterrows())
      * len(KELAS)
  )
  pct_kelas = (
      (total_realisasi_jam_guru / total_slot_kbm_sekolah * 100)
      if total_slot_kbm_sekolah > 0
      else 100.0
  )

  st.subheader("📊 Ringkasan Evaluasi Optimasi Jadwal")
  col1, col2, col3, col4 = st.columns(4)
  col1.metric("Total Guru Terjadwal", len(GURU))
  col2.metric("Total Kelas (Rombel)", len(KELAS))
  col3.metric(
      "Keterpenuhan Jam Wajib Guru",
      f"{pct_guru:.1f}%",
      f"{total_realisasi_jam_guru}/{total_target_jam_guru} Jam Terpenuhi",
  )
  col4.metric(
      "Keterisian Jam KBM Kelas",
      f"{pct_kelas:.1f}%",
      f"{total_realisasi_jam_guru}/{total_slot_kbm_sekolah} Slot Terisi",
  )

  if st.session_state["mgmp_selected_info"]:
    with st.expander(
        "👥 Penugasan Utusan Guru yang Harus Datang MGMP per Mata Pelajaran",
        expanded=True,
    ):
      df_info_mgmp = pd.DataFrame(st.session_state["mgmp_selected_info"])
      st.dataframe(df_info_mgmp, use_container_width=True)

  all_scheduled_slots = set(
      zip(df_hasil["Kelas"], df_hasil["Hari"], df_hasil["Jam_Ke"])
  )
  empty_slots = []
  for k in KELAS:
    for _, r in df_config.iterrows():
      h = str(r["Hari"]).strip()
      j_count = int(r["Jumlah_Jam_KBM"])
      for j in range(1, j_count + 1):
        if (k, h, j) not in all_scheduled_slots:
          empty_slots.append({"Kelas": k, "Hari": h, "Jam_Ke": j})
  df_empty_slots = pd.DataFrame(empty_slots)

  df_real_g = (
      df_hasil.groupby(["Guru", "Kelas"], as_index=False)["Jam_Ke"]
      .count()
      .rename(columns={"Jam_Ke": "Realisasi_Jam"})
  )
  df_unfulfilled_g = pd.merge(
      df_beban, df_real_g, on=["Guru", "Kelas"], how="left"
  ).fillna(0)
  df_unfulfilled_g["Kurang_Jam"] = (
      df_unfulfilled_g["Jam_Wajib"] - df_unfulfilled_g["Realisasi_Jam"]
  )
  df_unfulfilled_g = df_unfulfilled_g[df_unfulfilled_g["Kurang_Jam"] > 0]

  st.markdown("---")
  c_det1, c_det2 = st.columns(2)
  with c_det1:
    if not df_empty_slots.empty:
      st.warning(
          f"⚠️ **Terdapat {len(df_empty_slots)} Slot Jam Kosong pada Kelas**"
      )
      with st.expander("🔍 Lihat Rincian Jam Kosong Kelas per Hari"):
        st.dataframe(df_empty_slots, use_container_width=True)
    else:
      st.success(
          "✅ **100% Seluruh Slot Jam Kelas Terisi Penuh (Tidak Ada Jam"
          " Kosong)**"
      )

  with c_det2:
    if not df_unfulfilled_g.empty:
      st.error(
          f"❌ **Terdapat {len(df_unfulfilled_g)} Pasangan Guru-Kelas Belum"
          " Terpenuhi Jam Wajibnya**"
      )
      with st.expander("🔍 Lihat Rincian Guru yang Kurang Jam"):
        st.dataframe(
            df_unfulfilled_g[[
                "Guru",
                "Kelas",
                "Jam_Wajib",
                "Realisasi_Jam",
                "Kurang_Jam",
            ]],
            use_container_width=True,
        )
    else:
      st.success(
          "✅ **100% Jam Wajib Seluruh Guru Terpenuhi Sesuai Target**"
      )

  res_tab1, res_tab2, res_tab3, res_tab4 = st.tabs([
      "📅 Jadwal per Kelas",
      "👨‍🏫 Jadwal per Guru",
      "📊 Visualisasi Grafik Keterisian",
      "📋 Master Data & Export Excel",
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
    st.subheader("📈 Grafis & Visualisasi Keterisian Jam")

    st.markdown("##### 1. Grafik Keterpenuhan Jam Wajib per Guru")
    df_real_total_g = (
        df_hasil.groupby("Guru", as_index=False)["Jam_Ke"]
        .count()
        .rename(columns={"Jam_Ke": "Realisasi_Jam"})
    )
    df_target_g = (
        df_beban.groupby("Guru", as_index=False)["Jam_Wajib"]
        .sum()
        .rename(columns={"Jam_Wajib": "Target_Jam"})
    )
    df_chart_g = pd.merge(
        df_target_g, df_real_total_g, on="Guru", how="left"
    ).fillna(0)

    fig_g, ax_g = plt.subplots(figsize=(10, max(5, len(df_chart_g) * 0.4)))
    y_pos = np.arange(len(df_chart_g))
    bars_g = ax_g.barh(
        y_pos,
        df_chart_g["Realisasi_Jam"],
        color="#2b5c8f",
        height=0.55,
        label="Realisasi",
    )
    ax_g.scatter(
        df_chart_g["Target_Jam"],
        y_pos,
        color="#d9534f",
        s=70,
        zorder=5,
        label="Target Wajib",
    )

    for bar_i, target_i in zip(bars_g, df_chart_g["Target_Jam"]):
      w = bar_i.get_width()
      pct_i = int((w / target_i * 100)) if target_i > 0 else 100
      ax_g.text(
          w + 0.3,
          bar_i.get_y() + bar_i.get_height() / 2,
          f"{int(w)}/{int(target_i)} Jam ({pct_i}%)",
          ha="left",
          va="center",
          fontsize=9,
          fontweight="bold",
          color="#1a365d",
      )

    ax_g.set_yticks(y_pos)
    ax_g.set_yticklabels(df_chart_g["Guru"], fontsize=9.5, fontweight="bold")
    ax_g.invert_yaxis()
    ax_g.set_xlabel("Jumlah Jam Mengajar", fontsize=10, fontweight="bold")
    ax_g.set_title(
        "Keterpenuhan Jam Mengajar Guru (Target Wajib vs Realisasi)",
        fontsize=12,
        fontweight="bold",
    )
    ax_g.legend(loc="lower right")
    plt.tight_layout()
    st.pyplot(fig_g)
    plt.close(fig_g)

    st.divider()

    st.markdown("##### 2. Grafik Keterisian Jam KBM per Kelas (Target 31 Jam)")
    df_kelas_jam = (
        df_hasil.groupby("Kelas", as_index=False)["Jam_Ke"]
        .count()
        .rename(columns={"Jam_Ke": "Total_Jam"})
    )
    fig_k, ax_k = plt.subplots(figsize=(11, 5))
    bars_k = ax_k.bar(
        df_kelas_jam["Kelas"],
        df_kelas_jam["Total_Jam"],
        color="#4682b4",
        width=0.55,
    )
    ax_k.axhline(
        31,
        color="#d9534f",
        linestyle="--",
        linewidth=1.5,
        label="Target Full 31 Jam",
    )

    for b_k in bars_k:
      h_k = b_k.get_height()
      ax_k.text(
          b_k.get_x() + b_k.get_width() / 2,
          h_k / 2,
          f"{int(h_k)} Jam",
          ha="center",
          va="center",
          color="white",
          fontsize=9,
          fontweight="bold",
      )

    ax_k.set_xlabel("Rombongan Belajar (Kelas)", fontsize=10, fontweight="bold")
    ax_k.set_ylabel("Total Jam KBM Terisi", fontsize=10, fontweight="bold")
    ax_k.set_title(
        "Distribusi Total Jam KBM Terisi per Kelas",
        fontsize=12,
        fontweight="bold",
    )
    ax_k.legend(loc="upper right")
    plt.tight_layout()
    st.pyplot(fig_k)
    plt.close(fig_k)

  with res_tab4:
    st.subheader("📋 Master Data Penugasan & Download Excel")
    st.dataframe(df_master, use_container_width=True)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_master.to_excel(writer, sheet_name="Data_Penugasan", index=False)
      pivot_kelas.to_excel(writer, sheet_name="Jadwal_per_Kelas")
      pivot_guru.to_excel(writer, sheet_name="Jadwal_per_Guru")
      if not df_empty_slots.empty:
        df_empty_slots.to_excel(
            writer, sheet_name="Jam_Kosong_Kelas", index=False
        )
      if not df_unfulfilled_g.empty:
        df_unfulfilled_g.to_excel(
            writer, sheet_name="Guru_Kurang_Jam", index=False
        )
      if st.session_state["mgmp_selected_info"]:
        pd.DataFrame(st.session_state["mgmp_selected_info"]).to_excel(
            writer, sheet_name="Utusan_MGMP_Mapel", index=False
        )
    output.seek(0)

    st.download_button(
        label="📥 Download Paket Jadwal Excel Lengkap (.xlsx)",
        data=output,
        file_name="Jadwal_Guru_SMPN2Plered_Hasil.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
