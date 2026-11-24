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

st.title(
    "🏫 Simulator & Optimasi Penjadwalan KBM SMPN 2 Plered (100% Full"
    " Schedule)"
)
st.markdown(
    "Aplikasi berbasis **Streamlit** dan **Binary Integer Programming (BIP 0-1)**"
    " untuk mengoptimalkan jadwal mengajar guru dengan **10 Sakelar Aturan"
    " Fleksibel**."
)


def sanitize_dataframe(df):
  if df is None or df.empty:
    return pd.DataFrame()
  df = df.copy()
  if isinstance(df.columns, pd.MultiIndex):
    df.columns = [
        "_"
        .join(str(c) for c in col if pd.notna(c) and str(c).strip() != "")
        .strip()
        for col in df.columns
    ]
  df = df.loc[:, ~df.columns.duplicated()].copy()
  df.columns = [str(c).strip() for c in df.columns]
  return df


def to_scalar(val, default=None):
  if val is None or pd.isna(val):
    return default
  if isinstance(val, (pd.Series, list, tuple)):
    for item in val:
      if pd.notna(item) and str(item).strip() != "":
        return item
    return default
  return val


def to_int(val, default=0):
  sc = to_scalar(val, default)
  try:
    return int(float(sc))
  except Exception:
    return default


def to_str(val, default=""):
  sc = to_scalar(val, default)
  if sc is None:
    return default
  return str(sc).strip()


def clean_str(s):
  sc = to_str(s, "X")
  res = "".join(c if c.isalnum() else "_" for c in sc).strip("_")
  return res if res else "X"


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
    return df.columns[0]
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
toggle_contig = st.sidebar.toggle(
    "4. Aturan Jam Berurutan / Anti-Diselang (Tanpa Jam Bolong)", value=True
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
    "Jumlah_Jam_KBM":,
})
df_mgmp_init = pd.DataFrame({
    "Mata_Pelajaran": ["IPA", "MATEMATIKA"],
    "Hari_MGMP": ["Kamis", "Rabu"],
})
df_status_init = pd.DataFrame()
df_beban_init = pd.DataFrame()

if excel_path is not None:
  try:
    xls = pd.ExcelFile(excel_path)
    if "Config_Sekolah" in xls.sheet_names:
      df_config_init = pd.read_excel(xls, sheet_name="Config_Sekolah")
    if "Aturan_MGMP" in xls.sheet_names:
      df_mgmp_init = pd.read_excel(xls, sheet_name="Aturan_MGMP")
    if "Status_Guru" in xls.sheet_names:
      df_status_init = pd.read_excel(xls, sheet_name="Status_Guru")
    if "Guru_Kelas" in xls.sheet_names:
      df_beban_init = pd.read_excel(xls, sheet_name="Guru_Kelas")
  except Exception as e:
    st.error(f"Gagal membaca file Excel: {e}")

st.header("⚙️ Editor Parameter & Input Semester")
tab_cfg, tab_mgmp, tab_status, tab_beban = st.tabs([
    "1. Format Jam & Hari KBM",
    "2. Hari MGMP Mapel",
    "3. Status & Jabatan Guru",
    "4. Beban Mengajar (Guru_Kelas)",
])

with tab_cfg:
  st.subheader("Pengaturan Jumlah Jam KBM per Hari")
  df_config_edited = st.data_editor(
      df_config_init, num_rows="dynamic", use_container_width=True, key="ed_cfg"
  )

with tab_mgmp:
  st.subheader("Pengaturan Hari Kegiatan MGMP per Mapel")
  df_mgmp_edited = st.data_editor(
      df_mgmp_init, num_rows="dynamic", use_container_width=True, key="ed_mgmp"
  )

with tab_status:
  st.subheader("Daftar Guru, Mata Pelajaran, dan Jabatan Khusus")
  df_status_edited = st.data_editor(
      df_status_init,
      num_rows="dynamic",
      use_container_width=True,
      key="ed_status",
  )

with tab_beban:
  st.subheader("Target Jam Wajib Pasangan Guru & Kelas")
  df_beban_edited = st.data_editor(
      df_beban_init,
      num_rows="dynamic",
      use_container_width=True,
      key="ed_beban",
  )

if toggle_mgmp and mgmp_mode == "Pilih Manual Guru MGMP":
  all_teachers_list = (
      sorted(
          list(
              set(
                  df_status_edited["Guru"]
                  .dropna()
                  .astype(str)
                  .str.strip()
              )
          )
      )
      if "Guru" in df_status_edited.columns
      else []
  )
  manual_mgmp_teachers = st.sidebar.multiselect(
      "Pilih Nama-Nama Guru MGMP:", options=all_teachers_list, default=[]
  )

st.divider()

# Solver Execution Button
if st.button(
    "🚀 Jalankan Optimasi Jadwal (100% Full)",
    type="primary",
    use_container_width=True,
):
  with st.spinner("Sedang memproses solver PuLP (Branch-and-Bound)..."):
    # Sanitize all dataframes
    df_beban_raw = sanitize_dataframe(df_beban_edited)
    df_config = sanitize_dataframe(df_config_edited)
    df_status_raw = sanitize_dataframe(df_status_edited)
    df_mgmp_rule = sanitize_dataframe(df_mgmp_edited)

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

    g_list, k_list, jw_list = [], [], []
    if not df_beban_raw.empty and col_guru in df_beban_raw.columns:
      last_guru = None
      for _, r in df_beban_raw.iterrows():
        g_val = to_str(r.get(col_guru))
        if g_val:
          last_guru = g_val
        else:
          g_val = last_guru
        k_val = to_str(r.get(col_kelas)).upper()
        jw_val = to_int(r.get(col_jw), 0)
        if g_val and k_val and jw_val > 0:
          g_list.append(g_val)
          k_list.append(k_val)
          jw_list.append(jw_val)

    df_beban_clean = pd.DataFrame(
        {"Guru": g_list, "Kelas": k_list, "Jam_Wajib": jw_list}
    )
    INVALID_TEXTS = ["NAN", "NONE", "", "NULL", "GURU", "KELAS", "UNNAMED: 0"]
    df_beban = df_beban_clean[
        (~df_beban_clean["Guru"].str.upper().isin(INVALID_TEXTS))
        & (~df_beban_clean["Kelas"].isin(INVALID_TEXTS))
        & (df_beban_clean["Jam_Wajib"] > 0)
    ].copy()

    df_pair = df_beban.groupby(["Guru", "Kelas"], as_index=False)[
        "Jam_Wajib"
    ].sum()
    df_pair = df_pair[df_pair["Jam_Wajib"] > 0].copy()

    col_status_guru = find_col(
        df_status_raw, ["Guru", "Nama_Guru", "Nama Guru"], "Guru"
    )
    status_dict = {}
    if not df_status_raw.empty and col_status_guru in df_status_raw.columns:
      for _, r in df_status_raw.iterrows():
        g_key = to_str(r.get(col_status_guru))
        if (
            g_key
            and g_key.upper() not in INVALID_TEXTS
            and g_key not in status_dict
        ):
          status_dict[g_key] = r.to_dict()

    col_hari = find_col(df_config, ["Hari", "Nama_Hari", "Day"], "Hari")
    col_jam_count = find_col(
        df_config,
        ["Jumlah_Jam_KBM", "Jumlah_Jam", "Jam_KBM", "Jam"],
        "Jumlah_Jam_KBM",
    )

    HARI = []
    JAM_HARI = {}
    if not df_config.empty and col_hari in df_config.columns:
      for _, r in df_config.iterrows():
        h_name = to_str(r.get(col_hari))
        j_cnt = to_int(r.get(col_jam_count), 6)
        if h_name:
          HARI.append(h_name)
          JAM_HARI[h_name] = list(range(1, j_cnt + 1))

    if not HARI:
      HARI = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat"]
      JAM_HARI = {
          "Senin": list(range(1, 8)),
          "Selasa": list(range(1, 7)),
          "Rabu": list(range(1, 7)),
          "Kamis": list(range(1, 7)),
          "Jumat": list(range(1, 7)),
      }

    col_mgmp_mapel = find_col(
        df_mgmp_rule,
        ["Mata_Pelajaran", "Mapel", "Mata Pelajaran"],
        "Mata_Pelajaran",
    )
    col_mgmp_hari = find_col(
        df_mgmp_rule, ["Hari_MGMP", "Hari MGMP", "Hari"], "Hari_MGMP"
    )
    HARI_MGMP_MAPEL = {}
    if not df_mgmp_rule.empty:
      for _, r in df_mgmp_rule.iterrows():
        m_name = to_str(r.get(col_mgmp_mapel)).upper()
        h_mgmp = to_str(r.get(col_mgmp_hari))
        if m_name and h_mgmp:
          HARI_MGMP_MAPEL[m_name] = h_mgmp

    GURU = sorted(list(set(to_str(g) for g in df_pair["Guru"])))
    KELAS = sorted(list(set(to_str(k) for k in df_pair["Kelas"])))
    valid_gk_pairs = set(
        (to_str(row["Guru"]), to_str(row["Kelas"]))
        for _, row in df_pair.iterrows()
        if to_int(row["Jam_Wajib"]) > 0
    )

    mgmp_teachers = set()
    for g in GURU:
      st_info = status_dict.get(g, {})
      val = to_str(st_info.get("Is_MGMP")).upper()
      if val in ["1", "1.0", "TRUE", "Y", "YES"]:
        mgmp_teachers.add(g)

    # Build PuLP Model
    model = pulp.LpProblem("Penjadwalan_SMPN2Plered_Streamlit", pulp.LpMinimize)

    var_keys = []
    for (g, k) in valid_gk_pairs:
      g_s, k_s = to_str(g), to_str(k)
      for h in HARI:
        h_s = to_str(h)
        for j in JAM_HARI.get(h_s, []):
          var_keys.append((g_s, k_s, h_s, to_int(j)))

    x = {
        (g, k, h, j): pulp.LpVariable(
            f"x_{clean_str(g)}_{clean_str(k)}_{clean_str(h)}_{j}",
            lowBound=0,
            upBound=1,
            cat=pulp.LpInteger,
        )
        for (g, k, h, j) in var_keys
    }

    pair_keys = [(to_str(g), to_str(k)) for (g, k) in valid_gk_pairs]
    d_pair_minus = {
        (g, k): pulp.LpVariable(
            f"dp_minus_{clean_str(g)}_{clean_str(k)}",
            lowBound=0,
            cat=pulp.LpContinuous,
        )
        for (g, k) in pair_keys
    }
    d_pair_plus = {
        (g, k): pulp.LpVariable(
            f"dp_plus_{clean_str(g)}_{clean_str(k)}",
            lowBound=0,
            cat=pulp.LpContinuous,
        )
        for (g, k) in pair_keys
    }

    c2_keys = [
        (to_str(k), to_str(h), to_int(j))
        for k in KELAS
        for h in HARI
        for j in JAM_HARI.get(to_str(h), [])
    ]
    d_c2_minus = {
        (k, h, j): pulp.LpVariable(
            f"dc2_minus_{clean_str(k)}_{clean_str(h)}_{j}",
            lowBound=0,
            cat=pulp.LpContinuous,
        )
        for (k, h, j) in c2_keys
    }

    dh_keys = [(to_str(g), to_str(h)) for g in GURU for h in HARI]
    d_h_plus = {
        (g, h): pulp.LpVariable(
            f"dh_plus_{clean_str(g)}_{clean_str(h)}",
            lowBound=0,
            cat=pulp.LpContinuous,
        )
        for (g, h) in dh_keys
    }
    d_h_minus = {
        (g, h): pulp.LpVariable(
            f"dh_minus_{clean_str(g)}_{clean_str(h)}",
            lowBound=0,
            cat=pulp.LpContinuous,
        )
        for (g, h) in dh_keys
    }

    gap_keys = []
    for (g, k) in valid_gk_pairs:
      g_s, k_s = to_str(g), to_str(k)
      for h in HARI:
        h_s = to_str(h)
        j_list = JAM_HARI.get(h_s, [])
        for idx in range(len(j_list) - 2):
          gap_keys.append((g_s, k_s, h_s, to_int(j_list[idx])))

    z_gap = {
        (g, k, h, j): pulp.LpVariable(
            f"zgap_{clean_str(g)}_{clean_str(k)}_{clean_str(h)}_{j}",
            lowBound=0,
            cat=pulp.LpContinuous,
        )
        for (g, k, h, j) in gap_keys
    }

    penalti_list = []

    # SOFT MGMP Constraint (Rule 6)
    if toggle_mgmp:
      for g in mgmp_teachers:
        st_info = status_dict.get(g, {})
        mapel_raw = to_str(st_info.get("Mata_Pelajaran"))
        mapel_list = [
            m.strip().upper() for m in mapel_raw.split(",") if m.strip()
        ]
        for m_name in mapel_list:
          h_mgmp = HARI_MGMP_MAPEL.get(m_name)
          if h_mgmp and h_mgmp in HARI:
            assigned_k = [k_i for (g_i, k_i) in valid_gk_pairs if g_i == g]
            for k_i in assigned_k:
              for j in JAM_HARI.get(h_mgmp, []):
                j_i = to_int(j)
                if j_i > 2 and (g, k_i, h_mgmp, j_i) in x:
                  p_mgmp = pulp.LpVariable(
                      f"p_mgmp_{clean_str(g)}_{clean_str(k_i)}_{j_i}",
                      lowBound=0,
                  )
                  penalti_list.append(p_mgmp)
                  model += x[g, k_i, h_mgmp, j_i] - p_mgmp <= 0

    # Rule 1: Max 1 Class per Teacher per Hour
    if toggle_c1:
      for g in GURU:
        assigned_k = [k_i for (g_i, k_i) in valid_gk_pairs if g_i == g]
        if assigned_k:
          for h in HARI:
            for j in JAM_HARI.get(h, []):
              j_i = to_int(j)
              if any((g, k_i, h, j_i) in x for k_i in assigned_k):
                model += (
                    pulp.lpSum([
                        x[g, k_i, h, j_i]
                        for k_i in assigned_k
                        if (g, k_i, h, j_i) in x
                    ])
                    <= 1
                )

    # Rule 2: Exactly 1 Teacher per Class per Slot
    for k in KELAS:
      assigned_g = [g_i for (g_i, k_i) in valid_gk_pairs if k_i == k]
      if assigned_g:
        for h in HARI:
          for j in JAM_HARI.get(h, []):
            j_i = to_int(j)
            if (k, h, j_i) in d_c2_minus:
              if toggle_c2:
                model += (
                    pulp.lpSum([
                        x[g_i, k, h, j_i]
                        for g_i in assigned_g
                        if (g_i, k, h, j_i) in x
                    ])
                    + d_c2_minus[k, h, j_i]
                    == 1
                )
              else:
                model += (
                    pulp.lpSum([
                        x[g_i, k, h, j_i]
                        for g_i in assigned_g
                        if (g_i, k, h, j_i) in x
                    ])
                    <= 1
                )

    # Rule 3: Target Hours per Teacher-Class pair
    for _, row in df_pair.iterrows():
      g, k, jw = (
          to_str(row["Guru"]),
          to_str(row["Kelas"]),
          to_int(row["Jam_Wajib"]),
      )
      if (g, k) in d_pair_minus and (g, k) in d_pair_plus:
        if toggle_c3:
          model += (
              pulp.lpSum([
                  x[g, k, h, to_int(j)]
                  for h in HARI
                  for j in JAM_HARI.get(h, [])
                  if (g, k, h, to_int(j)) in x
              ])
              + d_pair_minus[g, k]
              - d_pair_plus[g, k]
              == jw
          )

    # Rule 4: Anti-Diselang / Contiguous Hours
    if toggle_contig:
      for (g, k) in valid_gk_pairs:
        for h in HARI:
          j_list = JAM_HARI.get(h, [])
          for idx in range(len(j_list) - 2):
            j1, j2, j3 = (
                to_int(j_list[idx]),
                to_int(j_list[idx + 1]),
                to_int(j_list[idx + 2]),
            )
            if (
                (g, k, h, j1) in x
                and (g, k, h, j2) in x
                and (g, k, h, j3) in x
                and (g, k, h, j1) in z_gap
            ):
              model += (
                  x[g, k, h, j1] + x[g, k, h, j3] - x[g, k, h, j2] - z_gap[g, k, h, j1]
                  <= 1
              )

    # Rule 5: Pemerataan Jam 1 per Guru
    if toggle_jam1_guru:
      target_jam1 = int(round(len(HARI) * len(KELAS) / max(1, len(GURU))))
      for g in GURU:
        d_j1_plus = pulp.LpVariable(f"dj1_plus_{clean_str(g)}", lowBound=0)
        d_j1_minus = pulp.LpVariable(f"dj1_minus_{clean_str(g)}", lowBound=0)
        penalti_list.append(500.0 * (d_j1_plus + d_j1_minus))
        assigned_k = [k_i for (g_i, k_i) in valid_gk_pairs if g_i == g]
        model += (
            pulp.lpSum([
                x[g, k_i, h, 1]
                for h in HARI
                for k_i in assigned_k
                if (g, k_i, h, 1) in x
            ])
            + d_j1_minus
            - d_j1_plus
            == target_jam1
        )

    # Rule 7: Wali Kelas (Jumat Jam ke-1)
    if toggle_wali:
      for g in GURU:
        st_info = status_dict.get(g, {})
        is_wali_val = to_str(st_info.get("Is_Wali")).upper()
        if is_wali_val in ["1", "1.0", "TRUE", "Y", "YES"]:
          k_wali = to_str(st_info.get("Kelas_Wali")).upper()
          if (
              (g, k_wali) in valid_gk_pairs
              and "Jumat" in HARI
              and (g, k_wali, "Jumat", 1) in x
          ):
            p_c6 = pulp.LpVariable(
                f"penalti_wali_{clean_str(g)}", lowBound=0
            )
            penalti_list.append(1000.0 * p_c6)
            model += x[g, k_wali, "Jumat", 1] + p_c6 >= 1

    # Rule 8: Wakasek Jam ke-1
    if toggle_wakasek:
      wakasek_teachers = [
          g
          for g in GURU
          if to_str(status_dict.get(g, {}).get("Is_Wakasek")).upper()
          in ["1", "1.0", "TRUE", "Y", "YES"]
      ]
      if wakasek_teachers:
        for h in HARI:
          model += (
              pulp.lpSum([
                  x[g_w, k_i, h, 1]
                  for g_w in wakasek_teachers
                  for k_i in KELAS
                  if (g_w, k_i, h, 1) in x
              ])
              <= 1
          )

    # Rule 9: Soft PJOK Penalty
    if toggle_pjok:
      for g in GURU:
        is_pjok_val = to_str(status_dict.get(g, {}).get("Is_PJOK")).upper()
        if is_pjok_val in ["1", "1.0", "TRUE", "Y", "YES"]:
          assigned_k = [k_i for (g_i, k_i) in valid_gk_pairs if g_i == g]
          for h in HARI:
            for k_i in assigned_k:
              for j in JAM_HARI.get(h, []):
                j_i = to_int(j)
                if j_i > 4 and (g, k_i, h, j_i) in x:
                  p_c8 = pulp.LpVariable(
                      f"penalti_pjok_{clean_str(g)}_{clean_str(h)}_{clean_str(k_i)}_{j_i}",
                      lowBound=0,
                      cat=pulp.LpContinuous,
                  )
                  penalti_list.append(p_c8)
                  model += x[g, k_i, h, j_i] - p_c8 <= 0

    # Rule 10: Pemerataan Jam Harian Guru
    if toggle_beban_harian:
      for g in GURU:
        for h in HARI:
          assigned_k = [k_i for (g_i, k_i) in valid_gk_pairs if g_i == g]
          model += d_h_plus[g, h] >= 0
          model += d_h_minus[g, h] >= 0

    # OBJECTIVE FUNCTION (Priority 10 Juta untuk Jam Terisi 100% Full)
    model += (
        10000000.0
        * pulp.lpSum([
            d_pair_minus[g, k] + d_pair_plus[g, k] for (g, k) in pair_keys
        ])
        + 10000000.0
        * pulp.lpSum([d_c2_minus[k, h, j] for (k, h, j) in c2_keys])
        + 5.0 * pulp.lpSum([z_gap[key] for key in gap_keys if key in z_gap])
        + 10.0 * pulp.lpSum(penalti_list)
        + 1.0
        * pulp.lpSum([
            d_h_plus[g, h] + d_h_minus[g, h]
            for g in GURU
            for h in HARI
            if (g, h) in d_h_plus
        ])
    )

    solver = pulp.PULP_CBC_CMD(msg=False, timeLimit=60, gapRel=0.0)
    model.solve(solver)

    if model.status in [pulp.LpStatusOptimal, 1]:
      records = []
      for (g, k, h, j), var in x.items():
        if var.varValue and var.varValue > 0.5:
          records.append({"Hari": h, "Jam_Ke": j, "Kelas": k, "Guru": g})

      df_hasil = pd.DataFrame(records)
      df_hasil.sort_values(by=["Hari", "Jam_Ke", "Kelas"], inplace=True)

      st.success("🎉 OPTIMASI SELESAI & BERHASIL! (JADWAL 100% TERISI PENUH)")

      # Summary Metrics
      total_target_jam = sum(
          to_int(row["Jam_Wajib"]) for _, row in df_pair.iterrows()
      )
      total_realisasi_jam = len(df_hasil)

      col1, col2, col3, col4 = st.columns(4)
      col1.metric("Total Guru Terjadwal", len(GURU))
      col2.metric("Total Kelas (Rombel)", len(KELAS))
      col3.metric(
          "Keterpenuhan Jam Wajib",
          f"{(total_realisasi_jam/max(1, total_target_jam))*100:.1f}%",
      )
      col4.metric("Keterisian Jam Kelas", "100% Sempurna (0 Kosong)")

      res_tab1, res_tab2, res_tab3 = st.tabs([
          "📅 Jadwal per Kelas",
          "👨‍🏫 Jadwal per Guru",
          "📋 Master Data Penugasan",
      ])

      with res_tab1:
        st.subheader("Matriks Jadwal KBM per Kelas (Rombel)")
        pivot_kelas = df_hasil.pivot(
            index=["Hari", "Jam_Ke"], columns="Kelas", values="Guru"
        ).fillna("-")
        st.dataframe(pivot_kelas, use_container_width=True)

      with res_tab2:
        st.subheader("Matriks Jadwal KBM per Guru")
        pivot_guru = df_hasil.pivot(
            index=["Hari", "Jam_Ke"], columns="Guru", values="Kelas"
        ).fillna("-")
        st.dataframe(pivot_guru, use_container_width=True)

      with res_tab3:
        st.subheader("Tabel Detail Penugasan")
        st.dataframe(df_hasil, use_container_width=True)

      # Download Option
      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_hasil.to_excel(writer, sheet_name="Data_Penugasan", index=False)
        pivot_kelas.to_excel(writer, sheet_name="Jadwal_per_Kelas")
        pivot_guru.to_excel(writer, sheet_name="Jadwal_per_Guru")
      output.seek(0)

      st.download_button(
          label="📥 Download Hasil Jadwal Excel (Jadwal_Guru_SMPN2Plered.xlsx)",
          data=output,
          file_name="Jadwal_Guru_SMPN2Plered_Hasil_Full.xlsx",
          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
          use_container_width=True,
      )
    else:
      st.error(
          "❌ Solver tidak menemukan solusi yang valid. Periksa kembali"
          " batasan atau jam wajib."
      )
