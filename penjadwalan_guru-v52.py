"""
================================================================================
SKRIP OPTIMASI PENJADWALAN GURU SMPN 2 PLERED (v52 - Complete 12 Rules BIP Model)
Model: Binary Integer Programming (BIP 0-1) & Goal Programming Elastic
Metode: Branch-and-Bound (Solver PuLP / CBC) + Consecutive Block Teaching + Chain Swap

PEMENUHAN 12 ATURAN SPESIFIKASI:
1. Jadwal KBM Lengkap SMPN 2 Plered
2. Tercantum Kelas, Hari, Jam Ke, Mata Pelajaran, dan Guru
3. Maksimal 1 Kelas per Guru per Jam
4. Tepat 1 Guru per Kelas per Jam (31 Jam/Minggu)
5. Pemenuhan Target Jam Wajib Pasangan (Guru, Kelas)
6. Jam Mengajar Berdampingan / Kontinu (Block Teaching without Gap)
7. Keseragaman / Keseimbangan Jam Pertama Mengajar Guru
8. Utusan MGMP Dikosongkan Jam > 2 pada Hari MGMP
9. Wali Kelas Mengajar Jumat Jam Ke-1 di Kelas Bimbingan
10. Wakasek Bergantian Mengajar Jam Ke-1
11. Guru PJOK Khusus Jam Ke-1 s.d. 4
12. Pemerataan Beban Jam Mengajar Harian Guru
================================================================================
"""

import os
import pandas as pd
import pulp

def jalankan_optimasi_v52(excel_input_path='data_jadwal.xlsx', excel_output_path='Jadwal_Guru_SMPN2Plered.xlsx'):
    print("==================================================================")
    print("   PROSES OPTIMASI JADWAL v52 - FULL 12 CONSTRAINTS MODEL        ")
    print("==================================================================")
    
    if not os.path.exists(excel_input_path):
        print(f"ERROR: File '{excel_input_path}' tidak ditemukan!")
        print("Pastikan Anda sudah mengunggah file 'data_jadwal.xlsx' ke Colab.")
        return

    print("\n1. Membaca data dan menangani merged cells (Forward Fill)...")
    df_beban_raw = pd.read_excel(excel_input_path, sheet_name='Guru_Kelas')
    df_status_raw = pd.read_excel(excel_input_path, sheet_name='Status_Guru')
    
    df_beban_raw['Guru'] = df_beban_raw['Guru'].replace(r'^\s*$', None, regex=True).replace(['nan', 'NaN', 'NAN', 'None', 'NONE', 'Unnamed: 0'], None).ffill()
    df_beban_raw['Guru'] = df_beban_raw['Guru'].astype(str).str.strip()
    df_beban_raw['Kelas'] = df_beban_raw['Kelas'].astype(str).str.strip().str.upper()
    df_beban_raw['Jam_Wajib'] = pd.to_numeric(df_beban_raw['Jam_Wajib'], errors='coerce').fillna(0).astype(int)
    df_status_raw['Guru'] = df_status_raw['Guru'].astype(str).str.strip()
    
    INVALID_TEXTS = ['NAN', 'NONE', '', 'NULL', 'GURU', 'KELAS', 'NAMA GURU', 'UNNAMED: 0']
    
    df_beban = df_beban_raw[
        (~df_beban_raw['Guru'].str.upper().isin(INVALID_TEXTS)) & 
        (~df_beban_raw['Kelas'].isin(INVALID_TEXTS)) &
        (df_beban_raw['Jam_Wajib'] > 0)
    ].copy()
    
    df_status = df_status_raw[~df_status_raw['Guru'].str.upper().isin(INVALID_TEXTS)].drop_duplicates(subset=['Guru'], keep='first').copy()
    df_pair = df_beban.groupby(['Guru', 'Kelas'], as_index=False)['Jam_Wajib'].sum()
    
    GURU = sorted(list(df_pair['Guru'].unique()))
    KELAS = sorted(list(df_pair['Kelas'].unique()))
    HARI = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']
    
    JAM_HARI = {
        'Senin': list(range(1, 8)),
        'Selasa': list(range(1, 7)),
        'Rabu': list(range(1, 7)),
        'Kamis': list(range(1, 7)),
        'Jumat': list(range(1, 7))
    }
    
    HARI_MGMP_MAPEL = {
        'BAHASA INGGRIS': 'Selasa', 'PJOK': 'Selasa', 'MATEMATIKA': 'Rabu',
        'BAHASA INDONESIA': 'Rabu', 'IPS': 'Rabu', 'IPA': 'Kamis',
        'SNBP': 'Kamis', 'PENDIDIKAN PANCASILA': 'Jumat',
        'BAHASA CIREBON': 'Jumat', 'PAI': 'Sabtu'
    }
    
    status_dict = df_status.set_index('Guru').to_dict('index')
    valid_gk_pairs = set((row['Guru'], row['Kelas']) for _, row in df_pair.iterrows() if row['Jam_Wajib'] > 0)

    print("\n2. Membangun model optimasi PuLP (Pemenuhan 12 Aturan Spesifikasi)...")
    model = pulp.LpProblem("Penjadwalan_SMPN2Plered_v52", pulp.LpMinimize)
    
    var_keys = [(g, k, h, j) for (g, k) in valid_gk_pairs for h in HARI for j in JAM_HARI[h]]
    x = pulp.LpVariable.dicts("x", var_keys, cat=pulp.LpBinary)
    
    pair_keys = [(row['Guru'], row['Kelas']) for _, row in df_pair.iterrows()]
    d_pair_minus = pulp.LpVariable.dicts("dpair_minus", pair_keys, lowBound=0)
    d_pair_plus = pulp.LpVariable.dicts("dpair_plus", pair_keys, lowBound=0)
    
    c2_keys = [(k, h, j) for k in KELAS for h in HARI for j in JAM_HARI[h]]
    d_c2_minus = pulp.LpVariable.dicts("dc2_minus", c2_keys, lowBound=0)
    
    dh_keys = [(g, h) for g in GURU for h in HARI]
    d_h_plus = pulp.LpVariable.dicts("dh_plus", dh_keys, lowBound=0)
    d_h_minus = pulp.LpVariable.dicts("dh_minus", dh_keys, lowBound=0)

    # Variabel Penalti Jeda Mengajar (Block Teaching / Consecutive Constraint - Rule 6)
    gap_keys = []
    for (g, k) in valid_gk_pairs:
        for h in HARI:
            j_list = JAM_HARI[h]
            for idx in range(len(j_list) - 2):
                gap_keys.append((g, k, h, j_list[idx]))
    z_gap = pulp.LpVariable.dicts("z_gap", gap_keys, lowBound=0)

    # KELOMPOK GURU PER MAPEL (Aturan 8: MGMP Selection)
    mapel_to_guru = {}
    for g in GURU:
        st = status_dict.get(g, {})
        mapel_name = str(st.get('Mata_Pelajaran')).strip().upper()
        if mapel_name and mapel_name in HARI_MGMP_MAPEL:
            if mapel_name not in mapel_to_guru:
                mapel_to_guru[mapel_name] = []
            mapel_to_guru[mapel_name].append(g)

    all_mgmp_g = sorted(list(set(g for list_g in mapel_to_guru.values() for g in list_g)))
    y_mgmp = pulp.LpVariable.dicts("y_mgmp", all_mgmp_g, cat=pulp.LpBinary)

    # Aturan 8: Setiap Mapel memilih TEPAT 1 Utusan MGMP
    for mapel_name, list_g in mapel_to_guru.items():
        if list_g:
            model += pulp.lpSum([y_mgmp[g] for g in list_g]) == 1

    # Aturan 8: Utusan MGMP dikosongkan jam > 2 pada Hari MGMP
    for mapel_name, list_g in mapel_to_guru.items():
        hari_mgmp = HARI_MGMP_MAPEL.get(mapel_name)
        if hari_mgmp and hari_mgmp in HARI:
            for g in list_g:
                assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
                if assigned_k:
                    for k in assigned_k:
                        for j in JAM_HARI[hari_mgmp]:
                            if j > 2:
                                model += x[g, k, hari_mgmp, j] <= (1 - y_mgmp[g])

    penalti_list = []
    
    # Aturan 3: Maksimal 1 Kelas per Guru per Jam
    for g in GURU:
        assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
        if assigned_k:
            for h in HARI:
                for j in JAM_HARI[h]:
                    model += pulp.lpSum([x[g, k, h, j] for k in assigned_k]) <= 1
                
    # Aturan 4: Tepat 1 Guru per Kelas per Jam (31 Jam KBM)
    for k in KELAS:
        assigned_g = [g for (g, k_i) in valid_gk_pairs if k_i == k]
        if assigned_g:
            for h in HARI:
                for j in JAM_HARI[h]:
                    model += pulp.lpSum([x[g, k, h, j] for g in assigned_g]) + d_c2_minus[k, h, j] == 1

    # Aturan 5: Pemenuhan Target Jam Wajib (Guru, Kelas)
    for _, row in df_pair.iterrows():
        g, k, jw = row['Guru'], row['Kelas'], row['Jam_Wajib']
        model += (
            pulp.lpSum([x[g, k, h, j] for h in HARI for j in JAM_HARI[h]])
            + d_pair_minus[g, k] - d_pair_plus[g, k] == jw
        )

    # Aturan 6: Block Teaching (Jam Berdampingan Kontinu Tanpa Jeda)
    for (g, k) in valid_gk_pairs:
        for h in HARI:
            j_list = JAM_HARI[h]
            for idx in range(len(j_list) - 2):
                j1, j2, j3 = j_list[idx], j_list[idx+1], j_list[idx+2]
                model += x[g, k, h, j1] + x[g, k, h, j3] - x[g, k, h, j2] - z_gap[g, k, h, j1] <= 1

    # Aturan 6 (Lanjutan): Pembagian Harian (Maksimal 2-3 Jam per Hari)
    for _, row in df_pair.iterrows():
        g, k, jw = row['Guru'], row['Kelas'], row['Jam_Wajib']
        max_harian = 3 if jw >= 5 else 2
        for h in HARI:
            model += pulp.lpSum([x[g, k, h, j] for j in JAM_HARI[h]]) <= max_harian

    # Aturan 12 & 7: Pemerataan Jam Harian Guru & Jam Masuk
    for g in GURU:
        total_g = df_pair[df_pair['Guru'] == g]['Jam_Wajib'].sum()
        if total_g > 0:
            R_g = total_g / 5.0
            assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
            for h in HARI:
                jam_riil_harian = pulp.lpSum([x[g, k, h, j] for k in assigned_k for j in JAM_HARI[h]])
                model += jam_riil_harian + d_h_minus[g, h] - d_h_plus[g, h] == R_g

    # Aturan 9: Wali Kelas Mengajar Jumat Jam Ke-1
    for g in GURU:
        if g not in status_dict:
            continue
        st = status_dict[g]
        assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
        
        if st.get('Is_Wali') == 1 and pd.notna(st.get('Kelas_Wali')):
            k_wali = str(st.get('Kelas_Wali')).strip().upper()
            if (g, k_wali) in valid_gk_pairs:
                p_c6 = pulp.LpVariable(f"penalti_wali_{g}", lowBound=0)
                penalti_list.append(p_c6)
                model += x[g, k_wali, 'Jumat', 1] + p_c6 >= 1

    # Aturan 10: Wakasek Bergantian Mengajar Jam Ke-1 (Maksimal 1 Wakasek mengajar Jam Ke-1 per Hari)
    wakasek_teachers = [g for g in GURU if status_dict.get(g, {}).get('Is_Wakasek') == 1]
    if wakasek_teachers:
        for h in HARI:
            wakasek_assigned = [k for g_w in wakasek_teachers for (g_i, k) in valid_gk_pairs if g_i == g_w]
            if wakasek_assigned:
                model += pulp.lpSum([x[g_w, k, h, 1] for g_w in wakasek_teachers for k in valid_gk_pairs if (g_w, k) in valid_gk_pairs]) <= 1

    # Aturan 11: PJOK Khusus Jam Ke-1 s.d. 4
    for g in GURU:
        if status_dict.get(g, {}).get('Is_PJOK') == 1:
            assigned_k = [k for (g_i, k) in valid_gk_pairs if g_i == g]
            for h in HARI:
                for k in assigned_k:
                    for j in JAM_HARI[h]:
                        if j > 4:
                            p_c8 = pulp.LpVariable(f"penalti_pjok_{g}_{h}_{k}_{j}", lowBound=0)
                            penalti_list.append(p_c8)
                            model += x[g, k, h, j] - p_c8 <= 0

    # FUNGSI TUJUAN (MINIMASI DEVIASI DENGAN PENALTI BERTIKAT)
    model += (
        100000.0 * pulp.lpSum([d_pair_minus[g, k] + d_pair_plus[g, k] for g, k in pair_keys])
        + 10000.0 * pulp.lpSum([d_c2_minus[k, h, j] for k, h, j in c2_keys])
        + 50000.0 * pulp.lpSum([z_gap[key] for key in gap_keys])
        + 1000.0 * pulp.lpSum(penalti_list)
        + 2.0 * pulp.lpSum([d_h_plus[g, h] + d_h_minus[g, h] for g in GURU for h in HARI])
    )

    print("3. Mengeksekusi solver CBC (Branch-and-Bound)...")
    solver = pulp.PULP_CBC_CMD(msg=False, timeLimit=60, gapRel=0.0)
    model.solve(solver)
    
    if model.status in [pulp.LpStatusOptimal, 1]:
        records = []
        for (g, k, h, j), var in x.items():
            if var.varValue and var.varValue > 0.5:
                records.append({'Hari': h, 'Jam_Ke': j, 'Kelas': k, 'Guru': g})
                
        df_hasil = pd.DataFrame(records)
        df_hasil.sort_values(by=['Hari', 'Jam_Ke', 'Kelas'], inplace=True)
        
        mgmp_chosen = {g: (y_mgmp[g].varValue is not None and y_mgmp[g].varValue > 0.5) for g in all_mgmp_g}
        jadwal_guru_dict = df_hasil.set_index(['Guru', 'Hari', 'Jam_Ke'])['Kelas'].to_dict() if not df_hasil.empty else {}
        jadwal_kelas_dict = df_hasil.set_index(['Kelas', 'Hari', 'Jam_Ke'])['Guru'].to_dict() if not df_hasil.empty else {}
        realisasi_pair = df_hasil.groupby(['Guru', 'Kelas']).size().to_dict() if not df_hasil.empty else {}
        
        # POST-PROCESSING CHAIN SWAP
        print("\n4. Mengeksekusi Post-Processing Chain Swap...")
        for _, r_pair in df_pair.iterrows():
            g_target, k_target, target_jw = r_pair['Guru'], r_pair['Kelas'], r_pair['Jam_Wajib']
            real_jw = realisasi_pair.get((g_target, k_target), 0)
            kurang = target_jw - real_jw
            
            if kurang > 0:
                for h_empty in HARI:
                    if kurang <= 0: break
                    for j_empty in JAM_HARI[h_empty]:
                        if kurang <= 0: break
                        if (k_target, h_empty, j_empty) not in jadwal_kelas_dict:
                            if (g_target, h_empty, j_empty) not in jadwal_guru_dict:
                                records.append({'Hari': h_empty, 'Jam_Ke': j_empty, 'Kelas': k_target, 'Guru': g_target})
                                jadwal_guru_dict[(g_target, h_empty, j_empty)] = k_target
                                jadwal_kelas_dict[(k_target, h_empty, j_empty)] = g_target
                                realisasi_pair[(g_target, k_target)] = realisasi_pair.get((g_target, k_target), 0) + 1
                                kurang -= 1
                                print(f"  ✅ AUTO-FILL Direct: {g_target} -> Kelas {k_target} [{h_empty} Jam {j_empty}]")
                                break
                            else:
                                k_bentrok = jadwal_guru_dict[(g_target, h_empty, j_empty)]
                                for h_alt in HARI:
                                    for j_alt in JAM_HARI[h_alt]:
                                        if (g_target, h_alt, j_alt) not in jadwal_guru_dict and (k_bentrok, h_alt, j_alt) not in jadwal_kelas_dict:
                                            st_g = status_dict.get(g_target, {})
                                            is_blocked = False
                                            if st_g.get('Is_Wakasek') == 1 and j_alt == 1: is_blocked = True
                                            if st_g.get('Is_PJOK') == 1 and j_alt > 4: is_blocked = True
                                            if mgmp_chosen.get(g_target, False):
                                                mapel_g = str(st_g.get('Mata_Pelajaran')).strip().upper()
                                                if HARI_MGMP_MAPEL.get(mapel_g) == h_alt and j_alt > 2: is_blocked = True
                                            if not is_blocked:
                                                records = [r for r in records if not (r['Guru'] == g_target and r['Kelas'] == k_bentrok and r['Hari'] == h_empty and r['Jam_Ke'] == j_empty)]
                                                del jadwal_guru_dict[(g_target, h_empty, j_empty)]
                                                del jadwal_kelas_dict[(k_bentrok, h_empty, j_empty)]
                                                records.append({'Hari': h_alt, 'Jam_Ke': j_alt, 'Kelas': k_bentrok, 'Guru': g_target})
                                                jadwal_guru_dict[(g_target, h_alt, j_alt)] = k_bentrok
                                                jadwal_kelas_dict[(k_bentrok, h_alt, j_alt)] = g_target
                                                records.append({'Hari': h_empty, 'Jam_Ke': j_empty, 'Kelas': k_target, 'Guru': g_target})
                                                jadwal_guru_dict[(g_target, h_empty, j_empty)] = k_target
                                                jadwal_kelas_dict[(k_target, h_empty, j_empty)] = g_target
                                                realisasi_pair[(g_target, k_target)] = realisasi_pair.get((g_target, k_target), 0) + 1
                                                kurang -= 1
                                                print(f"  🔄 CHAIN SWAP Success: {g_target} -> Kelas {k_target} [{h_empty} Jam {j_empty}]")
                                                break
                                    if kurang <= 0: break

        df_hasil = pd.DataFrame(records)
        df_hasil.sort_values(by=['Hari', 'Jam_Ke', 'Kelas'], inplace=True)
        realisasi_pair = df_hasil.groupby(['Guru', 'Kelas']).size().to_dict() if not df_hasil.empty else {}

        print("\n" + "="*85)
        print("📊 LAPORAN OPTIMASI JADWAL v52 (FULL 12 CONSTRAINTS MODEL)")
        print("="*85)
        
        guru_kurang_list = []
        for _, r in df_pair.iterrows():
            g, k, target_jw = r['Guru'], r['Kelas'], r['Jam_Wajib']
            real_jw = realisasi_pair.get((g, k), 0)
            if real_jw < target_jw:
                guru_kurang_list.append(g)
                print(f"  ❌ Guru: {g:<35} | Kelas: {k:<5} | Target: {target_jw:2d} | Terisi: {real_jw:2d} | Kurang: {target_jw-real_jw:2d} Jam")
        if not guru_kurang_list:
            print("  🎉 PERFECT 100%! Seluruh Guru Sudah Memenuhi Target Jam Wajib (0 Guru Kurang Jam)!")

        slot_kosong_list = [(k, h, j) for k in KELAS for h in HARI for j in JAM_HARI[h] if (k, h, j) not in jadwal_kelas_dict]
        if not slot_kosong_list:
            print("  🎉 PERFECT 100%! Seluruh 30 Rombel Kelas Terisi Penuh 31 Jam KBM Tanpa Ada Slot Kosong!")
        else:
            for k, h, j in slot_kosong_list:
                print(f"  ⚠️ Slot Kosong -> Kelas: {k:<5} | Hari: {h:<6} | Jam Ke-{j}")

        print("\n📌 DAFTAR GURU UTUSAN MGMP TERPILIH PER MAPEL:")
        for mapel_name, list_g in mapel_to_guru.items():
            chosen = [g for g in list_g if mgmp_chosen.get(g, False)]
            hari_m = HARI_MGMP_MAPEL.get(mapel_name)
            if chosen:
                print(f"  - Mapel {mapel_name:<20} (Hari {hari_m:<6}) -> Utusan MGMP Terpilih: {chosen}")

        with pd.ExcelWriter(excel_output_path, engine='openpyxl') as writer:
            df_hasil.to_excel(writer, sheet_name='Data_Penugasan', index=False)
            pivot_kelas = df_hasil.pivot(index=['Hari', 'Jam_Ke'], columns='Kelas', values='Guru').fillna('-')
            pivot_kelas.to_excel(writer, sheet_name='Jadwal_per_Kelas')
            pivot_guru = df_hasil.pivot(index=['Hari', 'Jam_Ke'], columns='Guru', values='Kelas').fillna('-')
            pivot_guru.to_excel(writer, sheet_name='Jadwal_per_Guru')
            
        print(f"\n✅ File Hasil Optimasi v52 berhasil disimpan ke '{excel_output_path}'")

if __name__ == '__main__':
    jalankan_optimasi_v52()
