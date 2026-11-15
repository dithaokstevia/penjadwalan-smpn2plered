"""
================================================================================
3. SKRIP VISUALISASI KETERPENUHAN JAM WAJIB GURU (DENGAN DETAIL MAPEL & JAM)
================================================================================
"""

import os
import pandas as pd
import matplotlib.pyplot as plt

def visualisasikan_keterpenuhan_guru(excel_output='Jadwal_Guru_SMPN2Plered.xlsx', excel_input='data_jadwal.xlsx'):
    if not os.path.exists(excel_output) or not os.path.exists(excel_input):
        print("❌ File input/output tidak ditemukan! Pastikan file Excel sudah tersedia.")
        return

    print("==================================================================")
    print("   VISUALISASI KETERPENUHAN JAM WAJIB GURU (DETAIL MULTI-MAPEL)   ")
    print("==================================================================")

    df_hasil = pd.read_excel(excel_output, sheet_name='Data_Penugasan')
    df_beban = pd.read_excel(excel_input, sheet_name='Guru_Kelas')
    df_status = pd.read_excel(excel_input, sheet_name='Status_Guru')

    df_hasil['Guru'] = df_hasil['Guru'].astype(str).str.strip()
    df_beban['Guru'] = df_beban['Guru'].astype(str).str.strip()
    df_status['Guru'] = df_status['Guru'].astype(str).str.strip()
    df_status['Mata_Pelajaran'] = df_status['Mata_Pelajaran'].astype(str).str.strip()

    mapel_utama = df_status.set_index('Guru')['Mata_Pelajaran'].to_dict()

    guru_summary = []
    
    for g, group in df_beban.groupby('Guru'):
        if g.upper() in ['NAN', 'NONE', '', 'NULL', 'GURU', 'KELAS']:
            continue
            
        total_target = group['Jam_Wajib'].sum()
        real_terisi = len(df_hasil[df_hasil['Guru'] == g])
        
        mapel_primary = mapel_utama.get(g, 'KBM')
        rincian_mapel_text = f"{mapel_primary} ({total_target} Jam)"
        
        guru_summary.append({
            'Guru': g,
            'Mapel_Detail': rincian_mapel_text,
            'Target_Jam': total_target,
            'Terisi_Jam': real_terisi,
            'Status': '100% TERPENUHI' if real_terisi == total_target else f'Kurang {total_target-real_terisi} Jam'
        })
        
    df_summary = pd.DataFrame(guru_summary).sort_values(by='Target_Jam', ascending=False)
    
    fig, ax = plt.subplots(figsize=(15, len(df_summary)*0.35 + 2))
    ax.axis('off')
    
    table_data = []
    for _, row in df_summary.iterrows():
        table_data.append([
            row['Guru'],
            row['Mapel_Detail'],
            f"{row['Target_Jam']} Jam",
            f"{row['Terisi_Jam']} Jam",
            f"✅ {row['Status']}" if row['Terisi_Jam'] == row['Target_Jam'] else f"❌ {row['Status']}"
        ])
        
    col_labels = ['Nama Guru', 'Rincian Mata Pelajaran & Jam', 'Target Jam Wajib', 'Realisasi Jam Terisi', 'Status Keterpenuhan']
    
    table = ax.table(
        cellText=table_data,
        colLabels=col_labels,
        cellLoc='left',
        loc='center',
        colColours=['#1F4E78']*5
    )
    
    table.auto_set_font_size(False)
    table.set_fontsize(9)
    table.scale(1.2, 1.8)
    
    for (row_idx, col_idx), cell in table.get_celld().items():
        if row_idx == 0:
            cell.set_text_props(weight='bold', color='white', fontsize=10, ha='center')
        else:
            if col_idx == 4:
                cell.set_text_props(weight='bold', color='#1E4620' if '100%' in cell.get_text().get_text() else '#C00000')
                cell.set_facecolor('#E2EFDA' if '100%' in cell.get_text().get_text() else '#FCE4D6')
            elif col_idx in [2, 3]:
                cell.set_text_props(ha='center')
                
    plt.title("REKAP AUDIT KETERPENUHAN JAM WAJIB GURU (SMPN 2 PLERED)", fontsize=14, fontweight='bold', pad=20)
    plt.tight_layout()
    plt.show()

if __name__ == '__main__':
    visualisasikan_keterpenuhan_guru()
