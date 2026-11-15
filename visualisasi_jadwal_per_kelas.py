"""
================================================================================
1. SKRIP VISUALISASI JADWAL PELAJARAN PER KELAS (HARI, JAM KE, GURU & MAPEL)
================================================================================
"""

import os
import pandas as pd
import matplotlib.pyplot as plt

def visualisasikan_jadwal_per_kelas(excel_output='Jadwal_Guru_SMPN2Plered.xlsx', excel_input='data_jadwal.xlsx'):
    if not os.path.exists(excel_output):
        print(f"❌ File '{excel_output}' tidak ditemukan! Jalankan optimasi jadwal terlebih dahulu.")
        return

    print("==================================================================")
    print("       MENAMPILKAN JADWAL VISUAL PER KELAS (MAPEL + GURU)          ")
    print("==================================================================")

    df_hasil = pd.read_excel(excel_output, sheet_name='Data_Penugasan')
    
    mapel_dict = {}
    if os.path.exists(excel_input):
        try:
            df_status = pd.read_excel(excel_input, sheet_name='Status_Guru')
            df_status['Guru'] = df_status['Guru'].astype(str).str.strip()
            df_status['Mata_Pelajaran'] = df_status['Mata_Pelajaran'].astype(str).str.strip()
            mapel_dict = df_status.set_index('Guru')['Mata_Pelajaran'].to_dict()
        except Exception as e:
            print(f"Catatan: Mapel dibaca default ({e})")

    df_hasil['Guru'] = df_hasil['Guru'].astype(str).str.strip()
    df_hasil['Kelas'] = df_hasil['Kelas'].astype(str).str.strip().str.upper()
    df_hasil['Mapel'] = df_hasil['Guru'].map(mapel_dict).fillna('KBM')
    
    def format_label(row):
        guru_singkat = row['Guru'].split(',')[0].strip()
        return f"{row['Mapel']}\n({guru_singkat})"
        
    df_hasil['Label'] = df_hasil.apply(format_label, axis=1)

    HARI_LIST = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']
    semua_kelas = sorted(list(df_hasil['Kelas'].unique()))
    
    print(f"Memproses Tampilan Matriks untuk {len(semua_kelas)} Kelas...\n")

    for nama_kelas in semua_kelas:
        df_k = df_hasil[df_hasil['Kelas'] == str(nama_kelas)]
        if df_k.empty:
            continue
            
        grid = pd.DataFrame(index=range(1, 8), columns=HARI_LIST).fillna('-')
        
        for _, row in df_k.iterrows():
            grid.loc[row['Jam_Ke'], row['Hari']] = row['Label']
            
        fig, ax = plt.subplots(figsize=(13, 7))
        ax.axis('off')
        
        table = ax.table(
            cellText=grid.values,
            rowLabels=[f"Jam Ke-{i}" for i in range(1, 8)],
            colLabels=HARI_LIST,
            cellLoc='center',
            loc='center',
            colColours=['#1F4E78']*5,
            rowColours=['#D9E1F2']*7
        )
        
        table.auto_set_font_size(False)
        table.set_fontsize(8.5)
        table.scale(1.25, 2.3)
        
        for (row_idx, col_idx), cell in table.get_celld().items():
            if row_idx == 0:
                cell.set_text_props(weight='bold', color='white', fontsize=11)
            elif col_idx == -1:
                cell.set_text_props(weight='bold', color='#1F4E78', fontsize=10)
            else:
                val = cell.get_text().get_text()
                if val == '-':
                    cell.set_facecolor('#FCE4D6')
                    cell.set_text_props(color='#C00000', weight='bold')
                else:
                    cell.set_facecolor('#E2EFDA')
                    cell.set_text_props(weight='bold', color='#1E4620')
                    
        plt.title(f"JADWAL PELAJARAN KELAS {nama_kelas} - SMPN 2 PLERED", fontsize=15, fontweight='bold', pad=20)
        plt.tight_layout()
        plt.show()

if __name__ == '__main__':
    visualisasikan_jadwal_per_kelas()
