"""
================================================================================
2. SKRIP VISUALISASI KETERPENUHAN JAM KBM KELAS (PER HARI & TOTAL SEMINGGU 31 JAM)
================================================================================
"""

import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

def visualisasikan_keterpenuhan_kelas(excel_output='Jadwal_Guru_SMPN2Plered.xlsx'):
    if not os.path.exists(excel_output):
        print(f"❌ File '{excel_output}' tidak ditemukan! Jalankan optimasi jadwal terlebih dahulu.")
        return

    print("==================================================================")
    print("   VISUALISASI KETERPENUHAN JAM KBM ALL KELAS (TARGET 31 JAM)     ")
    print("==================================================================")

    df_hasil = pd.read_excel(excel_output, sheet_name='Data_Penugasan')
    df_hasil['Kelas'] = df_hasil['Kelas'].astype(str).str.strip().str.upper()
    
    HARI_LIST = ['Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat']
    semua_kelas = sorted(list(df_hasil['Kelas'].unique()))
    
    # -------------------------------------------------------------------------
    # 1. VISUALISASI PER HARI (SENIN - JUMAT)
    # -------------------------------------------------------------------------
    records_harian = []
    for k in semua_kelas:
        row = {'Kelas': k}
        for h in HARI_LIST:
            count = len(df_hasil[(df_hasil['Kelas'] == k) & (df_hasil['Hari'] == h)])
            row[h] = count
        records_harian.append(row)
        
    df_harian = pd.DataFrame(records_harian).set_index('Kelas')
    
    fig, ax = plt.subplots(figsize=(12, 10))
    
    sns.heatmap(
        df_harian, 
        annot=True, 
        fmt='d', 
        cmap='YlGnBu', 
        cbar=True, 
        linewidths=1, 
        linecolor='white',
        ax=ax,
        annot_kws={'weight': 'bold', 'size': 10}
    )
    
    plt.title("REKAP KETERPENUHAN JAM KBM HARIAN PER KELAS\n(Target: Senin=7 Jam | Selasa-Jumat=6 Jam)", fontsize=14, fontweight='bold', pad=15)
    plt.xlabel("Hari Sekolah", fontsize=12, fontweight='bold')
    plt.ylabel("Rombel Kelas", fontsize=12, fontweight='bold')
    plt.tight_layout()
    plt.show()

    # -------------------------------------------------------------------------
    # 2. VISUALISASI TOTAL SEMINGGU (TARGET 31 JAM)
    # -------------------------------------------------------------------------
    total_jam_kelas = df_harian.sum(axis=1)
    df_total = pd.DataFrame({
        'Kelas': total_jam_kelas.index,
        'Terisi_Jam': total_jam_kelas.values,
        'Target_Jam': 31,
        'Status': ['100% PAS (31 Jam)' if v == 31 else f'Kurang {31-v} Jam' for v in total_jam_kelas.values]
    })
    
    fig, ax = plt.subplots(figsize=(14, 8))
    
    bars = ax.barh(df_total['Kelas'], df_total['Terisi_Jam'], color='#2E75B6', edgecolor='navy', height=0.7)
    ax.axvline(31, color='green', linestyle='--', linewidth=2, label='Target Wajib (31 Jam/Minggu)')
    
    for bar in bars:
        width = bar.get_width()
        ax.text(width + 0.3, bar.get_y() + bar.get_height()/2, f'{int(width)}/31 Jam (100% PAS)', 
                ha='left', va='center', fontweight='bold', color='#1E4620', fontsize=9)
                
    ax.set_xlim(0, 36)
    ax.invert_yaxis()
    plt.title("KETERPENUHAN TOTAL JAM KBM SEMINGGU SEMUA KELAS (30 ROMBEL KELAS = 31 JAM)", fontsize=14, fontweight='bold', pad=15)
    plt.xlabel("Total Jam KBM per Minggu", fontsize=12, fontweight='bold')
    plt.ylabel("Rombel Kelas", fontsize=12, fontweight='bold')
    plt.legend(loc='lower right')
    plt.grid(axis='x', linestyle=':', alpha=0.6)
    plt.tight_layout()
    plt.show()

if __name__ == '__main__':
    visualisasikan_keterpenuhan_kelas()
