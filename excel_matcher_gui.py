#!/usr/bin/env python3
"""
Excel Name Matcher - GUI Version
User-friendly interface for matching names between two Excel files
"""

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import re
from pathlib import Path
import threading

class NameMatcherGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel İsim Eşleştirme")
        self.root.geometry("900x650")
        self.root.resizable(True, True)
        
        # Variables
        self.master_file = tk.StringVar()
        self.messy_file = tk.StringVar()
        self.master_name_col = tk.StringVar(value="A")  # A column
        self.master_surname_col = tk.StringVar(value="B")  # B column
        self.messy_name_col = tk.StringVar(value="E")  # E column
        
        self.setup_ui()
    
    def excel_col_to_index(self, col_str):
        """Convert Excel column letter(s) to 0-based index (A=0, B=1, Z=25, AA=26, etc.)"""
        col_str = col_str.upper().strip()
        result = 0
        for char in col_str:
            result = result * 26 + (ord(char) - ord('A') + 1)
        return result - 1
    
    def index_to_excel_col(self, index):
        """Convert 0-based index to Excel column letter(s) (0=A, 1=B, 25=Z, 26=AA, etc.)"""
        result = ""
        index += 1  # Convert to 1-based
        while index > 0:
            index -= 1
            result = chr(index % 26 + ord('A')) + result
            index //= 26
        return result
    
    def setup_ui(self):
        # Title
        title_label = tk.Label(
            self.root, 
            text="Excel İsim Eşleştirme", 
            font=("Arial", 18, "bold"),
            pady=20
        )
        title_label.pack()
        
        # Main Frame
        main_frame = tk.Frame(self.root, padx=30, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Master Excel Selection
        master_frame = tk.LabelFrame(main_frame, text="📄 Ana Liste (Master Excel)", padx=15, pady=15)
        master_frame.pack(fill=tk.X, pady=10)
        
        # File selection row
        file_row = tk.Frame(master_frame)
        file_row.pack(fill=tk.X, pady=5)
        tk.Label(file_row, text="Excel Dosyası:", width=15, anchor='w').pack(side=tk.LEFT)
        tk.Entry(file_row, textvariable=self.master_file, state='readonly', width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(file_row, text="📁 Dosya Seç", command=self.select_master_file, bg="#2196F3", fg="white", width=12).pack(side=tk.LEFT)
        
        # Column settings
        col_frame = tk.Frame(master_frame)
        col_frame.pack(fill=tk.X, pady=5)
        tk.Label(col_frame, text="İsim Sütunu (A,B,C...):", width=20, anchor='w').pack(side=tk.LEFT)
        tk.Entry(col_frame, textvariable=self.master_name_col, width=8).pack(side=tk.LEFT, padx=5)
        tk.Label(col_frame, text="Soyisim Sütunu (A,B,C...):", width=25, anchor='w').pack(side=tk.LEFT, padx=(20,0))
        tk.Entry(col_frame, textvariable=self.master_surname_col, width=8).pack(side=tk.LEFT, padx=5)
        
        # Messy Excel Selection
        messy_frame = tk.LabelFrame(main_frame, text="📝 Kontrol Edilecek Liste (Messy Excel)", padx=15, pady=15)
        messy_frame.pack(fill=tk.X, pady=10)
        
        # File selection row
        file_row2 = tk.Frame(messy_frame)
        file_row2.pack(fill=tk.X, pady=5)
        tk.Label(file_row2, text="Excel Dosyası:", width=15, anchor='w').pack(side=tk.LEFT)
        tk.Entry(file_row2, textvariable=self.messy_file, state='readonly', width=40).pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        tk.Button(file_row2, text="📁 Dosya Seç", command=self.select_messy_file, bg="#2196F3", fg="white", width=12).pack(side=tk.LEFT)
        
        # Column settings
        col_frame2 = tk.Frame(messy_frame)
        col_frame2.pack(fill=tk.X, pady=5)
        tk.Label(col_frame2, text="İsim Sütunu (A,B,C...):", width=20, anchor='w').pack(side=tk.LEFT)
        tk.Entry(col_frame2, textvariable=self.messy_name_col, width=8).pack(side=tk.LEFT, padx=5)
        
        # Run Button
        run_btn = tk.Button(
            main_frame,
            text="🔍 Eşleştirmeyi Başlat",
            command=self.run_matching,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 12, "bold"),
            pady=10,
            cursor="hand2"
        )
        run_btn.pack(pady=20, fill=tk.X)
        
        # Progress Frame
        self.progress_frame = tk.LabelFrame(main_frame, text="📊 Sonuçlar", padx=15, pady=15)
        self.progress_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        self.result_text = tk.Text(self.progress_frame, height=10, wrap=tk.WORD, state='disabled')
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(self.result_text)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_text.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.result_text.yview)
    
    def select_master_file(self):
        self.root.update()  # Force UI update before dialog
        filename = filedialog.askopenfilename(
            parent=self.root,
            title="Ana Liste Excel Dosyasını Seç",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialdir=Path.home() / "Downloads"
        )
        if filename:
            self.master_file.set(filename)
            self.log(f"✅ Ana liste seçildi: {Path(filename).name}")
    
    def select_messy_file(self):
        self.root.update()  # Force UI update before dialog
        filename = filedialog.askopenfilename(
            parent=self.root,
            title="Kontrol Edilecek Excel Dosyasını Seç",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialdir=Path.home() / "Downloads"
        )
        if filename:
            self.messy_file.set(filename)
            self.log(f"✅ Kontrol listesi seçildi: {Path(filename).name}")
    
    def log(self, message):
        self.result_text.config(state='normal')
        self.result_text.insert(tk.END, message + "\n")
        self.result_text.see(tk.END)
        self.result_text.config(state='disabled')
        self.root.update()
    
    def clear_log(self):
        self.result_text.config(state='normal')
        self.result_text.delete(1.0, tk.END)
        self.result_text.config(state='disabled')
    
    def normalize_name(self, text):
        """Normalize name for comparison"""
        if pd.isna(text) or text is None:
            return ""
        
        text = str(text).strip()
        text = text.replace('ı', 'I').replace('i', 'I').replace('İ', 'I').replace('I', 'I')
        text = text.upper()
        text = re.sub(r'\s+', '', text)
        text = re.sub(r'[^A-ZÇĞÖŞÜI]', '', text)
        
        return text
    
    def run_matching(self):
        # Validate inputs
        if not self.master_file.get():
            messagebox.showerror("Hata", "Lütfen ana liste Excel dosyasını seçin!")
            return
        
        if not self.messy_file.get():
            messagebox.showerror("Hata", "Lütfen kontrol edilecek Excel dosyasını seçin!")
            return
        
        try:
            # Convert Excel column letters to indices
            master_name_idx = self.excel_col_to_index(self.master_name_col.get())
            master_surname_idx = self.excel_col_to_index(self.master_surname_col.get())
            messy_name_idx = self.excel_col_to_index(self.messy_name_col.get())
        except Exception as e:
            messagebox.showerror("Hata", f"Geçersiz sütun harfi!\nÖrnek: A, B, C, AA, AB\n\nHata: {e}")
            return
        
        # Run in separate thread to avoid freezing UI
        thread = threading.Thread(
            target=self.process_matching,
            args=(master_name_idx, master_surname_idx, messy_name_idx)
        )
        thread.daemon = True
        thread.start()
    
    def process_matching(self, master_name_idx, master_surname_idx, messy_name_idx):
        try:
            self.clear_log()
            self.log("🚀 İşlem başlatılıyor...\n")
            
            # Load master list
            self.log("📖 Ana liste yükleniyor...")
            master_df = pd.read_excel(self.master_file.get(), header=None)
            self.log(f"✅ Ana liste yüklendi: {len(master_df)} kayıt")
            self.log(f"📍 İsim: {self.index_to_excel_col(master_name_idx)} sütunu, Soyisim: {self.index_to_excel_col(master_surname_idx)} sütunu\n")
            
            # Get column names
            name_col = master_df.columns[master_name_idx]
            surname_col = master_df.columns[master_surname_idx]
            
            # Create normalized names
            master_df['full_name_original'] = master_df[name_col].astype(str) + ' ' + master_df[surname_col].astype(str)
            master_df['normalized_name'] = (master_df[name_col].astype(str) + master_df[surname_col].astype(str)).apply(self.normalize_name)
            
            # Load messy data
            self.log("📖 Kontrol listesi yükleniyor...")
            messy_df = pd.read_excel(self.messy_file.get(), header=None)
            self.log(f"✅ Kontrol listesi yüklendi: {len(messy_df)} kayıt")
            self.log(f"📍 İsim: {self.index_to_excel_col(messy_name_idx)} sütunu\n")
            
            # Get column name
            messy_col = messy_df.columns[messy_name_idx]
            
            # Create normalized names
            messy_df['name_original'] = messy_df[messy_col].astype(str)
            messy_df['normalized_name'] = messy_df[messy_col].apply(self.normalize_name)
            
            # Match names
            self.log("🔍 Eşleşme kontrolü yapılıyor...")
            merged = messy_df.merge(
                master_df,
                on='normalized_name',
                how='left',
                indicator=True,
                suffixes=('_messy', '_master')
            )
            
            matched = merged[merged['_merge'] == 'both'].copy()
            unmatched = merged[merged['_merge'] == 'left_only'].copy()
            
            # Prepare output
            if len(matched) > 0:
                matched_output = matched[['name_original', 'normalized_name', 'full_name_original']].copy()
                matched_output.columns = ['Girilen İsim', 'Normalize Edilmiş', 'Eşleşen Ana Liste İsmi']
            else:
                matched_output = pd.DataFrame(columns=['Girilen İsim', 'Normalize Edilmiş', 'Eşleşen Ana Liste İsmi'])
            
            if len(unmatched) > 0:
                unmatched_output = unmatched[['name_original', 'normalized_name']].copy()
                unmatched_output.columns = ['Girilen İsim', 'Normalize Edilmiş']
            else:
                unmatched_output = pd.DataFrame(columns=['Girilen İsim', 'Normalize Edilmiş'])
            
            # Create summary
            total_records = len(matched) + len(unmatched)
            summary_data = [
                ['Metrik', 'Değer'],
                ['Toplam Kontrol Edilen Kayıt', total_records],
                ['Eşleşen Kayıt Sayısı', len(matched)],
                ['Eşleşmeyen Kayıt Sayısı', len(unmatched)],
                ['Eşleşme Oranı', f"%{len(matched)/total_records*100:.1f}" if total_records > 0 else "%0.0"]
            ]
            summary_df = pd.DataFrame(summary_data[1:], columns=summary_data[0])
            
            # Save to Excel (same folder as Master Excel)
            self.log("\n💾 Rapor kaydediliyor...")
            output_path = Path(self.master_file.get()).parent / 'isim_eslestirme_raporu.xlsx'
            
            file_path = Path(output_path)
            mode = 'a' if file_path.exists() else 'w'
            if_sheet_exists = 'new' if mode == 'a' else None
            
            with pd.ExcelWriter(output_path, mode=mode, engine='openpyxl', if_sheet_exists=if_sheet_exists) as writer:
                summary_df.to_excel(writer, sheet_name='Özet', index=False)
                matched_output.to_excel(writer, sheet_name='Eşleşenler', index=False)
                unmatched_output.to_excel(writer, sheet_name='Eşleşmeyenler', index=False)
            
            # Display results
            self.log(f"✅ Rapor kaydedildi:\n   {output_path}\n")
            self.log("=" * 50)
            self.log("📊 SONUÇLAR")
            self.log("=" * 50)
            self.log(f"✅ Eşleşen: {len(matched)} kayıt")
            self.log(f"❌ Eşleşmeyen: {len(unmatched)} kayıt")
            self.log(f"📊 Eşleşme Oranı: {len(matched)/total_records*100:.1f}%")
            self.log("=" * 50)
            
            messagebox.showinfo(
                "Başarılı!", 
                f"İşlem tamamlandı!\n\n"
                f"✅ Eşleşen: {len(matched)} kayıt\n"
                f"❌ Eşleşmeyen: {len(unmatched)} kayıt\n\n"
                f"Rapor şuraya kaydedildi:\n{output_path.parent}\n\n"
                f"Dosya adı: {output_path.name}"
            )
            
        except Exception as e:
            self.log(f"\n❌ HATA: {str(e)}")
            messagebox.showerror("Hata", f"İşlem sırasında hata oluştu:\n\n{str(e)}")

def main():
    root = tk.Tk()
    app = NameMatcherGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()

