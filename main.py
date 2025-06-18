import cv2
import numpy as np
import qrcode
from PIL import Image, ImageTk
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from collections import defaultdict
import pandas as pd
from datetime import datetime
import pygame

pygame.mixer.init()

class QRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Absensi Mahasiswa Polinema PSDKU KEDIRI")
        self.root.geometry("900x800")
        self.root.resizable(False, False)
        self.mode_absen = "masuk"

        self.data_mahasiswa = defaultdict(lambda: {
            "nama": "",
            "kelas": "",
            "log_absensi": {}  
        })


        self.cap = None
        self.detector = cv2.QRCodeDetector()
        self.scanning = False
        self.current_qr_img = None

        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use("default")
        style.configure("TFrame", background="#f0f4ff")
        style.configure("TLabel", background="#f0f4ff", font=("Segoe UI", 10))
        style.configure("TLabelFrame", background="#f0f4ff", font=("Segoe UI", 10, "bold"))
        style.configure("TButton", font=("Segoe UI", 9), padding=6)

        title_label = ttk.Label(self.root, text="ABSENSI MAHASISWA D3 Manajemen Informatika", font=("Segoe UI", 16, "bold"), anchor="center")
        title_label.pack(pady=15)

        notebook = ttk.Notebook(self.root)
        notebook.pack(expand=True, fill='both')

        tab1 = ttk.Frame(notebook)
        notebook.add(tab1, text="Generate QR Code")

        self.setup_generate_tab(tab1)

        tab2 = ttk.Frame(notebook)
        notebook.add(tab2, text="Scan & Absensi")

        self.setup_scan_tab(tab2)

        tab3 = ttk.Frame(notebook)
        notebook.add(tab3, text="Export Data")

        self.setup_export_tab(tab3)

    def setup_generate_tab(self, parent):
        frame_input = ttk.LabelFrame(parent, text="Input Data Mahasiswa")
        frame_input.pack(fill='x', padx=10, pady=10)

        labels = ["NIM:", "Nama:", "Kelas:"]
        entries = []

        for i, label in enumerate(labels):
            ttk.Label(frame_input, text=label).grid(row=i, column=0, padx=5, pady=5, sticky="e")
            entry = ttk.Entry(frame_input, width=30)
            entry.grid(row=i, column=1, padx=5, pady=5)
            entries.append(entry)

        self.nim_entry_gen, self.nama_entry_gen, self.kelas_entry_gen = entries

        btn_input = ttk.Frame(frame_input)
        btn_input.grid(row=3, column=0, columnspan=2, pady=10)
        btn_input.columnconfigure((0, 1), weight=1)

        ttk.Button(btn_input, text="Tambah", command=self.tambah_data_gen).grid(row=0, column=0, padx=5)
        ttk.Button(btn_input, text="Update", command=self.update_data_gen).grid(row=0, column=1, padx=5)

        btn_bawah = ttk.Frame(parent)
        btn_bawah.pack(pady=5)

        ttk.Button(btn_bawah, text="Generate QR", command=self.generate_qr_gen).pack(side='left', padx=10)
        ttk.Button(btn_bawah, text="Hapus Mahasiswa", command=self.hapus_mahasiswa_gen).pack(side='left', padx=10)


        self.qr_label_gen = ttk.Label(frame_input)
        self.qr_label_gen.grid(row=4, column=0, columnspan=2, pady=10)

        frame_list = ttk.LabelFrame(parent, text="Daftar Mahasiswa")
        frame_list.pack(fill='both', expand=True, padx=10, pady=10)

        self.list_mahasiswa_gen = tk.Listbox(frame_list, height=8)
        self.list_mahasiswa_gen.pack(side='left', fill='both', expand=True, padx=(0, 5), pady=5)

        scrollbar = ttk.Scrollbar(frame_list, orient='vertical', command=self.list_mahasiswa_gen.yview)
        scrollbar.pack(side='right', fill='y')
        self.list_mahasiswa_gen.config(yscrollcommand=scrollbar.set)
        self.update_list_mahasiswa_gen()


    def setup_scan_tab(self, parent):
        frame_scan = ttk.LabelFrame(parent, text="Scan QR dan Absensi")
        frame_scan.pack(fill='x', padx=10, pady=10)

        ttk.Label(frame_scan, text="Hari:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.hari_var = tk.StringVar()
        hari_options = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat"]
        self.hari_combo = ttk.Combobox(frame_scan, textvariable=self.hari_var, values=hari_options, state="readonly")
        self.hari_combo.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        self.hari_combo.current(0)  

        ttk.Label(frame_scan, text="Kelas:").grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.kelas_var = tk.StringVar()
        kelas_options = ["1A", "1B", "1C", "1D", "1E"]
        self.selected_kelas = tk.StringVar()
        self.kelas_combo = ttk.Combobox(frame_scan, textvariable=self.kelas_var, values=kelas_options, state="readonly")
        self.kelas_combo.grid(row=0, column=3, sticky='w', padx=5, pady=5)
        self.kelas_combo.current(0)

        ttk.Label(frame_scan, text="Mata Kuliah:").grid(row=0, column=4, sticky='w', padx=5, pady=5)
        self.matkul_var = tk.StringVar()
        self.selected_matkul = tk.StringVar()
        self.matkul_combo = ttk.Combobox(frame_scan, textvariable=self.matkul_var, state="readonly")
        self.matkul_combo.grid(row=0, column=5, sticky='w', padx=5, pady=5)

        self.hari_combo.bind("<<ComboboxSelected>>", self.update_matkul_options)
        self.kelas_combo.bind("<<ComboboxSelected>>", self.update_matkul_options)

        self.update_matkul_options()

        self.scan_btn = ttk.Button(frame_scan, text="Start Scan", command=self.toggle_scan)
        self.scan_btn.grid(row=1, column=0, columnspan=6, pady=5, sticky='ew')

        btn_absen_frame = ttk.Frame(frame_scan)
        btn_absen_frame.grid(row=2, column=0, columnspan=6, pady=5)

        self.btn_masuk = ttk.Button(btn_absen_frame, text="Absen Masuk", command=self.set_mode_masuk)
        self.btn_masuk.pack(side='left', padx=5)

        self.btn_pulang = ttk.Button(btn_absen_frame, text="Absen Pulang", command=self.set_mode_pulang)
        self.btn_pulang.pack(side='left', padx=5)

        self.video_frame = ttk.Label(frame_scan)
        self.video_frame.grid(row=3, column=0, columnspan=6, pady=5)

        frame_absen = ttk.LabelFrame(parent, text="Data Absensi")
        frame_absen.pack(fill='both', expand=True, padx=10, pady=10)
        

        self.listbox = tk.Listbox(frame_absen, height=15, font=("Courier New", 10))
        self.listbox.pack(fill='both', padx=10, pady=5)

        self.update_listbox_absensi()

        ttk.Button(frame_absen, text="Reset Absensi", command=self.reset_absensi).pack(pady=5)


    def setup_export_tab(self, parent):
        frame_export = ttk.LabelFrame(parent, text="Export Data Absensi")
        frame_export.pack(fill='both', expand=True, padx=10, pady=10)

        ttk.Button(frame_export, text="Export ke Excel", command=self.export_excel).pack(pady=20)

    def tambah_data_gen(self):
        nim = self.nim_entry_gen.get().strip()
        nama = self.nama_entry_gen.get().strip()
        kelas = self.kelas_entry_gen.get().strip() 
        if not nim or not nama or not kelas:
            messagebox.showwarning("Input Kosong", "Isi semua data terlebih dahulu.")
            return
        if nim in self.data_mahasiswa:
            messagebox.showerror("Gagal", f"NIM {nim} sudah terdaftar.")
            return
        self.data_mahasiswa[nim]["nama"] = nama
        self.data_mahasiswa[nim]["kelas"] = kelas
        messagebox.showinfo("Sukses", f"Data untuk NIM {nim} berhasil ditambahkan.")
        self.update_list_mahasiswa_gen()
        self.update_listbox_absensi()


    def update_list_mahasiswa_gen(self):
        self.list_mahasiswa_gen.delete(0, tk.END)

        self.list_mahasiswa_gen.insert(tk.END, "Daftar Mahasiswa:")
        self.list_mahasiswa_gen.insert(tk.END, "-" * 55)

        for nim, data in self.data_mahasiswa.items():
            nama = data.get("nama", "-")
            kelas = data.get("kelas", "-")
            line = f"{nim} - {nama} - {kelas}"
            self.list_mahasiswa_gen.insert(tk.END, line)

    
    def hapus_mahasiswa_gen(self):
        selected = self.list_mahasiswa_gen.curselection()
        if not selected:
            messagebox.showwarning("Peringatan", "Pilih mahasiswa yang ingin dihapus.")
            return

        item = self.list_mahasiswa_gen.get(selected[0])
        nim = item.split(" - ")[0]

        confirm = messagebox.askyesno("Konfirmasi", f"Apakah Anda yakin ingin menghapus mahasiswa dengan NIM {nim}?")
        if confirm and nim in self.data_mahasiswa:
            del self.data_mahasiswa[nim]
            self.update_list_mahasiswa_gen()
            self.update_listbox_absensi() 
            messagebox.showinfo("Berhasil", f"Mahasiswa {nim} berhasil dihapus.")


    def update_data_gen(self):
        nim = self.nim_entry_gen.get().strip()
        nama = self.nama_entry_gen.get().strip()
        kelas = self.kelas_entry_gen.get().strip()
        if not nim or not nama or not kelas:
            messagebox.showwarning("Input Kosong", "Isi semua data terlebih dahulu.")
            return
        if nim not in self.data_mahasiswa:
            messagebox.showerror("Gagal", f"NIM {nim} belum terdaftar.")
            return
        self.data_mahasiswa[nim]["nama"] = nama
        self.data_mahasiswa[nim]["kelas"] = kelas
        messagebox.showinfo("Sukses", f"Data untuk NIM {nim} berhasil diperbarui.")
        self.update_list_mahasiswa_gen()
        self.update_listbox_absensi()
        

    def generate_qr_gen(self):
        selected = self.list_mahasiswa_gen.curselection()
        if not selected:
            messagebox.showwarning("Tidak Ada Pilihan", "Pilih salah satu mahasiswa dari daftar terlebih dahulu.")
            return

        item_text = self.list_mahasiswa_gen.get(selected[0])
        nim = item_text.split(" - ")[0].strip()  

        qr_img = qrcode.make(nim)
        self.current_qr_img = qr_img
        qr_tk = ImageTk.PhotoImage(qr_img.resize((200, 200)))
        self.qr_label_gen.config(image=qr_tk)
        self.qr_label_gen.image = qr_tk

        file_path = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG", "*.png")])
        if file_path:
            qr_img.save(file_path)
            messagebox.showinfo("Sukses", "QR Code berhasil disimpan.")


    def update_matkul_options(self, event=None):

        jadwal = {
            "1A": {
                "Senin": ["Matematika", "Bahasa Indonesia"],
                "Selasa": ["Fisika", "Agama"],
                "Rabu": ["Kimia", "Sejarah"],
                "Kamis": ["Biologi", "Seni Budaya"],
                "Jumat": ["Pendidikan Jasmani", "Bahasa Inggris"]
            },
            "1B": {
                "Senin": ["Bahasa Inggris", "Matematika"],
                "Selasa": ["Fisika", "Kimia"],
                "Rabu": ["Biologi", "Agama"],
                "Kamis": ["Sejarah", "Pendidikan Jasmani"],
                "Jumat": ["Seni Budaya", "Bahasa Indonesia"]
            },
            "1C": {
                "Senin": ["Bahasa Inggris", "Matematika"],
                "Selasa": ["Fisika", "Kimia"],
                "Rabu": ["Biologi", "Agama"],
                "Kamis": ["Sejarah", "Pendidikan Jasmani"],
                "Jumat": ["Seni Budaya", "Bahasa Indonesia"]
            },
            "1D": {
                "Senin": ["Bahasa Inggris", "Matematika"],
                "Selasa": ["Fisika", "Kimia"],
                "Rabu": ["Biologi", "Agama"],
                "Kamis": ["Sejarah", "Pendidikan Jasmani"],
                "Jumat": ["Seni Budaya", "Bahasa Indonesia"]
            },
            "1E": {
                "Senin": ["Bahasa Inggris", "Matematika"],
                "Selasa": ["Fisika", "Kimia"],
                "Rabu": ["Biologi", "Agama"],
                "Kamis": ["Sejarah", "Pendidikan Jasmani"],
                "Jumat": ["Seni Budaya", "Bahasa Indonesia"]
            },

        }

        kelas = self.kelas_var.get()
        hari = self.hari_var.get()

        if kelas in jadwal and hari in jadwal[kelas]:
            matkul_list = jadwal[kelas][hari]
        else:
            matkul_list = []

        self.matkul_combo['values'] = matkul_list
        if matkul_list:
            self.matkul_combo.current(0)
        else:
            self.matkul_combo.set('')


    def toggle_scan(self):
        self.scanning = not self.scanning
        self.scan_btn.config(text="Stop Scan" if self.scanning else "Start Scan")
        if self.scanning:
            self.cap = cv2.VideoCapture(0)
            self.scan_loop()
        else:
            if self.cap: 
                self.cap.release()
                self.cap = None
            self.video_frame.config(image='')

    def set_mode_masuk(self):
        self.mode_absen = "masuk"
        messagebox.showinfo("Mode Absen", "Mode absen masuk diaktifkan.")

    def set_mode_pulang(self):
        self.mode_absen = "pulang"
        messagebox.showinfo("Mode Absen", "Mode absen pulang diaktifkan.")

    def scan_loop(self):
        if not self.scanning:
            return

        ret, frame = self.cap.read()
        frame = cv2.flip(frame, 1)
        if not ret:
            self.video_frame.after(10, self.scan_loop)
            return

        try:
            data, bbox, _ = self.detector.detectAndDecode(frame)
        except cv2.error as e:
            print("QR Decode Error:", e)
            self.video_frame.after(10, self.scan_loop)
            return

        if bbox is not None:
            bbox = bbox.astype(int)
            for i in range(len(bbox[0])):
                pt1 = tuple(bbox[0][i])
                pt2 = tuple(bbox[0][(i + 1) % len(bbox[0])])
                cv2.line(frame, pt1, pt2, (0, 255, 0), 3)

        if data:
            nim = data.strip()
            if nim in self.data_mahasiswa:
                self.proses_absensi(nim)
            else:
                pygame.mixer.music.load("not_valid.mp3")
                pygame.mixer.music.play()
                messagebox.showerror(nim , "Tidak Ada Silahkan Verifikasi admin")

        cv2image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGBA)
        img = Image.fromarray(cv2image)
        imgtk = ImageTk.PhotoImage(image=img)
        self.video_frame.imgtk = imgtk
        self.video_frame.config(image=imgtk)

        self.video_frame.after(10, self.scan_loop)

    def proses_absensi(self, nim):
        mahasiswa = self.data_mahasiswa[nim]
        now = datetime.now()
        hari_ini = self.hari_var.get()
        matkul_dipilih = self.matkul_var.get()
        kelas_dipilih = self.kelas_var.get()

        if mahasiswa["kelas"] != kelas_dipilih:
            pygame.mixer.music.load("val_kelas.mp3")
            pygame.mixer.music.play()
            messagebox.showerror("Kelas Tidak Cocok", f"{mahasiswa['nama']} ({nim}) bukan dari kelas {kelas_dipilih}.")
            return

        log_hari = mahasiswa.get("log_absensi", {})
        log_key = f"{hari_ini}_{matkul_dipilih}"
        absensi = log_hari.get(log_key, {})

        if self.mode_absen == "masuk":
            if "masuk" in absensi and "pulang" in absensi:
                pygame.mixer.music.load("val_absen.mp3")
                pygame.mixer.music.play()
                messagebox.showinfo("Info", f"{mahasiswa['nama']} ({nim}) sudah absen untuk {matkul_dipilih} hari {hari_ini}.")

            elif "masuk" in absensi:
                pygame.mixer.music.load("error_absen_in.mp3")
                pygame.mixer.music.play()
                messagebox.showinfo("Info", f"{mahasiswa['nama']} ({nim}) sudah absen masuk.")

            else:
                waktu_masuk = now.strftime("%H:%M:%S")
                absensi["masuk_full"] = now.isoformat()
                absensi["masuk"] = waktu_masuk
                log_hari[log_key] = absensi
                mahasiswa["log_absensi"] = log_hari
                pygame.mixer.music.load("absen_in_succes.mp3")
                pygame.mixer.music.play()
                messagebox.showinfo("Hadir", f"{mahasiswa['nama']} ({nim}) berhasil absen masuk.")
                self.update_listbox_absensi()

        elif self.mode_absen == "pulang":
            if "masuk" not in absensi:
                pygame.mixer.music.load("error_to_absen_out_in.mp3")
                pygame.mixer.music.play()
                messagebox.showwarning("Belum Absen Masuk", f"{mahasiswa['nama']} ({nim}) belum melakukan absen masuk.")
            elif "pulang" in absensi:
                pygame.mixer.music.load("error_absen_out.mp3")
                pygame.mixer.music.play()
                messagebox.showinfo("Info", f"{mahasiswa['nama']} ({nim}) sudah absen pulang.")
            else:
                now = datetime.now()
                waktu_pulang = now.strftime("%H:%M:%S")
                waktu_pulang_full = now 

                masuk_full = datetime.fromisoformat(absensi["masuk_full"])
                total = waktu_pulang_full - masuk_full

                total_detik = int(total.total_seconds())
                jam = total_detik // 3600
                menit = (total_detik % 3600) // 60
                durasi = f"{jam} jam {menit} menit"

                absensi["pulang"] = waktu_pulang
                absensi["durasi"] = durasi
                log_hari[log_key] = absensi
                mahasiswa["log_absensi"] = log_hari

                pygame.mixer.music.load("absen_out_succes.mp3")
                pygame.mixer.music.play()
                messagebox.showinfo("Pulang", f"{mahasiswa['nama']} ({nim}) berhasil absen pulang.\nTotal: {durasi}")
                self.update_listbox_absensi()

    def update_listbox_absensi(self):
        self.listbox.delete(0, tk.END)
        self.listbox.insert(tk.END, f"{'NIM':<15}{'Nama':<25}{'Kelas':<10}{'Hadir':<7}{'Hari':<10}{'Matkul':<20}{'Masuk':<15}{'Pulang':<15}{'Durasi':<10}")
        self.listbox.insert(tk.END, "-"*130)

        for nim, data in self.data_mahasiswa.items():
            log_absen = data.get("log_absensi", {})
            nama = data.get("nama", "-")
            kelas = data.get("kelas", "-")

            if log_absen:
                for log_key, absensi in log_absen.items():
                    hari_matkul = log_key.split("_")
                    hari = hari_matkul[0] if len(hari_matkul) > 0 else "-"
                    matkul = hari_matkul[1] if len(hari_matkul) > 1 else "-"
                    masuk = absensi.get("masuk", "-")
                    pulang = absensi.get("pulang", "-")
                    durasi = absensi.get("durasi", "-")
                    hadir = "Ya" if "masuk" in absensi else "Tidak"

                    line = f"{nim:<15}{nama:<25}{kelas:<10}{hadir:<7}{hari:<10}{matkul:<20}{masuk:<15}{pulang:<15}{durasi:<10}"
                    self.listbox.insert(tk.END, line)
            else:
                line = f"{nim:<15}{nama:<25}{kelas:<10}{'Tidak':<7}{'-':<10}{'-':<20}{'-':<15}{'-':<15}{'-':<10}"
                self.listbox.insert(tk.END, line)


    def reset_absensi(self):
        for nim in self.data_mahasiswa:
            if "log_absensi" in self.data_mahasiswa[nim]:
                self.data_mahasiswa[nim]["log_absensi"] = {}
        self.update_listbox_absensi()
        messagebox.showinfo("Reset", "Data absensi telah direset.")


    def export_excel(self):
        if not self.data_mahasiswa:
            messagebox.showwarning("Data Kosong", "Belum ada data untuk diexport.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            export_data = []

            for nim, data in self.data_mahasiswa.items():
                nama = data.get("nama", "")
                kelas = data.get("kelas", "")
                log_absensi = data.get("log_absensi", {})

                for log_key, absen in log_absensi.items():
                    hari, matkul = log_key.split("_", 1) if "_" in log_key else (log_key, "")
                    masuk = absen.get("masuk", "")
                    pulang = absen.get("pulang", "")
                    durasi = absen.get("durasi", "")
                    export_data.append({
                        "NIM": nim,
                        "Nama": nama,
                        "Kelas": kelas,
                        "Hari": hari,
                        "Mata Kuliah": matkul,
                        "Waktu Masuk": masuk,
                        "Waktu Pulang": pulang,
                        "Durasi": durasi
                    })

            df = pd.DataFrame(export_data)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Sukses", f"Data berhasil diexport ke {file_path}")
        
        except Exception as e:
            messagebox.showerror("Error", f"Gagal export data:\n{str(e)}")


root = tk.Tk()
app = QRApp(root)
root.mainloop()
