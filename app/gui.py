from __future__ import annotations

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk
import pandas as pd

from app.analyzer import AnalysisResult, extract_unique_courses, find_students_with_two_courses
from app.excel_reader import get_columns, read_api_data, read_excel_file
from app.exporter import export_results_to_excel

NO_STUDENT_NO_OPTION = "Ogrenci no yok"


class StudentCourseCheckerApp(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")

        self.title("Ders Kesisim Analizi")
        self.geometry("980x700")
        self.minsize(900, 650)

        self.df: pd.DataFrame | None = None
        self.result_df: pd.DataFrame | None = None
        self.selected_file: str = ""

        self.data_source_var = tk.StringVar(value="Excel")
        self.api_url_var = tk.StringVar()
        self.student_no_col_var = tk.StringVar()
        self.first_name_col_var = tk.StringVar()
        self.last_name_col_var = tk.StringVar()
        self.course_col_var = tk.StringVar()
        self.first_course_var = tk.StringVar()
        self.second_course_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Lutfen bir Excel dosyasi secin.")

        self._build_ui()

    def _build_ui(self) -> None:
        top_frame = ctk.CTkFrame(self)
        top_frame.pack(fill="x", padx=16, pady=(16, 10))

        title_label = ctk.CTkLabel(
            top_frame,
            text="Ders Kesisim Analizi",
            font=ctk.CTkFont(size=26, weight="bold"),
        )
        title_label.pack(anchor="w", padx=12, pady=(12, 2))

        description_label = ctk.CTkLabel(
            top_frame,
            text=(
                "Secilen bir dersi alan ogrenciler icinde baska bir dersi de alan "
                "ogrencileri tespit eder."
            ),
            font=ctk.CTkFont(size=14),
        )
        description_label.pack(anchor="w", padx=12, pady=(0, 12))

        controls_frame = ctk.CTkFrame(self)
        controls_frame.pack(fill="x", padx=16, pady=8)

        ctk.CTkLabel(controls_frame, text="Veri Kaynagi").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        source_menu = ctk.CTkComboBox(
            controls_frame,
            variable=self.data_source_var,
            values=["Excel", "API"],
            width=120,
            state="readonly",
            command=self.on_source_change,
        )
        source_menu.grid(row=0, column=1, padx=8, pady=8, sticky="w")

        ctk.CTkLabel(controls_frame, text="API URL").grid(row=0, column=2, padx=8, pady=8, sticky="w")
        self.api_url_entry = ctk.CTkEntry(controls_frame, textvariable=self.api_url_var, placeholder_text="https://...")
        self.api_url_entry.grid(row=0, column=3, padx=8, pady=8, sticky="ew")

        self.api_fetch_button = ctk.CTkButton(
            controls_frame,
            text="API Verisini Cek",
            command=self.fetch_api_data,
            width=150,
        )
        self.api_fetch_button.grid(row=0, column=4, padx=8, pady=8, sticky="e")

        file_button = ctk.CTkButton(
            controls_frame,
            text="Excel Dosyasi Sec",
            command=self.select_excel_file,
            width=180,
        )
        file_button.grid(row=1, column=0, padx=8, pady=8, sticky="w")
        self.file_button = file_button

        self.file_label = ctk.CTkLabel(
            controls_frame,
            text="Dosya secilmedi",
            anchor="w",
        )
        self.file_label.grid(row=1, column=1, columnspan=4, padx=8, pady=8, sticky="ew")
        controls_frame.grid_columnconfigure(3, weight=1)

        mapping_frame = ctk.CTkFrame(self)
        mapping_frame.pack(fill="x", padx=16, pady=8)

        ctk.CTkLabel(mapping_frame, text="Ogrenci No Sutunu").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        ctk.CTkLabel(mapping_frame, text="Ad Sutunu").grid(row=0, column=2, padx=8, pady=8, sticky="w")
        ctk.CTkLabel(mapping_frame, text="Soyad Sutunu").grid(row=0, column=4, padx=8, pady=8, sticky="w")

        self.student_no_menu = ctk.CTkComboBox(mapping_frame, variable=self.student_no_col_var, values=[""])
        self.student_no_menu.grid(row=0, column=1, padx=8, pady=8, sticky="ew")

        self.first_name_menu = ctk.CTkComboBox(mapping_frame, variable=self.first_name_col_var, values=[""])
        self.first_name_menu.grid(row=0, column=3, padx=8, pady=8, sticky="ew")

        self.last_name_menu = ctk.CTkComboBox(mapping_frame, variable=self.last_name_col_var, values=[""])
        self.last_name_menu.grid(row=0, column=5, padx=8, pady=8, sticky="ew")

        ctk.CTkLabel(mapping_frame, text="Ders Adi Sutunu").grid(row=1, column=0, padx=8, pady=8, sticky="w")
        self.course_menu = ctk.CTkComboBox(mapping_frame, variable=self.course_col_var, values=[""])
        self.course_menu.grid(row=1, column=1, padx=8, pady=8, sticky="ew")

        confirm_columns_btn = ctk.CTkButton(
            mapping_frame,
            text="Sutunlari Onayla",
            command=self.confirm_columns,
            width=180,
        )
        confirm_columns_btn.grid(row=1, column=3, padx=8, pady=8, sticky="e")

        mapping_frame.grid_columnconfigure(1, weight=1)
        mapping_frame.grid_columnconfigure(3, weight=1)
        mapping_frame.grid_columnconfigure(5, weight=1)

        courses_frame = ctk.CTkFrame(self)
        courses_frame.pack(fill="x", padx=16, pady=8)

        ctk.CTkLabel(courses_frame, text="Birinci Ders").grid(row=0, column=0, padx=8, pady=8, sticky="w")
        self.first_course_menu = ctk.CTkComboBox(courses_frame, variable=self.first_course_var, values=[""])
        self.first_course_menu.grid(row=0, column=1, padx=8, pady=8, sticky="ew")

        ctk.CTkLabel(courses_frame, text="Ikinci Ders").grid(row=0, column=2, padx=8, pady=8, sticky="w")
        self.second_course_menu = ctk.CTkComboBox(courses_frame, variable=self.second_course_var, values=[""])
        self.second_course_menu.grid(row=0, column=3, padx=8, pady=8, sticky="ew")

        analyze_button = ctk.CTkButton(courses_frame, text="Analiz Et", command=self.run_analysis, width=140)
        analyze_button.grid(row=0, column=4, padx=8, pady=8, sticky="e")

        export_button = ctk.CTkButton(
            courses_frame, text="Sonucu Excel'e Aktar", command=self.export_result, width=180
        )
        export_button.grid(row=0, column=5, padx=8, pady=8, sticky="e")

        clear_button = ctk.CTkButton(courses_frame, text="Temizle", command=self.clear_form, width=100)
        clear_button.grid(row=0, column=6, padx=8, pady=8, sticky="e")

        courses_frame.grid_columnconfigure(1, weight=1)
        courses_frame.grid_columnconfigure(3, weight=1)

        result_frame = ctk.CTkFrame(self)
        result_frame.pack(fill="both", expand=True, padx=16, pady=(8, 16))

        self.result_info_label = ctk.CTkLabel(result_frame, text="", anchor="w")
        self.result_info_label.pack(fill="x", padx=10, pady=(10, 6))

        table_container = tk.Frame(result_frame)
        table_container.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        columns = ("Ogrenci No", "Ad Soyad", "Aldigi Ilgili Dersler")
        self.result_table = ttk.Treeview(table_container, columns=columns, show="headings")
        for col in columns:
            self.result_table.heading(col, text=col)
            self.result_table.column(col, anchor="w", width=220)

        y_scroll = ttk.Scrollbar(table_container, orient="vertical", command=self.result_table.yview)
        x_scroll = ttk.Scrollbar(table_container, orient="horizontal", command=self.result_table.xview)
        self.result_table.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        self.result_table.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right", fill="y")
        x_scroll.pack(side="bottom", fill="x")

        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.pack(fill="x", padx=16, pady=(0, 12))
        footer_frame.grid_columnconfigure(0, weight=1)

        status_label = ctk.CTkLabel(footer_frame, textvariable=self.status_var, anchor="w")
        status_label.grid(row=0, column=0, sticky="w")

        ownership_label = ctk.CTkLabel(
            footer_frame,
            text="Tarik Aykan tarafindan hazirlanmistir | betanova.tech",
            anchor="e",
            text_color="gray40",
        )
        ownership_label.grid(row=0, column=1, sticky="e")
        self.on_source_change(self.data_source_var.get())

    def on_source_change(self, source: str) -> None:
        is_excel = source == "Excel"
        self.file_button.configure(state="normal" if is_excel else "disabled")
        self.api_url_entry.configure(state="disabled" if is_excel else "normal")
        self.api_fetch_button.configure(state="disabled" if is_excel else "normal")
        if is_excel:
            self.status_var.set("Excel modu aktif. Lutfen bir Excel dosyasi secin.")
        else:
            self.status_var.set("API modu aktif. API URL girip veriyi cekin.")

    def select_excel_file(self) -> None:
        if self.data_source_var.get() != "Excel":
            messagebox.showwarning("Veri Kaynagi", "Excel secimi icin veri kaynagini Excel yapin.")
            return

        file_path = filedialog.askopenfilename(
            title="Excel Dosyasi Sec",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")],
        )
        if not file_path:
            return

        try:
            self.df = read_excel_file(file_path)
        except ValueError as exc:
            messagebox.showerror("Dosya Hatasi", str(exc))
            self.status_var.set("Excel dosyasi okunamadi.")
            return

        self.selected_file = file_path
        self.file_label.configure(text=file_path)
        self.status_var.set("Excel dosyasi yuklendi. Sutunlari secip onaylayin.")
        self.result_df = None
        self._clear_table()
        self.result_info_label.configure(text="")

        columns = get_columns(self.df)
        student_no_options = [NO_STUDENT_NO_OPTION] + columns
        self.student_no_menu.configure(values=student_no_options)
        self.first_name_menu.configure(values=columns)
        self.last_name_menu.configure(values=columns)
        self.course_menu.configure(values=columns)

        if columns:
            self.student_no_col_var.set(student_no_options[0])
            self.first_name_col_var.set(columns[min(1, len(columns) - 1)])
            self.last_name_col_var.set(columns[min(2, len(columns) - 1)])
            self.course_col_var.set(columns[min(3, len(columns) - 1)])

    def fetch_api_data(self) -> None:
        if self.data_source_var.get() != "API":
            messagebox.showwarning("Veri Kaynagi", "API cekimi icin veri kaynagini API yapin.")
            return

        api_url = self.api_url_var.get().strip()
        if not api_url:
            messagebox.showwarning("Eksik Bilgi", "Lutfen API URL girin.")
            return

        try:
            self.df = read_api_data(api_url)
        except ValueError as exc:
            messagebox.showerror("API Hatasi", str(exc))
            self.status_var.set("API verisi alinamadi.")
            return

        self.selected_file = f"API: {api_url}"
        self.file_label.configure(text=self.selected_file)
        self.status_var.set("API verisi yuklendi. Sutunlari secip onaylayin.")
        self.result_df = None
        self._clear_table()
        self.result_info_label.configure(text="")

        columns = get_columns(self.df)
        student_no_options = [NO_STUDENT_NO_OPTION] + columns
        self.student_no_menu.configure(values=student_no_options)
        self.first_name_menu.configure(values=columns)
        self.last_name_menu.configure(values=columns)
        self.course_menu.configure(values=columns)

        if columns:
            self.student_no_col_var.set(student_no_options[0])
            self.first_name_col_var.set(columns[min(1, len(columns) - 1)])
            self.last_name_col_var.set(columns[min(2, len(columns) - 1)])
            self.course_col_var.set(columns[min(3, len(columns) - 1)])

    def confirm_columns(self) -> None:
        if self.df is None:
            messagebox.showwarning("Eksik Bilgi", "Lutfen once bir Excel dosyasi secin.")
            return

        student_no_col = self.student_no_col_var.get().strip()
        if student_no_col == NO_STUDENT_NO_OPTION:
            student_no_col = ""
        first_name_col = self.first_name_col_var.get().strip()
        last_name_col = self.last_name_col_var.get().strip()
        course_col = self.course_col_var.get().strip()
        if not first_name_col or not last_name_col or not course_col:
            messagebox.showwarning("Eksik Bilgi", "Lutfen Ad, Soyad ve Ders Adi sutunlarini secin.")
            return

        try:
            courses = extract_unique_courses(self.df, course_col)
        except KeyError:
            messagebox.showerror("Sutun Hatasi", "Secilen sutunlardan biri dosyada bulunamadi.")
            return

        if not courses:
            messagebox.showwarning("Veri Bulunamadi", "Secilen Ders Adi sutununda gecerli ders bulunamadi.")
            return

        self.first_course_menu.configure(values=courses)
        self.second_course_menu.configure(values=courses)
        self.first_course_var.set(courses[0])
        self.second_course_var.set(courses[min(1, len(courses) - 1)])
        self.status_var.set("Sutunlar onaylandi. Dersleri secip analiz yapabilirsiniz.")

    def run_analysis(self) -> None:
        if self.df is None:
            messagebox.showwarning("Eksik Bilgi", "Lutfen once bir Excel dosyasi secin.")
            return

        student_no_col = self.student_no_col_var.get().strip()
        if student_no_col == NO_STUDENT_NO_OPTION:
            student_no_col = ""
        first_name_col = self.first_name_col_var.get().strip()
        last_name_col = self.last_name_col_var.get().strip()
        course_col = self.course_col_var.get().strip()
        first_course = self.first_course_var.get().strip()
        second_course = self.second_course_var.get().strip()

        if not first_name_col or not last_name_col or not course_col:
            messagebox.showwarning("Eksik Bilgi", "Lutfen once sutun secimlerini yapip onaylayin.")
            return
        if not first_course or not second_course:
            messagebox.showwarning("Eksik Bilgi", "Lutfen iki dersi de secin.")
            return
        if first_course == second_course:
            messagebox.showwarning("Gecersiz Secim", "Lutfen farkli iki ders secin.")
            return

        try:
            analysis_result: AnalysisResult = find_students_with_two_courses(
                df=self.df,
                student_no_col=student_no_col,
                first_name_col=first_name_col,
                last_name_col=last_name_col,
                course_col=course_col,
                first_course=first_course,
                second_course=second_course,
            )
        except KeyError:
            messagebox.showerror("Sutun Hatasi", "Secilen sutunlardan biri dosyada bulunamadi.")
            return
        except Exception as exc:
            messagebox.showerror("Analiz Hatasi", f"Analiz sirasinda hata olustu: {exc}")
            return

        self.result_df = analysis_result.result_df
        self._populate_table(self.result_df)

        if analysis_result.total_count == 0:
            text = f"{first_course} dersini alan ogrenciler icinde {second_course} dersini de alan ogrenci bulunamadi."
        else:
            text = (
                f"{first_course} dersini alan ogrenciler icinde {second_course} dersini de alan "
                f"{analysis_result.total_count} ogrenci bulundu."
            )
        self.result_info_label.configure(text=text)
        self.status_var.set("Analiz tamamlandi.")

    def export_result(self) -> None:
        if self.result_df is None:
            messagebox.showwarning("Eksik Bilgi", "Disa aktarmak icin once analiz yapin.")
            return
        if self.result_df.empty:
            messagebox.showwarning("Veri Yok", "Disa aktarilacak sonuc bulunmuyor.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Sonucu Excel'e Kaydet",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx"), ("Excel 97-2003", "*.xls")],
        )
        if not save_path:
            return

        try:
            export_results_to_excel(self.result_df, save_path)
        except ValueError as exc:
            messagebox.showerror("Kaydetme Hatasi", str(exc))
            return

        messagebox.showinfo("Basarili", "Sonuc dosyasi basariyla kaydedildi.")
        self.status_var.set("Sonuc Excel dosyasina aktarildi.")

    def clear_form(self) -> None:
        self.df = None
        self.result_df = None
        self.selected_file = ""
        self.file_label.configure(text="Dosya secilmedi")
        self.student_no_menu.configure(values=[NO_STUDENT_NO_OPTION])
        self.first_name_menu.configure(values=[""])
        self.last_name_menu.configure(values=[""])
        self.course_menu.configure(values=[""])
        self.first_course_menu.configure(values=[""])
        self.second_course_menu.configure(values=[""])
        self.api_url_var.set("")
        self.student_no_col_var.set(NO_STUDENT_NO_OPTION)
        self.first_name_col_var.set("")
        self.last_name_col_var.set("")
        self.course_col_var.set("")
        self.first_course_var.set("")
        self.second_course_var.set("")
        self.result_info_label.configure(text="")
        self._clear_table()
        self.status_var.set("Form temizlendi. Yeni bir veri kaynagi secin.")

    def _clear_table(self) -> None:
        for item in self.result_table.get_children():
            self.result_table.delete(item)

    def _populate_table(self, result_df: pd.DataFrame) -> None:
        self._clear_table()
        if result_df.empty:
            return
        for _, row in result_df.iterrows():
            self.result_table.insert(
                "",
                "end",
                values=(
                    row.get("Ogrenci No", ""),
                    row.get("Ad Soyad", ""),
                    row.get("Aldigi Ilgili Dersler", ""),
                ),
            )
