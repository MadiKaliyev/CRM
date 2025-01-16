import os
import sys
import tkinter as tk
from tkinter import messagebox
from functions import save_to_docs
from functionsPOZH import save_to_docs_POZH
from functionsBIOTRAB import save_to_docs_biot
from functionsBIOT import save_to_docs_biot2

class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("Ввод данных компании и сотрудников")
        self.geometry("1000x700+0+0")

        # Создаем Canvas и вертикальный Scrollbar
        self.canvas = tk.Canvas(self, borderwidth=0)
        self.scrollbar = tk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        # Создаем Frame внутри Canvas
        self.frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.frame, anchor="nw")

        self.frame.bind("<Configure>", lambda event: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        # Привязываем прокрутку мыши к Canvas
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        self.company_name_var = tk.StringVar()

        # Переменные для промышленной безопасности
        self.chislo_chelovek_var = tk.StringVar(value='')  # Пустое значение по умолчанию
        self.chislo_otv_var = tk.StringVar(value='')  # Пустое значение по умолчанию

        # Переменные для пожарной безопасности
        self.chislo_chelovek_var_pozh = tk.StringVar(value='')  # Пустое значение по умолчанию
        self.chislo_otv_var_pozh = tk.StringVar(value='')  # Пустое значение по умолчанию

        # Переменные для охраны труда
        self.chislo_chelovek_var_ohrana = tk.StringVar(value='')  # Пустое значение по умолчанию
        self.chislo_otv_var_ohrana = tk.StringVar(value='')  # Пустое значение по умолчанию

        self.chislo_chelovek_var_ohrana2 = tk.StringVar(value='')  # Пустое значение по умолчанию
        self.chislo_otv_var_ohrana = tk.StringVar(value='')
        padx_value = 25

        tk.Label(self.frame, text="Название компании:").grid(row=0, column=0, padx=padx_value, pady=10, sticky="w")
        tk.Entry(self.frame, textvariable=self.company_name_var, width=50).grid(row=0, column=1, padx=padx_value, pady=10, sticky="ew")

        self.prom_bez_var = tk.BooleanVar(value=False)
        self.pozh_bez_var = tk.BooleanVar(value=False)
        self.ohrana_var = tk.BooleanVar(value=False)
        self.ohrana_var2 = tk.BooleanVar(value=False)

        # Чекбоксы для выбора областей безопасности
        tk.Checkbutton(self.frame, text="Промышленная безопасность", variable=self.prom_bez_var, command=self.toggle_prom_bez_fields).grid(row=1, column=0, padx=padx_value, pady=5, sticky="w")
        tk.Checkbutton(self.frame, text="Пожарная безопасность", variable=self.pozh_bez_var, command=self.toggle_pozh_bez_fields).grid(row=2, column=0, padx=padx_value, pady=5, sticky="w")
        tk.Checkbutton(self.frame, text="Охрана труда для рабочих", variable=self.ohrana_var, command=self.toggle_ohrana_fields).grid(row=3, column=0, padx=padx_value, pady=5, sticky="w")
        tk.Checkbutton(self.frame, text="Охрана труда (специалисты)", variable=self.ohrana_var2, command=self.toggle_ohrana_fields2).grid(row=4, column=0, padx=padx_value, pady=5, sticky="w")

        # Поля для промышленной безопасности (скрыты по умолчанию)
        self.prom_fields_frame = tk.Frame(self.frame)
        self.prom_fields_frame.grid(row=5, column=0, columnspan=2, padx=padx_value, pady=10, sticky="nsew")
        self.prom_fields_frame.grid_remove()

        tk.Label(self.prom_fields_frame, text="Количество работников (Пром. Безопасность):").grid(row=0, column=0, padx=padx_value, pady=5, sticky="w")
        tk.Entry(self.prom_fields_frame, textvariable=self.chislo_chelovek_var, width=10).grid(row=0, column=1, padx=padx_value, pady=5, sticky="w")

        tk.Label(self.prom_fields_frame, text="Количество ответственных (Пром. Безопасность):").grid(row=1, column=0, padx=padx_value, pady=5, sticky="w")
        tk.Entry(self.prom_fields_frame, textvariable=self.chislo_otv_var, width=10).grid(row=1, column=1, padx=padx_value, pady=5, sticky="w")

        # Поля для пожарной безопасности (скрыты по умолчанию)
        self.pozh_fields_frame = tk.Frame(self.frame)
        self.pozh_fields_frame.grid(row=6, column=0, columnspan=2, padx=padx_value, pady=10, sticky="nsew")
        self.pozh_fields_frame.grid_remove()

        tk.Label(self.pozh_fields_frame, text="Количество работников (Пожарная Безопасность):").grid(row=0, column=0, padx=padx_value, pady=5, sticky="w")
        tk.Entry(self.pozh_fields_frame, textvariable=self.chislo_chelovek_var_pozh, width=10).grid(row=0, column=1, padx=padx_value, pady=5, sticky="w")

        tk.Label(self.pozh_fields_frame, text="Количество ответственных (Пожарная Безопасность):").grid(row=1, column=0, padx=padx_value, pady=5, sticky="w")
        tk.Entry(self.pozh_fields_frame, textvariable=self.chislo_otv_var_pozh, width=10).grid(row=1, column=1, padx=padx_value, pady=5, sticky="w")

        # Поля для охраны труда (скрыты по умолчанию)
        self.ohrana_fields_frame = tk.Frame(self.frame)
        self.ohrana_fields_frame.grid(row=7, column=0, columnspan=2, padx=padx_value, pady=10, sticky="nsew")
        self.ohrana_fields_frame.grid_remove()

        tk.Label(self.ohrana_fields_frame, text="Количество работников (Охрана труда):").grid(row=0, column=0, padx=padx_value, pady=5, sticky="w")
        tk.Entry(self.ohrana_fields_frame, textvariable=self.chislo_chelovek_var_ohrana, width=10).grid(row=0, column=1, padx=padx_value, pady=5, sticky="w")

        self.ohrana_fields_frame2 = tk.Frame(self.frame)
        self.ohrana_fields_frame2.grid(row=8, column=0, columnspan=2, padx=padx_value, pady=10, sticky="nsew")
        self.ohrana_fields_frame2.grid_remove()

        tk.Label(self.ohrana_fields_frame2, text="Количество Специалистов (Охрана труда):").grid(row=0, column=0, padx=padx_value, pady=5, sticky="w")
        tk.Entry(self.ohrana_fields_frame2, textvariable=self.chislo_chelovek_var_ohrana2, width=10).grid(row=0, column=1, padx=padx_value, pady=5, sticky="w")

        # Кнопка для подтверждения и запуска создания полей ввода
        tk.Button(self.frame, text="Подтвердить", command=self.check_and_proceed).grid(row=9, column=1, padx=10, pady=10, sticky="w")

        self.entries_frame = tk.Frame(self.frame)
        self.entries_frame.grid(row=10, column=0, columnspan=2, padx=padx_value, pady=10, sticky="nsew")

        self.entries_frame1 = tk.Frame(self.frame)
        self.entries_frame1.grid(row=11, column=0, columnspan=2, padx=padx_value, pady=10, sticky="nsew")

        self.entries_frame2 = tk.Frame(self.frame)
        self.entries_frame2.grid(row=12, column=0, columnspan=2, padx=padx_value, pady=10, sticky="nsew")

        self.entries_frame3 = tk.Frame(self.frame)
        self.entries_frame3.grid(row=13, column=0, columnspan=2, padx=padx_value, pady=10, sticky="nsew")

        # Кнопка сброса полей
        tk.Button(self.frame, text="Сброс", command=self.reset_fields).grid(row=9, column=0, padx=10, pady=10, sticky="w")

    def toggle_prom_bez_fields(self):
        if self.prom_bez_var.get():
            self.prom_fields_frame.grid()
        else:
            self.prom_fields_frame.grid_remove()

    def toggle_pozh_bez_fields(self):
        if self.pozh_bez_var.get():
            self.pozh_fields_frame.grid()
        else:
            self.pozh_fields_frame.grid_remove()

    def toggle_ohrana_fields(self):
        if self.ohrana_var.get():
            self.ohrana_fields_frame.grid()
        else:
            self.ohrana_fields_frame.grid_remove()

    def toggle_ohrana_fields2(self):
        if self.ohrana_var2.get():
            self.ohrana_fields_frame2.grid()
        else:
            self.ohrana_fields_frame2.grid_remove()

    def create_entry_with_menu(self, parent, width=20):
        entry = tk.Entry(parent, width=width)
        menu = tk.Menu(entry, tearoff=0)
        menu.add_command(label="Вставить", command=lambda: entry.event_generate("<<Paste>>"))
        menu.add_command(label="Копировать", command=lambda: entry.event_generate("<<Copy>>"))
        menu.add_command(label="Вырезать", command=lambda: entry.event_generate("<<Cut>>"))

        entry.bind("<Button-3>", lambda event: menu.post(event.x_root, event.y_root))

        return entry

    def check_and_proceed(self):
        if self.prom_bez_var.get():
            self.create_entries()
        if self.pozh_bez_var.get():
            self.create_entries1()
        if self.ohrana_var.get():
            self.create_entries2()
        if self.ohrana_var2.get():
            self.create_entries3()

    def create_entries(self):
        for widget in self.entries_frame.winfo_children():
            widget.destroy()

        try:
            num_workers = int(self.chislo_chelovek_var.get()) if self.chislo_chelovek_var.get() else 0
            num_responsibles = int(self.chislo_otv_var.get()) if self.chislo_otv_var.get() else 0
        except ValueError:
            messagebox.showerror("Ошибка ввода", "Пожалуйста, введите числовое значение.")
            return

        total_people = num_workers + num_responsibles

        if total_people == 0:
            messagebox.showinfo("Внимание", "Пожалуйста, введите количество работников или ответственных лиц.")
            return

        self.entries = []


        tk.Label(self.entries_frame, text="ФИО", width=30).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        tk.Label(self.entries_frame, text="Должность", width=20).grid(row=0, column=2, padx=5, pady=5, sticky="w")
        tk.Label(self.entries_frame, text="Образование").grid(row=0, column=3, padx=5, pady=5, sticky="w")
        tk.Label(self.entries_frame, text="Ответственность").grid(row=0, column=4, padx=5, pady=5, sticky="w")

        for i in range(num_workers):
            tk.Label(self.entries_frame, text=f"Работник {i + 1}:").grid(row=i + 1, column=0, sticky="w")
            fio_entry = self.create_entry_with_menu(self.entries_frame, width=20)
            fio_entry.grid(row=i + 1, column=1, sticky="ew")
            position_entry = self.create_entry_with_menu(self.entries_frame, width=20)
            position_entry.grid(row=i + 1, column=2, sticky="ew")
            education_entry = self.create_entry_with_menu(self.entries_frame, width=20)
            education_entry.grid(row=i + 1, column=3, sticky="ew")
            responsibility_entry = self.create_entry_with_menu(self.entries_frame, width=20)
            responsibility_entry.grid(row=i + 1, column=4, sticky="ew")

            self.entries.append([fio_entry, position_entry, education_entry, responsibility_entry])

        for i in range(num_responsibles):
            tk.Label(self.entries_frame, text=f"Ответственный {i + 1}:").grid(row=num_workers + i + 1, column=0, sticky="w")
            fio_entry = self.create_entry_with_menu(self.entries_frame, width=20)
            fio_entry.grid(row=num_workers + i + 1, column=1, sticky="ew")
            position_entry = self.create_entry_with_menu(self.entries_frame, width=20)
            position_entry.grid(row=num_workers + i + 1, column=2, sticky="ew")
            education_entry = self.create_entry_with_menu(self.entries_frame, width=20)
            education_entry.grid(row=num_workers + i + 1, column=3, sticky="ew")
            responsibility_entry = self.create_entry_with_menu(self.entries_frame, width=20)
            responsibility_entry.grid(row=num_workers + i + 1, column=4, sticky="ew")

            self.entries.append([fio_entry, position_entry, education_entry, responsibility_entry])

        tk.Button(self.entries_frame, text="Сохранить", command=self.save_data).grid(row=total_people + 1, column=0, columnspan=5, pady=10)

    def create_entries1(self):
        for widget in self.entries_frame1.winfo_children():
            widget.destroy()

        try:
            num_workers = int(self.chislo_chelovek_var_pozh.get()) if self.chislo_chelovek_var_pozh.get() else 0
            num_responsibles = int(self.chislo_otv_var_pozh.get()) if self.chislo_otv_var_pozh.get() else 0
        except ValueError:
            messagebox.showerror("Ошибка ввода", "Пожалуйста, введите числовое значение.")
            return

        total_people = num_workers + num_responsibles

        if total_people == 0:
            messagebox.showinfo("Внимание", "Пожалуйста, введите количество работников или ответственных лиц.")
            return

        self.entries1 = []

        tk.Label(self.entries_frame1, text="ФИО", width=30).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        tk.Label(self.entries_frame1, text="Должность", width=20).grid(row=0, column=2, padx=5, pady=5, sticky="w")

        for i in range(num_workers):
            tk.Label(self.entries_frame1, text=f"Работник {i + 1}:").grid(row=i + 1, column=0, sticky="w")
            fio_entry = self.create_entry_with_menu(self.entries_frame1, width=20)
            fio_entry.grid(row=i + 1, column=1, sticky="ew")
            position_entry = self.create_entry_with_menu(self.entries_frame1, width=20)
            position_entry.grid(row=i + 1, column=2, sticky="ew")

            self.entries1.append([fio_entry, position_entry])

        for i in range(num_responsibles):
            tk.Label(self.entries_frame1, text=f"Ответственный {i + 1}:").grid(row=num_workers + i + 1, column=0, sticky="w")
            fio_entry = self.create_entry_with_menu(self.entries_frame1, width=20)
            fio_entry.grid(row=num_workers + i + 1, column=1, sticky="ew")
            position_entry = self.create_entry_with_menu(self.entries_frame1, width=20)
            position_entry.grid(row=num_workers + i + 1, column=2, sticky="ew")

            self.entries1.append([fio_entry, position_entry])

        tk.Button(self.entries_frame1, text="Сохранить", command=self.save_data1).grid(row=total_people + 1, column=0, columnspan=5, pady=10)

    def create_entries2(self):
        for widget in self.entries_frame2.winfo_children():
            widget.destroy()

        try:
            num_workers = int(self.chislo_chelovek_var_ohrana.get()) if self.chislo_chelovek_var_ohrana.get() else 0
            num_responsibles = int(self.chislo_otv_var_ohrana.get()) if self.chislo_otv_var_ohrana.get() else 0
        except ValueError:
            messagebox.showerror("Ошибка ввода", "Пожалуйста, введите числовое значение.")
            return

        total_people = num_workers + num_responsibles

        if total_people == 0:
            messagebox.showinfo("Внимание", "Пожалуйста, введите количество работников или ответственных лиц.")
            return

        self.entries2 = []

        tk.Label(self.entries_frame2, text="ФИО", width=30).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        tk.Label(self.entries_frame2, text="Должность", width=20).grid(row=0, column=2, padx=5, pady=5, sticky="w")

        for i in range(num_workers):
            tk.Label(self.entries_frame2, text=f"Работник {i + 1}:").grid(row=i + 1, column=0, sticky="w")
            fio_entry = self.create_entry_with_menu(self.entries_frame2, width=20)
            fio_entry.grid(row=i + 1, column=1, sticky="ew")
            position_entry = self.create_entry_with_menu(self.entries_frame2, width=20)
            position_entry.grid(row=i + 1, column=2, sticky="ew")

            self.entries2.append([fio_entry, position_entry])

        for i in range(num_responsibles):
            tk.Label(self.entries_frame2, text=f"Ответственный {i + 1}:").grid(row=num_workers + i + 1, column=0, sticky="w")
            fio_entry = self.create_entry_with_menu(self.entries_frame2, width=20)
            fio_entry.grid(row=num_workers + i + 1, column=1, sticky="ew")
            position_entry = self.create_entry_with_menu(self.entries_frame2, width=20)
            position_entry.grid(row=num_workers + i + 1, column=2, sticky="ew")

            self.entries2.append([fio_entry, position_entry])

        tk.Button(self.entries_frame2, text="Сохранить", command=self.save_data2).grid(row=total_people + 1, column=0, columnspan=5, pady=10)

    def create_entries3(self):
        for widget in self.entries_frame3.winfo_children():
            widget.destroy()

        try:
            num_workers = int(self.chislo_chelovek_var_ohrana2.get()) if self.chislo_chelovek_var_ohrana2.get() else 0
        except ValueError:
            messagebox.showerror("Ошибка ввода", "Пожалуйста, введите числовое значение.")
            return

        if num_workers == 0:
            messagebox.showinfo("Внимание", "Пожалуйста, введите количество работников.")
            return

        self.entries3 = []

        tk.Label(self.entries_frame3, text="ФИО", width=30).grid(row=0, column=1, padx=5, pady=5, sticky="w")
        tk.Label(self.entries_frame3, text="Должность", width=20).grid(row=0, column=2, padx=5, pady=5, sticky="w")

        for i in range(num_workers):
            tk.Label(self.entries_frame3, text=f"Работник {i + 1}:").grid(row=i + 1, column=0, sticky="w")
            fio_entry = self.create_entry_with_menu(self.entries_frame3, width=20)
            fio_entry.grid(row=i + 1, column=1, sticky="ew")
            position_entry = self.create_entry_with_menu(self.entries_frame3, width=20)
            position_entry.grid(row=i + 1, column=2, sticky="ew")

            self.entries3.append([fio_entry, position_entry])

        tk.Button(self.entries_frame3, text="Сохранить", command=self.save_data3).grid(row=num_workers + 1, column=0, columnspan=5, pady=10)

    def reset_fields(self):
        """Сброс всех полей и переменных в начальное состояние."""
        self.company_name_var.set('')
        self.chislo_chelovek_var.set('')
        self.chislo_otv_var.set('')
        self.chislo_chelovek_var_pozh.set('')
        self.chislo_otv_var_pozh.set('')
        self.chislo_chelovek_var_ohrana.set('')
        self.chislo_chelovek_var_ohrana2.set('')
        self.chislo_otv_var_ohrana.set('')

        self.prom_bez_var.set(False)
        self.pozh_bez_var.set(False)
        self.ohrana_var.set(False)
        self.ohrana_var2.set(False)

        for frame in [self.entries_frame, self.entries_frame1, self.entries_frame2, self.entries_frame3]:
            for widget in frame.winfo_children():
                widget.destroy()

        self.prom_fields_frame.grid_remove()
        self.pozh_fields_frame.grid_remove()
        self.ohrana_fields_frame.grid_remove()
        self.ohrana_fields_frame2.grid_remove()

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def paste_data(self, event=None):
        clipboard_data = self.clipboard_get()
        rows = clipboard_data.strip().split('\n')

        for r_idx, row in enumerate(rows):
            if r_idx >= len(self.entries):
                break
            cells = row.split('\t')
            for c_idx, cell in enumerate(cells):
                if c_idx >= len(self.entries[r_idx]):
                    break
                # Удаляем текущий текст только если ячейка пустая
                if not self.entries[r_idx][c_idx].get().strip():
                    self.entries[r_idx][c_idx].delete(0, tk.END)
                    self.entries[r_idx][c_idx].insert(0, cell)

    def save_data(self):
        chislo_chelovek = self.chislo_chelovek_var.get().strip()
        chislo_chelovek = int(chislo_chelovek) if chislo_chelovek.isdigit() else 0

        self.people_data = []
        for entries in self.entries:
            fio = entries[0].get().strip()
            position = entries[1].get().strip()
            education = entries[2].get().strip()
            responsibility = entries[3].get().strip()

            self.people_data.append((fio, position, education, responsibility))

        save_to_docs(self.company_name_var.get().strip(), self.people_data, chislo_chelovek)

    def save_data1(self):
        chislo_chelovek_pozh = self.chislo_chelovek_var_pozh.get().strip()
        chislo_chelovek_pozh = int(chislo_chelovek_pozh) if chislo_chelovek_pozh.isdigit() else 0

        self.people_data = []
        for entries in self.entries1:
            fio = entries[0].get().strip()
            position = entries[1].get().strip()

            self.people_data.append((fio, position))

        save_to_docs_POZH(self.company_name_var.get().strip(), self.people_data, chislo_chelovek_pozh)

    def save_data2(self):
        chislo_chelovek_ohrana = self.chislo_chelovek_var_ohrana.get().strip()
        chislo_chelovek_ohrana = int(chislo_chelovek_ohrana) if chislo_chelovek_ohrana.isdigit() else 0

        self.people_data = []
        for entries in self.entries2:
            fio = entries[0].get().strip()
            position = entries[1].get().strip()

            self.people_data.append((fio, position))

        save_to_docs_biot(self.company_name_var.get().strip(), self.people_data, chislo_chelovek_ohrana)

    def save_data3(self):
        chislo_chelovek_ohrana2 = self.chislo_chelovek_var_ohrana2.get().strip()
        chislo_chelovek_ohrana2 = int(chislo_chelovek_ohrana2) if chislo_chelovek_ohrana2.isdigit() else 0

        self.people_data = []
        for entries in self.entries3:
            fio = entries[0].get().strip()
            position = entries[1].get().strip()

            self.people_data.append((fio, position))

        save_to_docs_biot2(self.company_name_var.get().strip(), self.people_data, chislo_chelovek_ohrana2)


if __name__ == "__main__":
    app = App()
    app.mainloop()
