"""Десктопный интерфейс лабораторной №2 на Tkinter."""
from __future__ import annotations

import logging
import os
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from ticket_core import DataError, ResultJournal, load_students, load_tickets, students_for_group

BASE = Path(__file__).resolve().parent


class TicketApp:
    def __init__(self, root: tk.Tk, directory: Path = BASE):
        self.root, self.directory = root, directory
        self.groups, self.tickets, self.current_students = {}, {}, []
        self.journal = ResultJournal(directory / "results.xlsx")
        self.result_window = None
        root.title("Генератор экзаменационных билетов • Лабораторные 1–2")
        root.geometry("700x480")
        root.minsize(620, 440)
        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("TLabel", font=("Segoe UI", 11))
        style.configure("TButton", font=("Segoe UI", 10), padding=8)
        style.configure("Title.TLabel", font=("Segoe UI", 21, "bold"))
        frame = ttk.Frame(root, padding=28)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Экзаменационные билеты", style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame, text="Выберите группу и студента, чтобы получить билет.").pack(anchor="w", pady=(8, 20))
        ttk.Label(frame, text="Группа").pack(anchor="w")
        self.group = ttk.Combobox(frame, state="readonly", font=("Segoe UI", 12))
        self.group.pack(fill="x", pady=(5, 12))
        self.group.bind("<<ComboboxSelected>>", self.select_group)
        ttk.Label(frame, text="Студент").pack(anchor="w")
        self.student = ttk.Combobox(frame, state="readonly", font=("Segoe UI", 12))
        self.student.pack(fill="x", pady=(5, 16))
        self.student.bind("<<ComboboxSelected>>", self.update_button)
        self.generate_button = ttk.Button(frame, text="Сгенерировать билет", command=self.generate, state="disabled")
        self.generate_button.pack(fill="x")
        self.status = tk.StringVar(value="Загрузка источников…")
        ttk.Label(frame, textvariable=self.status, wraplength=620).pack(anchor="w", pady=(12, 8))
        actions = ttk.Frame(frame)
        actions.pack(fill="x", side="bottom")
        ttk.Button(actions, text="Обновить файлы", command=self.reload).pack(side="left")
        ttk.Button(actions, text="Создать учебные примеры", command=self.create_examples).pack(side="left", padx=6)
        ttk.Button(actions, text="Открыть журнал", command=self.open_results).pack(side="right")
        root.after_idle(self.reload)

    def retry(self, operation):
        while True:
            try:
                return True, operation()
            except (OSError, DataError) as error:
                logging.getLogger(__name__).warning("Ошибка операции: %s", error)
                if isinstance(error, PermissionError):
                    detail = "Файл занят или недоступен для записи. Закройте его в Excel/Word и проверьте права на папку."
                else:
                    detail = str(error)
                if not messagebox.askretrycancel("Не удалось выполнить операцию", detail, parent=self.root):
                    return False, None

    def reload(self):
        self.groups, self.tickets, self.current_students = {}, {}, []
        self.group.set("")
        self.student.set("")
        self.group.configure(values=())
        self.student.configure(values=())
        self.update_button()

        def load():
            groups = load_students(self.directory / "students.xlsx")
            tickets = load_tickets(self.directory / "tickets.docx")
            self.journal.ensure_exists()
            return groups, tickets

        success, data = self.retry(load)
        if success:
            self.groups, self.tickets = data
            self.group.configure(values=list(self.groups))
            self.status.set(f"Групп: {len(self.groups)} · Корректных билетов: {len(self.tickets)}")
        else:
            self.status.set("Источники не загружены. Добавьте файлы и нажмите «Обновить файлы» либо создайте учебные примеры.")

    def select_group(self, _event=None):
        self.student.set("")
        self.current_students = students_for_group(self.groups, self.group.get())
        self.student.configure(values=[student.label for student in self.current_students])
        self.status.set("Выберите студента." if self.current_students else "В этой группе нет студентов.")
        self.update_button()

    def update_button(self, _event=None):
        enabled = bool(self.tickets and self.group.get() in self.groups
                       and 0 <= self.student.current() < len(self.current_students))
        self.generate_button.configure(state="normal" if enabled else "disabled")

    def generate(self):
        index = self.student.current()
        if not self.tickets or not 0 <= index < len(self.current_students):
            return
        success, assignment = self.retry(lambda: self.journal.prepare(self.current_students[index], self.tickets))
        if not success:
            return
        # Повтор сохранения замыкает один Assignment: номер и время не меняются.
        success, _ = self.retry(lambda: self.journal.append(assignment))
        if success:
            self.status.set("Запись сохранена в results.xlsx. " + ("Повторная выдача." if assignment.repeated else "Первая выдача."))
            self.show_result(assignment)

    def show_result(self, assignment):
        window = tk.Toplevel(self.root)
        self.result_window = window
        window.title(f"Билет № {assignment.ticket.number}")
        window.geometry("680x450")
        window.minsize(540, 360)
        window.transient(self.root)
        frame = ttk.Frame(window, padding=26)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text=f"Ваш билет № {assignment.ticket.number}", style="Title.TLabel").pack(anchor="w")
        ttk.Label(frame, text=f"{assignment.student.group} · {assignment.student.label}").pack(anchor="w", pady=(8, 16))
        body = tk.Text(frame, wrap="word", font=("Segoe UI", 12), height=8, relief="flat", padx=12, pady=12)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=body.yview)
        body.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        body.pack(fill="both", expand=True)
        body.insert("1.0", "\n\n".join(f"{i}. {question}" for i, question in enumerate(assignment.ticket.questions, 1)))
        body.configure(state="disabled")
        ttk.Button(frame, text="Следующий студент · ESC", command=self.close_result).pack(fill="x", pady=(16, 0))
        window.protocol("WM_DELETE_WINDOW", self.close_result)
        window.bind("<Escape>", lambda _event: self.close_result())
        window.grab_set()
        window.focus_set()

    def close_result(self):
        if self.result_window is not None:
            self.result_window.grab_release()
            self.result_window.destroy()
            self.result_window = None
        self.student.set("")
        self.update_button()
        self.student.focus_set()

    def create_examples(self):
        from create_examples import create_examples
        success, created = self.retry(lambda: create_examples(self.directory))
        if success:
            messagebox.showinfo("Учебные примеры", "Созданы: " + ", ".join(created) if created else
                                "Входные файлы уже существуют. Они не изменены.", parent=self.root)
            self.reload()

    def open_results(self):
        success, _ = self.retry(self.journal.ensure_exists)
        if success:
            self.retry(lambda: os.startfile(self.journal.path))


def main():
    logging.basicConfig(filename=BASE / "app.log", encoding="utf-8", level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s")
    root = tk.Tk()
    TicketApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
