"""
Парсер вакансий /
"""
import os
import sys
import requests
import pandas as pd
from datetime import datetime, timedelta
import time
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QLabel,
                               QLineEdit, QPushButton, QTextEdit)
from PySide6.QtCore import QThread, Signal
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo


class ParserThread(QThread):
    update_log = Signal(str)
    finished = Signal()

    def __init__(self, profession):
        super().__init__()
        self.profession = profession
        self.running = True

    def run(self):
        try:
            end_date = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
            start_date = end_date - timedelta(days=30)

            all_vacancies = []

            for date_from, date_to in self.date_range(start_date, end_date):
                if not self.running:
                    return

                date_from_str = date_from.isoformat()
                date_to_str = date_to.isoformat()

                self.update_log.emit(f"Парсинг за период: {date_from.date()} - {date_to.date()}")

                vacancies = self.get_vacancies(
                    text=self.profession,
                    date_from=date_from_str,
                    date_to=date_to_str
                )

                all_vacancies.extend(vacancies)
                self.update_log.emit(f"Найдено вакансий: {len(vacancies)}")
                time.sleep(1)

            if all_vacancies:
                self.save_to_excel(all_vacancies)
                self.update_log.emit("\nСохранение результатов...")
                self.update_log.emit(f"Всего собрано вакансий: {len(all_vacancies)}")
            else:
                self.update_log.emit("\nНе найдено подходящих вакансий")

        except Exception as e:
            self.update_log.emit(f"Ошибка: {str(e)}")
        finally:
            self.finished.emit()

    def get_vacancies(self, text, date_from, date_to):
        base_url = "https://api.hh.ru/vacancies"
        params = {
            "text": text,
            "area": 113,  # Поиск только по России
            "date_from": date_from,
            "date_to": date_to,
            "per_page": 100,
            "page": 0
        }

        vacancies = []

        while self.running:
            try:
                response = requests.get(base_url, params=params)
                response.raise_for_status()
                data = response.json()

                if "items" not in data:
                    break

                vacancies.extend(data["items"])

                pages = data["pages"]
                params["page"] += 1

                if params["page"] >= pages:
                    break

                time.sleep(0.5)

            except Exception as e:
                self.update_log.emit(f"Ошибка запроса: {str(e)}")
                break

        return vacancies

    def date_range(self, start_date, end_date, delta=timedelta(days=7)):
        current_date = start_date
        while current_date < end_date:
            next_date = min(current_date + delta, end_date)
            yield current_date, next_date
            current_date = next_date

    def save_to_excel(self, vacancies):
        data = []
        for item in vacancies:
            salary = item.get("salary")
            employer = item.get("employer", {})

            row = {
                "Название": item.get("name"),
                "Компания": employer.get("name"),
                "Зарплата от": salary.get("from") if salary else None,
                "Зарплата до": salary.get("to") if salary else None,
                "Валюта": salary.get("currency") if salary else None,
                "Регион": item.get("area", {}).get("name"),
                "Дата публикации": item.get("published_at"),
                "Ссылка": item.get("alternate_url")
            }
            data.append(row)

        if not data:
            self.update_log.emit("Нет данных для сохранения")
            return

        safe_name = "".join([c if c.isalnum() or c in ('_', '-') else '_'
                             for c in self.profession]).rstrip('_')
        file_name = f"{safe_name}_вакансии.xlsx"
        file_path = os.path.join(os.getcwd(), file_name)

        try:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                pd.DataFrame(data).to_excel(writer, index=False, sheet_name='Вакансии')

            wb = load_workbook(file_path)
            ws = wb.active

            # Добавляем таблицу с фильтром (без автоподбора ширины)
            tab = Table(displayName="VacanciesTable", ref=f"A1:H{len(data) + 1}")
            style = TableStyleInfo(
                name="TableStyleMedium9",
                showFirstColumn=False,
                showLastColumn=False,
                showRowStripes=True,
                showColumnStripes=False
            )
            tab.tableStyleInfo = style
            ws.add_table(tab)

            wb.save(file_path)
            self.update_log.emit(f"Файл успешно сохранен: {file_path}")

        except Exception as e:
            self.update_log.emit(f"Ошибка при сохранении файла: {str(e)}")

    def stop(self):
        self.running = False


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()
        self.parser_thread = None

    def initUI(self):
        self.setWindowTitle("Парсер вакансий hh.ru")
        self.setGeometry(300, 300, 600, 400)

        layout = QVBoxLayout()

        self.label = QLabel("Введите профессию для поиска:")
        layout.addWidget(self.label)

        self.profession_input = QLineEdit()
        self.profession_input.setPlaceholderText("Например: сварщик")
        layout.addWidget(self.profession_input)

        self.start_btn = QPushButton("Начать поиск")
        self.start_btn.clicked.connect(self.start_parsing)
        layout.addWidget(self.start_btn)

        self.log_output = QTextEdit()
        self.log_output.setReadOnly(True)
        layout.addWidget(self.log_output)

        self.setLayout(layout)

    def start_parsing(self):
        profession = self.profession_input.text().strip()
        if not profession:
            self.log_output.append("Введите название профессии!")
            return

        if self.parser_thread and self.parser_thread.isRunning():
            self.log_output.append("Парсинг уже выполняется!")
            return

        self.log_output.clear()
        self.parser_thread = ParserThread(profession)
        self.parser_thread.update_log.connect(self.log_output.append)
        self.parser_thread.finished.connect(self.on_finished)
        self.parser_thread.start()
        self.start_btn.setEnabled(False)
        self.start_btn.setText("Поиск...")

    def on_finished(self):
        self.start_btn.setEnabled(True)
        self.start_btn.setText("Начать поиск")
        self.parser_thread = None

    def closeEvent(self, event):
        if self.parser_thread and self.parser_thread.isRunning():
            self.parser_thread.stop()
            self.parser_thread.quit()
            self.parser_thread.wait()
        event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
