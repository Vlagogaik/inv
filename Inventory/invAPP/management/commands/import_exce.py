import openpyxl
from django.core.management.base import BaseCommand
from openpyxl.reader.excel import load_workbook
# Странно но факт тут должна быть ошибка(от компилятора)
from invAPP.models import Workers, Inventory1C, Characteristics, Relation3, Relation4, TypeOfEquipment

from datetime import datetime


class Command(BaseCommand):
    help = 'Импортирует данные из двух Excel-файлов в базы данных Django'

    def add_arguments(self, parser):
        parser.add_argument('excel_files', nargs=2, type=str, help="Пути к двум файлам Excel.")

    def handle(self, *args, **options):
        file1, file2 = options['excel_files']
        try:
            data1 = self.load_excel(file1)  # 1C
            data2 = self.load_excel(file2)  # 101.34 на 31.12.22 v2

            self.stdout.write(f"Обработка данных из файлов {file1} и {file2}")

            types_of_equipment = [
                "иное", "ПК", "моноблок", "Ноутбук", "Сервер", "Планшет",
                "КПК, смартфон", "Моб.телефон", "Факс", "ИБП", "Монитор",
                "Телевизор", "Видеомагнитофон", "Видеокамера аналог.", "Звуковое оборудование", "Микшерный пульт",
                "Проектор", "Интерактивная доска", "Интерактивная панель",
                "Принтер", "МФУ", "Копир", "Сканер", "Документ-проектор",
                "АТС", "Видеорегистратор", "СКУД", "Коммутатор, сетевое"
            ]
            types_of_equipment = [type_name.lower() for type_name in types_of_equipment]
            for type_name in types_of_equipment:
                TypeOfEquipment.objects.get_or_create(type_name=type_name)

            type_mapping = {
                "ПК 2021".lower(): "ПК",
                "ПК отечественный".lower(): "ПК".lower(),
                "Ноутбук отечественный".lower(): "Ноутбук".lower(),
                "Моноблок отечественный".lower(): "Моноблок".lower(),
                "Комп 2022".lower(): "ПК".lower(),
                "комп 2021".lower(): "ПК".lower(),
                "Комп. не более 5 лет".lower(): "ПК".lower(),
                "Презен. об. не более 5 лет".lower(): "Проектор".lower(),
                "Пр., мфу, коп., скан. не более 5 лет".lower(): "Принтер".lower(),
            }

            inventory_map = {}
            self.stdout.write("Идем по второму файлу")
            for row in data2:
                inv_number = row.get('Инвентарный номер')
                if inv_number:
                    inventory_map[inv_number] = row
                fio = row.get('ФИО')
                department = row.get('Подразделение')
                if not fio or not department:
                    fio = "Неизвестно"
                    department = "Неизвестно"

                worker, _ = Workers.objects.get_or_create(
                    full_name=fio,
                    department=department
                )

                initial_cost = row.get('Стоимость первоначальна')
                if not initial_cost:
                    initial_cost = 0

                type_obj = None
                for col, value in row.items():
                    col = col.strip().lower()
                    if col in types_of_equipment:
                        if str(value).strip() == "1":
                            type_name = type_mapping.get(col, col)
                            type_obj, _ = TypeOfEquipment.objects.get_or_create(type_name=type_name)
                            break

                # self.stdout.write(f"Обработка данных из файлов {fio} и {department} и {type_obj}")
                if type_obj is None:
                    type_obj, _ = TypeOfEquipment.objects.get_or_create(type_name="Иное")

                date_acceptance = row.get('Дата принятия к учету')
                date_decommission = row.get('Дата списания')
                if not date_decommission:
                    date_decommission = ""
                if not row.get('ОС'):
                    continue

                Inventory1C.objects.get_or_create(
                    inventory_decimal=inv_number,
                    defaults={
                        'id_workers': worker,
                        'accounting_name': row.get('ОС'),
                        'date_acceptance_accounting': date_acceptance,
                        'initial_cost': initial_cost,
                        'id_type': type_obj,
                        "real_name": "",
                        "date_of_decommission": date_decommission
                    }
                )

            self.stdout.write("Идем по первому файлу")
            for row in data1:
                inv_number = row.get('Инвентарный номер')
                if not inv_number:
                    continue

                fio = inventory_map.get(inv_number, {}).get('ФИО')
                department = inventory_map.get(inv_number, {}).get('Подразделение')

                if not fio or not department:
                    fio = "Неизвестно"
                    department = "Неизвестно"

                worker, _ = Workers.objects.get_or_create(
                    full_name=fio,
                    department=department
                )
                initial_cost = row.get('Стоимость первоначальна')
                if not initial_cost:
                    initial_cost = 0

                type_obj = None
                for type_name in types_of_equipment:
                    if type_name in row and str(row[type_name]).strip():
                        type_obj, _ = TypeOfEquipment.objects.get_or_create(type_name=type_name)
                        break

                if type_obj is None:
                    type_obj, _ = TypeOfEquipment.objects.get_or_create(type_name="Иное")

                date_acceptance = row.get('Дата принятия к учету')

                if not row.get('Основное средство'):
                    continue
                # self.stdout.write(f"Обработка данных из файлов {fio} и {department} и {type_obj}")
                Inventory1C.objects.get_or_create(
                    inventory_decimal=inv_number,
                    defaults={
                        'id_workers': worker,
                        'accounting_name': row.get('Основное средство'),
                        'date_acceptance_accounting': date_acceptance,
                        'initial_cost': initial_cost,
                        'id_type': type_obj,
                        "real_name": "",
                        "date_of_decommission": ""
                    }
                )

            self.stdout.write(self.style.SUCCESS("Импорт данных завершён."))

        except Exception as e:
            self.stderr.write(f"Ошибка: {e}")

    def load_excel(self, file):
        wb = openpyxl.load_workbook(file)
        sheet = wb.active
        headers = [cell.value for cell in sheet[1]]
        data = [
            {headers[idx]: cell for idx, cell in enumerate(row)}
            for row in sheet.iter_rows(min_row=2, values_only=True)
        ]
        return data

    def get_merged_cell_value(self, row, column_name):
        column_idx = list(row.keys()).index(column_name)
        cell_value = row.get(column_name)

        if cell_value is None:
            return None

        return cell_value

# Пока дату преобразовал в string так что метод не используется
def convert_date(date_string):
    if date_string:
        try:
            return datetime.strptime(date_string.strip(), '%d.%m.%Y').date()
        except ValueError:
            print(f"Ошибка преобразования даты: {date_string}")
            return None
    return None