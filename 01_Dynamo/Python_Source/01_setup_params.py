# ПИТОН СКРИПТ: СОЗДАНИЕ СПЕЦИФИКАЦИИ НАГРУЗОК
import clr
import sys
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument

# ===================== НАСТРОЙКИ =====================
# Исправлено: имя параметра соответствует файлу общих параметров
PARAM_NAME = "ADSK_Нагрузка_Полезная"
SCHEDULE_NAME = "00_Контроль нагрузок (Авто)"
CATEGORY = Category.GetCategory(doc, BuiltInCategory.OST_Rooms)

# ===================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =====================
def schedule_exists(schedule_name, category_id):
    """Проверка существования спецификации с указанным именем и категорией"""
    try:
        collector = FilteredElementCollector(doc).OfClass(ViewSchedule)
        for schedule in collector:
            if schedule.Name == schedule_name and schedule.Definition.CategoryId == category_id:
                return schedule
        return None
    except:
        return None

def create_new_schedule():
    """Создание новой спецификации для помещений"""
    try:
        # Проверяем, поддерживает ли категория спецификации
        if not CategoryAllowsSchedules(CATEGORY.Id):
            raise Exception("Категория '{}' не поддерживает создание спецификаций".format(CATEGORY.Name))
        
        # Создаем спецификацию
        schedule = ViewSchedule.CreateSchedule(doc, CATEGORY.Id)
        schedule.Name = SCHEDULE_NAME
        
        # Базовые настройки
        schedule_def = schedule.Definition
        schedule_def.ShowTitle = True
        schedule_def.ShowHeaders = True
        schedule_def.ShowGridLines = True
        
        return schedule, "Спецификация '{}' успешно создана".format(SCHEDULE_NAME)
    
    except Exception as e:
        raise Exception("Ошибка создания спецификации: {}".format(str(e)))

def CategoryAllowsSchedules(category_id):
    """Проверка, поддерживает ли категория создание спецификаций"""
    try:
        # Проверяем, можно ли создать спецификацию для этой категории
        schedule = ViewSchedule.CreateSchedule(doc, category_id)
        doc.Delete(schedule.Id)
        return True
    except:
        return False

def configure_schedule_fields(schedule):
    """Настройка полей спецификации по ГОСТ 21.501-2018 (форма 2)"""
    try:
        schedule_def = schedule.Definition
        
        # Удаляем все существующие поля
        field_order = schedule_def.GetFieldOrder()
        for field_id in list(field_order):  # Создаем копию списка для безопасного удаления
            schedule_def.RemoveField(field_id)
        
        # Добавляем обязательные поля по ГОСТ
        # 1. Номер помещения
        room_number_field = schedule_def.AddField(
            ScheduleFieldType.Instance, 
            ElementId(BuiltInParameter.ROOM_NUMBER)
        )
        room_number_field.ColumnHeading = "№ п/п"
        room_number_field.Width = 30
        
        # 2. Наименование помещения
        room_name_field = schedule_def.AddField(
            ScheduleFieldType.Instance, 
            ElementId(BuiltInParameter.ROOM_NAME)
        )
        room_name_field.ColumnHeading = "Наименование помещения"
        room_name_field.Width = 150
        
        # 3. Полезная нагрузка (ищем параметр из общего файла)
        param_found = False
        try:
            # Ищем параметр по имени, как указано в файле общих параметров
            binding_map = doc.ParameterBindings
            iterator = binding_map.ForwardIterator()
            while iterator.MoveNext():
                if iterator.Key.Name == PARAM_NAME:
                    load_field = schedule_def.AddField(
                        ScheduleFieldType.Instance, 
                        iterator.Key.Id
                    )
                    load_field.ColumnHeading = "Нагрузка, кг/м²"
                    load_field.Width = 60
                    load_field.DisplayType = ScheduleFieldDisplayType.Decimal
                    load_field.Accuracy = 0.1
                    load_field.HorizontalAlignment = HorizontalAlignmentStyle.Right
                    param_found = True
                    break
        
        except Exception as e:
            # Не прерываем работу, если параметр не найден
            print("Предупреждение: {}".format(str(e)))
        
        # Если параметр не найден, добавляем примечание
        if not param_found:
            note_field = schedule_def.AddCalculatedField("Примечание")
            note_field.Formula = "\"Параметр '" + PARAM_NAME + "' не найден. Добавьте его из общего файла параметров\""
            note_field.ColumnHeading = "Примечание"
            note_field.Width = 200
            note_field.HorizontalAlignment = HorizontalAlignmentStyle.Left
        
        return True, "Поля спецификации настроены"
    
    except Exception as e:
        return False, "Ошибка настройки полей: {}".format(str(e))

def format_schedule_table(schedule):
    """Форматирование таблицы спецификации"""
    try:
        # Принудительно обновляем данные
        schedule.Definition.Refresh()
        
        # Настраиваем форматирование
        table_data = schedule.GetTableData()
        if not table_data:
            return True, "Данные таблицы недоступны для форматирования"
        
        # Настраиваем заголовки
        title_section = table_data.GetSectionData(SectionType.Header)
        if title_section:
            title_section.SetColumnWidth(0, 30)  # № п/п
            title_section.SetColumnWidth(1, 150)  # Наименование помещения
            if title_section.NumberOfColumns > 2:
                title_section.SetColumnWidth(2, 60)  # Нагрузка
        
        return True, "Форматирование таблицы выполнено"
    
    except Exception as e:
        return False, "Ошибка форматирования таблицы: {}".format(str(e))

# ===================== ОСНОВНОЙ БЛОК КОДА =====================
try:
    # Проверяем, существует ли уже спецификация
    existing_schedule = schedule_exists(SCHEDULE_NAME, CATEGORY.Id)
    
    if existing_schedule:
        # Если спецификация существует, просто активируем ее
        TransactionManager.Instance.EnsureInTransaction(doc)
        try:
            # Обновляем поля существующей спецификации
            field_success, field_result = configure_schedule_fields(existing_schedule)
            format_success, format_result = format_schedule_table(existing_schedule)
            TransactionManager.Instance.TransactionTaskDone()
            
            # Активируем вид
            uidoc.ActiveView = existing_schedule
            
            OUT = {
                "status": "success",
                "messages": [
                    "Спецификация '{}' уже существует".format(SCHEDULE_NAME),
                    field_result,
                    format_result
                ],
                "schedule_id": existing_schedule.Id.ToString(),
                "schedule_name": SCHEDULE_NAME
            }
        except Exception as e:
            TransactionManager.Instance.ForceCloseTransaction()
            raise e
    
    else:
        # Создаем новую спецификацию
        TransactionManager.Instance.EnsureInTransaction(doc)
        try:
            # Шаг 1: Создаем спецификацию
            schedule, create_result = create_new_schedule()
            
            # Шаг 2: Настраиваем поля
            field_success, field_result = configure_schedule_fields(schedule)
            
            # Шаг 3: Форматируем таблицу
            format_success, format_result = format_schedule_table(schedule)
            
            TransactionManager.Instance.TransactionTaskDone()
            
            # Активируем спецификацию в интерфейсе
            try:
                uidoc.ActiveView = schedule
            except:
                pass
            
            OUT = {
                "status": "success",
                "messages": [
                    create_result,
                    field_result,
                    format_result
                ],
                "schedule_id": schedule.Id.ToString(),
                "schedule_name": SCHEDULE_NAME
            }
            
        except Exception as e:
            TransactionManager.Instance.ForceCloseTransaction()
            raise e

except Exception as e:
    error_msg = "Ошибка выполнения скрипта: {}".format(str(e))
    OUT = {
        "status": "error",
        "error_message": error_msg,
        "stack_trace": str(sys.exc_info()[2])
    }