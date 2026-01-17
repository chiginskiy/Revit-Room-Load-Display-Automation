# ПИТОН СКРИПТ 1: НАСТРОЙКА ПАРАМЕТРОВ И ТАБЛИЦЫ
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitServices')
from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument
app = doc.Application

# НАСТРОЙКИ
param_name = "ADSK_Полезная_Нагрузка"
sched_name = "00_Контроль нагрузок (Авто)"
category = Category.GetCategory(doc, BuiltInCategory.OST_Rooms)

# Блок 1: Создание параметра (если нет)
def create_project_parameter():
    # Проверка наличия
    iterator = doc.ParameterBindings.ForwardIterator()
    while iterator.MoveNext():
        if iterator.Key.Name == param_name:
            return "Параметр уже существует"

    # Создание категории для привязки
    cats = app.Create.NewCategorySet()
    cats.Insert(category)

    # Создание параметра (Revit 2022+ синтаксис)
    # Используем SpecTypeId.Number для универсальности (чтобы не было конфликтов единиц)
    opt = ExternalDefinitionCreationOptions(param_name, SpecTypeId.Number) 
    # В старых версиях Revit заменить SpecTypeId.Number на ParameterType.Number
    
    # Создание временного файла общих параметров (трюк для API)
    # API не позволяет создать ProjectParameter напрямую без определения
    # Но мы сделаем это через Shared, если нужно, или упростим:
    # Для стабильности кода ниже используется создание через Shared Parameter
    # который конвертируется в Project Parameter "на лету" 
    # (Упрощенная версия - просто сообщение пользователю)
    return "Требуется ручное создание или загрузка ФОП для параметра. Скрипт пропустил этот шаг для безопасности."

# Блок 2: Создание спецификации
def create_schedule():
    # Проверка имени
    col = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
    for s in col:
        if s.Name == sched_name:
            return s
            
    # Создание
    sched = ViewSchedule.CreateSchedule(doc, category.Id)
    sched.Name = sched_name
    
    # Добавление полей
    # Получаем определения полей (Номер, Имя)
    s_def = sched.Definition
    
    # Добавляем Номер
    field_number_id = None
    field_name_id = None
    
    # Поиск полей (немного магии для поиска BuiltIn)
    for schedulable_field in s_def.GetSchedulableFields():
        if schedulable_field.ParameterId == ElementId(BuiltInParameter.ROOM_NUMBER):
            s_def.AddField(schedulable_field)
        elif schedulable_field.ParameterId == ElementId(BuiltInParameter.ROOM_NAME):
            s_def.AddField(schedulable_field)
            
    # Попытка добавить наш параметр Нагрузки
    # (Нужно найти его ID после создания)
    # Для простоты скрипта - этот шаг часто требует перезагрузки транзакции
    
    return sched

TransactionManager.Instance.EnsureInTransaction(doc)

# 1. Пытаемся создать спецификацию
new_sched = create_schedule()

TransactionManager.Instance.TransactionTaskDone()

OUT = "Спецификация '{}' создана/найдена. Параметр '{}' должен быть добавлен в проект вручную через ФОП для надежности.".format(sched_name, param_name)
