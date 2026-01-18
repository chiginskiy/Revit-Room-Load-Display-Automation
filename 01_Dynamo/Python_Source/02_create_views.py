# ПИТОН СКРИПТ 2: ГЕНЕРАЦИЯ ЛИСТОВ И ЛЕГЕНД (ФИНАЛЬНАЯ ВЕРСИЯ)
import clr
clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager
import System
from System import Enum
import time

# === НАСТРОЙКИ ===
SUFFIX = "_НАГРУЗКИ"
SHEET_PREFIX = "Н-"
DEFAULT_SHEET_NAME = "План нагрузок"
MIN_SHEET_NAME_LENGTH = 3
MAX_RENAME_ATTEMPTS = 10
LOAD_PARAMETER_GUID = "88aea8e7-1818-4d65-8037-5c445ba7c5c3"  # GUID из SharedParameters.txt
LOAD_PARAMETER_NAME = "ADSK_Нагрузка_Полезная"
LOAD_DISPLAY_NAME = "Легенда нагрузок"
# =================

doc = DocumentManager.Instance.CurrentDBDocument
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
active_view = doc.ActiveView

result = []
transaction_started = False

try:
    # 1. Создаем копию активного вида как план помещений
    TransactionManager.Instance.EnsureInTransaction(doc)
    transaction_started = True
    
    # Проверяем, является ли активный вид планом этажа
    if active_view.ViewType != ViewType.FloorPlan:
        raise Exception("Активный вид должен быть планом этажа для создания плана помещений.")
    
    # Дублируем вид
    new_view_id = active_view.Duplicate(ViewDuplicateOption.Duplicate)
    new_view = doc.GetElement(new_view_id)
    
    # Переименовываем
    base_name = active_view.Name
    new_name = base_name + SUFFIX
    attempt = 1
    while True:
        try:
            new_view.Name = new_name
            break
        except:
            if attempt > MAX_RENAME_ATTEMPTS:
                raise Exception(f"Не удалось назначить имя после {MAX_RENAME_ATTEMPTS} попыток")
            new_name = f"{base_name}{SUFFIX}_{attempt}"
            attempt += 1
    
    result.append(f"✅ Вид создан: {new_name}")
    
    # Включаем отображение помещений
    room_param = new_view.get_Parameter(BuiltInParameter.VIEW_ROOMS)
    if room_param and not room_param.IsReadOnly:
        room_param.Set(1)
        result.append("✅ Отображение помещений включено")
    else:
        raise Exception("Не удалось включить отображение помещений")
    
    # 2. Проверяем наличие параметра нагрузки
    load_param_def = None
    binding_map = doc.ParameterBindings
    iterator = binding_map.ForwardIterator()
    while iterator.MoveNext():
        if iterator.Key.Name == LOAD_PARAMETER_NAME and iterator.Key.GUID.ToString() == LOAD_PARAMETER_GUID:
            load_param_def = iterator.Key
            break
    
    if load_param_def is None:
        raise Exception(f"Параметр '{LOAD_PARAMETER_NAME}' не найден. Выполните сначала скрипт создания спецификации.")
    
    result.append(f"✅ Параметр нагрузки найден: {LOAD_PARAMETER_NAME}")
    
    # 3. Настройка цветовой схемы
    try:
        # Включаем цветовое заполнение
        fill_param = new_view.get_Parameter(BuiltInParameter.VIEWER_ZONE_COLOR_FILL)
        if fill_param and not fill_param.IsReadOnly:
            fill_param.Set(1)
            result.append("✅ Цветовое заполнение включено")
        
        # Выбираем тип отображения (Background)
        cs_param = new_view.get_Parameter(BuiltInParameter.VIEWER_COLOR_SCHEME_LOCATION)
        if cs_param and not cs_param.IsReadOnly:
            cs_param.Set(1)
            result.append("✅ Тип отображения цветовой схемы установлен")
        
        # Привязываем параметр нагрузки
        scheme_param = new_view.get_Parameter(BuiltInParameter.VIEW_COLOR_SCHEME_PARAMETER)
        if scheme_param and not scheme_param.IsReadOnly:
            scheme_param.Set(load_param_def.Id)
            result.append(f"✅ Цветовая схема привязана к параметру '{LOAD_PARAMETER_NAME}'")
        
        # Обязательно обновляем вид
        doc.Regenerate()
        time.sleep(0.5)
        
        # Важно: пользователь должен создать легенду вручную через:
        # Вид → Изменить → Легенды цветовых обозначений → Создать легенду
        result.append("ℹ️ ЦВЕТОВАЯ СХЕМА НАСТРОЕНА. Создайте легенду через: Вид → Изменить → Легенды цветовых обозначений")
        result.append("ℹ️ Затем разместите легенду на листе вручную")
        
    except Exception as e:
        result.append(f"⚠️ Ошибка настройки цветовой схемы: {str(e)}")
    
    # 4. Создание Листа
    title_blocks = FilteredElementCollector(doc)\
        .OfCategory(BuiltInCategory.OST_TitleBlocks)\
        .WhereElementIsElementType()\
        .ToElements()
    
    if not title_blocks:
        raise Exception("Не найдены загруженные семейства Основных надписей (TitleBlocks).")
    
    selected_title_block = title_blocks[0]
    
    # Создаем лист
    new_sheet = ViewSheet.Create(doc, selected_title_block.Id)
    if new_sheet is None:
        raise Exception("Не удалось создать лист.")
    
    # Генерация уникального номера листа
    sheet_number_base = SHEET_PREFIX + (active_view.Name[:MIN_SHEET_NAME_LENGTH] if len(active_view.Name) >= MIN_SHEET_NAME_LENGTH else active_view.Name)
    current_sheet_number = sheet_number_base
    sheet_attempt = 1
    
    while True:
        try:
            new_sheet.SheetNumber = current_sheet_number
            break
        except:
            if sheet_attempt > MAX_RENAME_ATTEMPTS:
                raise Exception(f"Не удалось назначить номер листу после {MAX_RENAME_ATTEMPTS} попыток")
            current_sheet_number = f"{sheet_number_base}_{sheet_attempt}"
            sheet_attempt += 1
    
    # Уникальное имя листа
    sheet_name = DEFAULT_SHEET_NAME
    name_attempt = 1
    while True:
        try:
            new_sheet.Name = sheet_name
            break
        except:
            if name_attempt > MAX_RENAME_ATTEMPTS:
                raise Exception(f"Не удалось назначить имя листу после {MAX_RENAME_ATTEMPTS} попыток")
            sheet_name = f"{DEFAULT_SHEET_NAME} {name_attempt}"
            name_attempt += 1
    
    result.append(f"✅ Лист создан: {current_sheet_number} - {sheet_name}")
    
    # 5. Размещение вида на листе
    if Viewport.CanAddViewToSheet(doc, new_sheet.Id, new_view.Id):
        sheet_outline = new_sheet.Outline
        min_pt = sheet_outline.Min
        max_pt = sheet_outline.Max
        
        # Центр листа
        center_x = (min_pt.X + max_pt.X) / 2
        center_y = (min_pt.Y + max_pt.Y) / 2
        center_pt = XYZ(center_x, center_y, 0)
        
        vp = Viewport.Create(doc, new_sheet.Id, new_view.Id, center_pt)
        if vp:
            vp.ChangeLabelOffset(XYZ(0.5, -0.5, 0))
            result.append("✅ Вид размещен на листе по центру")
        else:
            result.append("⚠️ Вид размещен неудачно")
    else:
        result.append("⚠️ Вид не может быть размещен на листе")
    
    # 6. Автоматическое размещение спецификации (если она существует)
    try:
        # Ищем спецификацию по имени
        schedule_name = "00_Контроль нагрузок (Авто)"
        schedules = FilteredElementCollector(doc).OfClass(ViewSchedule)
        schedule = None
        
        for s in schedules:
            if s.Name == schedule_name:
                schedule = s
                break
        
        if schedule:
            # Позиция справа от вида
            sheet_outline = new_sheet.Outline
            schedule_pt = XYZ(sheet_outline.Max.X - 2.5, sheet_outline.Max.Y - 2.5, 0)
            
            # Создаем видовой фрагмент
            Viewport.Create(doc, new_sheet.Id, schedule.Id, schedule_pt)
            result.append("✅ Спецификация размещена на листе")
        else:
            result.append("⚠️ Спецификация не найдена. Создайте её сначала")
    except Exception as e:
        result.append(f"⚠️ Ошибка размещения спецификации: {str(e)}")
    
    # 7. Инструкции для пользователя
    result.append("\n")
    result.append("==============================================")
    result.append("ИНСТРУКЦИЯ ПО ЗАВЕРШЕНИЮ НАСТРОЙКИ:")
    result.append("1. В меню: Вид → Изменить → Легенды цветовых обозначений")
    result.append("2. Нажмите 'Создать легенду'")
    result.append("3. В диалоге укажите:")
    result.append("   - Заголовок: 'Легенда нагрузок'")
    result.append("   - Параметр: 'ADSK_Нагрузка_Полезная'")
    result.append("   - Добавьте значения: 500,00 кг/м² и 1000,00 кг/м²")
    result.append("   - Выберите соответствующие цвета")
    result.append("4. Разместите легенду на листе вручную")
    result.append("5. При необходимости отредактируйте таблицу экспликации")
    result.append("==============================================")

except Exception as e:
    error_msg = f"❌ КРИТИЧЕСКАЯ ОШИБКА: {str(e)}"
    result.append(error_msg)
    import traceback
    error_details = traceback.format_exc()
    result.append(f"Подробности: {error_details}")
    
finally:
    # Гарантированное завершение транзакции
    if transaction_started:
        try:
            TransactionManager.Instance.TransactionTaskDone()
        except:
            pass

OUT = result