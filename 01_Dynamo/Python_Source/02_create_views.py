# ПИТОН СКРИПТ 2: ГЕНЕРАЦИЯ ЛИСТОВ И ЛЕГЕНД
import clr
clr.AddReference('RevitAPI')
from Autodesk.Revit.DB import *
from RevitServices.Persistence import DocumentManager
from RevitServices.Transactions import TransactionManager

doc = DocumentManager.Instance.CurrentDBDocument
uidoc = DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
active_view = doc.ActiveView

# НАСТРОЙКИ
suffix = "_НАГРУЗКИ"
legend_category_id = ElementId(BuiltInCategory.OST_Rooms)

result = []

TransactionManager.Instance.EnsureInTransaction(doc)

try:
    # 1. Дублирование Вида
    # DuplicateOption.Duplicate (без детализации) или WithDetailing
    new_view_id = active_view.Duplicate(ViewDuplicateOption.Duplicate)
    new_view = doc.GetElement(new_view_id)
    
    # Переименование
    try:
        new_view.Name = active_view.Name + suffix
    except:
        new_view.Name = active_view.Name + suffix + "_Copy"
        
    # 2. Создание Листа (TitleBlock берется первый попавшийся)
    tb_collector = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_TitleBlocks).WhereElementIsElementType().FirstElement()
    
    if tb_collector:
        new_sheet = ViewSheet.Create(doc, tb_collector.Id)
        new_sheet.Name = "План нагрузок"
        new_sheet.SheetNumber = "Н-" + active_view.Name[:3] # Генерируем номер
        
        # 3. Размещение вида на листе (по центру, условно)
        # Получаем центр листа (примерно)
        center_pt = XYZ(1.5, 1.0, 0) # В футах! 1.5 фута ~ 450мм
        
        if Viewport.CanAddViewToSheet(doc, new_sheet.Id, new_view.Id):
            vp = Viewport.Create(doc, new_sheet.Id, new_view.Id, center_pt)
            result.append("Вид и Лист созданы")
    else:
        result.append("Ошибка: Не загружены семейства Основных надписей (TitleBlocks)")

    # 4. Включение цветовой схемы (Workaround)
    # Мы не можем легко создать схему, но можем включить отображение
    # Получаем параметр схемы
    color_scheme_param = new_view.get_Parameter(BuiltInParameter.VIEWER_COLOR_SCHEME_LOCATION)
    if color_scheme_param and not color_scheme_param.IsReadOnly:
        # Включаем "Background" (Фон)
        color_scheme_param.Set(1) 
        result.append("Цветовая схема активирована (настройте цвета вручную)")

    # 5. Попытка создать Легенду (Color Fill Legend)
    # Это возможно только в новых версиях API и требует точных координат
    # Обычно это делают вручную, так как API часто дает сбой на Create
    # place_legend = ColorFillLegend.Create(doc, new_view.Id, legend_category_id, XYZ(0,0,0))

except Exception as e:
    result.append("Ошибка: " + str(e))

TransactionManager.Instance.TransactionTaskDone()

OUT = result