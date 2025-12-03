# Класс clsCalendarDate

## Описание
Класс `clsCalendarDate` предоставляет функциональность для добавления календаря выбора даты к текстовому полю в пользовательской форме Excel VBA. Позволяет пользователям удобно выбирать даты, месяцы или годы через визуальный интерфейс.

## Функциональные возможности
- Выбор даты, месяца или года
- Настройка форматов отображения даты
- Ограничение диапазона доступных дат
- Настройка внешнего вида (цвета, иконки)
- Поддержка различных режимов календаря

## Типы перечислений

### enTypeCalendar
Определяет режим работы календаря:
- `enDay` - выбор дня
- `enMonth` - выбор месяца
- `enYear` - выбор года

### enIconCalendar
Содержит различные иконки для календаря:
- `icCalendar` - стандартный значок календаря
- `icCalendarDay` - значок дня календаря
- `icCalendarMirrored` - отраженный значок календаря
- `icCalendarReply` - значок ответа календаря
- `icCalendarSolid` - сплошной значок календаря
- `icCalendarWeek` - значок недели календаря

### enCalendarColors
Стандартизованные цветовые константы:
- `ccBlack`, `ccWhite`, `ccRed`, `ccGreen`, `ccBlue`, `ccYellow`, `ccOrange`, `ccBrown`, `ccGray`, `ccLightGray`
- `ccSelectedDate` - цвет выбранной даты
- `ccMoveHighlight` - цвет подсветки при наведении
- `ccTitleBar` - цвет заголовка

## Методы

### addDatePicker
Добавляет календарь к текстовому полю.

**Параметры:**
- `TextBox` - текстовое поле, которому добавляется календарь
- `SetDate` - устанавливаемая дата (по умолчанию - текущая дата)
- `minDate` - минимальная дата (по умолчанию - 0)
- `maxDate` - максимальная дата (по умолчанию - 0)
- `TypeCalendar` - тип календаря (по умолчанию - enDay)
- `FormatDate` - формат даты (по умолчанию - vbNullString)
- `ForeColorBtn` - цвет кнопки (по умолчанию - rgbBlack)
- `RightBtn` - положение кнопки (по умолчанию - True)
- `VisibleBtn` - видимость кнопки (по умолчанию - True)
- `ForeColorTitle` - цвет заголовка (по умолчанию - ccTitleBar)
- `IconCode` - код иконки (по умолчанию - icCalendar)
- `ColorSelectedDate` - цвет выбранной даты (по умолчанию - ccSelectedDate)
- `ColorMove` - цвет при наведении (по умолчанию - ccMoveHighlight)
- `ForeColorDayYes` - цвет доступных дней (по умолчанию - rgbBlack)
- `ForeColorDayNo` - цвет недоступных дней (по умолчанию - ccGray)

## Свойства

### ChoseDate
Получение или установка выбранной даты.

### TypeCalendar
Получение или установка типа календаря (день/месяц/год).

### ColorSelectedDate
Цвет выбранной даты.

### ColorMove
Цвет подсветки при наведении.

### ForeColorDayYes/ForeColorDayNo
Цвета для доступных и недоступных дней.

### FormatDate
Формат отображения даты.

### VisibleBtn/RightBtn
Управление видимостью и положением кнопки.

### IconCode
Код иконки календаря.

### Version
Информация о версии класса.

## Совместимость
Класс разработан для использования с Microsoft Excel VBA и требует наличия пользовательской формы `frmDatepickerform`.

## Пример использования

```vba
Dim clDate As clsCalendarDate
Set clDate = New clsCalendarDate
Call clDate.addDatePicker(TextBox1, VBA.Date(), VBA.Date() - 365, VBA.Date() + 30, , "yyyy/mmm/dd", rgbBrown, False, True, rgbBlueViolet, icCalendarWeek, rgbGoldenrod, rgbLightGreen, rgbGreen, rgbRed)