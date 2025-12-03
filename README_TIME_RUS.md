# Класс clsCalendarTime

## Описание
Класс `clsCalendarTime` предоставляет функциональность для добавления выбора времени к текстовому полю в пользовательской форме Excel VBA. Позволяет пользователям удобно выбирать часы и минуты через визуальный интерфейс с помощью стрелок вверх/вниз.

## Функциональные возможности
- Выбор времени (часов и минут)
- Настройка форматов отображения времени
- Настройка внешнего вида (цвета, иконки)
- Управление видимостью и положением кнопки
- Поддержка различных иконок для кнопки выбора времени

## Типы перечислений

### enTimeColors
Содержит стандартные цветовые константы для элементов времени:
- `tcBlack`, `tcWhite`, `tcRed`, `tcGreen`, `tcBlue`, `tcYellow`, `tcOrange`, `tcBrown`, `tcGray`, `tcLightGray`
- `tcSelectedTime` - цвет выбранного времени
- `tcMoveHighlight` - цвет подсветки при наведении
- `tcTitleBar` - цвет заголовка

### enIconTime
Содержит различные иконки для времени:
- `icAddTo` - иконка добавления
- `icEmojiTabFavorites` - иконка избранных
- `icHistory` - иконка истории
- `icRecent` - иконка недавних
- `icSetHistoryStatus` - иконка статуса истории
- `icSetHistoryStatus2` - альтернативная иконка статуса истории
- `icStopwatch` - иконка секундомера

## Методы

### addTimePicker
Добавляет выбор времени к текстовому полю.

**Параметры:**
- `TextBox` - текстовое поле, которому добавляется выбор времени
- `SetTime` - устанавливаемое время (по умолчанию - текущее время)
- `ForeColorBtn` - цвет кнопки (по умолчанию - tcBlack)
- `RightBtn` - положение кнопки (по умолчанию - True, справа)
- `VisibleBtn` - видимость кнопки (по умолчанию - True)
- `IconCode` - код иконки (по умолчанию - icStopwatch)
- `ForeColorHours` - цвет отображения часов (по умолчанию - tcBlack)
- `ForeColorMinutes` - цвет отображения минут (по умолчанию - tcBlack)
- `ForeColorColon` - цвет двоеточия-разделителя (по умолчанию - tcBlack)
- `ForeColorArrowUpHours` - цвет стрелки вверх для часов (по умолчанию - tcBlack)
- `ForeColorArrowDownHours` - цвет стрелки вниз для часов (по умолчанию - tcBlack)
- `ForeColorArrowUpMinutes` - цвет стрелки вверх для минут (по умолчанию - tcBlack)
- `ForeColorArrowDownMinutes` - цвет стрелки вниз для минут (по умолчанию - tcBlack)

## Свойства

### ChoseDate
Получение или установка выбранного времени.

### Color свойства
- `ForeColorHours/ForeColorMinutes` - цвета для отображения часов и минут
- `ForeColorColon` - цвет двоеточия-разделителя
- `ForeColorArrowUpHours/ForeColorArrowDownHours` - цвета стрелок для часов
- `ForeColorArrowUpMinutes/ForeColorArrowDownMinutes` - цвета стрелок для минут
- `ForeColorBtn` - цвет кнопки

### IconCode
Код иконки кнопки выбора времени.

### VisibleBtn/RightBtn
Управление видимостью и положением кнопки.

### Version
Информация о версии класса.

## Совместимость
Класс разработан для использования с Microsoft Excel VBA и требует наличия пользовательской формы с элементами управления.

## Пример использования

```vba
Dim clTime As clsCalendarTime
Set clTime = New clsCalendarTime
Call clTime.addTimePicker(TextBox1, VBA.Time(), tcBlue, True, True, icStopwatch, tcRed, tcGreen, tcBlack, tcGray, tcGray, tcLightGray, tcLightGray)