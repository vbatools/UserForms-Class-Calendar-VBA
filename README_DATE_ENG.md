# clsCalendarDate Class

## Description
The `clsCalendarDate` class provides functionality for adding a date picker calendar to a text box in an Excel VBA user form. It allows users to conveniently select dates, months, or years through a visual interface.

## Features
- Date, month, or year selection
- Date format customization
- Date range limitation
- Appearance customization (colors, icons)
- Support for various calendar modes

## Enum Types

### enTypeCalendar
Defines the calendar mode:
- `enDay` - day selection
- `enMonth` - month selection
- `enYear` - year selection

### enIconCalendar
Contains various calendar icons:
- `icCalendar` - standard calendar icon
- `icCalendarDay` - calendar day icon
- `icCalendarMirrored` - mirrored calendar icon
- `icCalendarReply` - calendar reply icon
- `icCalendarSolid` - solid calendar icon
- `icCalendarWeek` - calendar week icon

### enCalendarColors
Standardized color constants:
- `ccBlack`, `ccWhite`, `ccRed`, `ccGreen`, `ccBlue`, `ccYellow`, `ccOrange`, `ccBrown`, `ccGray`, `ccLightGray`
- `ccSelectedDate` - selected date color
- `ccMoveHighlight` - hover highlight color
- `ccTitleBar` - title bar color

## Methods

### addDatePicker
Adds a calendar to a text box.

**Parameters:**
- `TextBox` - the text box to which the calendar is added
- `SetDate` - the date to set (default is current date)
- `minDate` - minimum date (default is 0)
- `maxDate` - maximum date (default is 0)
- `TypeCalendar` - calendar type (default is enDay)
- `FormatDate` - date format (default is vbNullString)
- `ForeColorBtn` - button color (default is rgbBlack)
- `RightBtn` - button position (default is True)
- `VisibleBtn` - button visibility (default is True)
- `ForeColorTitle` - title color (default is ccTitleBar)
- `IconCode` - icon code (default is icCalendar)
- `ColorSelectedDate` - selected date color (default is ccSelectedDate)
- `ColorMove` - hover color (default is ccMoveHighlight)
- `ForeColorDayYes` - available days color (default is rgbBlack)
- `ForeColorDayNo` - unavailable days color (default is ccGray)

## Properties

### ChoseDate
Get or set the selected date.

### TypeCalendar
Get or set the calendar type (day/month/year).

### ColorSelectedDate
Selected date color.

### ColorMove
Hover highlight color.

### ForeColorDayYes/ForeColorDayNo
Colors for available and unavailable days.

### FormatDate
Date display format.

### VisibleBtn/RightBtn
Button visibility and position control.

### IconCode
Calendar icon code.

### Version
Class version information.

## Improvements in Version 1.0.2

1. Fixed naming typos:
   - `enMount` changed to `enMonth`
   - `olorMove` changed to `ColorMove`
   - `olorSelectedDate` changed to `ColorSelectedDate`

2. Improved error handling:
   - Fixed error message formation in the `addDatePicker` method

3. Optimized control creation:
   - Added check for existing control before creating a new one

4. Improved resource management:
   - Implemented check for existing form instance before creating a new one

5. Enhanced date format handling:
   - Added helper functions `IsValidDateFormat` and `GetDefaultFormat`
   - Implemented validation of date formats

6. Standardized color constants:
   - Added `enCalendarColors` enumeration with named constants
   - Updated constructor parameters using new constants

## Compatibility
The class is designed for use with Microsoft Excel VBA and requires the user form `frmDatepickerform`.

## Example Usage

```vba
Dim clDate As clsCalendarDate
Set clDate = New clsCalendarDate
Call clDate.addDatePicker(TextBox1, VBA.Date(), VBA.Date() - 365, VBA.Date() + 30, , "yyyy/mmm/dd", rgbBrown, False, True, rgbBlueViolet, icCalendarWeek, rgbGoldenrod, rgbLightGreen, rgbGreen, rgbRed)