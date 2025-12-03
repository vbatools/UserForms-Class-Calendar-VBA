# clsCalendarTime Class

## Description
The `clsCalendarTime` class provides functionality for adding a time picker to a text box in an Excel VBA user form. It allows users to conveniently select hours and minutes through a visual interface with up/down arrows.

## Features
- Time selection (hours and minutes)
- Time format customization
- Appearance customization (colors, icons)
- Button visibility and position control
- Support for various time picker icons

## Enum Types

### enTimeColors
Contains standard color constants for time elements:
- `tcBlack`, `tcWhite`, `tcRed`, `tcGreen`, `tcBlue`, `tcYellow`, `tcOrange`, `tcBrown`, `tcGray`, `tcLightGray`
- `tcSelectedTime` - selected time color
- `tcMoveHighlight` - hover highlight color
- `tcTitleBar` - title bar color

### enIconTime
Contains various time icons:
- `icAddTo` - add icon
- `icEmojiTabFavorites` - favorites icon
- `icHistory` - history icon
- `icRecent` - recent icon
- `icSetHistoryStatus` - history status icon
- `icSetHistoryStatus2` - alternative history status icon
- `icStopwatch` - stopwatch icon

## Methods

### addTimePicker
Adds a time picker to a text box.

**Parameters:**
- `TextBox` - the text box to which the time picker is added
- `SetTime` - the time to set (default is current time)
- `ForeColorBtn` - button color (default is tcBlack)
- `RightBtn` - button position (default is True, on the right)
- `VisibleBtn` - button visibility (default is True)
- `IconCode` - icon code (default is icStopwatch)
- `ForeColorHours` - hours display color (default is tcBlack)
- `ForeColorMinutes` - minutes display color (default is tcBlack)
- `ForeColorColon` - colon separator color (default is tcBlack)
- `ForeColorArrowUpHours` - up arrow color for hours (default is tcBlack)
- `ForeColorArrowDownHours` - down arrow color for hours (default is tcBlack)
- `ForeColorArrowUpMinutes` - up arrow color for minutes (default is tcBlack)
- `ForeColorArrowDownMinutes` - down arrow color for minutes (default is tcBlack)

## Properties

### ChoseDate
Get or set the selected time.

### Color Properties
- `ForeColorHours/ForeColorMinutes` - colors for hours and minutes display
- `ForeColorColon` - colon separator color
- `ForeColorArrowUpHours/ForeColorArrowDownHours` - arrow colors for hours
- `ForeColorArrowUpMinutes/ForeColorArrowDownMinutes` - arrow colors for minutes
- `ForeColorBtn` - button color

### IconCode
Time picker button icon code.

### VisibleBtn/RightBtn
Button visibility and position control.

### Version
Class version information.

## Compatibility
The class is designed for use with Microsoft Excel VBA and requires user form controls to be present.

## Example Usage

```vba
Dim clTime As clsCalendarTime
Set clTime = New clsCalendarTime
Call clTime.addTimePicker(TextBox1, VBA.Time(), tcBlue, True, True, icStopwatch, tcRed, tcGreen, tcBlack, tcGray, tcGray, tcLightGray, tcLightGray)