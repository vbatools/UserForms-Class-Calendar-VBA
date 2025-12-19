Attribute VB_Name = "modShowCells"
Option Explicit
Option Private Module

'========================================================================================
' Module: modShowCells
' Author: VBATools
' Description: Module for displaying a calendar in an Excel cell with the ability to select date, month, or year
'========================================================================================

' Subroutine to display a calendar by default in cell B2 (Cells(2, 3))
Public Sub showCalendar()
    ' Call showCalendarCell subroutine with default parameters for cell B2
    Call showCalendarCell(Cells(2, 3))
End Sub

' Subroutine to display a calendar with default colors in cell B2 (Cells(2, 3))
Public Sub showCalendarColor()
    ' Call showCalendarCell subroutine with color parameters
    Call showCalendarCell(Cells(2, 3), , , , , rgbOrangeRed, rgbRoyalBlue, rgbBrown, rgbGold, rgbBlueViolet)
End Sub

' Subroutine to display a calendar in a specified cell with customizable parameters
' Parameters:
' Target - Range object representing the cell where the calendar will be displayed
' minDate - Minimum allowed date for selection (default: 0 - no restriction)
' maxDate - Maximum allowed date for selection (default: 0 - no restriction)
' TypeCalendar - Calendar type: day/month/year (default: enDay)
' FormatDate - Date display format (default: vbNullString)
' ForeColorTitle - Calendar title background color (default: &H80&)
' ColorSelectedDate - Color of selected date (default: 12632319)
' ColorMove - Color on mouse hover (default: 14737632)
' ForeColorDayYes - Color of current month (default: rgbBlack)
' ForeColorDayNo - Color of other month days (default: 8421504)
Public Sub showCalendarCell(ByRef Target As Range, _
        Optional minDate As Date = 0, _
        Optional maxDate As Date = 0, _
        Optional TypeCalendar As enTypeCalendar = enTypeCalendar.enDay, _
        Optional FormatDate As String = vbNullString, _
        Optional ForeColorTitle As XlRgbColor = &H80&, _
        Optional ColorSelectedDate As XlRgbColor = 12632319, _
        Optional ColorMove As XlRgbColor = 14737632, _
        Optional ForeColorDayYes As XlRgbColor = rgbBlack, _
        Optional ForeColorDayNo As XlRgbColor = 8421504)

    Dim sVal        As String                       ' Temporary variable for cell value
    Dim sRes        As String                       ' Date formatting result
    Dim frmCalendar As frmDatepickerform            ' Variable for calendar form

    ' Create an instance of the calendar form
    Set frmCalendar = New frmDatepickerform

    With frmCalendar
        ' Set the background color of the calendar title
        .redBG.BackColor = ForeColorTitle

        ' Get value from target cell
        sVal = Target.Value

        ' If value is empty or not a date, use current date
        If sVal = vbNullString Then sVal = VBA.Date()
        If Not IsDate(sVal) Then sVal = VBA.Date()

        ' Convert and set the initial date in the calendar
        .DateGlobal = VBA.CDate(sVal)

        ' Set calendar mode (day/month/year)
        .pickerMode = TypeCalendar

        ' Set interface element colors
        .ColorSelectedDate = ColorSelectedDate      ' Color of selected date
        .ColorMove = ColorMove                      ' Color on cursor hover
        .ForeColorDayYes = ForeColorDayYes          ' Color of current month days
        .ForeColorDayNo = ForeColorDayNo            ' Color of other month days

        ' Set date restrictions
        .lbMinDate = minDate                        ' Minimum date
        .lbMaxDate = maxDate                        ' Maximum date

        ' Configure calendar depending on selection type
        Select Case TypeCalendar
            Case enTypeCalendar.enDay               ' Day selection mode
                .lbTypeCalendar.Caption = 0         ' Set calendar type indicator
            Case enTypeCalendar.enMonth             ' Month selection mode
                .lbTypeCalendar.Caption = 1         ' Set calendar type indicator
                .iSelectGlobal = VBA.Month(VBA.Date())  ' Set current month as default
            Case enTypeCalendar.enYear              ' Year selection mode
                .lbTypeCalendar.Caption = 1         ' Set calendar type indicator
                .iSelectGlobal = VBA.Year(VBA.Date())   ' Set current year as default
        End Select

        ' Position the calendar form relative to the target cell
        .lbLeft.Caption = Application.Left + Target.Left + Target.Width + 23  ' Horizontal position
        .lbTop.Caption = Application.Top + Target.Top + Application.CommandBars("Ribbon").Height + 15  ' Vertical position

        ' Display the modal calendar form
        Call .Show(1)

        ' Get the selected date from the form
        sVal = .lbDateChose.Caption

        ' If a date was selected, process and write to the cell
        If sVal <> vbNullString Then
            ' Format date according to the specified format
            sRes = VBA.Format$(sVal, FormatDate)

            ' If month selection mode is chosen, convert result to month name
            If TypeCalendar = enMonth Then
                Select Case FormatDate
                    Case "mm"                        ' Abbreviated month name format (e.g., "Jan")
                        sRes = VBA.MonthName(sVal, True)
                    Case "mm"                      ' Full month name format (e.g., "January")
                        sRes = VBA.MonthName(sVal, False)
                End Select
            End If

            ' Write the result to the target cell
            Target.Value2 = sRes
        End If
    End With

    ' Release memory
    Set frmCalendar = Nothing
End Sub
