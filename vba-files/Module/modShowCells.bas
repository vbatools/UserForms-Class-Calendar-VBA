Attribute VB_Name = "modShowCells"
Option Explicit
Option Private Module

Public Sub showCalendar()
    Call showCalendarCell(Cells(2, 3))
End Sub

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

    Dim sVal        As String
    Dim sRes        As String
    Dim frmCalendar As frmDatepickerform
    Set frmCalendar = New frmDatepickerform
    With frmCalendar
        .redBG.BackColor = ForeColorTitle
        sVal = Target.Value
        If sVal = vbNullString Then sVal = VBA.Date()
        If Not IsDate(sVal) Then sVal = VBA.Date()
        .DateGlobal = VBA.CDate(sVal)
        .pickerMode = TypeCalendar
        .ColorSelectedDate = ColorSelectedDate
        .ColorMove = ColorMove
        .ForeColorDayYes = ForeColorDayYes
        .ForeColorDayNo = ForeColorDayNo
        .lbMinDate = minDate
        .lbMaxDate = maxDate

        Select Case TypeCalendar
            Case enTypeCalendar.enDay
                .lbTypeCalendar.Caption = 0
            Case enTypeCalendar.enMonth
                .lbTypeCalendar.Caption = 1
                .iSelectGlobal = VBA.Month(VBA.Date())
            Case enTypeCalendar.enYear
                .lbTypeCalendar.Caption = 1
                .iSelectGlobal = VBA.Year(VBA.Date())
        End Select

        .lbLeft.Caption = Application.Left + Target.Left + Target.Width + 17
        .lbTop.Caption = Application.Top + Target.Top + Application.CommandBars("Ribbon").Height + 20

        Call .Show(1)
        sVal = .lbDateChose.Caption
        If sVal <> vbNullString Then
            sRes = VBA.Format$(sVal, FormatDate)
            If TypeCalendar = enMonth Then
                Select Case FormatDate
                    Case "mm"
                        sRes = VBA.MonthName(sVal, True)
                    Case "mmmm"
                        sRes = VBA.MonthName(sVal, False)
                End Select
            End If
            Target.Value2 = sRes
        End If
    End With
    Set frmCalendar = Nothing
End Sub
