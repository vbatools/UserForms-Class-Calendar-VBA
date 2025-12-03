VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalendarDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================================================================================
' Class: clsCalendarDate
' Author: VBATools
' Version: 1.0.2
' Creation Date: 18.11.2025 15:28
' Last Update Date: 03.12.2025 14:00
' License: Apache License
'========================================================================================

' Цветовые константы для календаря
Public Enum enCalendarColors
    ccBlack = 0
    ccWhite = 1677215
    ccRed = 255
    ccGreen = 65280
    ccBlue = 16711680
    ccYellow = 65535
    ccOrange = 42495
    ccBrown = 2763429
    ccGray = 8421504
    ccLightGray = 12632256
    ccSelectedDate = 12632319
    ccMoveHighlight = 14737632
    ccTitleBar = 128
End Enum

Public Enum enTypeCalendar
    enDay = 0
    enMonth
    enYear
End Enum

Public Enum enIconCalendar
    icCalendar = 59271
    icCalendarDay = 59583
    icCalendarMirrored = 60712
    icCalendarReply = 59637
    icCalendarSolid = 60041
    icCalendarWeek = 59584
End Enum

Private WithEvents mTextBox As MSForms.TextBox
Attribute mTextBox.VB_VarHelpID = -1
Private WithEvents mLbBtn As MSForms.Label
Attribute mLbBtn.VB_VarHelpID = -1
Private WithEvents mParent As MSForms.UserForm
Attribute mParent.VB_VarHelpID = -1

Private frmCalendar As frmDatepickerform
Private mForeColorTitle As XlRgbColor
Private mFormatDate As String
Private mRightBtn   As Boolean
Private mTypeCalendar As enTypeCalendar
Private mColorSelectedDate As XlRgbColor
Private mColorMove  As XlRgbColor
Private mForeColorDayYes As XlRgbColor
Private mForeColorDayNo As XlRgbColor
Private mMinDate As Date
Private mMaxDate As Date

' Property: Version
' Purpose: Gets version information about the class
Public Property Get Version() As String
    Version = "Version: 1.0.2" & vbNewLine & _
            "Author: VBATools" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Date of creation: 18.11.2025 15:28" & vbNewLine & _
            "Date of update: 03.12.2025 14:00"
End Property

Public Property Get TypeCalendar() As enTypeCalendar
    TypeCalendar = mTypeCalendar
End Property

Public Property Let TypeCalendar(ByVal TypeCalendar As enTypeCalendar)
    mTypeCalendar = TypeCalendar
End Property

Public Property Get ChoseDate() As Date
    If IsDate(mTextBox.Value) Then
        ChoseDate = VBA.CDate(mTextBox.Value)
    End If
End Property

Public Property Let ChoseDate(ByVal dtDate As Date)
    If IsDate(dtDate) Then
        mTextBox.Value = dtDate
    End If
End Property

Public Property Get ForeColorDayNo() As XlRgbColor
    ForeColorDayNo = mForeColorDayNo
End Property

Public Property Let ForeColorDayNo(ByVal Color As XlRgbColor)
    ForeColorDayNo = Color
End Property

Public Property Get ForeColorDayYes() As XlRgbColor
    ForeColorDayYes = mForeColorDayYes
End Property

Public Property Let ForeColorDayYes(ByVal Color As XlRgbColor)
    mForeColorDayYes = Color
End Property

Public Property Get ColorMove() As XlRgbColor
    ColorMove = mColorMove
End Property

Public Property Let ColorMove(ByVal Color As XlRgbColor)
    mColorMove = Color
End Property

Public Property Get ColorSelectedDate() As XlRgbColor
    ColorSelectedDate = mColorSelectedDate
End Property

Public Property Let ColorSelectedDate(ByVal Color As XlRgbColor)
    mColorSelectedDate = Color
End Property

Public Property Get ForeColorTitle() As XlRgbColor
    ForeColorTitle = mForeColorTitle
End Property

Public Property Let ForeColorTitle(ByVal Color As XlRgbColor)
    mForeColorTitle = Color
End Property

Public Property Get ForeColorBtn() As XlRgbColor
    ForeColorBtn = mLbBtn.ForeColor
End Property

Public Property Let ForeColorBtn(ByVal Color As XlRgbColor)
    mLbBtn.ForeColor = Color
End Property

Public Property Get IconCode() As Long
    IconCode = VBA.Asc(mLbBtn.Caption)
End Property

Public Property Let IconCode(ByVal IconCode As Long)
    mLbBtn.Caption = VBA.ChrW$(IconCode)
End Property

Public Property Get FormatDate() As String
    FormatDate = mFormatDate
End Property

Public Property Let FormatDate(ByVal FormatDate As String)
    mFormatDate = FormatDate
End Property

Public Property Get VisibleBtn() As Boolean
    VisibleBtn = mLbBtn.Visible
End Property

Public Property Let VisibleBtn(ByVal Value As Boolean)
    mLbBtn.Visible = Value
End Property

Public Property Get RightBtn() As Boolean
    RightBtn = mRightBtn
End Property

Public Property Let RightBtn(ByVal RightBtn As Boolean)
    mRightBtn = RightBtn
    If RightBtn Then
        mLbBtn.Left = mTextBox.Left + mTextBox.Width
    Else
        mLbBtn.Left = mTextBox.Left - mLbBtn.Width
    End If
End Property

Public Sub addDatePicker(ByRef TextBox As MSForms.TextBox, _
        Optional SetDate As Date = 0, _
        Optional minDate As Date = 0, _
        Optional maxDate As Date = 0, _
        Optional TypeCalendar As enTypeCalendar = enTypeCalendar.enDay, _
        Optional FormatDate As String = vbNullString, _
        Optional ForeColorBtn As XlRgbColor = rgbBlack, _
        Optional RightBtn As Boolean = True, _
        Optional VisibleBtn As Boolean = True, _
        Optional ForeColorTitle As XlRgbColor = &H80&, _
        Optional IconCode As enIconCalendar = enIconCalendar.icCalendar, _
        Optional ColorSelectedDate As XlRgbColor = 12632319, _
        Optional ColorMove As XlRgbColor = 14737632, _
        Optional ForeColorDayYes As XlRgbColor = rgbBlack, _
        Optional ForeColorDayNo As XlRgbColor = 8421504)
        
    If minDate > maxDate Then
        Call Err.Raise(vbObjectError + 101, "clsCalendarDate", "the minimum date:[" & minDate & "] is greater than the maximum:[" & maxDate & "]")
        Exit Sub
    End If

    Set mTextBox = TextBox
    Set mParent = mTextBox.Parent
    With mTextBox
        ' Проверяем, существует ли уже кнопка, прежде чем создавать новую
        On Error Resume Next
        Set mLbBtn = mParent.Controls(.Name & "_lb_btn")
        On Error GoTo 0
        If mLbBtn Is Nothing Then
            Set mLbBtn = mParent.Controls.Add("Forms.label.1", .Name & "_lb_btn", True)
        End If
    End With

    mForeColorTitle = ForeColorTitle
    mFormatDate = VBA.LCase$(FormatDate)
    mColorSelectedDate = ColorSelectedDate
    mColorMove = ColorMove
    mForeColorDayYes = ForeColorDayYes
    mForeColorDayNo = ForeColorDayNo
    mMinDate = minDate
    mMaxDate = maxDate

    ' Проверяем корректность формата даты
    If Not IsValidDateFormat(mFormatDate, TypeCalendar) Then
        mFormatDate = GetDefaultFormat(TypeCalendar)
    End If
    
    Select Case TypeCalendar
        Case enTypeCalendar.enDay
            If mFormatDate = vbNullString Then mFormatDate = "dd.mm.yyyy"
            If SetDate = 0 Then SetDate = VBA.Date()
            mTextBox.Value = VBA.Format$(SetDate, mFormatDate)
        Case enTypeCalendar.enMonth
            If SetDate = 0 Then
                mTextBox.Value = VBA.Month(VBA.Date())
            Else
                mTextBox.Value = VBA.Month(SetDate)
            End If
            Select Case mFormatDate
                Case "mm"
                    mTextBox.Value = VBA.MonthName(mTextBox.Value, True)
                Case "mmmm"
                    mTextBox.Value = VBA.MonthName(mTextBox.Value, False)
            End Select
        Case enTypeCalendar.enYear
            mFormatDate = vbNullString
            If SetDate = 0 Then
                mTextBox.Value = VBA.Year(VBA.Date())
            Else
                mTextBox.Value = VBA.Year(SetDate)
            End If
    End Select

    mRightBtn = RightBtn
    mTypeCalendar = TypeCalendar

    With mLbBtn
        .Height = mTextBox.Height
        .Width = .Height
        .Font.Size = .Height * 0.65
        .Font.Name = "Segoe MDL2 Assets"
        .Caption = VBA.ChrW$(IconCode)
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .Font.Bold = True
        .ForeColor = ForeColorBtn
        .Top = mTextBox.Top + 2
        If RightBtn Then
            .Left = mTextBox.Left + mTextBox.Width
        Else
            .Left = mTextBox.Left - .Width
        End If
        .ZOrder 0
        .Visible = VisibleBtn
    End With
End Sub

Private Sub mTextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    Call SetValueTextBox
End Sub

Private Sub mLbBtn_Click()
    Call SetValueTextBox
End Sub

' Проверяет корректность формата даты для заданного типа календаря
Private Function IsValidDateFormat(FormatDate As String, CalendarType As enTypeCalendar) As Boolean
    ' Для типа enDay проверяем, является ли формат допустимым форматом даты
    If CalendarType = enTypeCalendar.enDay Then
        If FormatDate = vbNullString Then
            IsValidDateFormat = True
            Exit Function
        End If
        ' Пробуем применить формат к текущей дате
        On Error GoTo ErrorHandler
        Dim testDate As Date
        testDate = VBA.Date
        Dim formatted As String
        formatted = VBA.Format$(testDate, FormatDate)
        IsValidDateFormat = (formatted <> vbNullString)
        Exit Function
    End If
    
    ' Для типа enMonth проверяем, является ли формат "mm" или "mm"
    If CalendarType = enTypeCalendar.enMonth Then
        IsValidDateFormat = (FormatDate = "mm" Or FormatDate = "mmmm")
        Exit Function
    End If
    
    ' Для типа enYear формат не используется, поэтому возвращаем True
    If CalendarType = enTypeCalendar.enYear Then
        IsValidDateFormat = True
        Exit Function
    End If
    
    IsValidDateFormat = True
    Exit Function

ErrorHandler:
    IsValidDateFormat = False
End Function

' Возвращает формат даты по умолчанию для заданного типа календаря
Private Function GetDefaultFormat(CalendarType As enTypeCalendar) As String
    Select Case CalendarType
        Case enTypeCalendar.enDay
            GetDefaultFormat = "dd.mm.yyyy"
        Case enTypeCalendar.enMonth
            GetDefaultFormat = "mmmm"
        Case enTypeCalendar.enYear
            GetDefaultFormat = ""
    End Select
End Function

Private Sub SetValueTextBox()
    Dim sVal        As String
    ' Проверяем, существует ли уже экземпляр формы
    If frmCalendar Is Nothing Then
        Set frmCalendar = New frmDatepickerform
    End If
    With frmCalendar
        .redBG.BackColor = mForeColorTitle
        sVal = mTextBox.Value
        If sVal = vbNullString Then sVal = VBA.Date()
        If Not IsDate(sVal) Then sVal = VBA.Date()
        .DateGlobal = VBA.CDate(sVal)
        .pickerMode = mTypeCalendar
        .ColorSelectedDate = mColorSelectedDate
        .ColorMove = mColorMove
        .ForeColorDayYes = mForeColorDayYes
        .ForeColorDayNo = mForeColorDayNo
        .lbMinDate = mMinDate
        .lbMaxDate = mMaxDate

        Select Case mTypeCalendar
            Case enTypeCalendar.enDay
                .lbTypeCalendar.Caption = 0
            Case enTypeCalendar.enMonth
                .lbTypeCalendar.Caption = 1
                .iSelectGlobal = VBA.Month(VBA.Date())
            Case enTypeCalendar.enYear
                .lbTypeCalendar.Caption = 1
                .iSelectGlobal = VBA.Year(VBA.Date())
        End Select

        Call .Show(1)
        sVal = .lbDateChose.Caption
        If sVal <> vbNullString Then
            mTextBox.Value = VBA.Format$(sVal, mFormatDate)
            If mTypeCalendar = enMonth Then
                Select Case mFormatDate
                    Case "mm"
                        mTextBox.Value = VBA.MonthName(sVal, True)
                    Case "mmmm"
                        mTextBox.Value = VBA.MonthName(sVal, False)
                End Select
            End If
        End If
    End With
    ' Не устанавливаем frmCalendar в Nothing, чтобы можно было повторно использовать экземпляр
    ' Set frmCalendar = Nothing
End Sub
