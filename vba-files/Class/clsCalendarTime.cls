VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCalendarTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========================================================================================
' Class:           clsModernStyle
' Author:          VBATools
' Version:         1.0.1
' Creation Date:   18.11.2025 15:28
' Update Date:     18.11.2025 15:28
'========================================================================================
Public Enum enIconTime
    icAddTo = 60616
    icEmojiTabFavorites = 60762
    icHistory = 59420
    icRecent = 59427
    icSetHistoryStatus = 63288
    icSetHistoryStatus2 = 63289
    icStopwatch = 59670
End Enum

Private WithEvents mTextBox As MSForms.TextBox
Attribute mTextBox.VB_VarHelpID = -1
Private WithEvents mLbArrowUpHours As MSForms.Label
Attribute mLbArrowUpHours.VB_VarHelpID = -1
Private WithEvents mLbArrowDownHours As MSForms.Label
Attribute mLbArrowDownHours.VB_VarHelpID = -1
Private WithEvents mLbArrowUpMinuts As MSForms.Label
Attribute mLbArrowUpMinuts.VB_VarHelpID = -1
Private WithEvents mLbArrowDownMinuts As MSForms.Label
Attribute mLbArrowDownMinuts.VB_VarHelpID = -1
Private WithEvents mLbBtn As MSForms.Label
Attribute mLbBtn.VB_VarHelpID = -1
Private WithEvents mParent As MSForms.UserForm
Attribute mParent.VB_VarHelpID = -1

Private mMainPanel  As MSForms.Frame
Private mLbMainColon As MSForms.Label
Private mLbHours    As MSForms.Label
Private mLbMinuts   As MSForms.Label
Private mRightBtn   As Boolean

Private Const SIZE_LABEL As Byte = 12
Private Const SIZE_WIDTH As Byte = 45

' Property: Version
' Purpose: Gets version information about the class
Public Property Get Version() As String
    Version = "Version: 1.0.1" & vbNewLine & _
            "Author: VBATools" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Date of creation: 18.11.2025 15:28" & vbNewLine & _
            "Date of update: 18.11.2025 15:28"
End Property

Public Property Get ForeColorArrowDownMinuts() As XlRgbColor
    ForeColorArrowDownMinuts = mLbArrowDownMinuts.ForeColor
End Property

Public Property Let ForeColorArrowDownMinuts(ByVal ValueColor As XlRgbColor)
    mLbArrowDownMinuts.ForeColor = ValueColor
End Property

Public Property Get ForeColorArrowUpMinuts() As XlRgbColor
    ForeColorArrowUpMinuts = mLbArrowUpMinuts.ForeColor
End Property

Public Property Let ForeColorArrowUpMinuts(ByVal ValueColor As XlRgbColor)
    mLbArrowUpMinuts.ForeColor = ValueColor
End Property

Public Property Get ForeColorArrowDownHours() As XlRgbColor
    ForeColorArrowDownHours = mLbArrowDownHours.ForeColor
End Property

Public Property Let ForeColorArrowDownHours(ByVal ValueColor As XlRgbColor)
    mLbArrowDownHours.ForeColor = ValueColor
End Property

Public Property Get ForeColorArrowUpHours() As XlRgbColor
    ForeColorArrowUpHours = mLbArrowUpHours.ForeColor
End Property

Public Property Let ForeColorArrowUpHours(ByVal ValueColor As XlRgbColor)
    mLbArrowUpHours.ForeColor = ValueColor
End Property

Public Property Get ForeColorColon() As XlRgbColor
    ForeColorColon = mLbMainColon.ForeColor
End Property

Public Property Let ForeColorColon(ByVal ValueColor As XlRgbColor)
    mLbMainColon.ForeColor = ValueColor
End Property

Public Property Get ForeColorMinuts() As XlRgbColor
    ForeColorMinuts = mLbMinuts.ForeColor
End Property

Public Property Let ForeColorMinuts(ByVal ValueColor As XlRgbColor)
    mLbMinuts.ForeColor = ValueColor
End Property

Public Property Get ForeColorHours() As XlRgbColor
    ForeColorHours = mLbHours.ForeColor
End Property

Public Property Let ForeColorHours(ByVal ValueColor As XlRgbColor)
    mLbHours.ForeColor = ValueColor
End Property

Public Property Get ForeColorBtn() As XlRgbColor
    ForeColorBtn = mLbBtn.ForeColor
End Property

Public Property Let ForeColorBtn(ByVal Color As XlRgbColor)
    mLbBtn.ForeColor = Color
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

Public Property Get IconCode() As Long
    IconCode = VBA.Asc(mLbBtn.Caption)
End Property

Public Property Let IconCode(ByVal IconCode As Long)
    mLbBtn.Caption = VBA.ChrW$(IconCode)
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

Public Sub addTimePicker(ByRef TextBox As MSForms.TextBox, _
        Optional SetTime As Date = 0, _
        Optional ForeColorBtn As XlRgbColor = rgbBlack, _
        Optional RightBtn As Boolean = True, _
        Optional VisibleBtn As Boolean = True, _
        Optional IconCode As enIconTime = icStopwatch, _
        Optional ForeColorHours As XlRgbColor = rgbBlack, _
        Optional ForeColorMinuts As XlRgbColor = rgbBlack, _
        Optional ForeColorColon As XlRgbColor = rgbBlack, _
        Optional ForeColorArrowUpHours As XlRgbColor = rgbBlack, _
        Optional ForeColorArrowDownHours As XlRgbColor = rgbBlack, _
        Optional ForeColorArrowUpMinuts As XlRgbColor = rgbBlack, _
        Optional ForeColorArrowDownMinuts As XlRgbColor = rgbBlack)

    Set mTextBox = TextBox
    Set mParent = mTextBox.Parent
    With mTextBox
        Set mMainPanel = mParent.Controls.Add("Forms.frame.1", .Name & "_lb_main", True)
        Set mLbBtn = mParent.Controls.Add("Forms.label.1", .Name & "_lb_btn", True)
    End With

    If SetTime = 0 Then SetTime = VBA.Time()
    mTextBox.Value = VBA.Format(SetTime, "hh:mm")
    mRightBtn = RightBtn

    With mMainPanel
        Set mLbMainColon = .Controls.Add("Forms.label.1", .Name & "_lb_colon", True)
        Set mLbHours = .Controls.Add("Forms.label.1", .Name & "_lb_hour", True)
        Set mLbMinuts = .Controls.Add("Forms.label.1", .Name & "_lb_minute", True)
        Set mLbArrowUpHours = .Controls.Add("Forms.label.1", .Name & "_lb_arrow_up_hour", True)
        Set mLbArrowDownHours = .Controls.Add("Forms.label.1", .Name & "_lb_arrow_down_hour", True)
        Set mLbArrowUpMinuts = .Controls.Add("Forms.label.1", .Name & "_lb_arrow_up_minute", True)
        Set mLbArrowDownMinuts = .Controls.Add("Forms.label.1", .Name & "_lb_arrow_down_minute", True)

        If mTextBox.Width < SIZE_WIDTH Then
            .Width = SIZE_WIDTH
        Else
            .Width = mTextBox.Width
        End If
        .Height = 40
        If mTextBox.Top + mTextBox.Height + .Height > mParent.InsideHeight Then
            .Top = mTextBox.Top - .Height - 2
        Else
            .Top = mTextBox.Top + mTextBox.Height + 2
        End If
        .Left = mTextBox.Left
    End With

    With mLbMainColon
        .Width = SIZE_LABEL
        .Height = SIZE_LABEL
        .Caption = ":"
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .Font.Bold = True
        .Top = mMainPanel.Height / 2 - .Height / 2
        .Left = mMainPanel.Width / 2 - .Width / 2
        .ForeColor = ForeColorColon
    End With

    With mLbHours
        .Width = SIZE_LABEL
        .Height = SIZE_LABEL
        .Caption = VBA.Format$(VBA.hour(SetTime), "00")
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .Font.Bold = True
        .Top = mLbMainColon.Top
        .Left = mLbMainColon.Left - mLbMainColon.Width
        .ForeColor = ForeColorHours
        .ZOrder 0
    End With

    With mLbMinuts
        .Width = SIZE_LABEL
        .Height = SIZE_LABEL
        .Caption = VBA.Format$(VBA.Minute(SetTime), "00")
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .Font.Bold = True
        .Top = mLbMainColon.Top
        .Left = mLbMainColon.Left + mLbMainColon.Width
        .ForeColor = ForeColorMinuts
        .ZOrder 0
    End With

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

    Call addArrowLabel(mLbArrowUpHours, mLbHours, 59150, -1, ForeColorArrowUpHours)
    Call addArrowLabel(mLbArrowDownHours, mLbHours, 59149, 1, ForeColorArrowDownHours)
    Call addArrowLabel(mLbArrowUpMinuts, mLbMinuts, 59150, -1, ForeColorArrowUpMinuts)
    Call addArrowLabel(mLbArrowDownMinuts, mLbMinuts, 59149, 1, ForeColorArrowDownMinuts)

    mMainPanel.Visible = False
End Sub

Private Sub addArrowLabel(ByRef lbArrow As MSForms.Label, ByRef lbParent As MSForms.Label, ByVal IconCode As Long, ByVal iUp As Integer, ByVal ForeColor As Long)
    With lbArrow
        .Width = SIZE_LABEL
        .Height = SIZE_LABEL
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .Font.Bold = True
        .Font.Name = "Segoe MDL2 Assets"
        .Font.Size = 10
        .Caption = VBA.ChrW$(IconCode)
        .Top = lbParent.Top + lbParent.Height * iUp
        .Left = lbParent.Left
        .ForeColor = ForeColor
    End With
End Sub

Private Sub mTextBox_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = True
    Call SetValueTextBox
End Sub

Private Sub mLbBtn_Click()
    Call SetValueTextBox
End Sub

Private Sub SetValueTextBox()
    mMainPanel.Visible = Not mMainPanel.Visible
    If mMainPanel.Visible Then
        If mTextBox.Value = vbNullString Then
            mLbHours.Caption = VBA.Format$(VBA.hour(VBA.Time()), "00")
            mLbMinuts.Caption = VBA.Format$(VBA.Minute(VBA.Time()), "00")
        End If
        mTextBox.Value = mLbHours.Caption & ":" & mLbMinuts.Caption
    End If
End Sub

Private Sub mParent_Click()
    mMainPanel.Visible = False
End Sub

Private Sub mLbArrowUpHours_Click()
    Call setTim(1, True)
End Sub

Private Sub mLbArrowDownHours_Click()
    Call setTim(-1, True)
End Sub

Private Sub mLbArrowUpMinuts_Click()
    Call setTim(1, False)
End Sub

Private Sub mLbArrowDownMinuts_Click()
    Call setTim(-1, False)
End Sub

Private Sub setTim(ByVal iUp As Integer, isHour As Boolean)
    Dim iHour       As Integer
    Dim iMinute     As Integer

    If isHour Then
        iHour = VBA.CInt(mLbHours.Caption) + iUp
        iMinute = VBA.CInt(mLbMinuts.Caption)
        Select Case iHour
            Case Is < 0
                iHour = 23
            Case Is > 23
                iHour = 0
            Case 0 To 23
        End Select
    Else
        iHour = VBA.CInt(mLbHours.Caption)
        iMinute = VBA.CInt(mLbMinuts.Caption) + iUp
        Select Case iMinute
            Case Is < 0
                iMinute = 59
                iHour = iHour + iUp
            Case Is > 59
                iMinute = 0
                iHour = iHour + iUp
            Case 0 To 59
        End Select
    End If
    mLbHours.Caption = VBA.Format$(iHour, "00")
    mLbMinuts.Caption = VBA.Format$(iMinute, "00")
    mTextBox.Value = mLbHours.Caption & ":" & mLbMinuts.Caption
End Sub
'
'Private Sub mParent_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    mLbArrowUpHours.ForeColor = BlendColor(mLbArrowUpHours.ForeColor, RGB(255, 255, 255), 0)
'End Sub
'
'Private Sub mMainPanel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    mLbArrowUpHours.ForeColor = BlendColor(mLbArrowUpHours.ForeColor, RGB(255, 255, 255), 0)
'End Sub
'
'Private Sub mLbArrowUpHours_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'    mLbArrowUpHours.ForeColor = BlendColor(mLbArrowUpHours.ForeColor, RGB(255, 255, 255), 155 / 255)
'End Sub
'
'' Function to blend colors for hover effect
'Private Function BlendColor(ByVal color1 As Long, color2 As Long, ByVal factor As Double) As Long
'    Dim r1 As Byte, g1 As Byte, b1 As Byte
'    Dim r2 As Byte, g2 As Byte, b2 As Byte
'    Dim r As Byte, g As Byte, b As Byte
'
'    ' Extract BGR components from the first color (VBA uses BGR format)
'    r1 = color1 Mod 256
'    g1 = (color1 \ 256) Mod 256
'    b1 = (color1 \ 6536) Mod 256
'
'    ' Extract BGR components from the second color (VBA uses BGR format)
'    r2 = color2 Mod 256
'    g2 = (color2 \ 256) Mod 256
'    b2 = (color2 \ 65536) Mod 256
'
'    ' Blend the colors
'    r = r1 + (r2 - r1) * factor
'    g = g1 + (g2 - g1) * factor
'    b = b1 + (b2 - b1) * factor
'
'    ' Return the blended color
'    BlendColor = RGB(r, g, b)
'End Function



