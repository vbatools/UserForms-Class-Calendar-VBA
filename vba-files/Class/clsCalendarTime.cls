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
' Class: clsCalendarTime
' Author: VBATools
' Version: 1.0.2
' Creation Date: 18.11.2025 15:28
' Last Update Date: 03.12.2025 14:00
' License: Apache License
'========================================================================================
' Time color constants
Public Enum enTimeColors
    tcBlack = 0
    tcWhite = 1677215
    tcRed = 255
    tcGreen = 65280
    tcBlue = 16711680
    tcYellow = 65535
    tcOrange = 42495
    tcBrown = 2763429
    tcGray = 8421504
    tcLightGray = 12632256
    tcSelectedTime = 12632319
    tcMoveHighlight = 14737632
    tcTitleBar = 128
End Enum

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
Private WithEvents mLbArrowUpMinutes As MSForms.Label
Attribute mLbArrowUpMinutes.VB_VarHelpID = -1
Private WithEvents mLbArrowDownMinutes As MSForms.Label
Attribute mLbArrowDownMinutes.VB_VarHelpID = -1
Private WithEvents mLbBtn As MSForms.Label
Attribute mLbBtn.VB_VarHelpID = -1
Private WithEvents mParent As MSForms.UserForm
Attribute mParent.VB_VarHelpID = -1

Private mMainPanel  As MSForms.Frame
Private mLbMainColon As MSForms.Label
Private mLbHours    As MSForms.Label
Private mLbMinutes   As MSForms.Label
Private mRightBtn   As Boolean

Private Const SIZE_LABEL As Byte = 12
Private Const SIZE_WIDTH As Byte = 45

' Property: Version
' Purpose: Gets version information about the class
Public Property Get Version() As String
    Version = "Version: 1.0.2" & vbNewLine & _
            "Author: VBATools" & vbNewLine & _
            "License: Apache" & vbNewLine & _
            "Date of creation: 18.11.2025 15:28" & vbNewLine & _
            "Date of update: 03.12.2025 14:00"
End Property

' Property: ForeColorArrowDownMinutes
' Purpose: Gets the color of the down arrow for minutes
Public Property Get ForeColorArrowDownMinutes() As XlRgbColor
    ForeColorArrowDownMinutes = mLbArrowDownMinutes.ForeColor
End Property

' Property: ForeColorArrowDownMinutes
' Purpose: Sets the color of the down arrow for minutes
Public Property Let ForeColorArrowDownMinutes(ByVal ValueColor As XlRgbColor)
    mLbArrowDownMinutes.ForeColor = ValueColor
End Property

' Property: ForeColorArrowUpMinutes
' Purpose: Gets the color of the up arrow for minutes
Public Property Get ForeColorArrowUpMinutes() As XlRgbColor
    ForeColorArrowUpMinutes = mLbArrowUpMinutes.ForeColor
End Property

' Property: ForeColorArrowUpMinutes
' Purpose: Sets the color of the up arrow for minutes
Public Property Let ForeColorArrowUpMinutes(ByVal ValueColor As XlRgbColor)
    mLbArrowUpMinutes.ForeColor = ValueColor
End Property

' Property: ForeColorArrowDownHours
' Purpose: Gets the color of the down arrow for hours
Public Property Get ForeColorArrowDownHours() As XlRgbColor
    ForeColorArrowDownHours = mLbArrowDownHours.ForeColor
End Property

' Property: ForeColorArrowDownHours
' Purpose: Sets the color of the down arrow for hours
Public Property Let ForeColorArrowDownHours(ByVal ValueColor As XlRgbColor)
    mLbArrowDownHours.ForeColor = ValueColor
End Property

' Property: ForeColorArrowUpHours
' Purpose: Gets the color of the up arrow for hours
Public Property Get ForeColorArrowUpHours() As XlRgbColor
    ForeColorArrowUpHours = mLbArrowUpHours.ForeColor
End Property

' Property: ForeColorArrowUpHours
' Purpose: Sets the color of the up arrow for hours
Public Property Let ForeColorArrowUpHours(ByVal ValueColor As XlRgbColor)
    mLbArrowUpHours.ForeColor = ValueColor
End Property

' Property: ForeColorColon
' Purpose: Gets the color of the colon separator
Public Property Get ForeColorColon() As XlRgbColor
    ForeColorColon = mLbMainColon.ForeColor
End Property

' Property: ForeColorColon
' Purpose: Sets the color of the colon separator
Public Property Let ForeColorColon(ByVal ValueColor As XlRgbColor)
    mLbMainColon.ForeColor = ValueColor
End Property

' Property: ForeColorMinutes
' Purpose: Gets the color of the minutes display
Public Property Get ForeColorMinutes() As XlRgbColor
    ForeColorMinutes = mLbMinutes.ForeColor
End Property

' Property: ForeColorMinutes
' Purpose: Sets the color of the minutes display
Public Property Let ForeColorMinutes(ByVal ValueColor As XlRgbColor)
    mLbMinutes.ForeColor = ValueColor
End Property

' Property: ForeColorHours
' Purpose: Gets the color of the hours display
Public Property Get ForeColorHours() As XlRgbColor
    ForeColorHours = mLbHours.ForeColor
End Property

' Property: ForeColorHours
' Purpose: Sets the color of the hours display
Public Property Let ForeColorHours(ByVal ValueColor As XlRgbColor)
    mLbHours.ForeColor = ValueColor
End Property

' Property: ForeColorBtn
' Purpose: Gets the color of the button
Public Property Get ForeColorBtn() As XlRgbColor
    ForeColorBtn = mLbBtn.ForeColor
End Property

' Property: ForeColorBtn
' Purpose: Sets the color of the button
Public Property Let ForeColorBtn(ByVal Color As XlRgbColor)
    mLbBtn.ForeColor = Color
End Property

' Property: ChoseDate
' Purpose: Gets or sets the selected time
Public Property Get ChoseDate() As Date
    If IsDate(mTextBox.Value) Then
        ChoseDate = VBA.CDate(mTextBox.Value)
    End If
End Property

' Property: ChoseDate
' Purpose: Gets or sets the selected time
Public Property Let ChoseDate(ByVal dtDate As Date)
    If IsDate(dtDate) Then
        mTextBox.Value = dtDate
    End If
End Property

' Property: IconCode
' Purpose: Gets or sets the icon code for the button
Public Property Get IconCode() As Long
    IconCode = VBA.Asc(mLbBtn.Caption)
End Property

' Property: IconCode
' Purpose: Gets or sets the icon code for the button
Public Property Let IconCode(ByVal IconCode As Long)
    mLbBtn.Caption = VBA.ChrW$(IconCode)
End Property

' Property: VisibleBtn
' Purpose: Gets or sets the visibility of the button
Public Property Get VisibleBtn() As Boolean
    VisibleBtn = mLbBtn.Visible
End Property

' Property: VisibleBtn
' Purpose: Gets or sets the visibility of the button
Public Property Let VisibleBtn(ByVal Value As Boolean)
    mLbBtn.Visible = Value
End Property

' Property: RightBtn
' Purpose: Gets or sets the position of the button (right or left)
Public Property Get RightBtn() As Boolean
    RightBtn = mRightBtn
End Property

' Property: RightBtn
' Purpose: Gets or sets the position of the button (right or left)
Public Property Let RightBtn(ByVal RightBtn As Boolean)
    mRightBtn = RightBtn
    If RightBtn Then
        mLbBtn.Left = mTextBox.Left + mTextBox.Width
    Else
        mLbBtn.Left = mTextBox.Left - mLbBtn.Width
    End If
End Property

' Method: addTimePicker
' Purpose: Adds a time picker control to a text box
' Parameters:
'   TextBox - The text box to attach the time picker to
'   SetTime - The initial time to set (default is current time)
'   ForeColorBtn - The color of the button icon (default is black)
'   RightBtn - Whether the button should be on the right side (default is True)
'   VisibleBtn - Whether the button should be visible (default is True)
'   IconCode - The icon to use for the button (default is stopwatch icon)
'   ForeColorHours - The color of the hours display (default is black)
'   ForeColorMinutes - The color of the minutes display (default is black)
'   ForeColorColon - The color of the colon separator (default is black)
'   ForeColorArrowUpHours - The color of the up arrow for hours (default is black)
'   ForeColorArrowDownHours - The color of the down arrow for hours (default is black)
'   ForeColorArrowUpMinutes - The color of the up arrow for minutes (default is black)
'   ForeColorArrowDownMinutes - The color of the down arrow for minutes (default is black)
Public Sub addTimePicker(ByRef TextBox As MSForms.TextBox, _
        Optional SetTime As Date = 0, _
        Optional ForeColorBtn As XlRgbColor = enTimeColors.tcBlack, _
        Optional RightBtn As Boolean = True, _
        Optional VisibleBtn As Boolean = True, _
        Optional IconCode As enIconTime = icStopwatch, _
        Optional ForeColorHours As XlRgbColor = enTimeColors.tcBlack, _
        Optional ForeColorMinutes As XlRgbColor = enTimeColors.tcBlack, _
        Optional ForeColorColon As XlRgbColor = enTimeColors.tcBlack, _
        Optional ForeColorArrowUpHours As XlRgbColor = enTimeColors.tcBlack, _
        Optional ForeColorArrowDownHours As XlRgbColor = enTimeColors.tcBlack, _
        Optional ForeColorArrowUpMinutes As XlRgbColor = enTimeColors.tcBlack, _
        Optional ForeColorArrowDownMinutes As XlRgbColor = enTimeColors.tcBlack)

    ' Check if TextBox is Nothing
    If TextBox Is Nothing Then
        Err.Raise vbObjectError + 101, "clsCalendarTime", "TextBox cannot be Nothing"
        Exit Sub
    End If
    
    Set mTextBox = TextBox
    Set mParent = mTextBox.Parent
    With mTextBox
        ' Check if controls already exist before creating new ones
        On Error Resume Next
        Set mMainPanel = mParent.Controls(.Name & "_lb_main")
        Set mLbBtn = mParent.Controls(.Name & "_lb_btn")
        On Error GoTo 0
        If mMainPanel Is Nothing Then
            Set mMainPanel = mParent.Controls.Add("Forms.frame.1", .Name & "_lb_main", True)
        End If
        If mLbBtn Is Nothing Then
            Set mLbBtn = mParent.Controls.Add("Forms.label.1", .Name & "_lb_btn", True)
        End If
    End With

    If SetTime = 0 Then SetTime = VBA.Time()
    mTextBox.Value = VBA.Format(SetTime, "hh:mm")
    mRightBtn = RightBtn

    With mMainPanel
        On Error Resume Next
        Set mLbMainColon = .Controls(.Name & "_lb_colon")
        Set mLbHours = .Controls(.Name & "_lb_hour")
        Set mLbMinutes = .Controls(.Name & "_lb_minute")
        Set mLbArrowUpHours = .Controls(.Name & "_lb_arrow_up_hour")
        Set mLbArrowDownHours = .Controls(.Name & "_lb_arrow_down_hour")
        Set mLbArrowUpMinutes = .Controls(.Name & "_lb_arrow_up_minute")
        Set mLbArrowDownMinutes = .Controls(.Name & "_lb_arrow_down_minute")
        On Error GoTo 0
        
        If mLbMainColon Is Nothing Then Set mLbMainColon = .Controls.Add("Forms.label.1", .Name & "_lb_colon", True)
        If mLbHours Is Nothing Then Set mLbHours = .Controls.Add("Forms.label.1", .Name & "_lb_hour", True)
        If mLbMinutes Is Nothing Then Set mLbMinutes = .Controls.Add("Forms.label.1", .Name & "_lb_minute", True)
        If mLbArrowUpHours Is Nothing Then Set mLbArrowUpHours = .Controls.Add("Forms.label.1", .Name & "_lb_arrow_up_hour", True)
        If mLbArrowDownHours Is Nothing Then Set mLbArrowDownHours = .Controls.Add("Forms.label.1", .Name & "_lb_arrow_down_hour", True)
        If mLbArrowUpMinutes Is Nothing Then Set mLbArrowUpMinutes = .Controls.Add("Forms.label.1", .Name & "_lb_arrow_up_minute", True)
        If mLbArrowDownMinutes Is Nothing Then Set mLbArrowDownMinutes = .Controls.Add("Forms.label.1", .Name & "_lb_arrow_down_minute", True)

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

    With mLbMinutes
        .Width = SIZE_LABEL
        .Height = SIZE_LABEL
        .Caption = VBA.Format$(VBA.Minute(SetTime), "00")
        .TextAlign = fmTextAlignCenter
        .BackStyle = fmBackStyleTransparent
        .Font.Bold = True
        .Top = mLbMainColon.Top
        .Left = mLbMainColon.Left + mLbMainColon.Width
        .ForeColor = ForeColorMinutes
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
    Call addArrowLabel(mLbArrowUpMinutes, mLbMinutes, 59150, -1, ForeColorArrowUpMinutes)
    Call addArrowLabel(mLbArrowDownMinutes, mLbMinutes, 59149, 1, ForeColorArrowDownMinutes)

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
            mLbMinutes.Caption = VBA.Format$(VBA.Minute(VBA.Time()), "00")
        End If
        mTextBox.Value = mLbHours.Caption & ":" & mLbMinutes.Caption
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

Private Sub mLbArrowUpMinutes_Click()
    Call setTim(1, False)
End Sub

Private Sub mLbArrowDownMinutes_Click()
    Call setTim(-1, False)
End Sub

' Method for cleaning up created controls
Public Sub Cleanup()
    On Error Resume Next
    If Not mMainPanel Is Nothing Then
        mMainPanel.Parent.Controls.Remove mMainPanel.Name
        Set mMainPanel = Nothing
    End If
    If Not mLbBtn Is Nothing Then
        mLbBtn.Parent.Controls.Remove mLbBtn.Name
        Set mLbBtn = Nothing
    End If
    Set mTextBox = Nothing
    Set mParent = Nothing
    Set mLbMainColon = Nothing
    Set mLbHours = Nothing
    Set mLbMinutes = Nothing
    Set mLbArrowUpHours = Nothing
    Set mLbArrowDownHours = Nothing
    Set mLbArrowUpMinutes = Nothing
    Set mLbArrowDownMinutes = Nothing
End Sub

Private Sub setTim(ByVal iUp As Integer, isHour As Boolean)
    Dim iHour       As Integer
    Dim iMinute     As Integer
    
    iHour = VBA.CInt(mLbHours.Caption)
    iMinute = VBA.CInt(mLbMinutes.Caption)

    If isHour Then
        iHour = iHour + iUp
        Select Case iHour
            Case Is < 0
                iHour = 23
            Case Is > 23
                iHour = 0
        End Select
    Else
        iMinute = iMinute + iUp
        Select Case iMinute
            Case Is < 0
                iMinute = 59
                ' When changing minutes by -1, if minutes become 59, decrease hours by 1
                If iUp < 0 Then iHour = iHour + iUp
            Case Is > 59
                iMinute = 0
                ' When changing minutes by +1, if minutes become 0, increase hours by 1
                If iUp > 0 Then iHour = iHour + iUp
        End Select
    End If
    
    ' Check hour boundaries when changing minutes
    If Not isHour Then
        If iHour < 0 Then iHour = 23
        If iHour > 23 Then iHour = 0
    End If
    
    mLbHours.Caption = VBA.Format$(iHour, "00")
    mLbMinutes.Caption = VBA.Format$(iMinute, "00")
    mTextBox.Value = mLbHours.Caption & ":" & mLbMinutes.Caption
End Sub
