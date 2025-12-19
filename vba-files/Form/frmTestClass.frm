VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTestClass 
   Caption         =   "Test class:"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10425
   OleObjectBlob   =   "frmTestClass.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTestClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clTime          As clsCalendarTime
Dim clTime_custom   As clsCalendarTime
Dim clDate          As clsCalendarDate
Dim clDate_custom   As clsCalendarDate
Dim clDate_mount    As clsCalendarDate
Dim clDate_year     As clsCalendarDate

Private Sub CheckBox1_Click()
    clTime_custom.VisibleBtn = CheckBox1.Value
    clDate_custom.VisibleBtn = Not clDate_custom.VisibleBtn
End Sub

Private Sub CheckBox2_Click()
    clTime_custom.RightBtn = CheckBox2.Value
    clDate_custom.RightBtn = Not clDate_custom.RightBtn
End Sub


Private Sub UserForm_Initialize()
    With Me
        .StartUpPosition = 0
        .Left = Application.Left + 0.5 * (Application.Width - .Width)
        .Top = Application.Top + 0.5 * (Application.Height - .Height)
    End With

    Set clTime = New clsCalendarTime
    Call clTime.addTimePicker(TextBox1)

    Set clTime_custom = New clsCalendarTime
    Call clTime_custom.addTimePicker(TextBox2, VBA.Time - 10, rgbYellowGreen, False, True, icRecent, rgbOrange, rgbBrown, rgbRed, rgbOrangeRed, rgbLawnGreen, rgbGold, rgbCoral)

    Set clDate = New clsCalendarDate
    Call clDate.addDatePicker(TextBox4)

    Set clDate_custom = New clsCalendarDate
    Call clDate_custom.addDatePicker(TextBox3, VBA.Date(), VBA.Date() - 365, VBA.Date() + 30, , "yyyy/mmm/dd", rgbBrown, False, True, rgbBlueViolet, icCalendarWeek, rgbGoldenrod, rgbLightGreen, rgbGreen, rgbRed)

    Set clDate_mount = New clsCalendarDate
    Call clDate_mount.addDatePicker(TextBox5, "01.01.2020", , , enMonth, "mmmm", rgbViolet)

    Set clDate_year = New clsCalendarDate
    Call clDate_year.addDatePicker(TextBox6, "01.01.2020", VBA.Date() - 365, VBA.Date() + 30, enYear, , , , , rgbGold, icCalendarMirrored)

    lbVersion.Caption = clDate.Version(enAll)
End Sub
