Attribute VB_Name = "modShowForms"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : Module1
'* Created    : 07-11-2025 10:06
'* Author     : VBATools
'* Copyright  : VBATools
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit

Public Sub showForm()
    frmTestClass.Show
End Sub

Public Sub fileConvert()
    Call fileToBase64("")
End Sub


Public Function fileToBase64(ByVal sFilePath As String) As String
    Dim l           As Long
    l = FileLen(sFilePath)
    ReDim ByteArr(0 To l) As Byte
    Open sFilePath For Binary As #1
    Get #1, 1, ByteArr
    Close #1
    Dim oBase       As Object
    Set oBase = CreateObject("MSXML2.DOMDocument").createElement("b64")
    With oBase
        .DataType = "bin.base64"
        .nodeTypedValue = ByteArr
        fileToBase64 = .Text
    End With
End Function


