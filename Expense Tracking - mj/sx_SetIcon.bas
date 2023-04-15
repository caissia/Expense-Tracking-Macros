Attribute VB_Name = "sx_SetIcon"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' modSetIcon
' By Chip Pearson, chip@cpearson.com, www.cpearson.com/SetIcon.aspx
' This module contains code to change the icon of the Excel main
' window. The code is compatible with 64-bit Office.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

#If VBA7 And Win64 Then
'''''''''''''''''''''''''''''
' 64 bit Excel
'''''''''''''''''''''''''''''
Private Declare PtrSafe Function SendMessageA Lib "user32" _
      (ByVal hWnd As LongPtr, ByVal wMsg As LongLong, ByVal wParam As LongLong, _
      ByVal lParam As LongLong) As LongPtr

Private Declare PtrSafe Function ExtractIconA Lib "shell32.dll" _
      (ByVal hInst As LongPtr, ByVal lpszExeFileName As String, _
      ByVal nIconIndex As LongPtr) As Long

Private Const ICON_SMALL = 0&
Private Const ICON_BIG = 1&
Private Const WM_SETICON = &H80

#Else
'''''''''''''''''''''''''''''
' 32 bit Excel
'''''''''''''''''''''''''''''
Private Declare PtrSafe Function SendMessageA Lib "user32" _
      (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, _
      ByVal lParam As Long) As Long

Private Declare PtrSafe Function ExtractIconA Lib "shell32.dll" _
      (ByVal hInst As Long, ByVal lpszExeFileName As String, _
      ByVal nIconIndex As Long) As Long

Private Const ICON_SMALL As Long = 0&
Private Const ICON_BIG As Long = 1&
Private Const WM_SETICON As LongPtr = &H80
#End If


Sub SetIcon(filename As String, Optional index As Long = 0)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' SetIcon
' This procedure sets the icon in the upper left corner of
' the main Excel window. FileName is the name of the file
' containing the icon. It may be an .ico file, an .exe file,
' or a .dll file. If it is an .ico file, Index must be 0
' or omitted. If it is an .exe or .dll file, Index is the
' 0-based index to the icon resource.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
#If VBA7 And Win64 Then
    ' 64 bit Excel
    Dim hWnd As LongPtr
    Dim HIcon As LongPtr
#Else
    ' 32 bit Excel
    Dim hWnd As Long
    Dim HIcon As Long
#End If
    Dim n As Long
    Dim s As String
    If Dir(filename, vbNormal) = vbNullString Then
        ' file not found, get out
        Exit Sub
    End If
    ' get the extension of the file.
    n = InStrRev(filename, ".")
    s = LCase(Mid(filename, n + 1))
    ' ensure we have a valid file type
    Select Case s
        Case "exe", "ico", "dll"
            ' OK
        Case Else
        Debug.Print s
            ' invalid file type
            err.Raise 5
    End Select
    hWnd = Application.hWnd
    If hWnd = 0 Then
        Exit Sub
    End If
    HIcon = ExtractIconA(0, filename, index)
    If HIcon <> 0 Then
        SendMessageA hWnd, WM_SETICON, ICON_SMALL, HIcon
    End If
End Sub

Sub ResetIconToExcel()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ResetIconToExcel
' This resets the Excel window's icon. It is assumed to
' be the first icon resource in the Excel.exe file.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim FName As String
    FName = Application.path & "\excel.exe"
    SetIcon FName
End Sub
