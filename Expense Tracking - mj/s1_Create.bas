Attribute VB_Name = "s1_Create"
Option Explicit

Sub CreateCopy()
'Jan, save as new workbook

    Dim a, b, c
    Dim FilePath, Title

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'create name of new workbook
    a = ActiveWorkbook.Name
    a = Mid(a, 1, InStrRev(a, ".") - 6) & " " & Year(Date)
    Title = a

    'confirm name of new workbook
    a = Space(4) & "Is the new workbook name below correct?"
    b = Title
    c = Len(a) - Len(b) + 2

    'acquire new wb name
    i = MsgBox(Space(6) & a & Chr(10) & Chr(10) _
        & Space(c) & b & Chr(10), 3, "Confirm")

    If i = 2 Then GoTo ext
    If i = 6 Then Title = Title & ".xlsm"
    If i = 7 Then

        i = MsgBox(Space(4) & "Would you like to enter in the new workbook name?", 1, "Enter Name")
        If i <> 1 Then GoTo ext

retry:
        'enter new wb name
        i = InputBox(Space(22) & "Enter the name of the new workbook" & Chr(10) _
                & Space(37) & "Exclude the extension", "New Name", Title)
        If i = "" Then GoTo ext
        If i = Title Or InStr(i, ".xl") <> 0 Then GoTo retry Else Title = i

        'confirm name of new wb
        a = Space(4) & "Is the new workbook name below correct?"
        b = i
        c = Len(a) - Len(b) + 2

        'confirm new wb name
        i = MsgBox(Space(6) & a & Chr(10) & Chr(10) _
            & Space(c) & b & Chr(10), 1, "Confirm")

        If i <> 1 Then GoTo ext
        Title = Title & ".xlsm"

    End If

    'acquire directory to save new workbook
    b = ActiveWorkbook.path
    b = Mid(b, 1, InStrRev(b, "\"))
    FilePath = b & Year(Date)

    'confirm name and directory of new workbook
    i = MsgBox(Chr(9) & Space(6) & "Is the new workbook directory below correct?" _
            & Chr(10) & Chr(10) & Space(7) & FilePath, 3, "Confirm")
    If i = 2 Then GoTo ext
    If i = 7 Then

        i = MsgBox(Space(4) & "Would you like to enter the new directory?", 1, "Enter File Path")
        If i <> 1 Then GoTo ext

redo:
        'enter new filepath
        i = InputBox(Space(36) & "Enter the new directory", "New File Path", b)
        If i = "" Then Exit Sub
        If i = b Then GoTo redo Else b = i

        i = MsgBox(Chr(9) & Space(10) & "Is the new directory below correct?" _
            & Chr(10) & Chr(10) & Space(4) & i, 3, "Confirm")
        If i = 2 Then GoTo ext
        If i = 7 Then GoTo redo
        FilePath = b

    End If

    'create directory path if it does not exist
    On Error GoTo err
    If Len(Dir(FilePath, 16)) = 0 Then
       MkDir FilePath
    End If
    On Error GoTo 0

    'save workbook as new title
    ActiveWorkbook.SaveAs filename:=FilePath & "\" & Title

    'set global variable to signal successful completion
    i = True

    'acquire msg data
    a = "A new copy of this workbook was created."
    b = "Title:   " & Title
    c = Len(a) - Len(b) + 6

    'alert of successful completion
    MsgBox Space(6) & a & Chr(10) & Chr(10) _
         & Space(c) & b & Chr(10), 0, "Success"

ext:

    'restore settings
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic

Exit Sub

err:

    MsgBox Space(4) & "An error occurred creating the folder." & Chr(10) _
        & Space(8) & "Check the file path and try again.", 0, "Error"
    On Error GoTo 0
    GoTo ext

End Sub
