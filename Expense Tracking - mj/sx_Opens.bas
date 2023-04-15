Attribute VB_Name = "sx_Opens"
Option Explicit
'Contains 2 macros: Restore, Setup

Sub Restore()
'sx, restores settings

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub Setup()
'sx, setup sheets at workbook open

    'basic setup to avoid unnecessary delays
    Application.WindowState = xlMaximized
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'remove toolbar
    Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"

    'change the caption from Excel and file name
    Application.Caption = "Ledger"

    'replace excel icon on caption bar
    SetIcon "C:\Users\imagine\Documents\personal\finances\credit card\misc\images\dollarsign.ico", 0

    Dim c, i

    'allows macros without disabling protection
    For i = 1 To 19

        With Sheets(i)

            .Activate
            .Protect , UserInterfaceOnly:=True
            Application.DisplayFormulaBar = False
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.DisplayHeadings = False
            Application.DisplayStatusBar = False

            If i < 13 Then
                Application.Goto Sheets(i).Range("A1"), True
                Application.Goto Sheets(i).Range("C4")
            End If

            If Sheets(i).Name = "Sum" Then
                Rows("1:110").EntireRow.Hidden = False
                Application.Goto Sheets(i).Range("A1"), True
                Application.Goto Sheets(i).Range("P32")
            End If

            If Sheets(i).Name = "View" Then
                Application.Goto Sheets("View").Range("A1"), True
                Application.Goto Sheets("View").Range("O17")
            End If

            If Sheets(i).Name = "Query" Then
                Application.Goto Sheets(i).Range("A1"), True
                Application.Goto Sheets(i).Range("L5")
            End If

            If Sheets(i).Name = "Codes" Then
                Application.Goto Sheets(i).Range("A1"), True
                Application.Goto Sheets(i).Range("Transfer").Offset(1, 1)
            End If

            If Sheets(i).Name = "Data" Then
                Application.Goto Sheets(i).Range("A1"), True
                Application.Goto Sheets(i).Range("J3")
                Sheets(i).Visible = xlVeryHidden
                Sheets(i).Unprotect
            End If

            If Sheets(i).Name = "Items" Then
                Application.Goto Sheets(i).Range("A1"), True
                Application.Goto Sheets(i).Range("B6")
                Sheets(i).Visible = xlVeryHidden
            End If

            If Sheets(i).Name = "Limit" Then
                Application.Goto Sheets(i).Range("A1"), True
                Application.Goto Sheets(i).Range("E6")
            End If

            Call SetScrollArea

        End With

    Next

    'select the current month
    Sheets(Month(Date)).Activate

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
