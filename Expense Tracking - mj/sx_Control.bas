Attribute VB_Name = "sx_Control"
Global i, t
Option Explicit
'Contains 17 macros: Clear, Clone, Creates, Deletes, Form, Formats, Indicator, Last, mLink,
'                    Order, PageFlip, Protects, Reset, SetScrollArea, Switch, Toggle, View

Sub Clear()
's1 - s12, delete sheet data

    'check if data exists else exit
    If Range("C4") = "" And Range("O4") = "" Then
        Exit Sub
    End If

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    'Application.EnableEvents = False

    i = MsgBox(Space(6) & "Do you want to delete all of the data on this sheet?" & Chr(10) _
                & Chr(9) & Space(12) & "This cannot be undone.", 4, "Reset")

    If i = 6 Then
    
        'clear data
        Range("B4:C203,E4:H203,O4:O203,Q4:T203").ClearContents

        'reset duplicate indicator: "Purchase" = duplicate found, "Purchases" = no duplicate
        Range("F3") = "Purchases"

        'set scroll area
        Call SetScrollArea

    End If

    'change indicator format
    ActiveSheet.Range("M3").Font.Size = 6
    ActiveSheet.Range("M3").Font.Name = "Wingdings"

    'restore settings
    Application.Goto Range("C4")
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    'Application.EnableEvents = True
    ActiveSheet.Protect

End Sub

Sub Clone()
's1 - s12, copies chosen sheet

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    For i = 1 To 11

        Sheets(i).Activate
        Sheets(i).Unprotect
        Sheets(i).ScrollArea = ""

        Cells.Select
        ActiveSheet.DrawingObjects.Select

        Selection.Delete
        Selection.Delete Shift:=xlUp

        Cells.Select
        Selection.Delete Shift:=xlUp

        Sheets("Dec").Activate
        Cells.Select
        Selection.Copy

        Sheets(i).Activate
        Sheets(i).Select

        Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteColumnWidths

        Range("A1").Select
        ActiveSheet.Paste

    Next

    Call AutoHyperlink

    'restore settings
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic

End Sub

Sub Creates()
's1 - s12, create sheets as needed

    Dim b, c, m, n, r

    n = "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov"

    b = Len(n) - Len(Replace(n, ",", ""))

    For i = 0 To b

        'acquire month
        m = Split(n, ",")(i)

        'change name of previous sheet
        'Sheets(m).Select
        'Sheets(m).Name = m & "-"

        'create sheet
        Sheets("Dec").Select
        Sheets("Dec").Copy Before:=Sheets(12)
        Sheets("Dec (2)").Select
        Sheets("Dec (2)").Name = m

        If Sheets(m & "-").Range("C4") <> "" Then

            Sheets(m & "-").Select

            'acquire last row of charge column
            r = Sheets(m & "-").Cells(Rows.Count, "C").End(xlUp).Row
            'Debug.Print ActiveSheet.Name & " last row  =  " & r

            'transfer data to new sheet
            Sheets(m).Range("B4:H" & r).Formula = Sheets(m & "-").Range("B4:H" & r).Formula

            'acquire last row of expense column
            r = Sheets(m & "-").Cells(Rows.Count, "P").End(xlUp).Row
            'Debug.Print ActiveSheet.Name & " last row  =  " & r

            'transfer data to new sheet
            Sheets(m).Range("O4:O" & r).Formula = Sheets(m & "-").Range("P4:P" & r).Formula
            Sheets(m).Range("Q4:R" & r).Formula = Sheets(m & "-").Range("R4:S" & r).Formula
            Sheets(m).Range("T4:T" & r).Formula = Sheets(m & "-").Range("T4:T" & r).Formula

            'replace hlinks and formulas for new sheet
            Sheets(m).Activate
            Call SingleLink
            Call Form

            're-acquire last row of charge column
            r = Sheets(m & "-").Cells(Rows.Count, "C").End(xlUp).Row

            'replace formulas with values where transaction code was overridden
            For Each c In Sheets(m & "-").Range("D4:D" & r)
                If InStr(c.Formula, "=") = 0 Then
                    Sheets(m & "-").Range(c.Address).Copy
                    Sheets(m).Range(c.Address).PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                End If
            Next

            're-acquire last row of expense column
            r = Sheets(m & "-").Cells(Rows.Count, "P").End(xlUp).Row

            'replace formulas with values where transaction code was overridden
            For Each c In Sheets(m & "-").Range("Q4:Q" & r)
                If InStr(c.Formula, "=") = 0 Then
                    Sheets(m & "-").Range(c.Address).Copy
                    Sheets(m).Range(c.Offset(0, -1).Address).PasteSpecial Paste:=xlPasteValues
                    Application.CutCopyMode = False
                End If
            Next

            'move income data to its column
            For Each c In Sheets(m).Range("P4:P" & r)
                If c = "I" Then
                    Sheets(m).Range(c.Offset(0, 3).Address) = Sheets(m).Range(c.Offset(0, 2).Address)
                    Sheets(m).Range(c.Offset(0, 2).Address).ClearContents
                End If
            Next

            're-acquire last row of charge column
            r = Sheets(m).Cells(Rows.Count, "C").End(xlUp).Row

            'place space in empty charge/expense cells
            For Each c In Sheets(m).Range("F4:F" & r)
                If c = "" Then Sheets(m).Range(c.Address) = "="""""
            Next c

            're-acquire last row of expense column
            r = Sheets(m).Cells(Rows.Count, "Q").End(xlUp).Row

            'place space in empty charge/expense cells
            For Each c In Sheets(m).Range("R4:R" & r)
                If c = "" Then Sheets(m).Range(c.Address) = "="""""
            Next c

            'remove any wrap-text of description
            Sheets(m).Range("E4:E203, Q4:Q203").WrapText = False

        End If

    Next

End Sub

Sub Deletes()
's1 - s12, delete sheets as needed

    Dim c, n, m

    n = "Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov"

    c = Len(n) - Len(Replace(n, ",", ""))

    Application.DisplayAlerts = False

    For i = 0 To c

        m = Split(n, ",")(i) & "-"

        Sheets(m).Delete

        'Debug.Print m

    Next

    Application.DisplayAlerts = True

End Sub

Sub Form()
Attribute Form.VB_ProcData.VB_Invoke_Func = "F\n14"
's1 - s12, refreshes formula for transaction codes
'shortcut: Ctrl + Shift + F

    ActiveSheet.Range("D4:D203").Formula = Sheets("Codes").Range("Form1").Formula
    ActiveSheet.Range("P4:P203").Formula = Sheets("Codes").Range("Form2").Formula

End Sub

Sub Formats()
'sx, copy/paste conditional formatting

    Dim n

    For i = 1 To 12

        If i < 9 Or i > 9 Then

            n = MonthName(i, True)

            With Sheets(n)

                .Activate
                .Unprotect

                .Range("J15").Select
                Selection.Locked = False
                Selection.FormulaHidden = False

'                Sheets("Sep").Select
'                Range("M4:M14").Select
'                Selection.Copy
'
'                Sheets(n).Select
'                Range("M4:M14").Select
'                ActiveSheet.Paste
'                Application.CutCopyMode = False

                Sheets(n).Select
                Cells.Select
                Cells.FormatConditions.Delete

                Sheets("Sep").Select
                Cells.Select
                Selection.Copy

                Sheets(n).Select
                Range("A1").Select
                Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
                    SkipBlanks:=False, Transpose:=False

            End With

        End If

    Next

End Sub

Sub Indicator()
's1 - s12, changes m3 format

    If Application.max(ActiveSheet.Range("A4:A203")) = 0 Then
        ActiveSheet.Range("M3").Font.Size = 6
        ActiveSheet.Range("M3").Font.Name = "Wingdings"
    Else
        ActiveSheet.Range("M3").Font.Size = 8
        ActiveSheet.Range("M3").Font.Name = "Century Gothic"
    End If

End Sub

Sub Last()
'Data, selects last transaction

    Dim c

    For Each c In Range("D5:D2404")
        If Trim(c) = "" Then Application.Goto c.Offset(-1): Exit For
    Next

End Sub

Sub mLink()
'Sum & View hyperlinks to s1 - s12

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

'''''''''''''''''''''''' View sheet ''''''''''''''''''''''''

    'hyperlink to s1 to s12 if data present
    If ActiveSheet.Name = "View" Then

        i = ActiveCell.Address

        If i <> "$O$17" Then
            Application.Goto Sheets("View").Range("O17")
        End If

        'if month clicked goto monthly sheet if data present
        If i = "$J$3" Then

            i = Left(Range(i), 3)

            If Sheets(i).Range("C4") <> "" Then
                Sheets(i).Unprotect
                Application.Goto Sheets(i).Range("A1"), True
                Application.Goto Sheets(i).Range("C4")
                Sheets(i).Protect
                GoTo ext
            End If

            GoTo ext

        End If

        'toggle between actual & projected expenses
        If i = "$M$3" Then

            If Left(Range("M3"), 1) = "A" Then

                Range("M3") = "Projected Expenses"

            ElseIf Left(Range("M3"), 1) = "P" Then

                Range("M3") = "Actual Expenses"

            End If

            GoTo ext

        End If

        GoTo ext

    End If

'''''''''''''''''''''''' Sum sheet ''''''''''''''''''''''''

    'hyperlink to s1 to s12 if Summary sheet or make visible sections of sum sheet
    If ActiveSheet.Name = "Sum" Then

        i = ActiveCell

        If Len(i) = 1 Then

            'if 1st section only is visible make entire sheet visible else make 2nd section visible
            If ActiveCell.Row = 4 Then
                If Rows(35).Hidden = True Then
                    ActiveSheet.ChartObjects("Chart 1").Visible = True
                    ActiveSheet.ChartObjects("Chart 2").Visible = True
                    ActiveSheet.ChartObjects("Chart 3").Visible = True
                    ActiveSheet.ChartObjects("Chart 4").Visible = True
                    ActiveSheet.ChartObjects("Chart 5").Visible = True
                    ActiveSheet.Shapes("Textbox 1").Visible = True
                    Rows("34:108").EntireRow.Hidden = False
                    GoTo ext
                Else
                    'make 2nd section visible
                    ActiveSheet.ChartObjects("Chart 5").Visible = False
                    ActiveSheet.Shapes("Textbox 1").Visible = False
                    Rows("58:108").EntireRow.Hidden = True
                    Rows("4:34").EntireRow.Hidden = True
                    Range("R35").Select
                    GoTo ext
                End If
            End If

            'make 3rd section visible
            If ActiveCell.Row = 35 Then
                Range("R59").Select
                Rows("4:108").EntireRow.Hidden = True
                Rows("59:73").EntireRow.Hidden = False
                ActiveSheet.Shapes("Textbox 1").Visible = True
                ActiveSheet.ChartObjects("Chart 5").Visible = True
                ActiveSheet.ChartObjects("Chart 4").Visible = False
                ActiveSheet.ChartObjects("Chart 3").Visible = False
                ActiveSheet.ChartObjects("Chart 2").Visible = False
                ActiveSheet.ChartObjects("Chart 1").Visible = False
                GoTo ext
            End If

            'make 4th section visible
            If ActiveCell.Row = 59 Then
                Range("R75").Select
                Rows("4:108").EntireRow.Hidden = True
                Rows("75:91").EntireRow.Hidden = False
                ActiveSheet.Shapes("Textbox 1").Visible = False
                ActiveSheet.ChartObjects("Chart 5").Visible = False
                GoTo ext
            End If

            'make 1st section visible
            If ActiveCell.Row = 75 Then
                ActiveSheet.ChartObjects("Chart 5").Visible = False
                ActiveSheet.Shapes("Textbox 1").Visible = False
                Rows("4:108").EntireRow.Hidden = True
                Rows("4:33").EntireRow.Hidden = False
                Application.Goto Range("A1"), True
                Range("R4").Select
                GoTo ext
            End If

        End If

        'goto monthly sheet if it contains data
        If Len(i) = 3 Then
            Application.Goto Range("P32")
            If Sheets(i).Range("C4") <> "" Then
                Sheets(i).Unprotect
                Application.Goto Sheets(i).Range("A1"), True
                Application.Goto Sheets(i).Range("C4")
                Sheets(i).Protect
                GoTo ext
            End If
        End If

        GoTo ext

    End If

ext:

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub Order()
's1 - s12 & Codes, sorts transactions

    'check if data is present
    If Range("C4") = "" Then ActiveCell.Offset(1).Select: Exit Sub

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'remove scrollarea limitation
    ActiveSheet.ScrollArea = ""

    If ActiveSheet.index < 13 Then

        'if request to sort charges
        If Split(ActiveCell.Address, "$")(1) = "C" Then

            'acquire last row of charge column
            i = Cells(Rows.Count, "C").End(xlUp).Row

            'sort charge column
            ActiveSheet.Range("B4:H" & i).Sort Key1:=Range("C4"), Order1:=xlAscending

            'rebuild charge list hlinks w/global var
            i = "Sheets(""" & ActiveSheet.Name & """).Range(""D4:D" & i & """)"
            Call SentLink
            Application.Goto Range("C4")

            'set scroll area
            Call SetScrollArea

        End If

        'if request to sort expenses
        If Split(ActiveCell.Address, "$")(1) = "O" Then

            'acquire last row of expense column
            i = Cells(Rows.Count, "O").End(xlUp).Row + 1

            'sort expense column
            ActiveSheet.Range("O4:T" & i).Sort Key1:=Range("O4"), Order1:=xlAscending

            'rebuild expense list hlinks w/global var
            i = "Sheets(""" & ActiveSheet.Name & """).Range(""P4:P103"")"
            Call SentLink

            'set cursor
            Application.Goto Range("O4")

            'set scroll area
            Call SetScrollArea

        End If

    ElseIf ActiveSheet.Name = "Codes" Then


        'set cursor
        Application.ScreenUpdating = True
        Application.Goto Range("Transfer").Offset(1, 1)
        Application.ScreenUpdating = False

        'reset coded transaction list
        For Each i In Range("B4:C1003")
            If Len(Trim(i)) = 0 Then Range(i.Address) = ""
        Next

        'sort coded transaction list
        Sheets("Codes").Range("B4:C1003").Sort Key1:=Range("C4"), Order1:=xlAscending
    
        'rebuild coded transaction list
        i = "Sheets(""Codes"").Range(""B4:C1003"")"     'global var to pass to macro
        Call SentLink

        'reset spaces for empty cells
        For Each i In Range("B4:C1003")
            If Len(Trim(i)) = 0 Then Range(i.Address) = Space(50)
        Next

        'set scroll area
        Call SetScrollArea

        'reset start settings
        Call View

    End If

    'restore settings
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic

End Sub

Sub PageFlip()
's1 - s12, flips between charge & expense sections

    Application.ScreenUpdating = False

    i = ActiveCell.Address
    t = ActiveSheet.Name

    If Split(i, "$")(1) = "J" Then

        Sheets(t).ScrollArea = ""

        Application.Goto Range("N1"), True
        Application.Goto Range("O4")

        Call SetScrollArea

    End If

    If Split(i, "$")(1) = "V" Then

        Application.Goto Range("A1"), True
        Application.Goto Range("C4")

    End If

    Application.ScreenUpdating = True

End Sub

Sub Protects()
'sx, (un)protect sheet

    If Split(ActiveCell.Address, "$")(1) = "E" Then

        Application.Goto Range("A1"), True
        Application.Goto Range("C4")

    End If

    If Split(ActiveCell.Address, "$")(1) = "Q" Then

        Application.Goto Range("N1"), True
        Application.Goto Range("O4")

    End If

    If ActiveSheet.ProtectContents = True Then
        ActiveSheet.Unprotect
    Else
        ActiveSheet.Protect
    End If

End Sub

Sub Reset()
's1 - s12, delete all data

    'exit if data missing
    For i = 1 To 12
        t = t + Len(Trim(Sheets(i).Range("C4")))
    Next
    If t = 0 Then Exit Sub

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'if called after creating new workbook skip confirmation
    If Not i Then

        'confirm deletion
        i = MsgBox(Space(5) & "Do you want to delete ALL of the data for the entire year?" & Chr(10) _
                    & Chr(9) & Space(19) & "This cannot be undone.", 4, "Complete Reset")
        If i <> 6 Then GoTo ext

    End If
    
    'delete all data from sheets
    For i = 12 To 1 Step -1

        Sheets(i).Unprotect
        Sheets(i).Range("B4:C203,E4:H203,O4:O203,Q4:T203").ClearContents
        Application.Goto Sheets(i).Range("C4")

        'refresh transaction code formula for sheets 1 - 12
        Sheets(i).Range("D4:D203").Formula = Sheets("Codes").Range("Form1").Formula
        Sheets(i).Range("P4:P203").Formula = Sheets("Codes").Range("Form2").Formula

        'change indicator format
        Sheets(i).Range("M3").Font.Size = 6
        Sheets(i).Range("M3").Font.Name = "Wingdings"

        Sheets(i).Protect

    Next

    'rebuild hlinks on monthly sheets
    Call AutoHyperlink

    'clear Query results
    Sheets("Query").Range("C5:J5").ClearContents
    Sheets("Query").Range("C6:J2005").ClearContents

    'clear Transaction list
    Sheets("Codes").Range("I4:I102").ClearContents

    'clear Data table
    Sheets("Data").Range("C5:I2404").ClearContents

    'clear Limit list
    Sheets("Limit").Range("A6:E205").ClearContents
    Sheets("Limit").Range("G6:J205").ClearContents
    Sheets("Limit").Range("L6:M205").ClearContents

    'alert user of successful refresh
    MsgBox Space(6) & "This workbook has been successfully refreshed." & Chr(10) _
        & Space(14) & "All data has been successfully cleared.", 0, "Success"

ext:

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub SetScrollArea()
'sx, set scroll area

    Dim c, i, n

    With ActiveSheet

        i = .index
        n = .Name

        'acquire last entry
        For Each c In .Range("C10:C2404")
            If Trim(c) = "" Then c = c.Row: Exit For
        Next

        If i < 13 Then .ScrollArea = "A1:X" & c
        If n = "Sum" Then .ScrollArea = "A1:R91"
        If n = "View" Then .ScrollArea = "A1:T44"
        If n = "Query" Then .ScrollArea = "A1:O" & c
        If n = "Codes" Then .ScrollArea = "A1:AJ" & c
        If n = "Items" Then .ScrollArea = "A1:CT" & c
        If n = "Limit" Then .ScrollArea = "A1:M" & c
        If n = "Data" Then .Sheets(i).Unprotect

        If i < 13 And c < 34 Then .ScrollArea = "A1:X34"
        If n = "Query" And c < 30 Then .ScrollArea = "A1:O30"

    End With

End Sub

Sub Switch()
's1 - s12 sheets
'determines macro required

    If ActiveSheet.Range("C4") = "" And ActiveSheet.Range("P4") = "" And ActiveSheet.Name <> "Jan" Then
        Call Acquire: Exit Sub
    End If

    If ActiveSheet.Name = "Jan" Then

        i = MsgBox(Space(4) & "Do you want to create a new workbook or handle data?" & Chr(10) _
                    & Chr(9) & Space(11) & "[ create = yes  |  data = no ]", 3, "Create / Handle")
        If i = 2 Then Exit Sub
        If i = 6 Then
            Call CreateCopy: If Not i Then Exit Sub
            Call Reset: Exit Sub
        End If

    End If

    i = MsgBox(Space(6) & "Do you want to acquire data or delete data?" & Chr(10) _
                    & Chr(9) & Space(2) & "[ acquire = yes  |  delete = no ]", 3, "Acquire / Delete")
    If i = 2 Then Exit Sub
    If i = 6 Then Call Acquire: Exit Sub
    If i = 7 And ActiveSheet.Name <> "Jan" Then Call Clear: Exit Sub

    If i = 7 And ActiveSheet.Name = "Jan" Then

        i = MsgBox(Space(4) & "Do you want to delete workbook data or sheet data?" & Chr(10) _
                    & Chr(9) & Space(9) & "[ book = yes  |  sheet = no ]", 3, "Book / Sheet")
        If i = 2 Then Exit Sub
        If i = 6 Then Call Reset: Exit Sub
        If i = 7 Then Call Clear: Exit Sub

    End If

End Sub

Sub Toggle()
's1 - s12, Codes sheets, toggles ribbon

    Application.ScreenUpdating = False

    If ActiveSheet.index < 13 Then

        If Split(ActiveCell.Address, "$")(1) = "M" Then

            Application.Goto Range("A1"), True
            Application.Goto Range("C4")

        End If

        If Split(ActiveCell.Address, "$")(1) = "N" Then

            Application.Goto Range("N1"), True
            Application.Goto Range("O4")

        End If

        Call Indicator

    End If

    If ActiveSheet.Name = "Codes" Then

        Application.Goto Range("A1"), True
        Application.Goto Range("Transfer").Offset(1, 1)

    End If

    If ActiveSheet.Name = "Data" Then

        Application.Goto Range("A1"), True
        Application.Goto Range("J3")
        ActiveSheet.Unprotect

    End If

    If ActiveSheet.Name = "Items" Then

        Application.Goto Range("A1"), True
        Application.Goto Range("B6")

    End If

    If Application.CommandBars("Ribbon").Visible = True And _
       ActiveSheet.ProtectContents = False Then

        ActiveWindow.DisplayHeadings = False
        Application.DisplayFormulaBar = False
        Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",False)"
        ActiveSheet.Protect

    ElseIf Application.CommandBars("Ribbon").Visible = True And _
       ActiveSheet.ProtectContents = True Then

       ActiveSheet.Unprotect

    Else

        Application.ExecuteExcel4Macro "show.toolbar(""Ribbon"",True)"
        If CommandBars("Ribbon").Height > 100 Then CommandBars.ExecuteMso "MinimizeRibbon"
        ActiveSheet.Unprotect

    End If

    If ActiveSheet.Name = "Data" Then ActiveSheet.Unprotect

    Application.ScreenUpdating = True

End Sub

Sub View()
'sx, scrolls to top of sheet & checks for transaction duplicates

    Dim addrs(), dates(), dscrp(), price(), temps()
    Dim a, b, c, d

    'basic setup to avoid unnecessary delays
    'Application.ScreenUpdating = False

    'reset each sheet to A1
    If Sheets("Codes").Range("E31") = "Automatic" Then

       'range of visible cells
        a = Windows(1).VisibleRange.Cells.Address

        'select cell
        If ActiveSheet.index < 13 Then

            Application.Goto Range("A1"), True
            Application.Goto Range("C4")

        ElseIf ActiveSheet.Name = "Sum" Then

            If ActiveCell.Address(0, 0) <> "T32" Then
                ActiveSheet.ChartObjects("Chart 1").Visible = True
                ActiveSheet.ChartObjects("Chart 2").Visible = True
                ActiveSheet.ChartObjects("Chart 3").Visible = True
                ActiveSheet.ChartObjects("Chart 4").Visible = True
                ActiveSheet.ChartObjects("Chart 5").Visible = True
                ActiveSheet.Shapes("Textbox 1").Visible = True
                Rows("1:110").EntireRow.Hidden = False
                Application.Goto Range("A1"), True
                Application.Goto Range("P32")
            End If

            Exit Sub

        ElseIf ActiveSheet.Name = "View" Then

            Application.Goto Range("A1"), True
            Application.Goto Range("O17")

            Exit Sub

        ElseIf ActiveSheet.Name = "Query" Then

            Application.Goto Range("A1"), True
            Application.Goto Range("L5")

        ElseIf ActiveSheet.Name = "Codes" Then

            Application.Goto Range("A1"), True
            Application.Goto Range("Transfer").Offset(1, 1)
    
            Exit Sub

        ElseIf ActiveSheet.Name = "Data" Then

            Application.Goto Range("A1"), True
            Application.Goto Range("J3")

            Exit Sub

        ElseIf ActiveSheet.Name = "Items" Then

            Application.Goto Range("A1"), True
            Application.Goto Range("B6")

            Exit Sub

        ElseIf ActiveSheet.Name = "Limit" Then

            Application.Goto Range("A1"), True
            Application.Goto Range("E6")

            Exit Sub

        End If

    End If

    'duplicate check - only monthly sheets
    If ActiveSheet.index < 13 Then

        'acquire data
        dates = Range("C4:C203")
        dscrp = Range("E4:E203")
        price = Range("F4:F203")

        'acquire number of rows in range
        c = Range("C4:C203").Rows.Count

        'redimension address array
        ReDim addrs(c)
        ReDim temps(c)

        'populate addresses array
        For Each a In Range("C4:C203")
            b = b + 1
            addrs(b) = a.Address
        Next

        'check for duplicates
        For a = 1 To c
            For b = 1 To c
                If dates(a, 1) = dates(b, 1) And dscrp(a, 1) = dscrp(b, 1) And price(a, 1) = price(b, 1) _
                 And dates(a, 1) <> "" And dscrp(a, 1) <> "" And price(a, 1) <> "" And addrs(a) <> addrs(b) Then
                    If IsError(Application.Match(dscrp(a, 1), temps, False)) Then    'if not in array already continue
                        'Debug.Print a; price(a, 1) & " - " & dscrp(a, 1) & " - " & dates(a, 1)
                        d = d + 1
                        temps(d) = dscrp(a, 1)
                    End If
                End If
            Next
        Next

        'paste to indicate if duplicates found
        If d > 0 Then Range("F3") = "Purchase" Else Range("F3") = "Purchases"

    End If

    'restore settings
    'Application.ScreenUpdating = True

End Sub
