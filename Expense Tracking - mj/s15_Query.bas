Attribute VB_Name = "s15_Query"
Option Explicit
'Contains 2 macros: Query, Result

Sub Query()
'Query, data search autofilters data

    Dim c       'last row of transactions to copy / loop var
    Dim p       'last row Data sheet to paste
    Dim n       'name of sheets for loop
    Dim q(8)    'query search terms

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Sheets("Query").ScrollArea = ""
    Sheets("Query").Unprotect

    With Sheets("Query")

        'exit if no query terms
        If Application.CountA(.Range("C5:J5")) = 0 Then GoTo ext
        
        'acquire query data
        For i = 1 To 8
            q(i) = .Range("B5").Offset(0, i)
        Next

        'delete residual query results
        .Range("C6:J1005").ClearContents

    End With

    'use auto-filter ~~> copy/paste to Query sheet
    On Error GoTo err
    With Sheets("Data")
        .AutoFilterMode = False
        .Range("C4:J2404").AutoFilter

        For i = 1 To 8

            If i > 2 Then p = i - 1 Else p = i              'skips 3 field since query is extra column larger than data set
            If q(i) <> "" Then

                If p = 1 Or p = 3 Or p = 4 Or p = 7 Then    'text
                    .Range("C5:J2404").AutoFilter Field:=p, Criteria1:="=*" & q(i) & "*", Operator:=xlAnd
                End If

                If p = 2 Then                               'dates
                    If i = 2 And q(i) <> "" And q(i + 1) <> "" Then     'between dates
                        .Range("C5:J2404").AutoFilter Field:=2, Criteria1:=">=" & q(i), Operator:=xlAnd, Criteria2:="<=" & q(i + 1)
                        i = i + 1
                    End If
                    If i = 2 And q(i) <> "" And q(i + 1) = "" Then      'after date
                        .Range("C5:J2404").AutoFilter Field:=2, Criteria1:=">=" & q(i), Operator:=xlAnd
                        i = i + 1
                    End If
                    If i = 3 And q(i - 1) = "" And q(i) <> "" Then      'before date
                        .Range("C5:J2404").AutoFilter Field:=2, Criteria1:="<=" & q(i), Operator:=xlAnd
                        i = i + 1
                    End If
                End If

                If p = 5 Then                               'money  - credit
                        .Range("C5:J2404").AutoFilter Field:=p, Criteria1:=">=" & q(i) - 1, Operator:=xlAnd, Criteria2:="<=" & q(i) + 1
                End If
                If p = 6 Then                               'money  - debit
                        .Range("C5:J2404").AutoFilter Field:=p, Criteria1:=">=" & q(i) - 1, Operator:=xlAnd, Criteria2:="<=" & q(i) + 1
                End If

            End If

        Next

        'unmerge paste area
        Sheets("Query").Range("D6:E2005").UnMerge
        p = Sheets("Data").Range("D" & .Rows.Count).End(xlUp).Row

        'paste results
        If p > 4 Then

            On Error GoTo ext
            Sheets("Data").Range("C5:D" & p).SpecialCells(xlVisible).Copy
            Sheets("Query").Range("C6").PasteSpecial xlPasteValues

            Sheets("Data").Range("E5:I" & p).SpecialCells(xlVisible).Copy
            Sheets("Query").Range("F6").PasteSpecial xlPasteValues
            On Error GoTo 0

            'uninitialize autofilter
            .AutoFilterMode = False
    
            'sort results based on date
            p = .Range("D" & .Rows.Count).End(xlUp).Row
            .Range("C5:J" & p).Sort Key1:=.Range("D5"), Order1:=xlAscending

        End If

    End With

    'sort, re-merge, and rebuild conditional formatting
    With Sheets("Query")

        'sort results based on date
        c = .Range("D" & .Rows.Count).End(xlUp).Row
        If c > 6 Then .Range("C6:J" & c).Sort Key1:=.Range("D6"), Order1:=xlAscending

        're-merge paste area
        For Each c In Range("D6:D2005")
            .Range(c.Address & ":" & c.Offset(0, 1).Address).Merge
        Next

        'remove conditional formatting
        .Range("C6:J2005").FormatConditions.Delete

'        '(1) fill in cells if data present
'        With .Range("C6:J2005").FormatConditions.Add(Type:=xlExpression, Formula1:="=COUNTA($C6:$J6)>0")
'            .Interior.Color = rgb(231, 228, 213)
'            .StopIfTrue = False
'        End With
'
'        '(2) create left border
'        With .Range("C6:C2005").FormatConditions.Add(Type:=xlExpression, Formula1:="=COUNTA($C6:$J6)>0")
'            .Borders(xlEdgeLeft).Color = rgb(196, 189, 151)
'            .StopIfTrue = False
'        End With
'
'        '(3) create right border
'        With .Range("J6:J2005").FormatConditions.Add(Type:=xlExpression, Formula1:="=COUNTA($C6:$J6)>0")
'            .Borders(xlEdgeRight).Color = rgb(196, 189, 151)
'            .StopIfTrue = False
'        End With

        '(4) create bottom border
        With .Range("C6:J2005").FormatConditions.Add(Type:=xlExpression, Formula1:="=COUNTA($C6:$J6)>0")
            .Borders(xlBottom).Color = rgb(196, 189, 151)
            .StopIfTrue = False
        End With

    End With

    'freeze panes for scrolling
    If Range("C6") <> "" Then
        ActiveWindow.FreezePanes = False
        Rows("6:6").Select
        ActiveWindow.FreezePanes = True
    End If

    'set focus on summary
    Application.Goto Sheets("Query").Range("N5")

    'reset scroll area
    Call SetScrollArea

ext:

    'restore settings
    Sheets("Data").AutoFilterMode = False
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Sheets("Query").Protect
    On Error GoTo 0

Exit Sub

err:

    Sheets("Data").AutoFilterMode = False
    On Error GoTo 0
    GoTo ext

End Sub

Sub Result()
'Query, uppercases search terms, clears search terms, or clears results

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    If InStr(i, "$") <> 0 Then

        'check for error
        If Range(i).Count > 1 Then GoTo ext
        If Range(i) = "" Or IsNumeric(Range(i)) Then GoTo ext
        If Range(i) = UCase(Range(i)) Then GoTo ext

        'uppercase the search term
        Range(i) = UCase(Range(i))

        'reset global variable
        i = ""

    ElseIf ActiveCell.Address(0, 0) = "L6" Then

        'clear search terms
        i = Application.CountA(Range("C5:J5"))
        If i > 0 Then Sheets("Query").Range("C5:J5").ClearContents
        If Range("C6") = "" Then
            Application.Goto Range("L5")
        Else
            Application.Goto Range("L7")
        End If

        'reset scroll area
        Call SetScrollArea

    ElseIf ActiveCell.Address(0, 0) = "L7" Then

        'clear data
        i = Application.CountA(Range("C6:J2005"))
        If i > 0 Then Sheets("Query").Range("C6:J2005").ClearContents
        If Application.CountA(Range("C5:J5")) = 0 Then
            Application.Goto Range("L5")
        Else
            Application.Goto Range("L6")
        End If

        'reset scroll area
        Sheets("Query").ScrollArea = "A1:L6"

        'unfreeze panes for new search
        ActiveWindow.FreezePanes = False

    End If

ext:

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
