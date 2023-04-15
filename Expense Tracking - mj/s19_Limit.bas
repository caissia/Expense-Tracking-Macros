Attribute VB_Name = "s19_Limit"
Option Explicit

Sub Limits()
'Limit, acquires mid-month transactions to check spending limits

    Dim a, b, c
    Dim file, path
    Dim temp, Target

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'remove scrollarea limitation
    Sheets("Limit").ScrollArea = ""

    'clear transfer range
    Range("C6:E205,G6:J205,L6:N205").ClearContents

    'acquire range of wb & ws to receive data
    Set a = ThisWorkbook.Sheets("Limit").Range("E6:E205")
    
    'acquire filepath of temp statements
    path = "C:\Users\imagine\Documents\personal\finances\credit card\misc\temp\"

    For i = 1 To 10

        'acquire file name of statement
        file = "temp" & i & ".xlsx"

        'check if next statement exists
        On Error Resume Next
        Set b = Workbooks.Open(path & file)
        On Error GoTo 0

        'return focus to this wb
        ThisWorkbook.Activate
    
        If Not b Is Nothing Then

            'reuse variable as integer
            b = 0

            'acquire the last empty row for data transfer
            For Each c In a
                If IsEmpty(c) Then
                    Set Target = Sheets(a.Parent.Name).Range(c.Address)
                    Exit For
                End If
            Next

            'acquire name of temp ws to process
            Set temp = Workbooks("temp" & i & ".xlsx").Sheets(1)
    
            'check cell A1 to determine processing module
            On Error Resume Next
            With temp
    
                If Left(.Range("A1"), 7) = "Account" Then
    
                    GoSub Module1
    
                ElseIf .Range("A1") = "Description" Then
    
                    GoSub Module2
    
                ElseIf .Range("A1") = "Status" Then
    
                    GoSub Module3
    
                ElseIf .Range("A1") = "Type" Then
    
                    GoSub Module4
    
                ElseIf Not IsEmpty(.Range("A1")) And IsDate(.Range("A1")) Then
    
                    If Not IsError(Day(.Range("A1"))) Then
    
                    GoSub Module5
                    
                    End If
    
                End If
    
            End With

        'close temp wb
        Workbooks(file).Close
    
        End If

    'set for next statement
    Set b = Nothing

    Next

    'return error handling
    On Error GoTo 0

    'remove wraptext
    Range(a.Address & ":" & a.Offset(0, 4).Address).WrapText = False

    'remove negative values
    For Each c In Range(a.Offset(0, 3).Address & "," & a.Offset(0, 4).Address)
        If Not IsEmpty(c) And Trim(c) <> "" Then Range(c.Address) = Abs(c)
    Next

    're-calc to move income to credit
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.Calculation = xlManual
    Application.ScreenUpdating = False

    'move income to credit column
    For Each c In a
        If c.Offset(0, 1) = "I" And Trim(c.Offset(0, 3)) <> "" Then
            Range(c.Offset(0, 4).Address).Value = Range(c.Offset(0, 3).Address).Value
            Range(c.Offset(0, 3).Address).ClearContents
        End If
    Next

    'add blanks to prevent text spillover
    For Each c In Range(a.Offset(0, 3).Address)
        If Trim(c.Offset(0, -1)) <> "" And Trim(c) = "" Then c.Value = Space(4)
    Next

    'remove transactions listed in watch list of Codes sh
    GoSub Watch

    'add space to move it away from column edge
    For Each c In a.Offset(0, 2)
        If Trim(c) <> "" Then c.Value = Space(2) & c
    Next

    'sort by date
    Set b = Range(a.Offset(0, -2).Address & ":" & a.Offset(0, 4).Address)
    b.Sort Key1:=a, Order1:=xlAscending

    'refresh formula
    c = "=IFERROR(IF(Trim(G6)="""","""",IF(AND(L6<>"""",LEN(L6)=1),L6,IFERROR(IFERROR(IFERROR(IFERROR(IFERROR(IFERROR(IFERROR(IFERROR(IFERROR(" & Chr(10) _
        & "INDEX(Code,MATCH(""*""&LEFT(Trim(G6),20)&""*"",Transact,0)),INDEX(Code,MATCH(""*""&LEFT(Trim(G6),8)&""*"",Transact,0))),INDEX(Code,MATCH(""*""&LEFT(Trim(G6),7)&""*"",Transact,0)))," & Chr(10) _
        & "INDEX(Code,MATCH(""*""&LEFT(Trim(G6),6)&""*"",Transact,0))),INDEX(Code,MATCH(""*""&LEFT(Trim(G6),5)&""*"",Transact,0))),INDEX(Code,MATCH(""*""&LEFT(Trim(G6),4)&""*"",Transact,0)))," & Chr(10) _
        & "INDEX(Code,MATCH(""*""&LEFT(Trim(G6),3)&""*"",Transact,0))),INDEX(Code,MATCH(""*""&LEFT(Trim(G6),2)&""*"",Transact,0))),INDEX(Code,MATCH(""*""&RIGHT(Trim(G6),6)&""*"",Transact,0))),""""))),"""")"

    Sheets("Limit").Range("F6:F205").Formula = c

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    On Error GoTo 0

    'reinstate scrollarea limitation
    Call SetScrollArea

Exit Sub

Module1:

    'transfer the transactions
    For Each c In temp.Range("B1:B200")
        If Not IsEmpty(c) And IsDate(c) Then
            If Not IsError(Day(c)) Then
                Target.Offset(b, -2).Value = "La Loma FCU"
                Target.Offset(b, 0).Value = c
                Target.Offset(b, 2).Value = c.Offset(0, 2)
                Target.Offset(b, 3).Value = c.Offset(0, 3)
                Target.Offset(b, 4).Value = c.Offset(0, 4)
                If Trim(c.Offset(0, 2)) = "" Then Target.Offset(b, 2).Value = c.Offset(0, 1)
                b = b + 1
            End If
        End If
    Next

    b = 0
    Return

Module2:

    'transfer the transactions
    For Each c In temp.Range("A1:A200")
        If Not IsEmpty(c) And IsDate(c) Then
            If Not IsError(Day(c)) And c.Offset(0, 2) <> "" Then
                Target.Offset(b, -2).Value = "Bank of America"
                Target.Offset(b, 0).Value = c
                Target.Offset(b, 2).Value = c.Offset(0, 1)
                Target.Offset(b, 3).Value = c.Offset(0, 2)
                If (c.Offset(0, 2)) > 0 Then
                    Target.Offset(b, 3).Value = c.Offset(0, 2)
                    Target.Offset(b, 4).Value = Space(4)
                End If
                b = b + 1
            End If
        End If
    Next

    b = 0
    Return

Module3:

    'transfer the transactions
    For Each c In temp.Range("B1:B200")
        If Not IsEmpty(c) And IsDate(c) Then
            If Not IsError(Day(c)) Then
                Target.Offset(b, -2).Value = "Citibank"
                Target.Offset(b, 0).Value = c
                Target.Offset(b, 2).Value = c.Offset(0, 1)
                Target.Offset(b, 3).Value = c.Offset(0, 2)
                Target.Offset(b, 4).Value = c.Offset(0, 3)
                b = b + 1
            End If
        End If
    Next

    b = 0
    Return

Module4:

    'transfer the transactions
    For Each c In temp.Range("B1:B200")
        If Not IsEmpty(c) And IsDate(c) Then
            If Not IsError(Day(c)) Then
                Target.Offset(b, -2).Value = "Chase Sapphire"
                Target.Offset(b, 0).Value = c
                Target.Offset(b, 2).Value = c.Offset(0, 2)
                Target.Offset(b, 3).Value = c.Offset(0, 3)
                Target.Offset(b, 4).Value = c.Offset(0, 4)
                If Not IsNumeric(c.Offset(0, 3)) Then Target.Offset(b, 3).Value = c.Offset(0, 4)
                If (c.Offset(0, 3)) > 0 Then Target.Offset(b, 4).Value = c.Offset(0, 3): Target.Offset(b, 3).Value = Space(4)
                b = b + 1
            End If
        End If
    Next

    b = 0
    Return

Module5:

    'transfer the transactions
    For Each c In temp.Range("A1:A200")
        If Not IsEmpty(c) And c <> "" Then
            Target.Offset(b, -2).Value = "Barclay"
            Target.Offset(b, 0).Value = c
            Target.Offset(b, 2).Value = c.Offset(0, 1)
            Target.Offset(b, 3).Value = c.Offset(0, 2)
            Target.Offset(b, 4).Value = c.Offset(0, 3)
            b = b + 1
        End If
    Next

    b = 0
    Return

Watch:

    'delete transactions on watch list
    For Each c In a.Offset(0, 2)
        If c <> "" And Len(c) > 0 Then
            For Each b In Sheets("Codes").Range("N4:N103")
                If Len(Trim(b)) <> 0 Then
                    i = Len(b)
                    If UCase(b) = UCase(Left(c, i)) Then
                        Sheets("Limit").Range("C" & c.Row & ":E" & c.Row).ClearContents
                        Sheets("Limit").Range("G" & c.Row & ":I" & c.Row).ClearContents
                    End If
                End If
            Next
        End If
    Next

    'delete other transactions not listed in watch list
    For Each c In a.Offset(0, 2)
        If c <> "" And Len(c) > 0 Then
            If Left(UCase(c), 5) = Left(UCase("Check"), 5) Or _
               Left(UCase(c), 5) = Left(UCase("BKOFAMERICA"), 5) Then
                Sheets("Limit").Range("C" & c.Row & ":E" & c.Row).ClearContents
                Sheets("Limit").Range("G" & c.Row & ":I" & c.Row).ClearContents
            End If
        End If
    Next

    Return

End Sub
