Attribute VB_Name = "sx_Acquire"
Option Explicit

Sub Acquire()
's1 - s12, acquire/transfer data from statements

    Dim y, m, d
    Dim a, b, c, f, i, n
    Dim aa, bb, lc, ls, cs, ct, jj, jm, pm, ps, sk, yr
    Dim FilePath, Found, Last, Remove, Start, Total

    'acquire month to process
    i = InputBox(Space(28) & "What month is to be processed?", "Month", "[ 1 - 12 ]")
    If Not IsNumeric(i) Or i < 1 Or i > 12 Then Exit Sub
    i = i * 1   'avoids mismatch type error

    'warn if month has data
    If Sheets(i).Range("C4") <> "" Or Sheets(i).Range("P4") <> "" Then
        c = MsgBox(Space(5) & "The month requested - " & MonthName(i) & " - appears to have been processed." & Chr(10) _
                & Space(39) & "Overwrite existing data?", 4, "Warning")
        If c <> 6 Then Exit Sub
    End If

    'remove scrollarea limit
    Sheets(i).ScrollArea = ""

    'select monthly sheet to process
    If ActiveSheet.Name <> Sheets(i).Name Then
        Application.ScreenUpdating = False
        Sheets(i).Activate
        Application.Goto Sheets(i).Range("A1"), True
        Application.Goto Sheets(i).Range("C4")
        Application.ScreenUpdating = True
    End If

    'check year for file path
    yr = Year(Date)
    If Month(Date) > 0 And Month(Date) < 4 Then
        c = MsgBox(Space(8) & "What statement year is to be processed?" & Chr(10) _
                   & Chr(9) & Space(5) & "[ " & Year(Date) - 1 & " = yes  |  " & Year(Date) & " = no ]", 3, "File Path")
        If c = 2 Then Exit Sub
        If c = 6 Then yr = Year(Date) - 1
    End If

    'acquire file path
    FilePath = "C:\Users\imagine\Documents\personal\finances\credit card\" & yr & "\Statements\" & i & "." & Right(yr, 2)

    'acquire prefix of workbook titles
    f = MonthName(i, True) & Right(yr, 2)

    'assign file names to vars
    aa = ActiveWorkbook.Name                'this workbook
    bb = f & " Boa.xlsx"                    'bank of america statement
    lc = f & " Llm C.xlsx"                  'la loma credit statement
    ls = f & " Llm S.xlsx"                  'la loma saving statement
    ct = f & " Citi.xlsx"                   'citibank - jm statement
    jj = f & " Jet J.xlsx"                  'jetblue - j statement
    jm = f & " Jet M.xlsx"                  'jetblue - m statement
    pm = f & " App M.xlsx"                  'apple   - m statement
    ps = f & " App S.xlsx"                  'apple   - s statement
    cs = f & " Chs.xlsx"                    'chase   - m statement (m, j, s)

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'clear transaction data of chosen month & unmatched transactions
    Sheets(i).Range("B4:C103,E4:H103,O4:O103,Q4:T103").ClearContents
    Sheets("Codes").Range("I4:I103") = Space(50)

    'create temp sheet
    ThisWorkbook.Sheets.Add(After:=Sheets(Sheets.Count)).Name = "Temp"

    'set space for msg to user
    n = 19

'~~> #1: acquire data for expenses/income

    'set variable to nothing for next statement
    Set c = Nothing

    'check if BoA statement exists
    On Error Resume Next
    Set c = Workbooks.Open(FilePath & "\" & bb)
    On Error GoTo 0

    If Not c Is Nothing Then

        'open bank statements
        Workbooks.Open (FilePath & "\" & bb)        'bank of america statement

        With Workbooks(bb)
            For Each c In .Sheets(1).Range("B:B")
                If Left(c, 3) = "Beg" Then
                    Start = c.Offset(1, -1).Address
                End If
                If c.Row > 10 And c = "" Then
                    Last = c.Offset(-1, 1).Address
                   Exit For
                End If
            Next
        End With

        'acquire number of rows and acquire data
        a = Range(Start & ":" & Last).Rows.Count
        Workbooks(aa).Sheets("Temp").Range("A1:C" & a).Value = Workbooks(bb).Sheets(1).Range(Start & ":" & Last).Value

        'close workbooks
        Workbooks(bb).Close False                   'bank of america statement

        'add to total for error check
        Total = Total + a

        'save the statement processed in order to alert user
        sk = Space(n) & "Bank of America Checking"

    End If

    'set variable to nothing for next statement
    Set c = Nothing

    'check if LLFCU checking statement exists
    On Error Resume Next
    Set c = Workbooks.Open(FilePath & "\" & lc)
    On Error GoTo 0

    If Not c Is Nothing Then

        'open bank statement
        Workbooks.Open (FilePath & "\" & lc)        'la loma checking statement

        With Workbooks(lc)
            For Each c In .Sheets(1).Range("B:B")
                If c = "Date" Then
                    Start = c.Offset(1).Address
                End If
                If c.Row > 5 And c = "" Then
                    Last = c.Offset(-1, 4).Address
                   Exit For
                End If
            Next
        End With

        'acquire number of rows and acquire data
        b = Range(Start & ":" & Last).Rows.Count
        Workbooks(aa).Sheets("Temp").Range("A" & a + 1 & ":A" & a + b).Value = Workbooks(lc).Sheets(1).Range(Start & ":" & Range(Start).Offset(b).Address).Value
        Workbooks(aa).Sheets("Temp").Range("B" & a + 1 & ":D" & a + b).Value = Workbooks(lc).Sheets(1).Range(Range(Start).Offset(0, 2).Address & ":" & Last).Value

        'close workbook
        Workbooks(lc).Close                         'la loma checking statement

        'add to total for error check
        Total = Total + b

        'save the statement processed in order to alert user
        If sk = "" Then sk = Space(n) & "La Loma Checking FCU" Else sk = sk & Chr(10) & Space(n) & "La Loma Checking FCU"

    End If

    'set variable to nothing for next statement
    Set c = Nothing

    'check if LLFCU savings statement exists
    On Error Resume Next
    Set c = Workbooks.Open(FilePath & "\" & ls)
    On Error GoTo 0

    If Not c Is Nothing Then

        'open bank statement
        Workbooks.Open (FilePath & "\" & ls)        'la loma savings statement

        With Workbooks(ls)
            For Each c In .Sheets(1).Range("B:B")
                If c = "Date" Then
                    Start = c.Offset(1).Address
                End If
                If c.Row > 5 And c = "" Then
                    Last = c.Offset(-1, 4).Address
                   Exit For
                End If
            Next
        End With

        'acquire number of rows and acquire data
        a = a + b
        b = Range(Start & ":" & Last).Rows.Count
        Workbooks(aa).Sheets("Temp").Range("A" & a + 1 & ":A" & a + b).Value = Workbooks(ls).Sheets(1).Range(Start & ":" & Range(Start).Offset(b).Address).Value
        Workbooks(aa).Sheets("Temp").Range("B" & a + 1 & ":D" & a + b).Value = Workbooks(ls).Sheets(1).Range(Range(Start).Offset(0, 2).Address & ":" & Last).Value

        'close workbook
        Workbooks(ls).Close                         'la loma savings statement

        'add to total for error check
        Total = Total + b

        'save the statement processed in order to alert user
        If sk = "" Then sk = Space(n) & "La Loma Savings FCU" Else sk = sk & Chr(10) & Space(n) & "La Loma Savings FCU"

    End If

'~~> #1a: paste data after finalizing format

    'paste if data present in temp sheet
    c = Application.CountA(Sheets("Temp").Range("A1:B200"))
    If c > 0 Then

        'move any credits in line with debits
        For Each c In Sheets("Temp").Range("D1:D200")
            If c <> "" Then
                c.Offset(0, -1) = c
                Range(c.Address) = ""
            End If
        Next

        'change all values to absolute values & indicate deposit
        For Each c In Sheets("Temp").Range("C1:C200")
            If c <> "" Then
                If c > 0 Then Range(c.Address).Offset(0, 1) = "+"
                Range(c.Address) = Abs(c)
            End If
        Next

        'add member to income
        For Each c In Sheets("Temp").Range("B1:B200")
            If c <> "" Then
                Set b = Range(c.Address).Find("Loma Linda Univ")
                If Not b Is Nothing Then
                    Range(b.Address).Offset(0, 2) = "M"
                End If
                Set a = Range(c.Address).Find("Guzman,Jose,E")
                If Not a Is Nothing Then
                    Range(a.Address).Offset(0, 2) = "J"
                End If
                Set a = Range(c.Address).Find("J E Guzman")
                If Not a Is Nothing Then
                    Range(a.Address).Offset(0, 2) = "J"
                End If
                Set a = Range(c.Address).Find("Desert Commun")
                If Not a Is Nothing Then
                    Range(a.Address).Offset(0, 2) = "J"
                End If
                Set a = Range(c.Address).Find("Claremont Grd Un")
                If Not a Is Nothing Then
                    Range(a.Address).Offset(0, 2) = "J"
                End If
            End If
        Next

        'remove transactions on the watch list
        GoSub Watch

        'sort by date
        GoSub sDate

        'acquire last cell
        For Each c In Sheets("Temp").Range("A:A")
            If c = "" Then
                a = c.Offset(-1).Row
                Exit For
            End If
        Next

        'remove leading/trailing spaces
        GoSub Clean

        'setup for paste - insert column, move income data, add blanks to prevent overflow text
        Sheets("Temp").Range("D1").EntireColumn.Insert
        For Each c In Sheets("Temp").Range("E1:E" & a)
            If Len(c) = 1 Then
                Sheets("Temp").Range(c.Offset(0, -1).Address) = Sheets("Temp").Range(c.Offset(0, -2).Address)
                Sheets("Temp").Range(c.Offset(0, -2).Address) = Space(1)
            End If
        Next

        'remove indicator that the transaction was a deposit
        For Each c In Sheets("Temp").Range("E1:E200")
            If c = "+" Then c.ClearContents
        Next

        'paste without selecting
        Sheets(i).Range("O4:O" & 4 + a - 1).Value = Sheets("Temp").Range("A1:A" & a).Value
        Sheets(i).Range("Q4:T" & 4 + a - 1).Formula = Sheets("Temp").Range("B1:E" & a).Formula

        'check for unmatched transactions without codes
        GoSub Check

        'remove wrap-text if any
        Workbooks(aa).Sheets(i).Range("O4:T" & a).WrapText = False

        'refresh transaction code formula
        Sheets(i).Range("P4:P203").Formula = Sheets("Codes").Range("Form2").Formula

        'clear temp area
        Sheets("Temp").Range("A:G").ClearContents

    End If

    'error check
    If Remove + a <> Total Then
        MsgBox Space(5) & "Re-check the expense/income entries since the totals do not match." & Chr(10) _
            & Space(35) & "There are " & Total - (Remove + a) & " missing transactions.", 0, "Warning"
    End If

'~~> #2: acquire data for credit charges

    'reset counters
    Total = 0
    Remove = 0

'1st credit statement

    'set variable to nothing for next statement
    Set c = Nothing

    'check if Citi statement exists
    On Error Resume Next
    Set c = Workbooks.Open(FilePath & "\" & ct)
    On Error GoTo 0

    If Not c Is Nothing Then

        'open credit card statement
        Workbooks.Open (FilePath & "\" & ct)        'citi statement

        'acquire range
        With Workbooks(ct)
            For Each c In .Sheets(1).Range("B:B")
                If c = "Date" Then
                    Start = c.Offset(1).Address
                End If
                If c.Row > 1 And c = "" Then
                    Last = c.Offset(-1, 4).Address
                    Exit For
                End If
            Next
        End With

        'acquire number of rows, acquire data and close workbook
        a = Range(Start & ":" & Last).Rows.Count
        Workbooks(aa).Sheets("Temp").Range("A1:E" & a).Value = Workbooks(ct).Sheets(1).Range(Start & ":" & Last).Value
        Workbooks(ct).Close                         'citi statement

        'add to total for error check
        Total = Total + a

        'clean up data
        With Workbooks(aa).Sheets("Temp")
            'remove wrap text
            .Range("A1:E" & a).WrapText = False
            'replace name w/initial
            For Each c In .Range("E:E")
                If c <> "" Then
                    .Range(c.Address) = Left(c, 1)
                End If
            Next
            'place credit card initial
            For Each c In .Range("E:E")
                If c <> "" Then
                    .Range(c.Offset(0, 1).Address) = "C"
                End If
            Next
        End With

        'save the statement processed in order to alert user
        If sk = "" Then sk = Space(n) & "CitiBank" Else sk = sk & Chr(10) & Space(n) & "CitiBank"

    End If

    'set variable to nothing for next statement
    Set c = Nothing

'2nd credit statement

    'check if JetBlue-M statement exists
    On Error Resume Next
    Set c = Workbooks.Open(FilePath & "\" & jm)
    On Error GoTo 0

    If Not c Is Nothing Then

        'open credit card statement
        Workbooks.Open (FilePath & "\" & jm)        'jet - m statement

        'acquire range
        With Workbooks(jm)
            For Each c In .Sheets(1).Range("A:A")
                If c.Address = "$A$1" And c <> "" Then
                    Start = c.Address
                End If
                If c.Row > 1 And c = "" Then
                    Last = c.Offset(-1, 3).Address
                    Exit For
                End If
            Next
        End With

        'acquire number of rows and acquire data
        With Workbooks(aa).Sheets("Temp")
            a = Application.CountA(.Range("A:A"))
        End With

        b = Range(Start & ":" & Last).Rows.Count
        Workbooks(aa).Sheets("Temp").Range("A" & a + 1 & ":D" & a + b).Value = Workbooks(jm).Sheets(1).Range(Start & ":" & Last).Value
        Workbooks(jm).Close                         'jet - m statement

        'add to total for error check
        Total = Total + b

        'clean up data
        With Workbooks(aa).Sheets("Temp")
            'add charge member
            For Each c In .Range("E" & a + 1 & ":E" & a + b)
                .Range(c.Address) = "M"
            Next
            'place credit card initial
            For Each c In .Range("E" & a + 1 & ":E" & a + b)
                If c <> "" Then
                    .Range(c.Offset(0, 1).Address) = "J"
                End If
            Next
        End With

        'save the statement processed in order to alert user
        If sk = "" Then sk = Space(n) & "JetBlue - m" Else sk = sk & Chr(10) & Space(n) & "JetBlue - m"

    End If

    'set variable to nothing for next statement
    Set c = Nothing

'3rd credit statement

    'check if JetBlue-J statement exists
    On Error Resume Next
    Set c = Workbooks.Open(FilePath & "\" & jj)
    On Error GoTo 0

    If Not c Is Nothing Then

        'open credit card statement
        Workbooks.Open (FilePath & "\" & jj)        'jet - j statement

        'acquire range
        With Workbooks(jj)
            For Each c In .Sheets(1).Range("A:A")
                If c.Address = "$A$1" And c <> "" Then
                    Start = c.Address
                End If
                If c.Row > 1 And c = "" Then
                    Last = c.Offset(-1, 3).Address
                    Exit For
                End If
            Next
        End With

        'acquire number of rows and acquire data
        With Workbooks(aa).Sheets("Temp")
            a = Application.CountA(.Range("A:A"))
        End With

        b = Range(Start & ":" & Last).Rows.Count
        Workbooks(aa).Sheets("Temp").Range("A" & a + 1 & ":D" & a + b).Value = Workbooks(jj).Sheets(1).Range(Start & ":" & Last).Value
        Workbooks(jj).Close                         'jet - j statement

        'add to total for error check
        Total = Total + b

        'clean up data
        With Workbooks(aa).Sheets("Temp")
            'add charge member
            For Each c In .Range("E" & a + 1 & ":E" & a + b)
                .Range(c.Address) = "J"
            Next
            'place credit card initial
            For Each c In .Range("E" & a + 1 & ":E" & a + b)
                If c <> "" Then
                    .Range(c.Offset(0, 1).Address) = "J"
                End If
            Next
        End With

        'save the statement processed in order to alert user
        If sk = "" Then sk = Space(n) & "JetBlue - j" Else sk = sk & Chr(10) & Space(n) & "JetBlue - j"

    End If

    'set variable to nothing for next statement
    Set c = Nothing

'4th credit statement

    'check if Apple-M statement exists
    On Error Resume Next
    Set c = Workbooks.Open(FilePath & "\" & pm)
    On Error GoTo 0

    If Not c Is Nothing Then

        'open credit card statement
        Workbooks.Open (FilePath & "\" & pm)        'app - m statement

        'acquire range
        With Workbooks(pm)
            For Each c In .Sheets(1).Range("A:A")
                If c.Address = "$A$1" And c <> "" Then
                    Start = c.Address
                End If
                If c.Row > 1 And c = "" Then
                    Last = c.Offset(-1, 3).Address
                    Exit For
                End If
            Next
        End With

        'acquire number of rows and acquire data
        With Workbooks(aa).Sheets("Temp")
            a = Application.CountA(.Range("A:A"))
        End With

        b = Range(Start & ":" & Last).Rows.Count
        Workbooks(aa).Sheets("Temp").Range("A" & a + 1 & ":D" & a + b).Value = Workbooks(pm).Sheets(1).Range(Start & ":" & Last).Value
        Workbooks(pm).Close                         'app - m statement

        'add to total for error check
        Total = Total + b

        'clean up data
        With Workbooks(aa).Sheets("Temp")
            'add charge member
            For Each c In .Range("E" & a + 1 & ":E" & a + b)
                .Range(c.Address) = "M"
            Next
            'place credit card initial
            For Each c In .Range("E" & a + 1 & ":E" & a + b)
                If c <> "" Then
                    .Range(c.Offset(0, 1).Address) = "A"
                End If
            Next
        End With

        'save the statement processed in order to alert user
        If sk = "" Then sk = Space(n) & "Apple - m" Else sk = sk & Chr(10) & Space(n) & "Apple - m"

    End If

    'set variable to nothing for next statement
    Set c = Nothing

'5th credit statement

    'check if Apple-S statement exists
    On Error Resume Next
    Set c = Workbooks.Open(FilePath & "\" & ps)
    On Error GoTo 0

    If Not c Is Nothing Then

        'open credit card statement
        Workbooks.Open (FilePath & "\" & ps)        'app - s statement

        'acquire range
        With Workbooks(ps)
            For Each c In .Sheets(1).Range("A:A")
                If c.Address = "$A$1" And c <> "" Then
                    Start = c.Address
                End If
                If c.Row > 1 And c = "" Then
                    Last = c.Offset(-1, 3).Address
                    Exit For
                End If
            Next
        End With

        'acquire number of rows and acquire data
        With Workbooks(aa).Sheets("Temp")
            a = Application.CountA(.Range("A:A"))
        End With

        b = Range(Start & ":" & Last).Rows.Count
        Workbooks(aa).Sheets("Temp").Range("A" & a + 1 & ":D" & a + b).Value = Workbooks(ps).Sheets(1).Range(Start & ":" & Last).Value
        Workbooks(ps).Close                         'app - s statement

        'add to total for error check
        Total = Total + b

        'clean up data
        With Workbooks(aa).Sheets("Temp")
            'add charge member
            For Each c In .Range("E" & a + 1 & ":E" & a + b)
                .Range(c.Address) = "S"
            Next
            'place credit card initial
            For Each c In .Range("E" & a + 1 & ":E" & a + b)
                If c <> "" Then
                    .Range(c.Offset(0, 1).Address) = "A"
                End If
            Next
        End With

        'save the statement processed in order to alert user
        If sk = "" Then sk = Space(n) & "Apple - s" Else sk = sk & Chr(10) & Space(n) & "Apple - s"

    End If

    'set variable to nothing for next statement
    Set c = Nothing

'6th credit statement

    'set variable to nothing for next statement
    Set c = Nothing

    'check if Chase statement exists
    On Error Resume Next
    Set c = Workbooks.Open(FilePath & "\" & cs)
    On Error GoTo 0

    If Not c Is Nothing Then

        'open credit card statement
        Workbooks.Open (FilePath & "\" & cs)        'chase statement

        'acquire range
        With Workbooks(cs)
            For Each c In .Sheets(1).Range("A:A")
                If c = "Transaction Date" Then
                    Start = c.Offset(1).Address
                End If
                If c.Row > 1 And c = "" Then
                    Last = c.Offset(-1, 5).Address
                    Exit For
                End If
            Next
        End With

        'acquire number of rows and acquire data
        With Workbooks(aa).Sheets("Temp")
            a = Application.CountA(Range("A:A"))
        End With

        'acquire number of rows, acquire data and close workbook
        b = Range(Start & ":" & Last).Rows.Count
        Workbooks(aa).Sheets("Temp").Range("A" & a + 1 & ":A" & a + b).Value = Workbooks(cs).Sheets(1).Range(Start & ":" & Range(Last).Offset(0, -5).Address).Value
        Workbooks(aa).Sheets("Temp").Range("B" & a + 1 & ":B" & a + b).Value = Workbooks(cs).Sheets(1).Range(Range(Start).Offset(0, 2).Address & ":" & Range(Last).Offset(0, -2).Address).Value
        Workbooks(aa).Sheets("Temp").Range("C" & a + 1 & ":C" & a + b).Value = Workbooks(cs).Sheets(1).Range(Range(Start).Offset(0, 5).Address & ":" & Range(Last).Offset(0, 0).Address).Value
        Workbooks(cs).Close                         'chase statement

        'add to total for error check
        Total = Total + b

        'clean up data
        With Workbooks(aa).Sheets("Temp")
            'remove wrap text
            .Range("A1:E" & a).WrapText = False
            'add more info to transactions
            For Each c In .Range("B" & a + 1 & ":B" & a + b)
                If c = "1236 REDLANDS" Then
                    .Range(c.Address) = c & " : Sephora"
                End If
            Next
            'move transactions to the right if needed
            For Each c In .Range("C" & a + 1 & ":C" & a + b)
                If Not IsNumeric(c) And IsNumeric(c.Offset(0, 1)) Then
                    .Range(c.Address) = c.Offset(0, 1)
                    .Range(c.Offset(0, 1).Address) = ""
                End If
            Next
            'add charge member
            For Each c In .Range("E" & a + 1 & ":E" & a + b)
                .Range(c.Address) = "M"
            Next
            'place credit card initial
            For Each c In .Range("E" & a + 1 & ":E" & a + b)
                If c <> "" Then
                    .Range(c.Offset(0, 1).Address) = "S"
                End If
            Next
            'place move returns/payments to appropriate column
            For Each c In .Range("C" & a + 1 & ":C" & a + b)
                If c <> "" And c > 0 Then
                    .Range(c.Offset(0, 1).Address) = c
                    .Range(c.Offset(0, 0).Address) = ""
                End If
            Next
            For Each c In .Range("D" & a + 1 & ":D" & a + b)
                If c <> "" And c < 0 Then
                    .Range(c.Offset(0, -1).Address) = c
                    .Range(c.Offset(0, 0).Address) = ""
                End If
            Next
        End With

        'save the statement processed in order to alert user
        If sk = "" Then sk = Space(n) & "Chase Sapphire CC" Else sk = sk & Chr(10) & Space(n) & "Chase Sapphire CC"

    End If

    'set variable to nothing for next statement
    Set c = Nothing

'** Add recurring charges not listed in monthly statements **

    'acquire last row of temp sheet to add recurring charges
    With Workbooks(aa).Sheets("Temp")
        a = Application.CountA(.Range("A:A"))
    End With

    b = 0
    'add recurring charges if present and not expired
    For Each c In Sheets("Codes").Range("Q36:Q45")

        If Len(c) > 1 Then

            m = i                                                               'month of ws processing
            d = DatePart("d", Sheets("Codes").Range("T36"))                     'day date of charge
            y = Mid(ThisWorkbook.Name, InStr(ThisWorkbook.Name, ".") - 4, 4)    'year of wb title

            m = m * 1
            y = y * 1

            If y <= Year(c.Offset(0, 3)) Then
                If m >= Month(c.Offset(0, 3)) Then

                    b = b + 1
                    Workbooks(aa).Sheets("Temp").Range("A" & a + b).Value = DateSerial(y, m, d)
                    Workbooks(aa).Sheets("Temp").Range("B" & a + b).Value = Workbooks(aa).Sheets("Codes").Range(c.Address).Offset(0, 0).Value
                    Workbooks(aa).Sheets("Temp").Range("C" & a + b).Value = Workbooks(aa).Sheets("Codes").Range(c.Address).Offset(0, 2).Value
                    Workbooks(aa).Sheets("Temp").Range("E" & a + b).Value = "M"
                    Workbooks(aa).Sheets("Temp").Range("F" & a + b).Value = "-"

                End If

            End If

        End If

    Next

    'add to total for error check
    Total = Total + b

'~~> #2a: paste data after finalizing format

    'paste if data present in temp sheet
    c = Application.CountA(Sheets("Temp").Range("A1:B200"))
    If c > 0 Then

        'remove transactions on the watch list
        GoSub Watch

        'sort by date
        GoSub sDate

        'acquire number of rows to copy
        With Workbooks(aa).Sheets("Temp")
            a = Application.CountA(.Range("A:A"))
        End With

        'remove leading/trailing spaces
        GoSub Clean

        'set transaction charges to abs values
        For Each c In Sheets("Temp").Range("C1:C200")
            If c <> "" Then
                Range(c.Address) = Abs(c)
            End If
        Next

        'set credits/returns to negative values
'        For Each c In Sheets("Temp").Range("D1:D200")      'commented out since the column indicates that it is a credit/return
'            If c <> "" Then                                'the negative value is not needed
'                Range(c.Address) = Abs(c) * -1             'the formulas have been updated to handle positive numbers in both columns:
'            End If                                         'charge column and the credit/returns column
'        Next

        'add blanks to prevent overflow text
        For Each c In Sheets("Temp").Range("C1:C" & a)
            If IsEmpty(c) Then Range(c.Address) = Space(1)
        Next

        'paste without selecting
        Sheets(i).Range("B4:B" & 4 + a - 1).Value = Sheets("Temp").Range("F1:F" & a).Value
        Sheets(i).Range("C4:C" & 4 + a - 1).Value = Sheets("Temp").Range("A1:A" & a).Value
        Sheets(i).Range("E4:H" & 4 + a - 1).Formula = Sheets("Temp").Range("B1:E" & a).Formula

        'remove wrap-text if any
        Workbooks(aa).Sheets(i).Range("B4:H" & a).WrapText = False

        'check for transactions not in list
        GoSub Check

        'copy/paste unmatched transactions
        GoSub Match

    End If

    'remove temp sheet
    Workbooks(aa).Sheets("Temp").Delete

    'refresh transaction code formula for charges
    Sheets(i).Range("D4:D203").Formula = Sheets("Codes").Range("Form1").Formula

    'set scroll area
    Call SetScrollArea

    'error check
    If Remove + a <> Total Then
        MsgBox Space(5) & "Re-check the credit entries since the totals do not match." & Chr(10) _
            & Space(25) & "There are " & Total - (Remove + a) & " missing transactions.", 0, "Warning"
    End If

    'report the statements processed
    c = Len(sk) - Len(Replace(sk, Chr(10), "")) + 1
    If c > 1 Then c = "statements were" Else c = "statement was"
    If sk <> "" Then
        If c > 1 Then
            MsgBox Space(5) & "The following " & c & " processed:" & Chr(10) & Chr(10) & sk, 0, "Processed Statements"
        End If
    End If

    'alert user of unmatched transactions
    If Found Then MsgBox Space(12) & "There are unmatched transactions without codes." & Chr(10) & _
                    Space(8) & "Assign codes to these transactions in the Codes sheet.", , "Warning"

ext:

    'change indicator format
    Sheets(i).Range("M3").Font.Size = 8
    Sheets(i).Range("M3").Font.Name = "Century Gothic"

    'restore settings
    If err.number = 0 Then Sheets(i).Unprotect: _
        Application.Goto Sheets(i).Range("A1"), _
        True: Application.Goto Sheets(i).Range("C4"): _
        Call SingleLink 'rebuild hlinks
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    On Error GoTo 0

    'transfer transactions based on transactions codes
    Call PostProcess

    'highlights duplicates
    Call View

Exit Sub

Clean:

    'remove leading and trailing spaces from transactions
    For Each c In Range("A1:G" & a)
        Range(c.Address) = Application.Trim(c.Value)
    Next
    Return

Watch:

    'delete transactions on Temp watch list
    For Each c In Sheets("Temp").Range("B1:B300")
        If c <> "" Then
            For Each b In Sheets("Codes").Range("N4:N103")
                If Len(Trim(b)) <> 0 Then
                    a = 10  'max length to compare
                    If UCase(Left(b, a)) = UCase(Left(c, a)) Then
                        Sheets("Temp").Range("A" & c.Row & ":G" & c.Row).ClearContents
                        Remove = Remove + 1
                    End If
                End If
            Next
        End If
    Next
    Return

Check:
'copy unmatched transactions to Codes sheet for inspection

    'check for transaction matches to delete
    For Each c In Sheets("Temp").Range("B1:B100")
        If c <> "" Then
            For Each b In Range("Transact")
                If b <> "" Then
                    If UCase(Left(b, 4)) = UCase(Left(c, 4)) Then
                        Sheets("Temp").Range(c.Address).Offset(0, 5) = "Found"
                    End If
                End If
            Next
        End If
    Next

    'delete matches
    For Each c In Sheets("Temp").Range("G1:G100")
        If c = "Found" Then
            Range("A" & c.Row & ":" & c.Address).ClearContents
        End If
    Next

    'sort by date to remove blanks
    GoSub sDate

    'copy to preserve
    For Each c In Sheets("Temp").Range("B1:B100")
        If c = "" And c.Row = 1 Then c = c.Row: Exit For
        If c = "" And c.Row > 1 Then c = c.Offset(-1).Row: Exit For
    Next
    For Each b In Sheets("Temp").Range("H1:H100")
        If b = "" Then b = b.Row: Exit For
    Next

    'move to preserve unmatched transactions
    If c <> "" Then
        Sheets("Temp").Range("H" & b & ":H" & b + c - 1).Value = Sheets("Temp").Range("B1:B" & c).Value
    End If

    'remove wrap-text if any
    With Workbooks(aa).Sheets("Temp")
        .Range("A:H").WrapText = False
    End With
    Return

Match:

    'sort unmatched transactions
    Sheets("Temp").Range("H1:H100").Sort Key1:=Range("H1"), Order1:=xlAscending

    'acquire last row unmatched transactions
    With Workbooks(aa).Sheets("Temp")
        c = Application.CountA(.Range("H1:H100"))
    End With

    'copy/paste unmatched transactions
    If c > 0 Then Workbooks(aa).Sheets("Codes").Range("I4:I" & c + 3).Value = Workbooks(aa).Sheets("Temp").Range("H1:H" & c).Value
    If c > 0 Then Found = True

    'remove wrap-text if any
    With Workbooks(aa).Sheets("Codes")
        .Range("I4:I103").WrapText = False
        Application.Goto .Range("Transfer").Offset(1, 1)
    End With
    Return

sDate:

    'sort by date
    Sheets("Temp").Range("A:G").Sort Key1:=Range("A1"), Order1:=xlAscending
    Return

End Sub



