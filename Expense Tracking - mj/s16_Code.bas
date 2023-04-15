Attribute VB_Name = "s16_Code"
Option Explicit
'Contains 4 macros: Defer, Scene, Scroll, Transfer

Sub Defer()
'Codes, delete/transfer single unmatched transactions to watch list
'delete/transfer coded transactions to unmatched list or watch list

    'exit if error or no entry
    If IsError(ActiveCell) Then Exit Sub
    If Trim(ActiveCell) = "" Then Exit Sub

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'delete or transfer coded transaction
    If Split(ActiveCell.Address, "$")(1) = "C" Then

        'confirm deletion/transfer
        i = (88 - (Len(Left(Trim(ActiveCell.Offset(0, 1)), 8)) + 9)) / 2

        i = MsgBox(Space(6) & "Do you want to transfer or delete the following entry?" & Chr(10) _
                    & Space(i) & UCase(Format(ActiveCell.Offset(0, -1), "mmm d")) _
                    & " - " & UCase(Left(Trim(ActiveCell.Offset(0, 0)), 8)) & Chr(10) & Chr(10) _
                    & Space(26) & "[ transfer = yes  |  delete = no ]", 3, "Transfer / Delete")

        If i = 2 Then GoTo ext
        If i = 6 Then

            i = MsgBox(Space(6) & "Transfer to the unmatched transaction list or watch list?" & Chr(10) _
                    & Space(16) & "[ unmatched list = yes  |  watch list = no ]", 3, "Transfer")

            If i = 2 Then GoTo ext
            If i = 6 Then

                'check if duplicate
                For Each i In Range("I4:I103")
                    If i = ActiveCell Then
                        MsgBox Space(5) & "The transaction is a duplicate and will not be transferred.", 0, "Duplicate"
                        GoTo ext
                    End If
                Next

                'copy/transfer transaction
                For Each i In Range("I4:I103")
                    If Len(Trim(i)) = 0 Then
                        Range(i.Address) = ActiveCell
                        Range("I4:I103").WrapText = False
                        Exit For
                    End If
                Next

                'delete transferred transaction
                ActiveCell.ClearContents
                Range(ActiveCell.Address).Offset(0, -1).ClearContents

                'sort, restore "-", rebuild hlinks
                GoSub cSort
                GoSub iSort

            End If

            If i = 7 Then

                'check if duplicate
                For Each i In Range("N4:N103")
                    If i = ActiveCell Then
                        MsgBox Space(5) & "The transaction is a duplicate and will not be transferred.", 0, "Duplicate"
                        GoTo ext
                    End If
                Next

                'copy/transfer transaction
                For Each i In Range("N4:N103")
                    If Len(Trim(i)) = 0 Then
                        Range(i.Address) = ActiveCell
                        Range("N4:N103").WrapText = False
                        Exit For
                    End If
                Next

                'delete transferred transaction
                ActiveCell.ClearContents
                Range(ActiveCell.Address).Offset(0, -1).ClearContents

                'sort, restore "-", rebuild hlinks
                GoSub cSort
                GoSub nSort

            End If

        End If

        If i = 7 Then

            'clear coded transaction
            ActiveCell.ClearContents
            Range(ActiveCell.Address).Offset(0, -1).ClearContents

            'sort, restore "-", rebuild hlinks
            GoSub cSort

        End If

        GoTo ext
    End If

    'delete or transfer/sort unmatched transaction
    If Split(ActiveCell.Address, "$")(1) = "I" Then

        'confirm deletion/transfer
        i = (92 - (Len(Left(Trim(ActiveCell), 14)))) / 2

        i = MsgBox(Space(6) & "Do you want to transfer to watch list or delete this entry?" & Chr(10) _
                & Space(i) & UCase(Left(Trim(ActiveCell), 14)) & Chr(10) & Chr(10) _
                & Space(29) & "[ transfer = yes  |  delete = no ]", 3, "Transfer / Delete")

        If i = 2 Then GoTo ext
        If i = 7 Then ActiveCell.ClearContents

        If i = 6 Then
            For Each i In Range("N4:N103")
                If Len(Trim(i)) > 0 Then
                    If i = ActiveCell Then
                        MsgBox Space(5) & "This transaction is already in the watch list.", 0, "Duplicate"
                        ActiveCell.ClearContents
                        i = True
                        Exit For
                    End If
                End If
            Next
            If Not i Then
                For Each i In Range("N4:N103")
                    If Len(Trim(i)) = 0 Then
                        Range(i.Address) = ActiveCell
                        Range("N4:N103").WrapText = False
                        ActiveCell.ClearContents
                        Exit For
                    End If
                Next
                GoSub nSort     'sort watch list
            End If
        End If

        'sort, restore "-", rebuild hlinks
        GoSub iSort

        GoTo ext
    End If

    'delete and sort watch list transaction
    If Split(ActiveCell.Address, "$")(1) = "N" Then

        'confirm deletion/transfer
        i = (88 - (Len(Left(Trim(ActiveCell), 14)))) / 2
    
        i = MsgBox(Space(4) & "Transfer to unmatched transactions or delete this entry?" & Chr(10) _
                & Space(i) & UCase(Left(Trim(ActiveCell), 14)) & Chr(10) & Chr(10) _
                & Space(26) & "[ transfer = yes  |  delete = no ]", 3, "Transfer / Delete")

        If i = 7 Then ActiveCell.ClearContents: GoSub nSort
        If i = 6 Then
            For Each i In Range("I4:I103")
                If Len(Trim(i)) > 0 Then
                    If i = ActiveCell Then
                        MsgBox Space(5) & "The item is already in the unmatched transaction list.", 0, "Duplicate"
                        i = True
                        GoTo ext
                    End If
                End If
            Next
            If Not i Then
                For Each i In Range("I4:I103")
                    If Len(Trim(i)) = 0 Then
                        Range(i.Address) = ActiveCell
                        Range("I4:I103").WrapText = False
                        ActiveCell.ClearContents
                        Exit For
                    End If
                Next
                GoSub iSort     'sort unmatched transaction list
                GoSub nSort     'sort watch list
            End If
        End If

    End If

ext:

    'restore settings
    Application.Goto Range("Transfer").Offset(1, 1)
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

Exit Sub

cSort:

    'reset coded transaction list
    For Each i In Range("B4:C1003")
        If Len(Trim(i)) = 0 Then
            Range(i.Address) = ""
        End If
    Next

    'sort coded transaction list
    Sheets("Codes").Range("B4:C1003").Sort Key1:=Range("C4"), Order1:=xlAscending

    'rebuild coded transaction list
    i = "Sheets(""Codes"").Range(""B4:C1003"")"     'global var to pass to macro
    Call SentLink

    'reset spaces for empty cells
    For Each i In Range("B4:C1003")
        If Len(Trim(i)) = 0 Then
            Range(i.Address) = Space(50)
        End If
    Next

    'set scroll area
    Call SetScrollArea

    Return

iSort:

    'reset unmatched transactions
    For Each i In Range("I4:I103")
        If Len(Trim(i)) = 0 Then
            Range(i.Address) = ""
        End If
    Next

    'sort unmatched transactions
    Sheets("Codes").Range("I4:K103").Sort Key1:=Range("I4"), Order1:=xlAscending

    'restore "-" to blank transactions
    For Each i In Range("I4:I103")
        If Len(Trim(i)) = 0 Then
            i.Offset(0, 2) = "-"
        End If
    Next

    'rebuild unmatched list hlinks
    i = "Sheets(""Codes"").Range(""I4:I103"")"     'global var to pass to macro
    Call SentLink

    Return

nSort:

    'reset watch list
    For Each i In Range("N4:N103")
        If Len(Trim(i)) = 0 Then
            Range(i.Address) = ""
        End If
    Next

    'sort watch list
    Sheets("Codes").Range("N4:N103").Sort Key1:=Range("N4"), Order1:=xlAscending

    'rebuild watch list hlinks
    i = "Sheets(""Codes"").Range(""N4:N103"")"     'global var to pass to macro
    Call SentLink

    Return

End Sub

Sub Scene()
'Codes, Data, Items; toggles sheets hidden property

    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    If ActiveSheet.Name = "Codes" Then

        Sheets("Codes").Range("Transfer").Offset(1, 1).Select

        'make Data visible
        Sheets("Items").Visible = xlSheetVisible
        Sheets("Data").Visible = xlSheetVisible
        Sheets("Data").Activate

        If ActiveCell.Address <> "$J$3" Then
            Application.Goto Sheets("Data").Range("J3")
        End If

    ElseIf ActiveSheet.Name = "Data" Or ActiveSheet.Name = "Items" Then

        Sheets("Codes").Activate
        Sheets("Data").Visible = xlVeryHidden
        Sheets("Items").Visible = xlVeryHidden
        If ActiveCell.Address = Range("Transfer").Offset(1, 1).Address Then GoTo ext

        Sheets("Codes").Unprotect
        Application.Goto Range("A1"), True
        Range("Transfer").Offset(1, 1).Select
        Sheets("Codes").Protect

    End If

ext:

    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub Scroll()
'Codes, toggle auto/manual, goto last coded transaction or top of sheet

    If ActiveCell.Address(0, 0) = "E23" Then

        'remove scrollarea
        Sheets("Codes").ScrollArea = ""

        'goto last coded transaction
        For Each i In Range("C4:C1004")
            If Trim(i) = "" Then
                i = i.Offset(-1).Row
                Sheets("Codes").Unprotect
                Application.Goto Range("B" & i)
                Exit For
            End If
        Next

        'reinstate scrollarea
        Call SetScrollArea

    ElseIf ActiveCell.Address(0, 0) = "E31" Then

        If ActiveCell = "Automatic" Then Range(ActiveCell.Address) = "Manual" Else Range(ActiveCell.Address) = "Automatic"

    Else

        'goto beginning of sheet
        Application.Goto Range("A1"), True
        Range("Transfer").Offset(1, 1).Select

    End If

End Sub

Sub Transfers()
'Codes, transfers unmatched list item to transaction list

    Dim a, b, c, d

    'set cursor
    Application.Goto Range("Transfer").Offset(1, 1)

    'check if data else exit
    For Each c In Range("I4:I103")
        If Len(Trim(c)) = 0 Then b = b + 1
        If c.Offset(0, 2) = "-" Then a = a + 1
    Next
    If a + b = 200 Then Exit Sub

    'confirm transfer
    i = MsgBox(Space(6) & "Do you want to transfer transactions that have a new code" & Chr(10) _
                & Chr(9) & Chr(9) & Chr(9) & Space(4) & "or" & Chr(10) _
                & Chr(9) & Space(8) & "delete all unmatched transactions?" & Chr(10) & Chr(10) _
                & Chr(9) & Space(11) & "[ transfer = yes  |  delete = no ]", 3, "Transfer")

    If i = 2 Then Exit Sub
    If i = 6 Then
        'check for new codes else exit
        For Each c In Range("K4:K103")
            If c.Offset(0, 2) = "-" Then a = a + 1
        Next
        If a = 100 Then Exit Sub
    End If

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Sheets("Codes").Unprotect

    'confirm deletion
    If i = 7 Then
        i = MsgBox(Space(7) & "Delete all unmatched transactions listed?", 4, "Confirm")
        If i = 6 Then a = "K4:K103": GoTo del Else GoTo ext
    End If

    a = ""
    'check for new codes else exit
    For Each c In Range("K4:K103")
        If c.Offset(0, 2) = "-" Then a = a + 1
    Next
    If a = 100 Then GoTo ext

    'check for regular processing or duplicate entry processing
    i = MsgBox(Space(12) & "Do you want to attempt to process all entries " & Chr(10) _
                & Chr(9) & Chr(9) & Chr(9) & Space(0) & "or" & Chr(10) _
                & Space(13) & "check duplicate entries for possible transfer?" & Chr(10) & Chr(10) _
                & Chr(9) & Space(6) & "[ process = yes  |  duplicate = no ]", 3, "Transfer Check")

    If i = 2 Then GoTo ext
    If i = 7 Then GoTo Dcheck

'****Regular Processing****

    'reset vars and scroll area
    a = "": d = "": i = True: Sheets("Codes").ScrollArea = ""

    'acquire range of transactions to transfer
    For Each c In Range("K4:K103")
        If Trim(c) <> "" And c <> "-" And Trim(c.Offset(0, -2)) <> "" Then
            For Each b In Range("C4:C2003")
                If Trim(b) <> "" And b = c.Offset(0, -2) Then
                    i = False
                    If d = "" Then
                        d = Chr(9) & Chr(9) & Space(10) & UCase(c.Offset(0, -2))
                    Else
                        d = d & Chr(10) & Chr(9) & Chr(9) & Space(10) & UCase(c.Offset(0, -2))
                    End If
                End If
            Next
            If i Then
                If a = "" Then a = c.Address Else a = a & "," & c.Address
            End If
            i = True
        End If
    Next

    'alert of duplicates
    i = Len(d) - Len(Replace(d, Chr(10), "")) + 1   'number of duplicates
    If a = "" And d <> "" Then
        If i = 1 Then MsgBox Space(5) & "The item is a duplicate and cannot be transferred.", 0, "Duplicate"
        If i > 1 Then MsgBox Space(5) & "There are no items to transfer since they are all duplicates.", 0, "Duplicates"
        GoTo ext
    End If

    'acquire last entry in proper code/transaction list
    For Each c In Range("B4:B2003")
        If Trim(c) = "" Then b = c.Address: Exit For
    Next

    'copy/paste transfers
    Range(a).Copy
    Range(b).PasteSpecial Paste:=xlPasteValues
    Range(a).Offset(0, -2).Copy
    Range(b).Offset(0, 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    'reset coded transaction list
    For Each i In Range("B4:C1003")
        If Len(Trim(i)) = 0 Then
            Range(i.Address) = ""
        End If
    Next

    'sort coded transaction list
    Sheets("Codes").Range("B4:C1003").Sort Key1:=Range("C4"), Order1:=xlAscending

    'rebuild coded transaction list
    i = "Sheets(""Codes"").Range(""B4:C1003"")"     'global var to pass to macro
    Call SentLink

    'reset spaces for empty cells
    For Each i In Range("B4:C1003")
        If Len(Trim(i)) = 0 Then
            Range(i.Address) = Space(50)
        End If
    Next

    'set scroll area
    Call SetScrollArea

del:

    'delete the transactions transferred
    Range(a) = "-"
    Range(a).Offset(0, -2).ClearContents

hlink:

    'reset unmatched transactions
    For Each c In Range("I4:I103")
        If Len(Trim(c)) = 0 Then
            Range(c.Address) = ""
        End If
    Next

    'sort unmatched transactions to remove spaces
    Sheets("Codes").Range("I4:K103").Sort Key1:=Range("I4"), Order1:=xlAscending

    'rebuild hlinks
    i = "Sheets(""Codes"").Range(""I4:I103"")"      'global var to pass to macro
    Call SentLink

    'set scroll area
    Call SetScrollArea

ext:

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    'alert user of duplicates
    If a <> "" And d <> "" Then
        MsgBox Space(5) & "The following are duplicate transactions that were not transferred:" & Chr(10) & Chr(10) _
                & d, 0, "Duplicates"
    End If

Exit Sub

'****Duplicate Processing****
Dcheck:

    'check for duplicates
    For Each a In Range("K4:K103")
        If Trim(a) <> "" And a <> "-" And Trim(a.Offset(0, -2)) <> "" Then
            For Each b In Range("C4:C2003")
                If Trim(b) <> "" And b = a.Offset(0, -2) Then 'duplicate found
                    c = a.Address
                    d = b.Offset(0, -1).Address
                    GoSub Dfound
                End If
            Next
        End If
    Next

    'alert user of completion
    MsgBox "All duplicates have been checked.", 0, "Complete"

    GoTo hlink

Dfound:

    'alert user of identical duplicate
    If Range(c) = Range(d) Then
        MsgBox Space(6) & "The following entry is identical and need not be transferred:" & Chr(10) & Chr(10) & _
               Space(18) & "Entry:    " & Left(Range(c).Offset(0, -2), 30) & Chr(10) & _
               Space(18) & "Transaction Code:    " & Range(d), 0, "Duplicate"
        Return
    End If

    'alert user of duplicate w/diff code
    If Range(c) <> Range(d) Then
        i = MsgBox(Space(6) & "The following entry is a duplicate with a different code:" & Chr(10) & Chr(10) & _
               Space(16) & "Entry:  " & Left(Range(c).Offset(0, -2), 28) & Chr(10) & _
               Space(16) & "New Transaction Code:   " & Range(c) & Chr(10) & _
               Space(16) & "Old   Transaction Code:   " & Range(d) & Chr(10) & Chr(10) & _
               Space(12) & "Replace the transaction code with the new one?", 3, "Replace Code?")

        If i = 6 Then

            'copy transaction code
            Range(d) = Range(c)

            'delete the entry from unmatched list
            Range(c).Offset(0, -3).ClearContents
            Range(c).Offset(0, -2).ClearContents
            Range(c) = "-"
        
        End If

        Return

    End If

    Return

End Sub
