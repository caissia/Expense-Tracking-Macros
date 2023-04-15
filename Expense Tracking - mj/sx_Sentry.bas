Attribute VB_Name = "sx_Sentry"
Option Explicit

Sub Sentry()
's1 - s12 sheets
'sort expense list
'or delete single entry
'or transfer to charge list
'or transfer to expense list
'or transfer to another sheet

    'exit if transaction not found
    If Split(ActiveCell.Address, "$")(1) = "D" Then
        i = ActiveCell.Row
        i = Application.CountA(Range("B" & i & ":" & "H" & i))
        If i < 2 Then Exit Sub
    End If

    If Split(ActiveCell.Address, "$")(1) = "P" Then
        i = ActiveCell.Row
        i = Application.CountA(Range("O" & i & ":" & "T" & i))
        If i < 2 Then Exit Sub
    End If

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

'''''''''''''''''''''''' s1 - s12 ''''''''''''''''''''''''

    'if charge column transaction
    If Split(ActiveCell.Address, "$")(1) = "D" Then

        'delete or transfer
        GoSub msg

        'delete entry
        If i = 1 Then

            'acquire range to clear
            i = ActiveCell.Offset(0, -2).Address & ":" & ActiveCell.Offset(0, -1).Address & ","
            i = i & ActiveCell.Offset(0, 1).Address & ":" & ActiveCell.Offset(0, 4).Address

            'clear content, sort, rebuild hlinks
            GoTo clr

        End If

        'transfer entry
        If i = 7 Then

            'check if transfer is to expense list or unmatched list in Codes sheet
            i = MsgBox(Space(7) & "Transfer to the expense list or unmatched list?" & Chr(10) _
                    & Space(12) & "[ expense list = yes  |  unmatched = no ]", 3, "Transfer")

            If i = 2 Then GoTo ext
            If i = 6 Then

                'check if duplicate
                For Each i In Range("Q4:Q203")
                    If i = ActiveCell.Offset(0, 1) And i.Offset(0, 1) = ActiveCell.Offset(0, 2) Then
                        MsgBox Space(5) & "The transaction is a duplicate and will not be transferred.", 0, "Duplicate"
                        GoTo ext
                    End If
                Next

                'acquire last row of expense column
                i = Cells(Rows.Count, "O").End(xlUp).Row + 1

                'copy/transfer transaction
                Range("O" & i) = ActiveCell.Offset(0, -1)
                Range("Q" & i) = ActiveCell.Offset(0, 1)
                Range("R" & i) = ActiveCell.Offset(0, 2)
                Range("S" & i) = ActiveCell.Offset(0, 3)

                'sort expense column
                Range("O4:T" & i).Sort Key1:=Range("O4"), Order1:=xlAscending

                'remove any wrap-text of description
                Range("O4:T" & i).WrapText = False

                'acquire range to clear
                i = ActiveCell.Offset(0, -2).Address & ":" & ActiveCell.Offset(0, -1).Address & ","
                i = i & ActiveCell.Offset(0, 1).Address & ":" & ActiveCell.Offset(0, 4).Address

                'clear content, sort, rebuild hlinks
                GoTo clr

            End If

            'transfer to codes sheet
            If i = 7 Then GoTo cds

        End If

    'if expense column transaction
    ElseIf Split(ActiveCell.Address, "$")(1) = "P" Then

        'delete or transfer
        GoSub msg

        'delete entry
        If i = 1 Then

            'acquire range to clear
            i = ActiveCell.Offset(0, -1).Address & ","
            i = i & ActiveCell.Offset(0, 1).Address & ":" & ActiveCell.Offset(0, 4).Address

            'clear content, sort, rebuild hlinks
            GoTo clr

        End If

        'transfer entry
        If i = 7 Then

            i = MsgBox(Space(7) & "Transfer to the charge list or unmatched list?" & Chr(10) _
                    & Space(16) & "[ charge list = yes  |  unmatched = no ]", 3, "Transfer")

            If i = 2 Then GoTo ext
            If i = 6 Then

                'check if duplicate
                For Each i In Range("E4:E203")
                    If i = ActiveCell.Offset(0, 1) And i.Offset(0, 1) = ActiveCell.Offset(0, 2) Then
                        MsgBox Space(5) & "The transaction is a duplicate and will not be transferred.", 0, "Duplicate"
                        GoTo ext
                    End If
                Next

                'acquire last row of charge column
                i = Cells(Rows.Count, "C").End(xlUp).Row + 1

                'copy/transfer transaction
                Range("B" & i) = "B"
                Range("C" & i) = ActiveCell.Offset(0, -1)
                Range("E" & i) = ActiveCell.Offset(0, 1)
                Range("F" & i) = ActiveCell.Offset(0, 2)
                Range("G" & i) = ActiveCell.Offset(0, 3)
                Range("H" & i) = ActiveCell.Offset(0, 4)

                'sort charge column
                Range("B4:H" & i).Sort Key1:=Range("C4"), Order1:=xlAscending

                'remove any wrap-text of description
                Range("B4:H" & i).WrapText = False

                'acquire range to clear
                i = ActiveCell.Offset(0, -1).Address & ","
                i = i & ActiveCell.Offset(0, 1).Address & ":" & ActiveCell.Offset(0, 4).Address

                'clear content, sort, rebuild hlinks
                GoTo clr

            End If

            'transfer to codes sheet
            If i = 7 Then GoTo cds

        End If

    End If

ext:

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    'reapply toggle settings
    GoSub ind

Exit Sub

msg: 'ask user if entry is to be transferred or deleted

    'delete or transfer
    i = (86 - (Len(Left(Trim(ActiveCell.Offset(0, 1)), 8)) + 9)) / 2 'spacing

    i = MsgBox(Space(6) & "Do you want to transfer or delete the following entry?" & Chr(10) & Chr(10) _
                & Space(i) & UCase(Format(ActiveCell.Offset(0, -1), "mmm d")) _
                & " - " & UCase(Left(Trim(ActiveCell.Offset(0, 1)), 10)) & Chr(10) & Chr(10) _
                & Space(26) & "[ transfer = yes  |  delete = no ]", 3, "Transfer / Delete")

    'if canceled exit
    If i = 2 Then
        GoTo ext

    'check if transfer is to another monthly sheet
    ElseIf i = 6 Then

        i = MsgBox(Space(16) & "Is the transfer to another month?", 3, "Monthly Transfer")
        If i = 2 Then GoTo ext
        If i = 6 Then GoTo mtransfer

    'confirm deletion
    ElseIf i = 7 Then

        i = (38 - (Len(Left(Trim(ActiveCell.Offset(0, 1)), 8)) + 9)) / 2 'spacing

        i = MsgBox(Space(6) & "Delete the following entry?" & Chr(10) & Chr(10) _
                & Space(i) & UCase(Format(ActiveCell.Offset(0, -1), "mmm d")) _
                & " - " & UCase(Left(Trim(ActiveCell.Offset(0, 1)), 10)), 1, "Confirm")

    End If

Return

clr: 'clear content, sort, rebuild hlinks

    'clear contents of selected entry
    Range(i).ClearContents

    'determine key and range for sorting
    i = Split(i, "$")(1) & "4:" & Mid(i, InStrRev(i, ":") + 2, 1) & Cells(Rows.Count, Split(i, "$")(1)).End(xlUp).Row

    'sort to remove cleared empty row
    ActiveSheet.Range(i).Sort Key1:=ActiveCell.Offset(0, -1), Order1:=xlAscending

    'reapply self-reference hyperlinks to active sheet code columns
    i = True
    Call SingleLink
    i = False

    'set scroll area, active sheet
    Call SetScrollArea

    'set scroll area, target sheet
    If Len(Trim(t)) = 3 Then

        'set scroll area
        For Each i In Sheets(t).Range("C10:C2404")
            If Trim(i) = "" Then i = i.Row: Exit For
        Next
        If i < 34 Then Sheets(t).ScrollArea = "A1:X34" _
        Else: Sheets(t).ScrollArea = "A1:X" & i

    End If

    'exit sub
    GoTo ext

ind: 'reapply toggle settings

    'set toggle indicator, active sheet
    Call Indicator

    'set toggle indicator, target sheet
    If Len(Trim(t)) = 3 Then

        If Application.max(Sheets(t).Range("A4:A203")) = 0 Then
            Sheets(t).Range("M3").Font.Size = 6
            Sheets(t).Range("M3").Font.Name = "Wingdings"
        Else
            Sheets(t).Range("M3").Font.Size = 8
            Sheets(t).Range("M3").Font.Name = "Century Gothic"
        End If

    End If

Return

cds: 'transfer entry to codes sheet

    With Sheets("Codes")

        'check if duplicate
        For Each i In .Range("I4:I103")
            If i = ActiveCell.Offset(0, 1) Then
                MsgBox Space(5) & "The transaction is a duplicate and will not be transferred.", 0, "Duplicate"
                GoTo ext
            End If
        Next

        'copy/transfer transaction
        For Each i In .Range("I4:I103")
            If Len(Trim(i)) = 0 Then
                .Range(i.Address) = ActiveCell.Offset(0, 1)
                .Range("I4:I103").WrapText = False
                Exit For
            End If
        Next

        'sort reset list
        For Each i In .Range("I4:I103")
            If Len(Trim(i)) = 0 Then
                .Range(i.Address) = ""
            End If
        Next

        'sort unmatched list
        Sheets("Codes").Range("I4:I103").Sort Key1:=.Range("I4"), Order1:=xlAscending

        'rebuild unmatched list hlinks
        i = "Sheets(""Codes"").Range(""I4:I103"")"      'global var to pass to macro
        Call SentLink

        'exit sub
        GoTo ext

    End With

mtransfer: 'transfer to another month

     'acquire month for transfer
ask: i = MsgBox(Space(6) & "Transfer to the previous month or the following month?" & Chr(10) & Chr(10) _
             & Space(26) & "[ previous = yes  |  following = no ]", 3, "Transfer / Delete")

    If i = 2 Then GoTo ext
    If i = 6 Then t = ActiveSheet.index - 1
    If i = 7 Then t = ActiveSheet.index + 1
    If t < 1 Or t > 12 Then GoTo ask

    'acquire sheet name
    t = MonthName(t, 1)

    'transfer the transaction
    With Sheets(t)

        'if charge entry transfer to new sheet's charge transactions
        If Split(ActiveCell.Address, "$")(1) = "D" Then

            'check if duplicate
            For Each i In .Range("E4:E203")
                If i = ActiveCell.Offset(0, 1) And i.Offset(0, 1) = ActiveCell.Offset(0, 2) And _
                    i.Offset(0, -2) = ActiveCell.Offset(0, -1) Then
                    MsgBox Space(5) & "The transaction is a duplicate and will not be transferred.", 0, "Duplicate"
                    GoTo ext
                End If
            Next

            'acquire last row of charge column
            i = .Cells(Rows.Count, "C").End(xlUp).Row + 1

            'copy/transfer transaction
            .Range("B" & i) = ActiveCell.Offset(0, -2)
            .Range("C" & i) = ActiveCell.Offset(0, -1)
            .Range("E" & i) = ActiveCell.Offset(0, 1)
            .Range("F" & i) = ActiveCell.Offset(0, 2)
            .Range("G" & i) = ActiveCell.Offset(0, 3)
            .Range("H" & i) = ActiveCell.Offset(0, 4)

            'sort charge column
            Sheets(t).Range("B4:H" & i).Sort Key1:=Sheets(t).Range("C4"), Order1:=xlAscending

            'remove any wrap-text of description
            Sheets(t).Range("B4:H" & i).WrapText = False

            'acquire range to clear
            i = ActiveCell.Offset(0, -2).Address & ":" & ActiveCell.Offset(0, -1).Address & ","
            i = i & ActiveCell.Offset(0, 1).Address & ":" & ActiveCell.Offset(0, 4).Address

            'clear content, sort, rebuild hlinks
            GoTo clr

        'if expense entry transfer to new sheet's expense transactions
        ElseIf Split(ActiveCell.Address, "$")(1) = "P" Then

            'check if duplicate
            For Each i In .Range("Q4:Q203")
                If i = ActiveCell.Offset(0, 1) And i.Offset(0, 1) = ActiveCell.Offset(0, 2) Then
                    MsgBox Space(5) & "The transaction is a duplicate and will not be transferred.", 0, "Duplicate"
                    GoTo ext
                End If
            Next

            'acquire last row of expense column
            i = .Cells(Rows.Count, "O").End(xlUp).Row + 1

            'copy/transfer transaction
            .Range("O" & i) = ActiveCell.Offset(0, -1)
            .Range("Q" & i) = ActiveCell.Offset(0, 1)
            .Range("R" & i) = ActiveCell.Offset(0, 2)
            .Range("S" & i) = ActiveCell.Offset(0, 3)
            .Range("T" & i) = ActiveCell.Offset(0, 4)

            'sort expense column
            Sheets(t).Range("O4:T" & i).Sort Key1:=Sheets(t).Range("O4"), Order1:=xlAscending

            'remove any wrap-text of description
            Sheets(t).Range("O4:T" & i).WrapText = False

            'acquire range to clear
            i = ActiveCell.Offset(0, -1).Address & ","
            i = i & ActiveCell.Offset(0, 1).Address & ":" & ActiveCell.Offset(0, 4).Address

            'clear content, sort, rebuild hlinks
            GoTo clr

        End If

    End With

    'exit sub
    GoTo ext

End Sub
