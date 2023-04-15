Attribute VB_Name = "sx_A_Proc"
Option Explicit

Sub PostProcess()
'transfer charge/expense transactions
'to appropriate section based on codes

    Dim i, c

    'basic setup to avoid unnecessary delays
    Application.ScreenUpdating = False

    'remove scroll area limit
    ActiveSheet.ScrollArea = ""

    'check/transfer to expense list
    For Each c In Range("D4:D203")
        If Len(Trim(c)) = 1 Then
            i = InStr("ABDEFGMPRSV", c)
            If i = 0 Then
                GoSub Transfer
            End If
        End If
    Next

    'check/transfer to charge list
    For Each c In Range("P4:P203")
        If Len(Trim(c)) = 1 Then
            i = InStr("ACHILTU", c)
            If i = 0 Then
                GoSub Transfer
            End If
        End If
    Next

    'sort charge & expense columns
    i = Cells(Rows.Count, "C").End(xlUp).Row + 1
    ActiveSheet.Range("B4:H" & i).Sort Key1:=Range("C4"), Order1:=xlAscending

    i = Cells(Rows.Count, "O").End(xlUp).Row + 1
    ActiveSheet.Range("O4:T" & i).Sort Key1:=Range("O4"), Order1:=xlAscending

    'remove wrap-text if any
    ActiveSheet.Range("B4:H" & i).WrapText = False
    ActiveSheet.Range("O4:T" & i).WrapText = False

    'reapply hyperlinks
    i = True
    Call SingleLink
    i = False

    'set scroll area
    Call SetScrollArea

    'restore settings
    Application.ScreenUpdating = True

Exit Sub

Transfer:

    If Split(c.Address, "$")(1) = "D" Then

        'acquire last row of expense column
        i = Cells(Rows.Count, "O").End(xlUp).Row + 1

        'copy/transfer transaction
        Range("O" & i).Value = Range("C" & c.Row).Value
        Range("Q" & i & ":" & "S" & i).Value = Range("E" & c.Row & ":" & "G" & c.Row).Value

        'clear the transferred transaction from the charge list
        Range("B" & c.Row & ":" & "C" & c.Row).ClearContents
        Range("E" & c.Row & ":" & "H" & c.Row).ClearContents

    End If

    If Split(c.Address, "$")(1) = "P" Then

        'acquire last row of charge column
        i = Cells(Rows.Count, "C").End(xlUp).Row + 1

        'copy/transfer transaction
        Range("B" & i) = "B": Range("H" & i) = "M"
        Range("C" & i).Value = Range("O" & c.Row).Value
        Range("E" & i & ":" & "G" & i).Value = Range("Q" & c.Row & ":" & "S" & c.Row).Value

        'clear the transferred transaction from the expense list
        Range("O" & c.Row & "," & "Q" & c.Row & ":" & "T" & c.Row).ClearContents

    End If

    Return

End Sub
