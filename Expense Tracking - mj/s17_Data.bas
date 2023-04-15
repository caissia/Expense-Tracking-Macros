Attribute VB_Name = "s17_Data"
Option Explicit

Sub Capture()
'Data, acquire all monthly data in table format

    Dim c       'last row of transactions to copy / loop var
    Dim p       'last row Data sheet to paste
    Dim n       'name of sheets for loop
    Dim s       'source sheet

    'exit if data missing
    For c = 1 To 12
        s = s + Len(Trim(Sheets(c).Range("C4")))
    Next
    If s = 0 Then Exit Sub

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'source sheet
    s = ActiveSheet.Name

    'clear data sheet to acquire fresh data
    Sheets("Data").Range("C5:I2404").ClearContents

    'transfer monthly transactions to Data sheet
    For i = 1 To 12

        'acquire/activate sheet
        n = MonthName(i, True)

        'check if month has data - charges
        If Sheets(n).Range("C4") <> "" Then

            'acquire last row of monthly sheets
            For Each c In Sheets(n).Range("C4:C203")
                If c = "" Then c = c.Offset(-1).Row: Exit For
            Next

            'find last row of Data sheet
            For Each p In Sheets("Data").Range("D5:D2404")
                If p = "" Then p = p.Row: Exit For
            Next

            'transfer data
            Sheets(n).Range("B4:H" & c).Copy
            Sheets("Data").Range("C" & p).PasteSpecial xlPasteValues
            Application.CutCopyMode = False

        End If

        'check if month has data - expenses
        If Sheets(n).Range("P4") <> "" Then

            'acquire last row of monthly sheets
            For Each c In Sheets(n).Range("O4:O203")
                If c = "" Then c = c.Offset(-1).Row: Exit For
            Next

            'find last row of Data sheet
            For Each p In Sheets("Data").Range("D5:D2404")
                If p = "" Then p = p.Row: Exit For
            Next

            'transfer data
            Sheets(n).Range("O4:T" & c).Copy
            Sheets("Data").Range("D" & p).PasteSpecial xlPasteValues
            Application.CutCopyMode = False

        End If

    Next

    'sort results based on date
    c = Sheets("Data").Range("D" & Sheets("Data").Rows.Count).End(xlUp).Row
    Sheets("Data").Range("C5:I" & c).Sort Key1:=Sheets("Data").Range("D5"), Order1:=xlAscending

    'set cursor
    Sheets("Data").Visible = xlSheetVisible
    Application.Goto Sheets("Data").Range("A1"), True
    Application.Goto Sheets("Data").Range("J3")
    Sheets("Data").Visible = xlVeryHidden
    Sheets(s).Activate

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub
