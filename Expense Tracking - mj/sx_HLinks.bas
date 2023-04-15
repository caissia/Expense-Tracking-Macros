Attribute VB_Name = "sx_HLinks"
Option Explicit
'Contains 4 macros: AutoHyperlink, SingleLink, SelectLink, SentLink

Sub AutoHyperlink()
'automatically insert links on all monthly sheets
'shortcut: Ctrl + Shift + A

    Dim a, s, z
    Dim c, f, n, t
    Dim r, g, b

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'acquire range for hlinks
    a = "D4:D203,M3,N3,P4:P203"

    'insert hyperlinks
    For i = 1 To 12

        'acquire sheet name
        s = MonthName(i, True)

        For Each c In Sheets(s).Range(a)

            'acquire cell data
            f = c.Formula
            n = c.Font.Name
            z = c.Font.Size
            t = c.Font.Color

            If f = "" Then f = c.Value
            If Trim(f) = "" Then f = Space(50)

            r = t And 255
            g = t \ 256 And 255
            b = t \ 256 ^ 2 And 255

            'insert link
            c.Hyperlinks.Add Anchor:=c, Address:="", SubAddress:=s & "!" & c.Address & "", TextToDisplay:=f

            'restore font size and name
            c.Font.Color = rgb(r, g, b)
            c.Font.Name = n
            c.Font.Size = z

        Next

        'remove underline
        Sheets(s).Range(a).Font.Underline = xlUnderlineStyleNone

    Next

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub SingleLink()
Attribute SingleLink.VB_Description = "Rebuilds hyperlinks... only for sheets 1 - 12"
Attribute SingleLink.VB_ProcData.VB_Invoke_Func = "L\n14"
'insert hlinks in activesheet, only for s1 - s12 sheets
'shortcut: Ctrl + Shift + L

    Dim a, s
    Dim c, f, n, z, t
    Dim r, g, b

    'exit if not monthly sheets
    s = ActiveSheet.index
    If s < 1 Or s > 12 Then Exit Sub

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'acquire sheet name
    s = ActiveSheet.Name

    'acquire range for hlinks
    a = "D4:D203,M3,N3,P4:P203"

    'insert hlinks
    For Each c In Sheets(s).Range(a)

        'acquire cell data
        f = c.Formula
        n = c.Font.Name
        z = c.Font.Size
        t = c.Font.Color

        If f = "" Then f = c.Value
        If Trim(f) = "" Then f = Space(50)

        r = t And 255
        g = t \ 256 And 255
        b = t \ 256 ^ 2 And 255

        'insert link
        c.Hyperlinks.Add Anchor:=c, Address:="", SubAddress:=s & "!" & c.Address & "", TextToDisplay:=f

        'restore font size and name
        c.Font.Size = z
        c.Font.Name = n
        c.Font.Color = rgb(r, g, b)

    Next

    'remove underline
    ActiveSheet.Range(a).Font.Underline = xlUnderlineStyleNone

    'skip if called from macro
    If i = True Then Exit Sub

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub SelectLink()
Attribute SelectLink.VB_Description = "Inserts self-reference hyperlinks in selection"
Attribute SelectLink.VB_ProcData.VB_Invoke_Func = "H\n14"
'insert hlinks on selected range
'shortcut: Ctrl + Shift + H

    Dim a, s, z
    Dim c, f, n, t
    Dim r, g, b

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'acquire range and sheet name
    a = Selection.Address
    s = ActiveSheet.Name

    'insert hlinks
    For Each c In Sheets(s).Range(a)

        'acquire cell data
        f = c.Formula
        n = c.Font.Name
        z = c.Font.Size
        t = c.Font.Color

        If f = "" Then f = c.Value
        If Trim(f) = "" Then f = Space(50)

        r = t And 255
        g = t \ 256 And 255
        b = t \ 256 ^ 2 And 255

        'insert link
        c.Hyperlinks.Add Anchor:=c, Address:="", SubAddress:=s & "!" & c.Address & "", TextToDisplay:=f

        'restore font size and name
        c.Font.Name = n
        c.Font.Size = z
        c.Font.Color = rgb(r, g, b)

    Next

    'remove underline
    ActiveSheet.Range(a).Font.Underline = xlUnderlineStyleNone

    'restore settings
    Application.Calculation = xlAutomatic
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True

End Sub

Sub SentLink()
'insert hlinks on sent/passed range (i)

    Dim a, s
    Dim c, f, n, t
    Dim r, g, b

    If Trim(i) = "" Or i = 0 Then Exit Sub
    If InStr(i, "Range") = 0 Then Exit Sub
    If InStr(i, "Sheet") = 0 Then Exit Sub

    'acquire range and sheet name
    s = Split(i, """")(1)
    a = Split(i, """")(3)

    'insert hlinks
    For Each c In Sheets(s).Range(a)

        'acquire cell data
        f = c.Formula
        n = c.Font.Name
        i = c.Font.Size
        t = c.Font.Color

        If f = "" Then f = c.Value
        If Trim(f) = "" Then f = Space(50)

        r = t And 255
        g = t \ 256 And 255
        b = t \ 256 ^ 2 And 255

        'insert link
        c.Hyperlinks.Add Anchor:=c, Address:="", SubAddress:=s & "!" & c.Address & "", TextToDisplay:=f

        'restore font size and name
        c.Font.Size = i
        c.Font.Name = n
        c.Font.Color = rgb(r, g, b)

    Next

    'remove underline
    Sheets(s).Range(a).Font.Underline = xlUnderlineStyleNone

End Sub
