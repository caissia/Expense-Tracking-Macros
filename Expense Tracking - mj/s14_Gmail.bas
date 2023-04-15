Attribute VB_Name = "s14_Gmail"
Option Explicit
'Contains 4 macros: Gmail, MonthLimit, Summary, YearLimit

Sub Gmail()
'View / Limit, switchboard to send email

    Dim a, b, c
    Dim i, s

    If Month(Date) = 1 Or Month(Date) = 12 Then

        'set vars for msg
        s = 6

        a = "Send a monthly or year-end spending limit report?"
        b = "[ monthly = yes  |  year-end = no ]"
        c = Len(a) - Len(b) + s

        'alert of successful completion
        i = MsgBox(Space(s) & a & Chr(10) _
                  & Space(c) & b & Chr(10), 3, "Report")
        If i = 6 Then Call MonthLimit
        If i = 7 Then Call YearLimit

    Else

        Call MonthLimit

    End If

End Sub

Sub MonthLimit()
'View, send emails via gmail of spending limits

    'create CDO object
    Dim NewMail As CDO.Message
    Set NewMail = New CDO.Message

    Dim s
    Dim msg, rng
    Dim diff, over, under
    Dim category, max, spent

    'set range
    Set rng = ThisWorkbook.Sheets("View").Range("Q63:Q72")

    'check if data present & check if current year matches wb year - else exit
    If Sheets("View").Range("L63") = "" Then Exit Sub
    If Year(Date) <> Mid(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 4, 4) * 1 Then Exit Sub

    'activate so macro can access data
    ThisWorkbook.Activate
    Sheets("View").Activate

    'enable SSL Authentication
    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

    'make SMTP authentication Enabled=true (1)
    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

    'set the SMTP server and port Details
    'to get these details you can get on Settings Page of your Gmail Account
    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"

    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

    'set your credentials of your Gmail Account
    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "caissia@gmail.com"

    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "!4gA02$3Sh00*2mI73%1Jo71"

    'update the configuration fields
    NewMail.Configuration.Fields.Update

    'acquire category data for msg
    For Each i In rng
        If i <> "Income" Then

            category = i
            spent = i.Offset(0, 1)
            max = i.Offset(0, 2)
            diff = Abs(i.Offset(0, 3))

            If category = "Home" Then category = "Housing"

            If spent > max Then

                GoSub frm
                over = over & Chr(10) & Chr(9) & category & s & " is over by  " & diff

            ElseIf spent >= (max * 0.8) Then

                GoSub frm
                under = under & Chr(10) & Chr(9) & category & s & " is under by  " & diff & "  -  " & Format((spent / max), "0%")

            End If

        End If
    Next

    'add headers to msg
    If Len(over) = 0 And Len(under) = 0 Then Exit Sub
    If Len(over) = 0 Then msg = "80% of Limit - Category" & Chr(10) & under
    If Len(under) = 0 Then msg = "Over Limit - Category" & Chr(10) & over
    If Len(over) <> 0 And Len(under) <> 0 Then _
        msg = "Over Limit - Category" & Chr(10) & over & Chr(10) & Chr(10) & _
              "80% of Limit - Category" & Chr(10) & under

    'Dec gift limit for xmas shopping based on yearly gift & shopping categories - add to msg
    If DatePart("m", Date) = 12 Then

        'check/add to max if gift limit reached
        For Each i In Sheets("View").Range("M18:M43")

            If Left(i.Offset(0, -1), 4) = "Gift" Then

                max = i * 12
                spent = Sheets("Sum").Range("I97")
                diff = Abs(max - spent)

            End If

        Next

        'add shopping to max if limit not reached
        For Each i In Sheets("View").Range("M18:M43")

            If i.Offset(0, -1) = "Shopping" Then

                s = Sheets("Sum").Range("I100")
                If s < (i * 12) Then
                    max = max + ((i * 12) - s)
                    diff = Abs(max - spent)
                End If

            End If

        Next

        'add xmas limit to msg
        If spent > max Then

            category = "Gifts"
            GoSub frm

            msg = msg & Chr(10) & Chr(10) & "* Xmas Spending Limit Reached *" & Chr(10) & Chr(10) & Chr(9) & _
                  category & s & " is over by  " & diff

        End If

        If spent < max Then

            category = "Gifts"
            GoSub frm

            msg = msg & Chr(10) & Chr(10) & "Xmas Spending" & Chr(10) & Chr(10) & Chr(9) & _
                  category & s & " is under by  " & diff & "  -  " & Format((spent / max), "0%")

        End If

    End If

    over = ""
    'yearly limit summmary - display during first week of month only - add to msg
    If Day(Date) >= 1 And Day(Date) <= 7 Then

        For Each i In Sheets("Items").Range("CM7:CM32")

            If i <> "" And i < 0 Then
    
                category = i.Offset(0, -4)
                diff = Abs(i)

                GoSub frm
                over = over & Chr(10) & Chr(9) & category & s & " is over by  " & diff
                
            End If

        Next

        msg = msg & Chr(10) & Chr(10) & "Yearly - Over Limit Items" & Chr(10) & over

    End If

    'set all email properties
    With NewMail
      .Subject = "Spending Update - " & Format(Date, "mmm d")
      .From = "caissia@gmail.com"
      .To = "caissia@gmail.com"
      .CC = ""
      .BCC = ""
      .TextBody = msg
      .ReplyTo = ""
    End With

    On Error Resume Next
    'send the email
    NewMail.Send

    'notify that it was sent
    'MsgBox ("Mail has been Sent")

    'set the NewMail Variable to Nothing
    Set NewMail = Nothing

    'clear & return error handling
    err.Clear
    On Error GoTo 0

    'activate original
    Sheets("Limit").Activate

Exit Sub

frm:

    'format data
    spent = Format(spent, "$0")
    max = Format(max, "$0")
    diff = Format(diff, "$0")

    'calc space to add
    s = ""

    If category = "Auto" Then s = s & "" & Space(14) & ""
    If category = "Care" Then s = s & "" & Space(14) & ""
    If category = "Entertain" Then s = s & "" & Space(7) & ""
    If category = "Food" Then s = s & "" & Space(13) & ""
    If category = "Housing" Then s = s & "" & Space(8) & ""
    If category = "Misc" Then s = s & "" & Space(14) & ""
    If category = "School" Then s = s & "" & Space(10) & ""
    If category = "Shopping" Then s = s & "" & Space(6) & ""
    If category = "Travel" Then s = s & "" & Space(12) & ""

    If category = "Gifts" Then s = s & "" & Space(13) & ""

    If category = "Auto Expenses" Then s = s & "" & Space(8) & ""
    If category = "Auto Insurance" Then s = s & "" & Space(8) & ""
    If category = "Auto Payment" Then s = s & "" & Space(10) & ""
    If category = "Cell Plan" Then s = s & "" & Space(18) & ""
    If category = "Electric Bill" Then s = s & "" & Space(15) & ""
    If category = "Entertainment" Then s = s & "" & Space(10) & ""
    If category = "Fee / Penalty" Then s = s & "" & Space(11) & ""
    If category = "Fuel - Auto" Then s = s & "" & Space(15) & ""
    If category = "Gas Bill" Then s = s & "" & Space(20) & ""
    If category = "Gift / Donation" Then s = s & "" & Space(10) & ""
    If category = "Groceries" Then s = s & "" & Space(17) & ""
    If category = "HOA" Then s = s & "" & Space(25) & ""
    If category = "Home Security" Then s = s & "" & Space(9) & ""
    If category = "Housing - Misc" Then s = s & "" & Space(9) & ""
    If category = "Internet / Phone" Then s = s & "" & Space(7) & ""
    If category = "Landscaping" Then s = s & "" & Space(12) & ""
    If category = "LL Academy" Then s = s & "" & Space(12) & ""
    If category = "Loan Edu" Then s = s & "" & Space(17) & ""
    If category = "Medical" Then s = s & "" & Space(20) & ""
    If category = "Mortgage" Then s = s & "" & Space(17) & ""
    If category = "Personal Care" Then s = s & "" & Space(9) & ""
    If category = "Restaurant" Then s = s & "" & Space(15) & ""
    If category = "Shopping" Then s = s & "" & Space(11) & ""
    If category = "Travel" Then s = s & "" & Space(10) & ""
    If category = "Tuition - Misc" Then s = s & "" & Space(11) & ""
    If category = "Water / Sewage" Then s = s & "" & Space(7) & ""

    Return

End Sub

Sub Summary()
'View, saves year-end summary as jpgs to email

    Dim i, h, w, rng, sht
    Dim FilePath As String
    Dim chtObj As ChartObject

    For i = 1 To 3

        'save location is the desktop
        FilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\YearEnd" & i & ".jpg"

        'set range and sheet
        If i = 1 Then rng = "CH4:CM33": sht = "Items"
        If i = 2 Then rng = "CO4:CT33": sht = "Items"
        If i = 3 Then rng = "C4:P32":   sht = "Sum"

        With ThisWorkbook.Sheets(sht)

            .Visible = xlSheetVisible
            .Activate
            .Unprotect
            .Range(rng).Select

            'copy selection & acquire size
            Selection.CopyPicture xlScreen, xlBitmap
            w = Selection.Width
            h = Selection.Height

            Set chtObj = .ChartObjects.Add(100, 30, 400, 250)
            chtObj.Name = "TemporaryPictureChart"

            'resize obj to picture size
            chtObj.Width = w
            chtObj.Height = h

            ActiveSheet.ChartObjects("TemporaryPictureChart").Activate
            ActiveChart.Paste

            ActiveChart.Export filename:=FilePath, FilterName:="jpg"

            chtObj.Delete

            Call View

            If ActiveSheet.Name = "Items" Then .Visible = xlVeryHidden
            If ActiveSheet.Name = "Sum" Then .Protect

        End With

    Next

    Application.ScreenUpdating = True

End Sub

Sub YearLimit()
'Limit, send emails via gmail of yearly spending

    Dim FilePath As String
    Dim NewMail As CDO.Message
    Set NewMail = New CDO.Message

    'basic setup to avoid unnecessary delays
    Application.Calculation = xlManual
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    'save year-end summary as jpgs to email
    Call Summary

    'set file path of documents to send
    FilePath = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    'check if data present & check if current year matches wb year - else exit
    If Sheets("View").Range("L63") = "" Then Exit Sub
    If Year(Date) <> Mid(ActiveWorkbook.Name, InStr(ActiveWorkbook.Name, ".") - 4, 4) * 1 Then Exit Sub

    'activate so macro can access data
    ThisWorkbook.Activate

    'enable SSL Authentication
    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

    'make SMTP authentication Enabled=true (1)
    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

    'set the SMTP server and port Details
    'to get these details you can get on Settings Page of your Gmail Account
    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"

    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

    'set your credentials of your Gmail Account
    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "caissia@gmail.com"

    NewMail.Configuration.Fields.Item _
    ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "!4gA02$3Sh00*2mI73%1Jo71"

    'update the configuration fields
    NewMail.Configuration.Fields.Update

    'set all email properties
    With NewMail
        .Subject = "Year-End Spending Update - " & Year(Date) - 1
        .From = "caissia@gmail.com"
        .To = "caissia@gmail.com"
        .CC = ""
        .BCC = ""
        .ReplyTo = ""
        .TextBody = "Attached is the spending summary for " & Year(Date) - 1
        .AddAttachment FilePath & "\YearEnd1.jpg"
        .AddAttachment FilePath & "\YearEnd2.jpg"
        .AddAttachment FilePath & "\YearEnd3.jpg"
    End With

    On Error Resume Next
    'send the email
    NewMail.Send

    'notify that it was sent
    'MsgBox ("Mail has been Sent")

    'set the NewMail Variable to Nothing
    Set NewMail = Nothing

    'clear & return error handling
    err.Clear
    On Error GoTo 0

    'activate original
    Sheets("Limit").Activate

    'rename and delete the sent files from desktop
    On Error Resume Next
    Name FilePath & "\YearEnd1.jpg" As FilePath & "\filename.dll": Kill FilePath & "\filename.dll"
    Name FilePath & "\YearEnd2.jpg" As FilePath & "\filename.dll": Kill FilePath & "\filename.dll"
    Name FilePath & "\YearEnd3.jpg" As FilePath & "\filename.dll": Kill FilePath & "\filename.dll"
    On Error GoTo 0

    'restore settings
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic

End Sub
