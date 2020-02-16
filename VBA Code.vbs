Sub helloworld()
Dim str As String, pos As Integer, com As String, count As Integer, Ldate As String, ind As Integer
Dim OutApp As Object, limit As Integer, hiber As String, hiber1 As String, lastrow As Integer
Dim st1 As String, st2 As String, st3 As String, gth1 As Integer, gth2 As Integer, ip As String
Dim OutMail As Object
Dim strbody As String
Ldate = Date
With ActiveSheet
    lastrow = .Cells(.Rows.count, "A").End(xlUp).Row
End With
For i = 2 To lastrow
        str1 = "nslookup"
        str2 = Cells(i, 1)
        com = "nslookup" & " " & Cells(i, 1)
        str = CreateObject("WScript.shell").Exec(com).StdOut.ReadAll
        Set objShell = CreateObject("Wscript.Shell")
        pos = InStr(str, "dns2.uchicago.edu")
        'count = Cells(15, 15)
        If pos = 0 Then
         Cells(i, 2) = "PC is OFF"
        Else
             Cells(i, 2) = "PC is ON"
             gth2 = InStr(str, "Address:")
             st2 = Mid(str, gth2 + 8, Len(str))
             gth1 = InStr(st2, "Address:")
             st3 = Mid(st2, gth1 + 8, Len(st2))
             Cells(i, 3) = Trim(st3)
             Cells(i, 3) = Left(Cells(i, 3), Len(Cells(i, 3)) - 4)
             hiber1 = "powershell.exe -noexit WMIC /node:" & Cells(i, 3) & " process call create 'powershell.exe /c shutdown -h'"
             hiber = CreateObject("WScript.Shell").Run(hiber1)
             Set OutApp = CreateObject("Outlook.Application")
                 Set OutMail = OutApp.CreateItem(0)
                 Total = Total - limit
                 With OutMail
                .To = Cells(i, 4)
                .CC = ""
                .BCC = ""
                .Subject = "Please switch off your PC before leaving"
                .HTMLBody = strbody
                .Send   'or use .Send'
                 End With
             Set OutMail = Nothing
             Set OutApp = Nothing
      End If
Next i
        Cells(1, 2) = Ldate
End Sub
