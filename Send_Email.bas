Attribute VB_Name = "Send_Email"
Sub Send_Email_Click()




Dim OutApp As Object
Dim OutMail As Object
Dim olInsp As Object
Dim xlSheet As Worksheet
Dim wdDoc As Object
Dim oRng As Object
    Set xlSheet = ActiveSheet
    xlSheet.Range("A1:T17").Copy

    On Error Resume Next
    Set OutApp = GetObject(, "Outlook.Application")
    If Err <> 0 Then Set OutApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    
    Set OutMail = OutApp.CreateItem(0)
    With OutMail
        .BodyFormat = 3
        .To = "abrockmeyer@tifs.com"
        .CC = ""
        .BCC = ""
        .Subject = "Today's Attendance Bingo Number"
        Set olInsp = .GetInspector
        Set wdDoc = olInsp.WordEditor
        Set oRng = wdDoc.Range
        oRng.collapse 1
        oRng.Paste
        .Display
        .Send
    End With
   
 
    Set OutMail = Nothing
    Set OutApp = Nothing
    Set olInsp = Nothing
    Set wdDoc = Nothing
    Set oRng = Nothing
    
Application.CutCopyMode = False


End Sub



