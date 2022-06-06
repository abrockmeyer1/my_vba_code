Attribute VB_Name = "Select_File"
Sub Select_File_Click()

    Dim R As Range
    Dim MonDate As String
    Dim TuDate As String
    Dim ThuDate As String
    Dim FriDate As String
    Dim SaDate As String
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim wb2 As Workbook
    Dim Srch As String
    Dim strFile As String
    MonDate = Range("C3").Text
    
    Set ws1 = Sheets("Parts List and Volumes (Modify)")
    
    ws1.Select
    Range("G2").Select
    Range(Selection, Selection.End(xlDown)).ClearContents
    
    strFile = Application.GetOpenFilename(FileFilter:="Excel files (*.xlsm*), *.xlsm*", Title:="Choose an Excel file to open", MultiSelect:=False)
    
    Workbooks.Open Filename:=strFile
    Set wb2 = ActiveWorkbook
    Set ws2 = Sheets("Data")
    ws2.Activate
    
        ws1.Activate
        Range("D2").Select
        Range(Selection, Selection.End(xlDown)).Select
        
        For Each Cell In Selection
            Srch = Cell.Text
            ws2.Activate
            On Error GoTo Skip1
            
             Set R = Cells.Find(What:=Srch, After:=Cells(1, 1), Lookat:=xlWhole)
             
             If R Is Nothing Then
                Err.Clear
                GoTo Skip1
            Else
                R.Select
            End If
            
            Set R = Cells.Find(What:=Srch, After:=R, Lookat:=xlWhole)
            If R Is Nothing Then
                Err.Clear
                GoTo Skip1
            Else
                R.Select
            End If
            
            Set R = Cells.Find(What:=Srch, After:=R, Lookat:=xlWhole)
            If R Is Nothing Then
                Err.Clear
                GoTo Skip1
            Else
                R.Select
            End If
Skip1:
        Next Cell
        
    
    wb2.Close
    
    
End Sub
