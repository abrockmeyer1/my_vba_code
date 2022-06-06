Attribute VB_Name = "Insert_Formula"
Sub Insert_Formula()

    Dim ws As Worksheet
    Dim Wb1 As Workbook
    Dim i As Integer
    Dim ShName As String
    
    Set Wb1 = Application.Workbooks("Rivian Supplier Capacity Data Verification Edit")
    
    i = 2
    
    Sheets("Supplier Part List").Activate
    
    For Each Cell In Range("J13:J43")
        ShName = Sheets(i).Name
        
        Cell.Formula = "='" & ShName & "'!H34"
        i = i + 1
    
    Next Cell

    
    

End Sub
