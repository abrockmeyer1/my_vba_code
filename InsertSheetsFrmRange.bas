Attribute VB_Name = "InsertSheetsFrmRange"
Sub Insert_Sheets_From_Range()

    Dim Wb1 As Workbook
    Dim Wb2 As Workbook
    Dim Sh1 As Worksheet
    Dim FrmSh As Worksheet
    Dim Val As String
    Dim ws As Worksheet
    Dim wsname As String
    Dim i As Variant
    Dim CopSht As String
    
    Set Wb1 = Application.Workbooks("Rivian Supplier Capacity Data Verification Edit")
    
    Wb1.Activate
    Set Sh1 = Wb1.Sheets(1)
    CopSht = Wb1.Sheets(2).Name
    Wb1.Sheets(CopSht).Activate
    For Each Cell In Sh1.Range("F13:F43")
        Val = Cell.Value
        
        ActiveSheet.Copy After:=Wb1.Sheets(CopSht)
        ActiveSheet.Name = Val
        CopSht = Val
        
        
    Next Cell
    

End Sub

