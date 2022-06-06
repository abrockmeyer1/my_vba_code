Attribute VB_Name = "Format_Sheets"
Sub Format_Sheets()
    Dim ws As Worksheet
    Dim Wb1 As Workbook
    Dim i As Integer
    
    Set Wb1 = Application.Workbooks("Rivian Supplier Capacity Data Verification Edit")
    
    i = 1
    For Each ws In Wb1.Sheets
        
        If ws.Name = "Supplier Part List" Then
            GoTo Skip1
        Else
            ws.Activate
            Range("J14:M33").Delete Shift:=xlToLeft
            Range("J14:M33").Copy
            Range("AH14:AK33").PasteSpecial xlPasteFormats
            For Each Cell In Range("H14:AK14")
                If Cell.Column Mod 2 = 0 Then
                    Cell.Formula = "Process " & i
                    i = i + 1
                Else
                End If
            
            Next Cell
            i = 1
        End If
            
    
    
    
Skip1:
    Next ws


End Sub
