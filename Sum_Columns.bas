Attribute VB_Name = "Sum_Columns"
Sub Sum_Columns()

    
    i = 17
    g = 7
    e = 1
    
    
    Do While i <= 200
        Columns(i).Select
        Selection.Insert
        Cells(1, i) = "Sum of Week " & e
        Cells(2, i).Select
        Cells(2, i).Formula = "=Sum(RC[-1]:RC[-6])"
        Selection.Copy
        Selection.Offset(1, 0).Select
        Range(Selection, Selection.Offset(385, 0)).Select
        ActiveSheet.Paste
        
        e = e + 1
        i = i + g
        g = g
        
    Loop

End Sub
