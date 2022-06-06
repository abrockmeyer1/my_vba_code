Attribute VB_Name = "Module1"
Sub Lock_Unlock()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        If ws.Name = "Assumptions" Then
            GoTo Skip1
        ElseIf ws.Name = "Parts list and Volumes" Then GoTo Skip1
        ElseIf ws.Name = "Master Sheet" Then GoTo Skip1
        ElseIf ws.Name = "Customer and Platform List" Then GoTo Skip1
        Else
            Range("A1:AX50").Locked = False
            Range("A1:AX2").Locked = True
            Range("G3:AX50").Locked = True
            Range("C3:E50").Locked = True
            ws.Protect Password:="TIFS#1", AllowFormattingCells:=True, AllowFormattingColumns:=True, _
                    AllowFormattingRows:=True, AllowInsertingColumns:=True, _
                    AllowInsertingRows:=True, AllowFiltering:=True, AllowDeletingRows:=True
        
        End If
Skip1:
    Next ws


End Sub
