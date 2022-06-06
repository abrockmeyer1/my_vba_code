Attribute VB_Name = "Goal_Seek"
Sub Goal_Seek()

    Dim ws As Worksheet
    
    Dim i As Long
    

            For Each Cell In Range("H15:AB15")
                If Cell = "FINAL ASSEMBLY" Then
                    GoTo Skip1
                Else
                    Cell.Select
                    Selection.Offset(18, 0).GoalSeek Goal:=300, ChangingCell:=Selection.Offset(2, 0)
                End If
            Next Cell

Skip1:


End Sub
