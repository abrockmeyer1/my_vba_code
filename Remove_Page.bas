Attribute VB_Name = "Remove_Page"
Sub Remove_Page_Click()
    Dim iLast As Integer
    Dim i As Integer
    Dim LastRow As Integer
    
    Application.ScreenUpdating = False
    
    i = 1
    iLast = Application.ExecuteExcel4Macro("GET.DOCUMENT(50)")
    If Not iLast = 4 Then
        Cells.Find(What:="CONTROLLED DOCUMENT", SearchOrder:=xlByRows, SearchDirection:=xlNext).Select
        Do While i < iLast
            Cells.FindNext(After:=ActiveCell).Activate
            i = i + 1
        Loop
        LastRow = Selection.Row
    
        Range(Selection, Selection.Offset(-14, 0)).EntireRow.Select
        Selection.Delete
        Call Page_Number
    
        Application.ScreenUpdating = True
        Range("A1").Select
    Else
        Application.ScreenUpdating = True
        Range("A1").Select
    End If
End Sub


Sub Page_Number()
    Application.ScreenUpdating = False
    Dim Sh1 As Worksheet
    Dim iLast As Integer
    Dim i As Integer
    Dim xVPC As Integer
    Dim xHPC As Integer
    Dim xVPB As VPageBreak
    Dim xHPB As HPageBreak
    Dim xNumPage As Integer
    xHPC = 1
    xVPC = 1
    i = 1
    iLast = Application.ExecuteExcel4Macro("GET.DOCUMENT(50)")
    Set Sh1 = ActiveSheet
    Cells.Find(What:="PAGE", SearchOrder:=xlByRows, SearchDirection:=xlNext).Select
    Do While i <= iLast + 1
        If ActiveSheet.PageSetup.Order = xlDownThenOver Then
            xHPC = ActiveSheet.HPageBreaks.Count + 1
        Else
            xVPC = ActiveSheet.VPageBreaks.Count + 1
        End If
        xNumPage = 1
        For Each xVPB In ActiveSheet.VPageBreaks
            If xVPB.Location.Column > ActiveCell.Column Then Exit For
            xNumPage = xNumPage + xHPC
        Next
        For Each xHPB In ActiveSheet.HPageBreaks
            If xHPB.Location.Row > ActiveCell.Row Then Exit For
            xNumPage = xNumPage + xVPC
        Next
        ActiveCell = "PAGE " & xNumPage & " OF " & Application.ExecuteExcel4Macro("GET.DOCUMENT(50)")
        Cells.FindNext(After:=ActiveCell).Activate
        i = i + 1
    Loop
    
    Application.ScreenUpdating = True
End Sub


