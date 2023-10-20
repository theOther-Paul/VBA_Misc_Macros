Private Sub TextBox1_Change()
    Application.ScreenUpdating = False
    ActiveSheet.ListObjects("Table1").Range.AutoFilter field:=2, Criteria1:=[c3] & "*", Operator:=xlFilterValues
    Application.ScreenUpdating = True
End Sub