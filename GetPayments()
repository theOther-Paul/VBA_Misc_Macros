Sub GetPayments()
On Error GoTo Err_clear
Dim wbOpen As Workbook
Dim shOpen As Worksheet, shMain As Worksheet
Dim vPath As Variant
Dim lCol1 As Long, lCol2 As Long, lCol3 As Long, lColM As Long
Dim l_LastR As Long, l_Row As Long, l_RowM As Long, l_LastRM As Long
Dim mDate As Date, rDate As Date
Dim i As Integer, j As Integer
Dim sSrc As String, sMod As String, sCol As String
Dim vColsM As Variant, vColsS As Variant, vColsD As Variant

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
    End With

Set shMain = ActiveSheet

vColsM = Array("Amex", "Remise CB", "Remise cheque", "Versement especes") 'main array
vColsS = Array("Amex", "Carte bancaire", "Chèque Auto", "Espèces") 'corresponding array
vColsD = Array("C", "D", "E", "F") 'column range for asigning data in the main sheet
vPath = Application.GetOpenFilename(FileFilter:="Master file, *.xls; *.xlsm; *.xlsb; *.xlsx;", MultiSelect:=False)
If vPath <> vbNullString Then
    Set wbOpen = Application.Workbooks.Open(CStr(vPath), False, True)
Else
    GoTo Err_clear
End If

Set shOpen = wbOpen.Worksheets("synth")
lCol1 = getColNr(shOpen, "Source.Name")
lCol2 = getColNr(shOpen, "Synthèse Modes de Paiements")
lCol3 = getColNr(shOpen, "Column2")

    If lCol1 = 0 Or lCol2 = 0 Or lCol3 = 0 Then
        MsgBox "Check Import Procedure!", vbExclamation
        GoTo Err_clear
    End If

l_LastR = shMain.UsedRange.Rows.Count
l_LastRM = shOpen.UsedRange.Rows.Count
    For l_Row = 2 To l_LastR
        If shMain.Range("B" & l_Row) = "Montant CM" Then
            mDate = CDate(shMain.Range("A" & l_Row))
            For l_RowM = 2 To l_LastRM
                sSrc = CStr(shOpen.Cells(l_RowM, lCol1))
                rDate = extractDt(sSrc)
                If rDate = mDate Then
                    sMod = CStr(shOpen.Cells(l_RowM, lCol2))
                    j = getVectorID(sMod, vColsS)
                        If j > -1 Then
                            sMod = CStr(vColsM(j))
                            lColM = getColNr(shMain, sMod)
                            shMain.Cells(l_Row, lColM) = shOpen.Cells(l_RowM, lCol3)
                        End If
                End If
            Next l_RowM
        End If
        If shMain.Range("B" & l_Row) = "Delta" Then
            For i = 0 To UBound(vColsD)
                sCol = CStr(vColsD(i))
                shMain.Range(sCol & l_Row).Formula = "=" & Replace(shMain.Range(sCol & l_Row - 2).Address, "$", "") & "-" & Replace(shMain.Range(sCol & l_Row - 1).Address, "$", "")
            Next i
        End If
    Next l_Row

MsgBox "Process done!", vbInformation

Err_clear:
    If Err.Number <> 0 Then
        MsgBox "Failed processing: " & Err.Description, vbCritical
        Err.Clear
        'Resume Next
    End If

    If Not wbOpen Is Nothing Then
        wbOpen.Close False
        Set wbOpen = Nothing
    End If

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
    End With

End Sub

Function getVectorID(sElem As String, vCols As Variant) As Integer
Dim i As Integer

getVectorID = -1
    For i = 0 To UBound(vCols)
        If CStr(vCols(i)) = sElem Then
            getVectorID = i
            Exit For
        End If
    Next i

End Function

Function getColNr(shLook As Worksheet, colName As String) As Long
Dim l_Col As Long
Dim l_Cols As Long
l_Cols = shLook.UsedRange.Columns.Count

    For l_Col = 1 To l_Cols
        If UCase(shLook.Cells(1, l_Col).Value) = UCase(colName) Then
            getColNr = l_Col
            Exit For
        End If
    Next l_Col

    If l_Col = l_Cols Then
        l_Col = 0
    End If
    
End Function

 Private Sub ProcessPayments()
 Dim shT As Worksheet, lastRT As Long, shM As Worksheet, lastRM As Long, dict As Object
 Dim arr, arrInt, arrT, i As Long, j As Long, k As Long, arrMeth, mtch
 
 arrMeth = Split("Amex, Remise CB, Remise cheque, Versement especes", ",")

 Set shT = ActiveSheet 'the white sheet
 lastRT = shT.Range("A" & shT.Rows.Count).End(xlUp).Row  'last row in A:A
 arrT = shT.Range("A2:F" & lastRT).Value 'place the range in an array for faster iteration
 
Dim path As String
     'FileDialogFilePicker
    'path = Application.GetOpenFilename(FileFilter:="Master file, *.xls; *.xlsm; *.xlsb; *.xlsx;", MultiSelect:=False)
     
    'With Application.FileDialog(msoFileDialogFilePicker)
       ' .AllowMultiSelect = False    'Forces to choose 1 file.
      '  If .Show = -1 Then    'Checks if OK button was clicked.
     '   path = .SelectedItems(1)
    '    End If
   'End With
 'If Not UCase$(MyFile) Like "*.XLSM" Then 'will open any kind of excel files
 
path = Application.GetOpenFilename(FileFilter:="Master file, *.xls; *.xlsm; *.xlsb; *.xlsx;", MultiSelect:=False)
If vPath <> vbNullString Then
    Set wbOpen = Application.Workbooks.Open(CStr(vPath), False, True)
Else
    Exit Sub
End If

Set wb = Workbooks.Open(path)

 Set shM = wbOpen.Worksheets("synth")
 lastRM = shM.Range("A" & shM.Rows.Count).End(xlUp).Row 
 arr = shM.Range("A2:C" & lastRM).Value '
 Set dict = CreateObject("Scripting.Dictionary") 
 For i = 1 To UBound(arr)    
    If Not dict.Exists(arr(i, 1)) Then 
        dict.Add arr(i, 1), Array(Array(arr(i, 2), arr(i, 3))) 'place the item as an array of two elements (method and value)
    Else                     
        arrInt = dict(arr(i, 1)): ReDim Preserve arrInt(UBound(arrInt) + 1)  
        arrInt(UBound(arrInt)) = Array(arr(i, 2), arr(i, 3))                 
        dict(arr(i, 1)) = arrInt                                                          
    End If
 Next i

 For i = 1 To UBound(arrT)                    
    If arrT(i, 2) = "Montant CM" Then         
        For j = 0 To dict.Count - 1             
            If arrT(i, 1) = dict.Keys()(j) Then 
                For k = 0 To UBound(dict.items()(j)) 
                  
                    mtch = Application.Match(dict.items()(j)(k)(0), arrMeth, 0)
                    arrT(i, mtch + 2) = dict.items()(j)(k)(1)   
                Next k
            End If
        Next j
    End If
 Next i
 
shT.Range("H2").Resize(UBound(arrT), UBound(arrT, 2)).Value = arrT
End Sub

Private Function FO()
Dim path As String
     'FileDialogFilePicker
    With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False    'Forces to choose 1 file.
        If .Show = -1 Then    'Checks if OK button was clicked.
        path = .SelectedItems(1)
        End If
    End With

'If Not UCase$(MyFile) Like "*.XLSM" Then 'will open any kind of excel files
Set wb = Workbooks.Open(path)
'End If
End Function

Private Function extractDt(fileName As String) As Date 'will extract numbers from a string and convert it into date data type

   Dim strNum As String, Dt As Date
    With CreateObject("VBScript.RegExp")
        .Pattern = "(\d{4,6})" 'number digits, from 4 to 6
        .Global = True
        If .test(fileName) Then
            strNum = CStr(.Execute(fileName)(0))
        End If
    End With
    If Len(strNum) = 6 Then
        Dt = DateSerial(20 & CLng(Right(strNum, 2)), CLng(Mid(strNum, 3, 2)), CLng(Left(strNum, 2)))
   ElseIf Len(strNum) = 5 Then
        Dt = DateSerial(20 & CLng(Right(strNum, 2)), CLng(Mid(strNum, 2, 2)), CLng(Left(strNum, 1)))
   ElseIf Len(strNum) = 4 Then
        Dt = DateSerial(20 & CLng(Right(strNum, 2)), CLng(Mid(strNum, 2, 1)), CLng(Left(strNum, 1)))
   End If
   extractDt = Dt
End Function

Private Sub insextDt() 'works perfecly and production ready
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range("B:B").Insert
    ActiveCell.FormulaR1C1 = "=extractDt(RC[-1])"
    ActiveCell.Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A100")
    ActiveCell.Range("A1:A100").Select
    Selection.NumberFormat = "m/d/yyyy"
    ActiveCell.Offset(0, -1).Range("A1").Select
End Sub


