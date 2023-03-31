Attribute VB_Name = "Module1"
Sub Refill()
Attribute Refill.VB_ProcData.VB_Invoke_Func = " \n14"
With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .AskToUpdateLinks = False
        .EnableEvents = False
        .Calculation = xlAutomatic
End With
ThisWorkbook.Date1904 = False
ActiveWindow.View = xlNormalView

Dim ws As Worksheet: Set ws = Sheets("Data")
'Dim Temp As Worksheet: Set Temp = Sheets("Temp")
Dim lr As Long, lc As Long
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

Dim rep1m3 As Range: Set rep1m3 = ws.Range("M3:Q9")
Dim rep1m2 As Range: Set rep1m2 = ws.Range("G3:K9")
Dim rep1m1 As Range: Set rep1m1 = ws.Range("A3:E9")

Dim rep2m3 As Range: Set rep2m3 = ws.Range("M14:Q16")
Dim rep2m2 As Range: Set rep2m2 = ws.Range("G14:K16")
Dim rep2m1 As Range: Set rep2m1 = ws.Range("A14:E16")

Dim rep3m3 As Range: Set rep3m3 = ws.Range("G22:H24")
Dim rep3m2 As Range: Set rep3m2 = ws.Range("D22:E24")
Dim rep3m1 As Range: Set rep3m1 = ws.Range("A22:B24")

Dim rep4m3 As Range: Set rep4m3 = ws.Range("G31:H33")
Dim rep4m2 As Range: Set rep4m2 = ws.Range("D31:E33")
Dim rep4m1 As Range: Set rep4m1 = ws.Range("A31:B33")

For i = 1 To 13
    Dim FileRep As String: FileRep = ThisWorkbook.Path & "/" & i & ".xlsx"
    If fso.FileExists(FileRep) Then
        Dim wb As Workbook
        Set wb = Workbooks.Open(FileRep)
        Dim wsTemp As Worksheet
        Set wsTemp = wb.Worksheets(1)
        Dim arr() As Variant
        arr = wsTemp.UsedRange.Value
        wb.Close False
        lr = UBound(arr, 1)
        lc = UBound(arr, 2)
        If i = 1 Then
            If lr = 7 And lc = 5 Then
                rep1m1.Value = rep1m2.Value
                rep1m2.Value = rep1m3.Value
                rep1m3.Value = arr
            Else
                Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
            End If
        If i = 2 Then
            If lr <= 3 And lc <= 5 Then
                rep2m1.Value = rep2m2.Value
                rep2m2.Value = rep2m3.Value
                rep2m3.Value = arr
            Else
                Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
            End If
        If i = 3 Then
            If arr.Range("A1").Value = "Производство" And arr.Range("A1").Value = "Количество необеспеченных" Then
                arr.Sort.SortFields.Add Key:=Range("B2:B" & lr) _
                    , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                With ActiveWorkbook.Worksheets("TDSheet").Sort
                    .SetRange Range("A1:D" & lr)
                    .Header = xlYes
                    .MatchCase = False
                    .Orientation = xlTopToBottom
                    .SortMethod = xlPinYin
                    .Apply
                End With
                rep3m1.Value = rep3m2.Value
                rep3m2.Value = rep3m3.Value
                rep3m3.Value = arr.Range("A3:B5").Value
            Else
                Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
            End If
        If i = 4 Then
            If Temp.Range("B1").Value = "Количество необеспеченных норм" And lc <= 2 Then
                rep4m1.Value = rep4m2.Value
                rep4m2.Value = rep4m3.Value
                rep4m3.Value = arr.Range("A3:B5").Value
            Else
                Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
            End If
    End If
Next i
 
With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .AskToUpdateLinks = True
        .DisplayAlerts = True
        .Calculation = xlAutomatic
        .StatusBar = False
    End With
ThisWorkbook.Date1904 = False
End Sub
