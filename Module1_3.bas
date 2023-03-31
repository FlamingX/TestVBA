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

Dim repRanges(1 To 13, 1 To 3) As Range
repRanges(1, 1) = ws.Range("A3:E9"): repRanges(1, 2) = ws.Range("G3:K9"): repRanges(1, 3) = ws.Range("M3:Q9")
repRanges(2, 1) = ws.Range("A14:E16"): repRanges(2, 2) = ws.Range("G14:K16"): repRanges(2, 3) = ws.Range("M14:Q16")
repRanges(3, 1) = ws.Range("A22:B24"): repRanges(3, 2) = ws.Range("D22:E24"): repRanges(3, 3) = ws.Range("G22:H24")
repRanges(4, 1) = ws.Range("A31:B33"): repRanges(4, 2) = ws.Range("D31:E33"): repRanges(4, 3) = ws.Range("G31:H33")
repRanges(5, 1) = ws.Range("A40:D43"): repRanges(5, 2) = ws.Range("F40:I43"): repRanges(5, 3) = ws.Range("K40:N43")
repRanges(6, 1) = ws.Range("A51:C51"): repRanges(6, 2) = ws.Range("E51:G51"): repRanges(6, 3) = ws.Range("I51:K51")
repRanges(7, 1) = ws.Range("A61:D63"): repRanges(7, 2) = ws.Range("F61:I63"): repRanges(7, 3) = ws.Range("K61:N63")
repRanges(8, 1) = ws.Range("A69:C74"): repRanges(8, 2) = ws.Range("E69:G74"): repRanges(8, 3) = ws.Range("I69:K74")
repRanges(9, 1) = ws.Range("A69:C74"): repRanges(9, 2) = Nothing: repRanges(9, 3) = Nothing
repRanges(10, 1) = ws.Range("A89:B91"): repRanges(10, 2) = ws.Range("D89:E91"): repRanges(10, 3) = ws.Range("G89:H91")
repRanges(11, 1) = ws.Range("A96:B98"): repRanges(11, 2) = ws.Range("D96:E98"): repRanges(11, 3) = ws.Range("G96:H98")
repRanges(12, 1) = ws.Range("A104"): repRanges(12, 2) = ws.Range("C104"): repRanges(12, 3) = ws.Range("E104")
repRanges(13, 1) = ws.Range("A109:D114"): repRanges(13, 2) = Nothing: repRanges(13, 3) = Nothing

For i = 1 To 12
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
        Select Case i
            Case 1
                If lr = 7 And lc = 5 Then
                    repRanges(1, 1).Value = repRanges(1, 2).Value
                    repRanges(1, 2).Value = repRanges(1, 3).Value
                    repRanges(1, 3).Value = arr
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
                End If
            Case 2
                If lr <= 3 And lc <= 5 Then
                    repRanges(2, 1).Value = repRanges(2, 2).Value
                    repRanges(2, 2).Value = repRanges(2, 3).Value
                    repRanges(2, 3).Value = arr
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
                End If
            Case 3
                If arr(1, 1) = "Производство" And arr(1, 2) = "Количество необеспеченных" Then
                    With wsTemp.Sort.SortFields
                        .Clear
                        .Add Key:=Range("B2:B" & lr), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                    End With
                    With wsTemp.Sort
                        .SetRange Range("A1:D" & lr)
                        .Header = xlYes
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        .SortMethod = xlPinYin
                        .Apply
                    End With
                    repRanges(3, 1).Value = repRanges(3, 2).Value
                    repRanges(3, 2).Value = repRanges(3, 3).Value
                    repRanges(3, 3).Value = wsTemp.Range("A3:B5").Value
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
                End If
            Case 4
                If arr(1, 2) = "Количество необеспеченных норм" And lc = 2 Then
                    repRanges(4, 1).Value = repRanges(4, 2).Value
                    repRanges(4, 2).Value = repRanges(4, 3).Value
                    repRanges(4, 3).Value = wsTemp.Range("A3:B5").Value
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
                End If
            Case 5
                If arr(1, 5) = "Просроченные выдачи" And arr(1, 4) = "Выдано в месяце" Then
                    repRanges(5, 1).Value = repRanges(5, 2).Value
                    repRanges(5, 2).Value = repRanges(5, 3).Value
                    repRanges(5, 3).Value = ""
                    Dim cell As Range
                    For Each cell In wsTemp.Range("B1:B" & wsTemp.Cells(wsTemp.Rows.Count, "B").End(xlUp).Row)
                        If cell.Value = "Костюмы" Then
                            repRanges(5, 3).Rows(1).Value = cell.Offset(0, 3).Value
                        ElseIf cell.Value = "Обувь" Then
                            repRanges(5, 3).Rows(2).Value = cell.Offset(0, 3).Value
                        ElseIf cell.Value = "Футболки" Then
                            repRanges(5, 3).Rows(3).Value = cell.Offset(0, 3).Value
                        ElseIf cell.Value = "Термобельё" Then
                            repRanges(5, 3).Rows(4).Value = cell.Offset(0, 3).Value
                        End If
                    Next cell
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
                End If
'ыыфвфвы фывфывфыфывфы фывфывфывф
'            Case 6
'                If arr(1, 2) = "Количество необеспеченных норм" And lc = 2 Then
'                    rep6m1.Value = rep6m2.Value
'                    rep6m2.Value = rep6m3.Value
'                    rep6m3.Value = ""
'                Else
'                    Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
'                End If
            Case 7
                If lr = 3 And lc = 4 Then
                    repRanges(7, 1).Value = repRanges(7, 2).Value
                    repRanges(7, 2).Value = repRanges(7, 3).Value
                    repRanges(7, 3).Value = arr
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
                End If
'                dsasddasdas
'            Case 8
            Case 9
                If lr = 5 And lc = 6 Then
                    repRanges(9, 1).Value = arr
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
                End If
            Case 10
                If lr = 3 And lc = 2 Then
                    repRanges(10, 1).Value = repRanges(10, 2).Value
                    repRanges(10, 2).Value = repRanges(10, 3).Value
                    repRanges(10, 3).Value = arr
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
                End If
            Case 11
                If lr = 3 And lc = 2 Then
                    repRanges(11, 1).Value = repRanges(11, 2).Value
                    repRanges(11, 2).Value = repRanges(11, 3).Value
                    repRanges(11, 3).Value = arr
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & FileRep)
                End If
            Case 12
                If lr = 6 And lc = 4 Then
                    repRanges(13, 1).Value = arr
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
