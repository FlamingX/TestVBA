Attribute VB_Name = "Module1"
Option Explicit

Sub Refill()
Attribute Refill.VB_ProcData.VB_Invoke_Func = " \n14"
With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .AskToUpdateLinks = False
        .EnableEvents = False
End With
ThisWorkbook.Date1904 = False
ActiveWindow.View = xlNormalView

Dim ws As Worksheet: Set ws = Sheets("Data")
'Dim Temp As Worksheet: Set Temp = Sheets("Temp")
Dim LastRow As Long, LastCol As Long, ReportIndex As Integer
Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")

Dim ReportRanges(1 To 13, 1 To 3) As Range
Set ReportRanges(1, 1) = ws.Range("A3:E9"): Set ReportRanges(1, 2) = ws.Range("G3:K9"): Set ReportRanges(1, 3) = ws.Range("M3:Q9")
Set ReportRanges(2, 1) = ws.Range("A14:E16"): Set ReportRanges(2, 2) = ws.Range("G14:K16"): Set ReportRanges(2, 3) = ws.Range("M14:Q16")
Set ReportRanges(3, 1) = ws.Range("A22:B24"): Set ReportRanges(3, 2) = ws.Range("D22:E24"): Set ReportRanges(3, 3) = ws.Range("G22:H24")
Set ReportRanges(4, 1) = ws.Range("A31:B33"): Set ReportRanges(4, 2) = ws.Range("D31:E33"): Set ReportRanges(4, 3) = ws.Range("G31:H33")
Set ReportRanges(5, 1) = ws.Range("A40:D43"): Set ReportRanges(5, 2) = ws.Range("F40:I43"): Set ReportRanges(5, 3) = ws.Range("K40:N43")
Set ReportRanges(6, 1) = ws.Range("A51:C51"): Set ReportRanges(6, 2) = ws.Range("E51:G51"): Set ReportRanges(6, 3) = ws.Range("I51:K51")
Set ReportRanges(7, 1) = ws.Range("A61:D63"): Set ReportRanges(7, 2) = ws.Range("F61:I63"): Set ReportRanges(7, 3) = ws.Range("K61:N63")
Set ReportRanges(8, 1) = ws.Range("A69:C74"): Set ReportRanges(8, 2) = ws.Range("E69:G74"): Set ReportRanges(8, 3) = ws.Range("I69:K74")
Set ReportRanges(9, 1) = ws.Range("A69:C74"): Set ReportRanges(9, 2) = Nothing: Set ReportRanges(9, 3) = Nothing
Set ReportRanges(10, 1) = ws.Range("A89:B91"): Set ReportRanges(10, 2) = ws.Range("D89:E91"): Set ReportRanges(10, 3) = ws.Range("G89:H91")
Set ReportRanges(11, 1) = ws.Range("A96:B98"): Set ReportRanges(11, 2) = ws.Range("D96:E98"): Set ReportRanges(11, 3) = ws.Range("G96:H98")
Set ReportRanges(12, 1) = ws.Range("A104"): Set ReportRanges(12, 2) = ws.Range("C104"): Set ReportRanges(12, 3) = ws.Range("E104")
Set ReportRanges(13, 1) = ws.Range("A109:D114"): Set ReportRanges(13, 2) = Nothing: Set ReportRanges(13, 3) = Nothing

For ReportIndex = 1 To 12
    Dim ReportName As String: ReportName = ThisWorkbook.Path & "/" & ReportIndex & ".xlsx"
    If fso.FileExists(ReportName) Then
        On Error Resume Next
        Dim wb As Workbook: Set wb = Workbooks.Open(ReportName)
        Dim wsTemp As Worksheet: Set wsTemp = wb.Worksheets(1)
        Dim TempArray() As Variant: TempArray = wsTemp.UsedRange.Value
        wb.Close False
        LastRow = UBound(TempArray, 1)
        LastCol = UBound(TempArray, 2)
        Select Case ReportIndex
            Case 1 To 12
                FillReportRange ReportIndex, TempArray, ReportRanges, LastRow, LastCol, ReportName, wsTemp
            Case Else
                Debug.Print "Ошибка заполнения: " & ReportIndex
        End Select
    End If
SkipFile:
Next ReportIndex

With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .AskToUpdateLinks = True
        .DisplayAlerts = True
        .StatusBar = False
    End With
ThisWorkbook.Date1904 = False
End Sub

Private Function FillReportRange(ReportIndex As Integer, TempArray As Variant, ReportRanges, LastRow As Long, LastCol As Long, ReportName As String, wsTemp As Worksheet)
Select Case ReportIndex
Case 1
                If LastRow = 7 And LastCol = 5 Then
                    ReportRanges(1, 1).Value = ReportRanges(1, 2).Value
                    ReportRanges(1, 2).Value = ReportRanges(1, 3).Value
                    ReportRanges(1, 3).Value = TempArray
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & ReportName)
                End If
            Case 2
                If LastRow <= 3 And LastCol <= 5 Then
                    ReportRanges(2, 1).Value = ReportRanges(2, 2).Value
                    ReportRanges(2, 2).Value = ReportRanges(2, 3).Value
                    ReportRanges(2, 3).Value = TempArray
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & ReportName)
                End If
            Case 3
                If TempArray(1, 1) = "Производство" And TempArray(1, 2) = "Количество необеспеченных" Then
                    With wsTemp.Sort.SortFields
                        .Clear
                        .Add Key:=Range("B2:B" & LastRow), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                    End With
                    With wsTemp.Sort
                        .SetRange Range("A1:D" & LastRow)
                        .Header = xlYes
                        .MatchCase = False
                        .Orientation = xlTopToBottom
                        .SortMethod = xlPinYin
                        .Apply
                    End With
                    ReportRanges(3, 1).Value = ReportRanges(3, 2).Value
                    ReportRanges(3, 2).Value = ReportRanges(3, 3).Value
                    ReportRanges(3, 3).Value = wsTemp.Range("A3:B5").Value
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & ReportName)
                End If
            Case 4
                If TempArray(1, 2) = "Количество необеспеченных норм" And LastCol = 2 Then
                    ReportRanges(4, 1).Value = ReportRanges(4, 2).Value
                    ReportRanges(4, 2).Value = ReportRanges(4, 3).Value
                    ReportRanges(4, 3).Value = wsTemp.Range("A3:B5").Value
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & ReportName)
                End If
            Case 5
                If TempArray(1, 5) = "Просроченные выдачи" And TempArray(1, 4) = "Выдано в месяце" Then
                    ReportRanges(5, 1).Value = ReportRanges(5, 2).Value
                    ReportRanges(5, 2).Value = ReportRanges(5, 3).Value
                    ReportRanges(5, 3).Value = ""
                    Dim cell As Range
                    For Each cell In wsTemp.Range("B1:B" & wsTemp.Cells(wsTemp.Rows.Count, "B").End(xlUp).Row)
                        If cell.Value = "Костюмы" Then
                            ReportRanges(5, 3).Rows(1).Value = cell.Resize(0, 3).Value
                        ElseIf cell.Value = "Обувь" Then
                            ReportRanges(5, 3).Rows(2).Value = cell.Resize(0, 3).Value
                        ElseIf cell.Value = "Футболки" Then
                            ReportRanges(5, 3).Rows(3).Value = cell.Resize(0, 3).Value
                        ElseIf cell.Value = "Термобельё" Then
                            ReportRanges(5, 3).Rows(4).Value = cell.Resize(0, 3).Value
                        End If
                    Next cell
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & ReportName)
                End If
            Case 6
                If TempArray(1, 2) = "Количество обращений план" And LastCol = 3 Then
                    ReportRanges(6, 1).Value = ReportRanges(6, 2).Value
                    ReportRanges(6, 2).Value = ReportRanges(6, 3).Value
                    ReportRanges(6, 3).Value = ""
                    For Each cell In wsTemp.Range("A4:A" & wsTemp.Cells(wsTemp.Rows.Count, "A").End(xlUp).Row)
                        ReportRanges(6, 3).Rows(1).Value = cell.Offset(0, 2).Resize(0, 3).Value
                        If cell.Value = "Работники" Then
                            ReportRanges(6, 3).Rows(2).Columns(2).Value = cell.Offset(0, 1).Resize(0, 2).Value
                        ElseIf cell.Value = "Футболки" Then
                            ReportRanges(5, 3).Rows(3).Columns(2).Value = cell.Offset(0, 1).Resize(0, 2).Value
                        ElseIf cell.Value = "Термобельё" Then
                            ReportRanges(5, 3).Rows(4).Columns(2).Value = cell.Offset(0, 1).Resize(0, 2).Value
                        End If
                    Next cell
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & ReportName)
                End If
            Case 7
                If LastRow = 3 And LastCol = 4 Then
                    ReportRanges(7, 1).Value = ReportRanges(7, 2).Value
                    ReportRanges(7, 2).Value = ReportRanges(7, 3).Value
                    ReportRanges(7, 3).Value = TempArray
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & ReportName)
                End If
            Case 8
                If TempArray(1, 3) = "% востребованности" And LastCol = 3 Then
                    ReportRanges(8, 1).Value = ReportRanges(8, 2).Value
                    ReportRanges(8, 2).Value = ReportRanges(8, 3).Value
                    ReportRanges(8, 3).Rows(1).Value = wsTemp.Range("A2:C4").Value
                    ReportRanges(8, 3).Rows(4).Value = wsTemp.Range("A" & LastRow - 3 & ",C" & LastRow - 1).Value
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & ReportName)
                End If
            Case 9
                If LastRow = 5 And LastCol = 6 Then
                    ReportRanges(9, 1).Value = TempArray
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & ReportName)
                End If
            Case 10
                If LastRow = 3 And LastCol = 2 Then
                    ReportRanges(10, 1).Value = ReportRanges(10, 2).Value
                    ReportRanges(10, 2).Value = ReportRanges(10, 3).Value
                    ReportRanges(10, 3).Value = TempArray
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & ReportName)
                End If
            Case 11
                If LastRow = 3 And LastCol = 2 Then
                    ReportRanges(11, 1).Value = ReportRanges(11, 2).Value
                    ReportRanges(11, 2).Value = ReportRanges(11, 3).Value
                    ReportRanges(11, 3).Value = TempArray
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & ReportName)
                End If
            Case 12
                If LastRow = 6 And LastCol = 4 Then
                    ReportRanges(13, 1).Value = TempArray
                Else
                    Debug.Print ("Не удалось распознать шаблон отчёта " & ReportName)
                End If
End Select
End Function
