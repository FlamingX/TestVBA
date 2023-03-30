Sub Refill()
With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .AskToUpdateLinks = False
        .EnableEvents = False
        .Calculation = xlAutomatic
End With
ThisWorkbook.Date1904 = False
ActiveWindow.View = xlNormalView

Dim ws As Worksheets: Set ws = Sheets("Data")
Dim Temp As Worksheet: Set Temp = Sheets("Temp")
Dim FileExists As String, Temprng As Range, lr As Long, lc As Long
Dim PathOnly As String: PathOnly = ThisWorkbook.Path

Dim FileRep1 As String: FileRep1 = PathOnly + "/1.xlsx"
Dim rep1m3 As Range: Set rep1m3 = ws.Range("M3:Q9")
Dim rep1m2 As Range: Set rep1m2 = ws.Range("G3:K9")
Dim rep1m1 As Range: Set rep1m1 = ws.Range("A3:E9")
FileExists = Dir(FileRep1)

Temp.Cells.Clear
If FileExists <> "" Then
    rep1m1.UnMerge
    rep1m2.UnMerge
    rep1m3.UnMerge
    Workbooks.Open FileRep1
    Cells.Copy Temp.Range("A1")
    lr = Temp.UsedRange.Rows(Temp.UsedRange.Rows.Count).Row
    lc = Temp.UsedRange.Columns(Temp.UsedRange.Columns.Count).Column
    If lr <> 7 Or lc <> 5 Then
        MsgBox ("Íå óäàëîñü ðàñïîçíàòü øàáëîí îò÷¸òà " & FileRep1 & " " & lr & " " & lc)
        Else:
            Set Temprng = Temp.Range("A1:E7")
            rep1m2.Copy rep1m1
            rep1m3.Copy rep1m2
            rep1m3.ClearContents
            Temprng.UnMerge
            Temprng.Copy
            rep1m3.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            Temp.UsedRange.Delete
            Temp.Cells.Clear
    End If
    Workbooks("1.xlsx").Close SaveChanges:=False
    Temp.UsedRange.Delete
    Temp.Cells.Clear
End If


Dim FileRep2 As String: FileRep2 = PathOnly + "/2.xlsx"
Dim rep2m3 As Range: Set rep2m3 = ws.Range("M14:Q16")
Dim rep2m2 As Range: Set rep2m2 = ws.Range("G14:K16")
Dim rep2m1 As Range: Set rep2m1 = ws.Range("A14:E16")
FileExists = Dir(FileRep2)

If FileExists <> "" Then
    rep2m1.UnMerge
    rep2m2.UnMerge
    rep2m3.UnMerge
    Workbooks.Open FileRep2
    Temp.Cells.Clear
    Cells.Copy Temp.Range("A1")
    Workbooks("2.xlsx").Close SaveChanges:=False
    lr = Temp.UsedRange.Rows(Temp.UsedRange.Rows.Count).Row
    lc = Temp.UsedRange.Columns(Temp.UsedRange.Columns.Count).Column
    If lr > 3 Or lc > 5 Then
        MsgBox ("Íå óäàëîñü ðàñïîçíàòü øàáëîí îò÷¸òà " & FileRep2)
        Else:
            Set Temprng = Temp.Range("A1:E3")
            rep2m2.Copy rep2m1
            rep2m3.Copy rep2m2
            rep2m3.ClearContents
            Temprng.UnMerge
            Temprng.Copy
            rep2m3.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End If
End If


Dim FileRep3 As String: FileRep3 = PathOnly + "/3.xlsx"
Dim rep3m3 As Range: Set rep3m3 = ws.Range("G22:H24")
Dim rep3m2 As Range: Set rep3m2 = ws.Range("D22:E24")
Dim rep3m1 As Range: Set rep3m1 = ws.Range("A22:B24")
FileExists = Dir(FileRep3)

If FileExists <> "" Then
    rep3m1.UnMerge
    rep3m2.UnMerge
    rep3m3.UnMerge
    Workbooks.Open FileRep3
    Temp.Cells.Clear
    Cells.Copy Temp.Range("A1")
    Workbooks("3.xlsx").Close SaveChanges:=False
    Set Temprng = Temp.Range("A3:B5")
    Temprng.UnMerge
    If Temp.Range("A1").Value <> "Ïðîèçâîäñòâî" Or Temp.Range("A1").Value <> "Êîëè÷åñòâî íåîáåñïå÷åííûõ" Then
        MsgBox ("Íå óäàëîñü ðàñïîçíàòü øàáëîí îò÷¸òà " & FileRep3)
        Else:
            lr = Temp.UsedRange.Rows(Temp.UsedRange.Rows.Count).Row
            Temp.Sort.SortFields.Add Key:=Range("B2:B" & lr) _
                , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            With ActiveWorkbook.Worksheets("TDSheet").Sort
                .SetRange Range("A1:D" & lr)
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
            rep3m2.Copy rep3m1
            rep3m3.Copy rep3m2
            rep3m3.ClearContents
            Set Temprng = Temp.Range("A3:B5")
            Temprng.Copy
            rep3m3.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End If
End If



Dim FileRep4 As String: FileRep4 = PathOnly + "/4.xlsx"
Dim rep4m3 As Range: Set rep4m3 = ws.Range("G31:H33")
Dim rep4m2 As Range: Set rep4m2 = ws.Range("D31:E33")
Dim rep4m1 As Range: Set rep4m1 = ws.Range("A31:B33")
FileExists = Dir(FileRep4)

If FileExists <> "" Then
    rep4m1.UnMerge
    rep4m2.UnMerge
    rep4m3.UnMerge
    Workbooks.Open FileRep4
    Temp.Cells.Clear
    Cells.Copy Temp.Range("A1")
    Workbooks("4.xlsx").Close SaveChanges:=False
    lr = Temp.UsedRange.Rows(Temp.UsedRange.Rows.Count).Row
    lc = Temp.UsedRange.Columns(Temp.UsedRange.Columns.Count).Column
    If Temp.Range("B1").Value <> "Êîëè÷åñòâî íåîáåñïå÷åííûõ íîðì" Or lc > 2 Then
        MsgBox ("Íå óäàëîñü ðàñïîçíàòü øàáëîí îò÷¸òà " & FileRep4)
        Else:
            Set Temprng = Temp.Range("A2:B4")
            rep4m2.Copy rep4m1
            rep4m3.Copy rep4m2
            rep4m3.ClearContents
            Temprng.UnMerge
            Temprng.Copy
            rep4m3.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End If
End If


'Dim rep5m3 As Range: Set rep5m3 = ws.Range("K40:N43")
'Dim rep5m2 As Range: Set rep5m2 = ws.Range("F40:I43")
'Dim rep5m1 As Range: Set rep5m1 = ws.Range("A40:D43")
'
'    rep5m2.Copy rep5m1
'    rep5m3.Copy rep5m2
'    rep5m3.ClearContents
'
'Dim rep6m3 As Range: Set rep6m3 = ws.Range("I51:K51")
'Dim rep6m2 As Range: Set rep6m2 = ws.Range("E51:G51")
'Dim rep6m1 As Range: Set rep6m1 = ws.Range("A51:C51")
'
'    rep6m2.Copy rep6m1
'    rep6m3.Copy rep6m2
'    rep6m3.ClearContents
'
'Dim rep6m3p2 As Range: Set rep6m3p2 = ws.Range("J52:K55")
'Dim rep6m2p2 As Range: Set rep6m2p2 = ws.Range("F52:G55")
'Dim rep6m1p2 As Range: Set rep6m1p2 = ws.Range("B52:C55")
'
'    rep6m2p2.Copy rep6m1p2
'    rep6m3p2.Copy rep6m2p2
'    rep6m3p2.ClearContents
'
'Dim rep7m3 As Range: Set rep7m3 = ws.Range("K61:N63")
'Dim rep7m2 As Range: Set rep7m2 = ws.Range("F61:I63")
'Dim rep7m1 As Range: Set rep7m1 = ws.Range("A61:D63")
'
'    rep7m2.Copy rep7m1
'    rep7m3.Copy rep7m2
'    rep7m3.ClearContents
'
'Dim rep8m3 As Range: Set rep8m3 = ws.Range("I69:K74")
'Dim rep8m2 As Range: Set rep8m2 = ws.Range("E69:G74")
'Dim rep8m1 As Range: Set rep8m1 = ws.Range("A69:C74")
'
'    rep8m2.Copy rep8m1
'    rep8m3.Copy rep8m2
'    rep8m3.ClearContents
'
'Dim rep9 As Range: Set rep9 = ws.Range("A69:C74")
'
'Dim rep10m3 As Range: Set rep10m3 = ws.Range("G89:H91")
'Dim rep10m2 As Range: Set rep10m2 = ws.Range("D89:E91")
'Dim rep10m1 As Range: Set rep10m1 = ws.Range("A89:B91")
'
'    rep10m2.Copy rep10m1
'    rep10m3.Copy rep10m2
'    rep10m3.ClearContents
'
'Dim rep11m3 As Range: Set rep11m3 = ws.Range("G96:H98")
'Dim rep11m2 As Range: Set rep11m2 = ws.Range("D96:E98")
'Dim rep11m1 As Range: Set rep11m1 = ws.Range("A96:B98")
'
'    rep11m2.Copy rep11m1
'    rep11m3.Copy rep11m2
'    rep11m3.ClearContents
'
'Dim rep12m3 As Range: Set rep12m3 = ws.Range("E104")
'Dim rep12m2 As Range: Set rep12m2 = ws.Range("C104")
'Dim rep12m1 As Range: Set rep12m1 = ws.Range("A104")
'
'    rep12m2.Copy rep12m1
'    rep12m3.Copy rep12m2
'    rep12m3.ClearContents
'
'Dim rep13 As Range: Set rep13 = ws.Range("A109:D114")
'
'
'
'
'Dim FileRep3 As String: FileRep3 = PathOnly + "/3.xlsx"
'Dim FileRep4 As String: FileRep4 = PathOnly + "/4.xlsx"
'Dim FileRep5 As String: FileRep5 = PathOnly + "/5.xlsx"
'Dim FileRep6 As String: FileRep6 = PathOnly + "/6.xlsx"
'Dim FileRep7 As String: FileRep7 = PathOnly + "/7.xlsx"
'Dim FileRep8 As String: FileRep8 = PathOnly + "/8.xlsx"
'Dim FileRep9 As String: FileRep9 = PathOnly + "/9.xlsx"
'Dim FileRep10 As String: FileRep10 = PathOnly + "/10.xlsx"
'Dim FileRep11 As String: FileRep11 = PathOnly + "/11.xlsx"
'Dim FileRep12 As String: FileRep12 = PathOnly + "/12.xlsx"
'Dim FileRep13 As String: FileRep13 = PathOnly + "/13.xlsx"

'Workbooks.Open FileRep3
'Workbooks.Open FileRep4
'Workbooks.Open FileRep5
'Workbooks.Open FileRep6
'Workbooks.Open FileRep7
'Workbooks.Open FileRep8
'Workbooks.Open FileRep9
'Workbooks.Open FileRep10
'Workbooks.Open FileRep11
'Workbooks.Open FileRep12
'Workbooks.Open FileRep13


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