Attribute VB_Name = "档案袋封面生成"
Sub A档案袋封面打印()
Attribute A档案袋封面打印.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 档案袋封面打印 宏
'

'
    '''清空内容
    Sheets("封面准备").Cells.ClearContents
    Sheets("封面打印").ResetAllPageBreaks
    With Sheets("封面打印").Cells '
        .ClearContents
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .RowHeight = 18.75
        .UnMerge
        .Font.Name = "微软雅黑"
        .Font.Size = 10.5
        .Font.Bold = False
        .WrapText = False
        .ShrinkToFit = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .NumberFormatLocal = "@"
    End With
    Columns(1).ColumnWidth = 12
    Columns(2).ColumnWidth = 30
    Columns(3).ColumnWidth = 12
    Columns(4).ColumnWidth = 23.2
    Columns(5).ColumnWidth = 8.5
    Columns(6).ColumnWidth = 12
    
    '''封面准备
    Sheets("宗地属性表").Select
    ActiveWindow.ScrollColumn = 10
    Range("K:K,L:L,J:J,Q:Q").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("封面准备").Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "不动产单元号:"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "不动产权利人:"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "联系电话:"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "不动产座落:"
    Cells.EntireColumn.AutoFit
    ActiveWorkbook.Worksheets("封面准备").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("封面准备").Sort.SortFields.Add2 Key:=Range("D2:D" & Range("D2").End(xlDown).Row) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("封面准备").Sort
        .SetRange Range("A1:D" & Range("A1").End(xlDown).Row)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '''封面打印
    Sheets("封面打印").Select
    i = 0
    Application.ScreenUpdating = False
    For j = 2 To Sheets("封面准备").Range("A1").End(xlDown).Row
        '''分页符
        If i > 0 And i Mod 8 = 0 Then
            Sheets("封面打印").HPageBreaks.Add Before:=Sheets("封面打印").Cells(i * 3 + 1, 1)
        End If
        '''第一行
        Rows(i * 3 + 1).RowHeight = 55
        Cells(i * 3 + 1, 1) = Sheets("封面准备").Cells(1, 1)
        Cells(i * 3 + 1, 2) = Sheets("封面准备").Cells(j, 1)
        Call 表内格式("B" & i * 3 + 1, "B" & i * 3 + 1)
        Cells(i * 3 + 1, 3) = Sheets("封面准备").Cells(1, 2)
        Cells(i * 3 + 1, 4) = Sheets("封面准备").Cells(j, 2)
        Call 权利人格式("D" & i * 3 + 1, "D" & i * 3 + 1)
        Cells(i * 3 + 1, 5) = Sheets("封面准备").Cells(1, 3)
        Cells(i * 3 + 1, 6) = Sheets("封面准备").Cells(j, 3)
        Call 表内格式("F" & i * 3 + 1, "F" & i * 3 + 1)
        '''第二行
        Rows(i * 3 + 2).RowHeight = 27.5
        Cells(i * 3 + 2, 1) = Sheets("封面准备").Cells(1, 4)
        Cells(i * 3 + 2, 2) = Sheets("封面准备").Cells(j, 4)
        Range("B" & i * 3 + 2 & ":F" & i * 3 + 2).Merge
        Call 表内格式("B" & i * 3 + 2, "F" & i * 3 + 2)
        '''裁剪线
        With Range("A" & i * 3 + 3 & ":F" & i * 3 + 3).Borders(xlEdgeBottom)
            .LineStyle = xlDot
            .ThemeColor = 1
            .TintAndShade = -0.249946592608417
            .Weight = xlThin
        End With
        i = i + 1
    Next
    Application.ScreenUpdating = True
    
    '''打印设置
    With ActiveSheet.PageSetup
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = 0
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
        .CenterHorizontally = True
        .CenterVertically = True
        .Zoom = 100
    End With
    ActiveWorkbook.Save
    mainDir = ThisWorkbook.Path
    
    '''另存
    '获取村名
    cunName = Left(Sheets("封面准备").Range("D2"), InStr(Sheets("封面准备").Range("D2"), "村"))
    ActiveWorkbook.SaveAs mainDir & "\" & cunName & "档案袋封面打印.xls"
    
    '''打印为PDF
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, ActivePrinter:="Microsoft Print to PDF", _
        PrintToFile:=True, Collate:=True, PrToFileName:=mainDir & "\" & cunName & "档案袋封面打印.pdf", IgnorePrintAreas:=True

End Sub
Sub 表内格式(i, j)
'
' 表内格式 宏
'

'
    With Range(i, j)
        .Font.Name = "仿宋"
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Borders(xlEdgeBottom).TintAndShade = 0
        .Borders(xlEdgeBottom).Weight = xlHairline
    End With
End Sub
Sub 权利人格式(i, j)
Attribute 权利人格式.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 权利人格式 宏
'

'
    With Range(i, j)
        .Font.Name = "仿宋"
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Borders(xlEdgeBottom).TintAndShade = 0
        .Borders(xlEdgeBottom).Weight = xlHairline
        .WrapText = True
        .ShrinkToFit = False
    End With
End Sub
