Attribute VB_Name = "��������������"
Sub A�����������ӡ()
Attribute A�����������ӡ.VB_ProcData.VB_Invoke_Func = " \n14"
'
' �����������ӡ ��
'

'
    '''�������
    Sheets("����׼��").Cells.ClearContents
    Sheets("�����ӡ").ResetAllPageBreaks
    With Sheets("�����ӡ").Cells '
        .ClearContents
        .Borders(xlInsideHorizontal).LineStyle = xlNone
        .RowHeight = 18.75
        .UnMerge
        .Font.Name = "΢���ź�"
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
    
    '''����׼��
    Sheets("�ڵ����Ա�").Select
    ActiveWindow.ScrollColumn = 10
    Range("K:K,L:L,J:J,Q:Q").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("����׼��").Select
    Range("A1").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "��������Ԫ��:"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "������Ȩ����:"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "��ϵ�绰:"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "����������:"
    Cells.EntireColumn.AutoFit
    ActiveWorkbook.Worksheets("����׼��").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("����׼��").Sort.SortFields.Add2 Key:=Range("D2:D" & Range("D2").End(xlDown).Row) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("����׼��").Sort
        .SetRange Range("A1:D" & Range("A1").End(xlDown).Row)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    '''�����ӡ
    Sheets("�����ӡ").Select
    i = 0
    Application.ScreenUpdating = False
    For j = 2 To Sheets("����׼��").Range("A1").End(xlDown).Row
        '''��ҳ��
        If i > 0 And i Mod 8 = 0 Then
            Sheets("�����ӡ").HPageBreaks.Add Before:=Sheets("�����ӡ").Cells(i * 3 + 1, 1)
        End If
        '''��һ��
        Rows(i * 3 + 1).RowHeight = 55
        Cells(i * 3 + 1, 1) = Sheets("����׼��").Cells(1, 1)
        Cells(i * 3 + 1, 2) = Sheets("����׼��").Cells(j, 1)
        Call ���ڸ�ʽ("B" & i * 3 + 1, "B" & i * 3 + 1)
        Cells(i * 3 + 1, 3) = Sheets("����׼��").Cells(1, 2)
        Cells(i * 3 + 1, 4) = Sheets("����׼��").Cells(j, 2)
        Call Ȩ���˸�ʽ("D" & i * 3 + 1, "D" & i * 3 + 1)
        Cells(i * 3 + 1, 5) = Sheets("����׼��").Cells(1, 3)
        Cells(i * 3 + 1, 6) = Sheets("����׼��").Cells(j, 3)
        Call ���ڸ�ʽ("F" & i * 3 + 1, "F" & i * 3 + 1)
        '''�ڶ���
        Rows(i * 3 + 2).RowHeight = 27.5
        Cells(i * 3 + 2, 1) = Sheets("����׼��").Cells(1, 4)
        Cells(i * 3 + 2, 2) = Sheets("����׼��").Cells(j, 4)
        Range("B" & i * 3 + 2 & ":F" & i * 3 + 2).Merge
        Call ���ڸ�ʽ("B" & i * 3 + 2, "F" & i * 3 + 2)
        '''�ü���
        With Range("A" & i * 3 + 3 & ":F" & i * 3 + 3).Borders(xlEdgeBottom)
            .LineStyle = xlDot
            .ThemeColor = 1
            .TintAndShade = -0.249946592608417
            .Weight = xlThin
        End With
        i = i + 1
    Next
    Application.ScreenUpdating = True
    
    '''��ӡ����
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
    
    '''���
    '��ȡ����
    cunName = Left(Sheets("����׼��").Range("D2"), InStr(Sheets("����׼��").Range("D2"), "��"))
    ActiveWorkbook.SaveAs mainDir & "\" & cunName & "�����������ӡ.xls"
    
    '''��ӡΪPDF
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, ActivePrinter:="Microsoft Print to PDF", _
        PrintToFile:=True, Collate:=True, PrToFileName:=mainDir & "\" & cunName & "�����������ӡ.pdf", IgnorePrintAreas:=True

End Sub
Sub ���ڸ�ʽ(i, j)
'
' ���ڸ�ʽ ��
'

'
    With Range(i, j)
        .Font.Name = "����"
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Borders(xlEdgeBottom).TintAndShade = 0
        .Borders(xlEdgeBottom).Weight = xlHairline
    End With
End Sub
Sub Ȩ���˸�ʽ(i, j)
Attribute Ȩ���˸�ʽ.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Ȩ���˸�ʽ ��
'

'
    With Range(i, j)
        .Font.Name = "����"
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).ColorIndex = xlAutomatic
        .Borders(xlEdgeBottom).TintAndShade = 0
        .Borders(xlEdgeBottom).Weight = xlHairline
        .WrapText = True
        .ShrinkToFit = False
    End With
End Sub
