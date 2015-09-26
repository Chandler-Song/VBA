Sub a整理()
'
' 1整理 Macro


'----------------<<<<<复制到临时表而非创建临时表--5.28>>>>>----------------'

'-------------复制筛选列到表:合同-临时数据 开始--------------------------------'
    Sheets("原数据").Select
    Cells.Select
    Selection.Copy

    Sheets("临时数据").Select
    Range("A1").Select
    ActiveSheet.Paste

    Range("A:A,B:B,C:C,D:D,G:G,H:H,J:J,K:K,O:O,P:P").Select
    Range("P1").Activate

    Range("A:A,B:B,C:C,D:D,G:G,H:H,J:J,K:K,O:O,P:P,Q:Q").Select
    Range("Q1").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft

 '-------------复制筛选列到表:合同-临时数据 结束--------------------------------'


''-------------复制筛选列到新表开始--------------------------------'
'    Sheets("原数据").Select
'    Cells.Select
'    Selection.Copy
'    Sheets.Add After:=Sheets(Sheets.Count)
'    ActiveSheet.Paste
'    Range("A:A,B:B,C:C,D:D,G:G,H:H,J:J,K:K,O:O,P:P").Select
'    Range("P1").Activate
'    ActiveWindow.ScrollColumn = 2
'    ActiveWindow.ScrollColumn = 3
'    Range("A:A,B:B,C:C,D:D,G:G,H:H,J:J,K:K,O:O,P:P,Q:Q").Select
'    Range("Q1").Activate
'    Application.CutCopyMode = False
'    Selection.Delete Shift:=xlToLeft
'    ActiveSheet.Name = "临时数据"
''-------------复制筛选列到新表结束--------------------------------'
    
'-------------初始总行数开始--------------------------------'
    Dim i As Integer
    i = 2
    Do While ActiveSheet.Cells(i, 1) <> ""
    i = i + 1
    Loop
    ActiveSheet.Cells(1, 7) = "总行数:"
    ActiveSheet.Cells(1, 8) = i - 1
'-------------初始总行数结束--------------------------------'

' 排序 Macro
'
    Dim finalcow As Integer
    finalcow = ActiveSheet.Cells(1, 8)
    'MsgBox finalcow
    tmp_Range = "G1:G" & finalcow
    
'-------------排序开始--------------------------------'
    Range("F1").Select
    Selection.AutoFilter
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields. _
    Add Key:=Range(tmp_Range), _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    DataOption:= _
        xlSortNormal
    With ActiveSheet.AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
'-------------排序结束--------------------------------'
    


'-------------需要删除的系统外代理商开始--------------------------------'
    Dim odl As Integer
    
    odl = 2
    
    If ActiveSheet.Cells(2, 7) <> "" Then
        
        Do While ActiveSheet.Cells(odl, 7) <> ""
            odl = odl + 1
        Loop
        
        odl = odl - 1
        
        tmp_odl = "2:" & odl
        Rows(tmp_odl).Select
        Selection.Delete Shift:=xlUp
        
        finalcow = finalcow - odl + 1
        
    Else
    
        odl = 0
        
    End If
    'MsgBox odl
    

    

   
    '-------修改删除后的总行数开始-----------'
    
    ActiveSheet.Cells(1, 8) = finalcow
'MsgBox finalcow
    '-------修改删除后的总行数结束-----------'
    
'-------------需要删除的系统外代理商结束--------------------------------'


'---------------修改列开始----------------------------------------------'
    ActiveSheet.Cells(1, 9) = "渠道"
    ActiveSheet.Cells(1, 10) = "Time"
    ActiveSheet.Cells(1, 11) = "SF"
    ActiveSheet.Cells(1, 12) = "Name"
    
    
'------------------------------------<<<<<优化排序--5.28>>>>>----------------------------------------------'

    '---------------插入公式开始-------------------------------'
    
    ActiveSheet.Cells(2, 9) = "=LOOKUP(E2,{0,1000,6000,9000,10000,11000,20000,60000,70000,70200,70300,70600,70900,72000,73000,74000,75000,80000,9000000},{""社内"",""IB"",""社内"",""IB"",""vcs"",""vcs"",""广东"",""社内"",""华东"",""华北"",""华东"",""华北"",""华东"",""华南"",""华北"",""华南"",""成都"",""社内""})"
    ActiveSheet.Cells(2, 10) = "=LOOKUP(B2,{0,6,12,18,24,99},{""6"",""6"",""12"",""18"",""24""})"
    ActiveSheet.Cells(2, 11) = "=LEFT(F2,1)"
    ActiveSheet.Cells(2, 12) = "=LEFT(C2,LEN(C2)-1)"
    
    Range("I2:L2").Select
    Selection.Copy

    finalcow = ActiveSheet.Cells(1, 8)
    
    fm_Range = "I2" & ":L" & finalcow

    Range(fm_Range).Select
    

    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    '---------------插入公式结束-------------------------------'
    
    
    
     '---------------根据公式结果修改开始--------------------'
        
    '--------------初级修改开始----------------'
    
    finalcow = ActiveSheet.Cells(1, 8)
    
    fm_Range_all = "I1" & ":I" & finalcow


'------排序-------'
    Range("H1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    ActiveWorkbook.Worksheets("临时数据").AutoFilter.Sort.SortFields.Clear
    
    ActiveWorkbook.Worksheets("临时数据").AutoFilter.Sort.SortFields. _
        Add Key:=Range(fm_Range_all), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption _
        :=xlSortTextAsNumbers
        
    With ActiveWorkbook.Worksheets("临时数据").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 '------排序-------'
 
 '------筛选-------'
 Range_all = "$A$1:$L$" & finalcow
 
 
 
    ActiveSheet.Range(Range_all).AutoFilter Field:=9, Criteria1:="#N/A"
    ActiveSheet.Range(Range_all).AutoFilter Field:=5, Criteria1:="=dl*", _
        Operator:=xlOr, Criteria2:="=p101*"
 '------筛选-------'
   '--------------------DL--------------------'
    ActiveSheet.UsedRange.Columns(9).SpecialCells(xlCellTypeVisible).Cells = "DL"
        
   '--------------------IB--------------------'
   '------取消筛选-------'
    ActiveSheet.Range(Range_all).AutoFilter Field:=5
    
    
    ActiveSheet.UsedRange.Columns(9).SpecialCells(xlCellTypeVisible).Cells = "IB"
    
    
   '------排序-------'
    Range("J1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    ActiveWorkbook.Worksheets("临时数据").AutoFilter.Sort.SortFields.Clear
    
    ActiveWorkbook.Worksheets("临时数据").AutoFilter.Sort.SortFields. _
        Add Key:=Range(fm_Range_all), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption _
        :=xlSortTextAsNumbers
        
    With ActiveWorkbook.Worksheets("临时数据").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 '------排序-------'

 
 '--------------初级修改结束----------------'

    
    
    
    
    
'    '---------------插入公式开始-------------------------------'
'
'    For i = 2 To finalcow Step 1
'    ActiveSheet.Cells(i, 9).Formula = _
'    "=LOOKUP(E" & i & ",{0,1000,6000,9000,10000,11000,20000,60000,70000,70200,70300,70600,70900,72000,73000,74000,75000,80000,9000000},{""社内"",""IB"",""社内"",""IB"",""vcs"",""vcs"",""广东"",""社内"",""华东"",""华北"",""华东"",""华北"",""华东"",""华南"",""华北"",""华南"",""成都"",""社内""})"
'    '---------------根据公式结果修改开始--------------------'
'    '--------初级修改开始-------'
'    'MsgBox ActiveSheet.Cells(i, 9)
'
'    ActiveSheet.Cells(i, 10).Formula = _
'    "=LOOKUP(B" & i & ",{0,6,12,18,24,99},{""6"",""6"",""12"",""18"",""24""})" '插入ordertime公式
'
'
'
'    ActiveSheet.Cells(i, 11).Formula = _
'    "=left(F" & i & ",1)" '插入starttype公式
'
'    ActiveSheet.Cells(i, 12).Formula = _
'    "=left(C" & i & ",len(C" & i & ")-1)" '插入starttype公式
'
'    If ActiveSheet.Cells(i, 9).Text = "#N/A" Then
'    'MsgBox "#N/A"
'        If Left(ActiveSheet.Cells(i, 5), 2) = "DL" Or Left(ActiveSheet.Cells(i, 5), 2) = "dl" Or Left(ActiveSheet.Cells(i, 5), 4) = "P101" Then
'            ActiveSheet.Cells(i, 9) = "DL"
'        Else
'            ActiveSheet.Cells(i, 9) = "IB"
'        End If
'    End If
'    Next
    '--------初级修改结束-------'
    
    
    '--------最终修改开始(仅主体表用)-------'
    'For i = 2 To Final Step 1
    'If ActiveSheet.Cells(i, 9) = "IB" And ActiveSheet.Cells(i, 9) Then
    '    ActiveSheet.Cells(i, 9) = "WEB"
    'End If
    'Next
    '--------最终修改结束(仅主体表用)-------'
    
    
     '---------------根据公式结果修改结束--------------------'
    '---------------修改列结束----------------------------------------------'

    
    Range("I1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    Range("L:L,D:D,I:I,J:J,K:K").Select
    Range("K1").Activate
    Selection.Copy
    Sheets("整理后数据").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    Columns("A:A").Select
    ActiveWorkbook.Names.Add Name:="ViewStart", RefersToR1C1:="=整理后数据!C1"
    Columns("B:B").Select
    ActiveWorkbook.Names.Add Name:="QD", RefersToR1C1:="=整理后数据!C2"
    Columns("C:C").Select
    ActiveWorkbook.Names.Add Name:="Time", RefersToR1C1:="=整理后数据!C3"
    Columns("D:D").Select
    ActiveWorkbook.Names.Add Name:="SF", RefersToR1C1:="=整理后数据!C4"
    Columns("E:E").Select
    ActiveWorkbook.Names.Add Name:="Name", RefersToR1C1:="=整理后数据!C5"

End Sub

Sub c清空()
'
' 清空 Macro
'
'
    Sheets("原数据").Select
    Cells.Select
    Selection.ClearContents
    
    Sheets("整理后数据").Select
    Cells.Select
    Selection.ClearContents
    
    Sheets("当日合同明细").Select
    Range("C3:G50").Select
    Selection.ClearContents
    
    Range("K3:O50").Select
    Selection.ClearContents
    Range("C55:G102").Select
    Selection.ClearContents
    
    Sheets("临时数据").Select
    Cells.Select
    Selection.ClearContents

    
    
End Sub

Sub b保存()
'
' 保存 Macro
'
'
    Sheets("目标表").Select
    Range("C3:G50").Select
    Selection.Copy
    Sheets("当日合同明细").Select
    ActiveWindow.SmallScroll Down:=-42
    Range("C3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("目标表").Select
    Sheets("目标表").Name = "目标表"
    Range("K3:O50").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("当日合同明细").Select
    Range("K3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("目标表").Select
    Range("C55:G102").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("当日合同明细").Select
    Range("C55").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
End Sub



