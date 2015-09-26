Sub a整理()
'
' 1整理 Macro

'-------------去重开始--------------------------------'

    ActiveWorkbook.Worksheets("原数据").Select
'-------------元数据初始总行数开始--------------------------------'
    Dim i_o As Long
    i_o = 2
    Do While ActiveSheet.Cells(i_o, 1) <> ""
    i_o = i_o + 1
    Loop
'-------------元数据总行数结束--------------------------------'

    Dim finalcow_o As Long
    finalcow_o = i_o - 1
    'MsgBox finalcow
    tmp_Range_o_1 = "T2:T" & finalcow_o
    tmp_Range_o_2 = "H2:H" & finalcow_o
    tmp_Range_o_all = "$A$1:$U$" & finalcow_o

    ActiveWorkbook.Worksheets("原数据").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("原数据").Sort.SortFields.Add Key:= _
        Range(tmp_Range_o_1) _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        
    ActiveWorkbook.Worksheets("原数据").Sort.SortFields.Add Key:= _
        Range(tmp_Range_o_2) _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("原数据").Sort
        .SetRange Range(tmp_Range_o_all)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveSheet.Range(tmp_Range_o_all).RemoveDuplicates Columns:=20, Header:= _
        xlYes

'-------------去重结束--------------------------------'


'----------------<<<<<复制到临时表而非创建临时表--5.28>>>>>----------------'

'-------------复制筛选列到表:合同-临时数据 开始--------------------------------'
    Sheets("原数据").Select
    Range("C:C,D:D,N:N,P:P,S:S").Select
    Selection.Copy

    Sheets("临时数据").Select
    Range("A1").Select
    ActiveSheet.Paste

   'Range("A:A,B:B,E:E,F:F,G:G,H:H,I:I,J:J,K:K,L:L,M:M,O:O,Q:Q,R:R,T:T,U:U").Select
'    Range("A1").Activate
'
'
'    Application.CutCopyMode = False
'    Selection.Delete Shift:=xlToLeft

    Sheets("原数据").Select

    Range("A:B,E:M,O:O,Q:R,T:U").Select
    Range("T1").Activate
    Selection.Copy
    Sheets("临时数据").Select
    Range("L1").Select
    ActiveSheet.Paste


'-------------初始总行数开始--------------------------------'
    Dim i As Long
    i = 2
    Do While ActiveSheet.Cells(i, 1) <> ""
    i = i + 1
    Loop
    ActiveSheet.Cells(1, 6) = "总行数:"
    ActiveSheet.Cells(1, 7) = i - 1
    
'-------------初始总行数结束--------------------------------'

    ActiveSheet.Cells(1, 8) = "订购类型"
    ActiveSheet.Cells(1, 9) = "到期版本"
    ActiveSheet.Cells(1, 10) = "续订类型"
    ActiveSheet.Cells(1, 11) = "渠道"
    
    
' 排序 Macro
'
    Dim finalcow As Long
    finalcow = ActiveSheet.Cells(1, 7)
    'MsgBox finalcow
    tmp_Range = "E1:E" & finalcow
    
'-------------排序开始--------------------------------'
    Range("D1").Select
    Selection.AutoFilter
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
    


'-------------需要删除的YKZ开始--------------------------------'


'-------------需要删除的YK开始--------------------------------'
    Range_all = "$A$1:$E$" & finalcow
  
    ActiveSheet.Range(Range_all).AutoFilter Field:=5, _
        Criteria1:="=k*", Operator:=xlOr, _
        Criteria2:="=y*", Operator:=xlOr
        
 '------筛选-------'
   '--------------------KY--------------------'
    ActiveSheet.UsedRange.Columns(5).SpecialCells(xlCellTypeVisible).Cells.Delete Shift:=xlUp

    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
   
    ActiveSheet.Cells(1, 1) = "订购类型"
    ActiveSheet.Cells(1, 2) = "到期版本"
    ActiveSheet.Cells(1, 3) = "订购类型1"
    ActiveSheet.Cells(1, 4) = "合同创建人"
    ActiveSheet.Cells(1, 5) = "FROMCODE"
    
   
    
    
    '-------------新总行数开始--------------------------------'
    
    i = 2
    Do While ActiveSheet.Cells(i, 1) <> ""
    i = i + 1
    Loop
    ActiveSheet.Cells(1, 6) = "总行数:"
    ActiveSheet.Cells(1, 7) = i - 1
'-------------新总行数结束--------------------------------'
'-------------需要删除的YK结束--------------------------------'

    
'-------------需要删除的z开始--------------------------------'
    
    finalcow = ActiveSheet.Cells(1, 7)
    Range_all = "$A$1:$E$" & finalcow
  
    ActiveSheet.Range(Range_all).AutoFilter Field:=5, _
        Criteria1:="=z*"
        
 '------筛选-------'
   '--------------------Z--------------------'
    ActiveSheet.UsedRange.Columns(5).SpecialCells(xlCellTypeVisible).Cells.Delete Shift:=xlUp

    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
   
    ActiveSheet.Cells(1, 1) = "订购类型"
    ActiveSheet.Cells(1, 2) = "到期版本"
    ActiveSheet.Cells(1, 3) = "订购类型1"
    ActiveSheet.Cells(1, 4) = "合同创建人"
    ActiveSheet.Cells(1, 5) = "FROMCODE"
   
    ActiveSheet.Cells(1, 12) = "合同号"
    ActiveSheet.Cells(1, 13) = "终止刊号"
    ActiveSheet.Cells(1, 14) = "到期合同金额"
    ActiveSheet.Cells(1, 15) = "到期开始类型"
    ActiveSheet.Cells(1, 16) = "OLD-*-NEW"
    ActiveSheet.Cells(1, 17) = "合同号1"
    ActiveSheet.Cells(1, 18) = "起始刊号"
    ActiveSheet.Cells(1, 19) = "合同创建日"
    ActiveSheet.Cells(1, 20) = "生日"
    ActiveSheet.Cells(1, 21) = "媒体代码"
    ActiveSheet.Cells(1, 20) = "版本"
    ActiveSheet.Cells(1, 23) = "合同创建日1"
    ActiveSheet.Cells(1, 24) = "解约刊号"
    ActiveSheet.Cells(1, 25) = "代理2"
    ActiveSheet.Cells(1, 26) = "PIN"
    ActiveSheet.Cells(1, 27) = "CONTACTID"


    
   
    
    '-------------新总行数开始--------------------------------'
    
    i = 2
    Do While ActiveSheet.Cells(i, 1) <> ""
    i = i + 1
    Loop
    ActiveSheet.Cells(1, 6) = "总行数:"
    ActiveSheet.Cells(1, 7) = i - 1
'-------------新总行数结束-----------------------
 '-------------需要删除的Z结束-------------------------------'
 
 
 '-------------需要删除的YKZ结束-------------------------------'
    
    
    '-------修改删除后的总行数开始-----------'

    finalcow = ActiveSheet.Cells(1, 7)

'---------------修改列开始----------------------------------------------'


    ActiveSheet.Cells(1, 8) = "订购类型"
    ActiveSheet.Cells(1, 9) = "到期版本"
    ActiveSheet.Cells(1, 10) = "续订类型"
    ActiveSheet.Cells(1, 11) = "渠道"
    
    
'------------------------------------<<<<<优化排序--5.28>>>>>----------------------------------------------'

    '---------------插入公式开始-------------------------------'
    
    
    ActiveSheet.Cells(2, 8) = "=LOOKUP(A2,{0,6,12,18,24,99},{""6"",""6"",""12"",""18"",""24""})"
    ActiveSheet.Cells(2, 9) = "=VLOOKUP(B2,对照表!$A:$B,2,0)"
    ActiveSheet.Cells(2, 10) = "=LOOKUP(C2,{0,6,12,18,24,99},{""6"",""6"",""12"",""18"",""24""})"
    ActiveSheet.Cells(2, 11) = "=LOOKUP(D2,{0,60000,70000},{""IB"",""OB"",""IB""})"
    
    Range("H2:K2").Select
    Selection.Copy

    finalcow = ActiveSheet.Cells(1, 7)
    
    fm_Range = "H2" & ":K" & finalcow

    Range(fm_Range).Select
    

    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    '---------------插入公式结束-------------------------------'
    
    
    
     '---------------根据公式结果修改开始--------------------'
        
    '--------------初级修改开始----------------'
    
    finalcow = ActiveSheet.Cells(1, 7)
    
    fm_Range_all = "K1" & ":K" & finalcow


'------排序-------'
    Range("K1").Select
    Selection.AutoFilter
    'Selection.AutoFilter
    
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
 Range_all = "$A$1:$AA$" & finalcow
 
 
 
    ActiveSheet.Range(Range_all).AutoFilter Field:=11, Criteria1:="#N/A"
    '------筛选-------'
   '--------------------IB--------------------'
    ActiveSheet.UsedRange.Columns(11).SpecialCells(xlCellTypeVisible).Cells = "IB"
        
    ActiveSheet.Cells(1, 11) = "渠道"
    
    '------取消筛选-------'
    ActiveSheet.Range(Range_all).AutoFilter Field:=11
    
   '------排序-------'
    Range("K1").Select
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

    
    
'    Range("H:H,I:I,J:J,K:K").Select
    Columns("H:AA").Select

    Selection.Copy
    Sheets("整理后数据").Select
    Range("A1").Select
    ActiveSheet.Paste
    
'    Columns("A:A").Select
'    ActiveWorkbook.Names.Add Name:="ViewStart", RefersToR1C1:="=整理后数据!C1"
'    Columns("B:B").Select
'    ActiveWorkbook.Names.Add Name:="QD", RefersToR1C1:="=整理后数据!C2"
'    Columns("C:C").Select
'    ActiveWorkbook.Names.Add Name:="Time", RefersToR1C1:="=整理后数据!C3"
'    Columns("D:D").Select
'    ActiveWorkbook.Names.Add Name:="SF", RefersToR1C1:="=整理后数据!C4"
'    Columns("E:E").Select
'    ActiveWorkbook.Names.Add Name:="Name", RefersToR1C1:="=整理后数据!C5"

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
    
'    Sheets("当日合同明细").Select
'    Range("C3:G50").Select
'    Selection.ClearContents
'
'    Range("K3:O50").Select
'    Selection.ClearContents
'    Range("C55:G102").Select
'    Selection.ClearContents
    
    Sheets("临时数据").Select
    Cells.Select
    Selection.ClearContents

    
    
End Sub
