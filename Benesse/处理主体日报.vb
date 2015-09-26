Sub b发货整理()
'
' 1发货-整理 Macro
   Sheets("发货原数据").Select

'----------------<<<<<排除非当月号和修改-V--4.27>>>>>----------------'

'-------------初始总行数开始--------------------------------'
    Dim i_O As Long
    i_O = 2
    Do While ActiveSheet.Cells(i_O, 1) <> ""
    i_O = i_O + 1
    Loop

    i_O = i_O - 1
'-------------初始总行数结束--------------------------------'


    tmp_Range_O = "C1:C" & i_O

'-------------排序开始--------------------------------'
    Range("F1").Select
    Selection.AutoFilter
    ActiveSheet.AutoFilter.Sort.SortFields.Clear
    ActiveSheet.AutoFilter.Sort.SortFields. _
    Add Key:=Range(tmp_Range_O), _
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



'-------------需要删除的非当月号--------------------------------'
    Dim odl_O As Long

    odl_O = 2
    'MsgBox ActiveSheet.Cells(2, 3)
    If ActiveSheet.Cells(2, 3) <> "201506" Then

        Do While ActiveSheet.Cells(odl_O, 3) <> "201506"
            odl_O = odl_O + 1
        Loop

        odl_O = odl_O - 1

        tmp_odl_O = "2:" & odl_O
        Rows(tmp_odl_O).Select
        Selection.Delete Shift:=xlUp

        i_O = i_O - odl_O + 1

    Else

        odl_O = 0

    End If
    'MsgBox odl
'-------------需要删除的非当月号--------------------------------'

'-------------需要删除的-v--------------------------------'
    Columns("L:L").Select
    Selection.Replace What:="-v", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
'-------------需要删除的-v--------------------------------'
'----------------<<<<<排除非当月号和修改-V--4.27>>>>>----------------'



'----------------<<<<<复制到临时表而非创建临时表--5.28>>>>>----------------'

'-------------复制筛选列到表:发货-临时数据 开始--------------------------------'
    Cells.Select
    Selection.Copy
    
    Sheets("发货-临时数据").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    Range("B:B,D:D,E:E,F:F,I:I,K:K,M:M,P:P,Q:Q,R:R,S:S").Select
    Range("S1").Activate
    Application.CutCopyMode = False
    Selection.Delete Shift:=xlToLeft
    
 '-------------复制筛选列到表:发货-临时数据 结束--------------------------------'
'
''-------------复制筛选列到新表开始--------------------------------'
'    Cells.Select
'    Selection.Copy
'    Sheets.Add After:=Sheets(Sheets.Count)
'    ActiveSheet.Paste
'
''    Range("A:A,G:G,H:H,J:J,L:L,N:N,O:O").Select
''    Range("O1").Activate
''    Selection.Copy
'
'    Range("B:B,D:D,E:E,F:F,I:I,K:K,M:M,P:P,Q:Q,R:R,S:S").Select
'    Range("S1").Activate
'    Application.CutCopyMode = False
'    Selection.Delete Shift:=xlToLeft
'    ActiveSheet.Name = "发货-临时数据"
''-------------复制筛选列到新表结束--------------------------------'

'-------------初始总行数开始--------------------------------'
    Dim i As Long
    i = 2
    Do While ActiveSheet.Cells(i, 1) <> ""
    i = i + 1
    Loop
    ActiveSheet.Cells(1, 9) = "总行数:"
    ActiveSheet.Cells(1, 10) = i - 1
'-------------初始总行数结束--------------------------------'

' 排序 Macro
'
    Dim finalcow As Long
    finalcow = ActiveSheet.Cells(1, 10)
    'MsgBox finalcow
    tmp_Range = "H1:H" & finalcow

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
    Dim odl As Long

    odl = 2

    If ActiveSheet.Cells(2, 8) <> "" Then '----------------<<<<<排除直接为空--4.27>>>>>----------------'


        Do While ActiveSheet.Cells(odl, 8) <> ""
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

    ActiveSheet.Cells(1, 10) = finalcow
'MsgBox finalcow
    '-------修改删除后的总行数结束-----------'

'-------------需要删除的系统外代理商结束--------------------------------'


'---------------修改列开始----------------------------------------------'
    ActiveSheet.Cells(1, 11) = "渠道"
    ActiveSheet.Cells(1, 12) = "Time"
    ActiveSheet.Cells(1, 13) = "SF"
    ActiveSheet.Cells(1, 14) = "Name"
   

    
'----------------<<<<<优化排序--5.27>>>>>----------------'

    '---------------插入公式开始-------------------------------'
    
    ActiveSheet.Cells(2, 11) = "=LOOKUP(G2,{0,1000,6000,9000,10000,11000,20000,60000,70000,70200,70300,70600,70900,72000,73000,74000,75000,80000,9000000},{""社内"",""IB"",""社内"",""IB"",""vcs"",""vcs"",""广东"",""社内"",""华东"",""华北"",""华东"",""华北"",""华东"",""华南"",""华北"",""华南"",""成都"",""社内""})"
    ActiveSheet.Cells(2, 12) = "=LOOKUP(C2,{0,6,12,18,24,99},{""6"",""6"",""12"",""18"",""24""})"
    ActiveSheet.Cells(2, 13) = "=VLOOKUP(F2,对照表!$G$2:$H$21,2,0)"
    ActiveSheet.Cells(2, 14) = "=VLOOKUP(A2,对照表!$A$2:$B$114,2,0)"
    
    Range("K2:N2").Select
    Selection.Copy

    finalcow = ActiveSheet.Cells(1, 10)
    
    fm_Range = "K2" & ":N" & finalcow

    Range(fm_Range).Select
    

    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    '---------------插入公式结束-------------------------------'
    
    
    
     '---------------根据公式结果修改开始--------------------'
        
    '--------------初级修改开始----------------'
    
    finalcow = ActiveSheet.Cells(1, 10)
    
    fm_Range_all = "K1" & ":K" & finalcow


'------排序-------'
    Range("J1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    ActiveWorkbook.Worksheets("发货-临时数据").AutoFilter.Sort.SortFields.Clear
    
    ActiveWorkbook.Worksheets("发货-临时数据").AutoFilter.Sort.SortFields. _
        Add Key:=Range(fm_Range_all), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption _
        :=xlSortTextAsNumbers
        
    With ActiveWorkbook.Worksheets("发货-临时数据").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 '------排序-------'
 
 '------筛选-------'
 Range_all = "$A$1:$N$" & finalcow
 
 
 
    ActiveSheet.Range(Range_all).AutoFilter Field:=11, Criteria1:="#N/A"
    ActiveSheet.Range(Range_all).AutoFilter Field:=7, Criteria1:="=dl*", _
        Operator:=xlOr, Criteria2:="=p101*"
 '------筛选-------'
   '--------------------DL--------------------'
    ActiveSheet.UsedRange.Columns(11).SpecialCells(xlCellTypeVisible).Cells = "DL"
        
   '--------------------IB--------------------'
   '------取消筛选-------'
    ActiveSheet.Range(Range_all).AutoFilter Field:=7
    
    
    ActiveSheet.UsedRange.Columns(11).SpecialCells(xlCellTypeVisible).Cells = "IB"
    
    
   '--------------------WEB---------------------'
   '------排序-------'
    Range("J1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    ActiveWorkbook.Worksheets("发货-临时数据").AutoFilter.Sort.SortFields.Clear
    
    ActiveWorkbook.Worksheets("发货-临时数据").AutoFilter.Sort.SortFields. _
        Add Key:=Range(fm_Range_all), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption _
        :=xlSortTextAsNumbers
        
    With ActiveWorkbook.Worksheets("发货-临时数据").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 '------排序-------'
 
 
 '--------------初级修改结束----------------'
 
 
'--------------最终修改开始(仅主体表用)----------------'
   
'------筛选-------'
    ActiveSheet.Range(Range_all).AutoFilter Field:=11, Criteria1:="IB"
    ActiveSheet.Range(Range_all).AutoFilter Field:=4, Criteria1:="=q*"
'------筛选-------'
    
    ActiveSheet.UsedRange.Columns(11).SpecialCells(xlCellTypeVisible).Cells = "WEB"

    
   '------排序-------'
    Range("J1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    ActiveWorkbook.Worksheets("发货-临时数据").AutoFilter.Sort.SortFields.Clear
    
    ActiveWorkbook.Worksheets("发货-临时数据").AutoFilter.Sort.SortFields. _
        Add Key:=Range(fm_Range_all), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption _
        :=xlSortTextAsNumbers
        
    With ActiveWorkbook.Worksheets("发货-临时数据").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 '------排序-------'
    ActiveSheet.Cells(1, 11) = "渠道"
    

 '--------------最终修改结束(仅主体表用)----------------'
    
 '----------------------根据公式结果修改结束------------------------'
     
   
     
    
'    For i = 2 To finalcow Step 1
'        ActiveSheet.Cells(i, 11).Formula = _
'        "=LOOKUP(G" & i & ",{0,1000,6000,9000,10000,11000,20000,60000,70000,70200,70300,70600,70900,72000,73000,74000,75000,80000,9000000},{""社内"",""IB"",""社内"",""IB"",""vcs"",""vcs"",""广东"",""社内"",""华东"",""华北"",""华东"",""华北"",""华东"",""华南"",""华北"",""华南"",""成都"",""社内""})"
'    '---------------根据公式结果修改开始--------------------'
'    '--------初级修改开始-------'
'    'MsgBox ActiveSheet.Cells(i, 9)
'
'        ActiveSheet.Cells(i, 12).Formula = _
'        "=LOOKUP(C" & i & ",{0,6,12,18,24,99},{""6"",""6"",""12"",""18"",""24""})" '插入ordertime公式
'
'
''        If Right(ActiveSheet.Cells(i, 6), 2) = "-V" Then
''            ActiveSheet.Cells(i, 13).Formula = _
''            "=VLOOKUP(LEFT(F" & i & ",(LEN(F" & i & ")-2)),对照表!$G$2:$H$21,2,0)"
''        Else
''            ActiveSheet.Cells(i, 13).Formula = _
''            "=VLOOKUP(F" & i & ",对照表!$G$2:$H$21,2,0)" '插入sendflag公式
''        End If
'
'        ActiveSheet.Cells(i, 13).Formula = _
'        "=VLOOKUP(F" & i & ",对照表!$G$2:$H$21,2,0)" '插入sendflag公式------手动替换-V
'
'
'
'
'        ActiveSheet.Cells(i, 14).Formula = _
'        "=VLOOKUP(A" & i & ",对照表!$A$2:$B$114,2,0)" '插入name公式
'
'        If ActiveSheet.Cells(i, 11).Text = "#N/A" Then
'    'MsgBox "#N/A"
'            If Left(ActiveSheet.Cells(i, 7), 2) = "DL" Or _
'                Left(ActiveSheet.Cells(i, 7), 2) = "dl" Or _
'                Left(ActiveSheet.Cells(i, 7), 4) = "P101" Then
'                ActiveSheet.Cells(i, 11) = "DL"
'            Else
'                ActiveSheet.Cells(i, 11) = "IB"
'            End If
'        End If
'    Next
'    '--------初级修改结束-------'
'
'
'    '--------最终修改开始(仅主体表用)-------'
'    For i = 2 To finalcow Step 1
'
'        If ActiveSheet.Cells(i, 11).Text = "IB" And _
'            Left(ActiveSheet.Cells(i, 4), 1) = "Q" Or _
'            Left(ActiveSheet.Cells(i, 4), 1) = "q" Then
'            ActiveSheet.Cells(i, 11) = "WEB"
'
'        End If
'    Next
'    '--------最终修改结束(仅主体表用)-------'


     '---------------根据公式结果修改结束--------------------'
    '---------------修改列结束----------------------------------------------'

    
    Range("I1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    Range("E:E,K:K,L:L,M:M,N:N,B:B").Select
    Range("K1").Activate
    Selection.Copy
    Sheets("发货整理后数据").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    
    

End Sub


Sub a合同整理()
'
' 1合同-整理 Macro
'
'----------------<<<<<复制到临时表而非创建临时表--5.28>>>>>----------------'

'-------------复制筛选列到表:合同-临时数据 开始--------------------------------'
    Sheets("合同原数据").Select
    Cells.Select
    Selection.Copy

    Sheets("合同-临时数据").Select
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
'    Sheets("合同原数据").Select
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
'    ActiveSheet.Name = "合同-临时数据"
''-------------复制筛选列到新表结束--------------------------------'
    
'-------------初始总行数开始--------------------------------'
    Dim i As Long
    i = 2
    Do While ActiveSheet.Cells(i, 1) <> ""
    i = i + 1
    Loop
    ActiveSheet.Cells(1, 7) = "总行数:"
    ActiveSheet.Cells(1, 8) = i - 1
'-------------初始总行数结束--------------------------------'

' 排序 Macro
'
    Dim finalcow As Long
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
    Dim odl As Long
    
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
    '---------------插入公式开始-------------------------------'
    
    
    
    
    
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
    
    ActiveWorkbook.Worksheets("合同-临时数据").AutoFilter.Sort.SortFields.Clear
    
    ActiveWorkbook.Worksheets("合同-临时数据").AutoFilter.Sort.SortFields. _
        Add Key:=Range(fm_Range_all), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption _
        :=xlSortTextAsNumbers
        
    With ActiveWorkbook.Worksheets("合同-临时数据").AutoFilter.Sort
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
    
    ActiveWorkbook.Worksheets("合同-临时数据").AutoFilter.Sort.SortFields.Clear
    
    ActiveWorkbook.Worksheets("合同-临时数据").AutoFilter.Sort.SortFields. _
        Add Key:=Range(fm_Range_all), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption _
        :=xlSortTextAsNumbers
        
    With ActiveWorkbook.Worksheets("合同-临时数据").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 '------排序-------'

 
 '--------------初级修改结束----------------'
 
 
'--------------最终修改开始(仅主体表用)----------------'
'--------------------WEB---------------------'

'------筛选-------'
    ActiveSheet.Range(Range_all).AutoFilter Field:=9, Criteria1:="IB"
    ActiveSheet.Range(Range_all).AutoFilter Field:=1, Criteria1:="=q*"
'------筛选-------'

    ActiveSheet.UsedRange.Columns(9).SpecialCells(xlCellTypeVisible).Cells = "WEB"


   '------排序-------'
    Range("H1").Select
    Selection.AutoFilter
    Selection.AutoFilter

    ActiveWorkbook.Worksheets("合同-临时数据").AutoFilter.Sort.SortFields.Clear

    ActiveWorkbook.Worksheets("合同-临时数据").AutoFilter.Sort.SortFields. _
        Add Key:=Range(fm_Range_all), _
        SortOn:=xlSortOnValues, _
        Order:=xlAscending, _
        DataOption _
        :=xlSortTextAsNumbers

    With ActiveWorkbook.Worksheets("合同-临时数据").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
 '------排序-------'
    ActiveSheet.Cells(1, 9) = "渠道"
    
    
    
    
    
    
    

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
'    '--------初级修改结束-------'
'
'
'    '--------最终修改开始(仅主体表用)-------'
'    For i = 2 To finalcow Step 1
'        If ActiveSheet.Cells(i, 9).Text = "IB" And _
'            Left(ActiveSheet.Cells(i, 1), 1) = "Q" Or _
'            Left(ActiveSheet.Cells(i, 1), 1) = "q" Then
'            ActiveSheet.Cells(i, 9) = "WEB"
'        End If
'    Next
'    '--------最终修改结束(仅主体表用)-------'
'
'
'     '---------------根据公式结果修改结束--------------------'
'    '---------------修改列结束----------------------------------------------'

    
    Range("I1").Select
    Selection.AutoFilter
    Selection.AutoFilter
    
    Range("L:L,D:D,I:I,J:J,K:K").Select
    Range("K1").Activate
    Selection.Copy
    Sheets("合同整理后数据").Select
    Range("A1").Select
    ActiveSheet.Paste
    


End Sub

Sub e清空命名空间()
'
' 清空命名空间Macro
'
'
    ActiveWorkbook.Names("Day_Name").Delete
    ActiveWorkbook.Names("Day_QD").Delete
    ActiveWorkbook.Names("Day_SF").Delete
    ActiveWorkbook.Names("Day_Time").Delete
    ActiveWorkbook.Names("Day_ViewStart").Delete
    ActiveWorkbook.Names("Month_Name").Delete
    ActiveWorkbook.Names("Month_ProID").Delete
    ActiveWorkbook.Names("Month_QD").Delete
    ActiveWorkbook.Names("Month_SF").Delete
    ActiveWorkbook.Names("Month_Time").Delete
    ActiveWorkbook.Names("Month_ViewStart").Delete
    
    
End Sub


Sub c添加命名空间()
'
' 添加命名空间Macro
'
'
    Sheets("发货原数据").Select
    Columns("A:A").Select
    ActiveWorkbook.Names.Add Name:="Month_ProID", RefersToR1C1:="=发货整理后数据!C1"
    Columns("B:B").Select
    ActiveWorkbook.Names.Add Name:="Month_ViewStart", RefersToR1C1:="=发货整理后数据!C2"
    Columns("C:C").Select
    ActiveWorkbook.Names.Add Name:="Month_QD", RefersToR1C1:="=发货整理后数据!C3"
    Columns("D:D").Select
    ActiveWorkbook.Names.Add Name:="Month_Time", RefersToR1C1:="=发货整理后数据!C4"
    Columns("E:E").Select
    ActiveWorkbook.Names.Add Name:="Month_SF", RefersToR1C1:="=发货整理后数据!C5"
    Columns("F:F").Select
    ActiveWorkbook.Names.Add Name:="Month_Name", RefersToR1C1:="=发货整理后数据!C6"
    
    Sheets("合同整理后数据").Select
    Columns("A:A").Select
    ActiveWorkbook.Names.Add Name:="Day_ViewStart", RefersToR1C1:="=合同整理后数据!C1"
    Columns("B:B").Select
    ActiveWorkbook.Names.Add Name:="Day_QD", RefersToR1C1:="=合同整理后数据!C2"
    Columns("C:C").Select
    ActiveWorkbook.Names.Add Name:="Day_Time", RefersToR1C1:="=合同整理后数据!C3"
    Columns("D:D").Select
    ActiveWorkbook.Names.Add Name:="Day_SF", RefersToR1C1:="=合同整理后数据!C4"
    Columns("E:E").Select
    ActiveWorkbook.Names.Add Name:="Day_Name", RefersToR1C1:="=合同整理后数据!C5"
    
    
End Sub




Sub f清空数据()
'
' 清空数据 Macro
'
'


    Sheets("发货原数据").Select
    Cells.Select
    Selection.ClearContents
    
    Sheets("合同原数据").Select
    Cells.Select
    Selection.ClearContents



    Sheets("发货整理后数据").Select
    Cells.Select
    Selection.ClearContents
    Sheets("合同整理后数据").Select
    Cells.Select
    Selection.ClearContents
   
    Sheets("发货-临时数据").Select
    Cells.Select
    Selection.ClearContents

    
    Sheets("合同-临时数据").Select
    Cells.Select
    Selection.ClearContents
 
    
End Sub

Sub d保存()
'
' 保存 Macro
'
'
'

    '-------填充当日数据----------'
    Sheets("目标表").Select
    Range("C6:F41").Select
    Selection.Copy
    Sheets("主体日报").Select
    Range("C6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("目标表").Select
    Range("L6:P41").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("主体日报").Select
    Range("L6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("目标表").Select
    Range("U6:AD41").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("主体日报").Select
    Range("U6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    '-------填充当日数据----------'



    '-------填充每月数据----------'
    '--------------不含团购    新规----------------'
'--------name-t---------'
    Sheets("目标表").Select
    Range("C48:F76").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("主体日报").Select
    Range("C48").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
        :=False, Transpose:=False
    Sheets("目标表").Select
    Range("C78:F83").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("主体日报").Select
    Range("C78").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
        :=False, Transpose:=False
'--------name-t---------'

'--------name-qd---------'
    Sheets("目标表").Select
    Range("U48:AD76").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("主体日报").Select
    Range("U48").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
        :=False, Transpose:=False
    Sheets("目标表").Select
    Range("U78:AD83").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("主体日报").Select
    Range("U78").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
        :=False, Transpose:=False
'--------name-qd---------'

'--------t-qd---------'

    Sheets("目标表").Select
    Range("U88:AD91").Select
    Selection.Copy
    Sheets("主体日报").Select
    Range("U88").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
        :=False, Transpose:=False



''--------t-qd---------'
'--------不含团购    新规---------'

   
'--------不含团购    全部---------'
'--------name-t---------'
    Sheets("目标表").Select
    Range("C89:F117").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("主体日报").Select
    Range("C89").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
        :=False, Transpose:=False
    Sheets("目标表").Select
    Range("C119:F124").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("主体日报").Select
    Range("C119").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
        :=False, Transpose:=False
'--------name-t---------'
    
    
'--------name-sf---------'
    Sheets("目标表").Select
    Range("C133:M161").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("主体日报").Select
    Range("C133").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
        :=False, Transpose:=False
    Sheets("目标表").Select
    Range("C163:M168").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("主体日报").Select
    Range("C163").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
        :=False, Transpose:=False
'--------name-sf---------'
'--------不含团购    全部---------'
    '-------填充每月数据----------'
    
    
    
End Sub


