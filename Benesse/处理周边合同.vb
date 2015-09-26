Sub a整理()
'
'' 1整理 Macro
'
'
'----------------<<<<<复制到临时表而非创建临时表--5.28>>>>>----------------'
'
'-------------复制筛选列到表:临时数据 开始--------------------------------'

    '----清空历史数据-----'
    Sheets("整理后数据").Select
    Cells.Select
    Selection.ClearContents
    
    Sheets("临时数据").Select
    Cells.Select
    Selection.ClearContents
    
    
    
    '----清空历史数据-----'
    

    Sheets("原数据").Select
    Cells.Select
    Selection.Copy

    Sheets("临时数据").Select
    Range("A1").Select
    ActiveSheet.Paste

    Range("A:A,C:C,D:D,F:F,G:G").Select
    Range("A1").Activate
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

'-------------<<<<<优化总行数--6.3>>>>>----------------------'
    
'-------------初始总行数开始--------------------------------'

   
    row = ActiveSheet.UsedRange.Rows.Count
'    Dim i As Integer
'    i = 2
'    Do While ActiveSheet.Cells(i, 1) <> ""
'    i = i + 1
'    Loop
'    ActiveSheet.Cells(1, 3) = "总行数:"
'    ActiveSheet.Cells(1, 4) = i - 1
'-------------初始总行数结束--------------------------------'


'---------------修改列开始----------------------------------------------'
    ActiveSheet.Cells(1, 3) = "渠道"
    
'------------------------------------<<<<<优化排序--5.28>>>>>----------------------------------------------'

    '---------------插入公式开始-------------------------------'
    
    ActiveSheet.Cells(2, 3) = "=LOOKUP(B2,{0,10000,15200,15299,60000,70000,9000000},{""IB"",""IB"",""VCS"",""IB"",""OB"",""SC""})"
    
    Range("C2").Select
    Selection.Copy


    
    fm_Range = "C2" & ":C" & row

    Range(fm_Range).Select
    

    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    '---------------插入公式结束-------------------------------'
    
    
    
     '---------------根据公式结果修改开始--------------------'
        
    '--------------初级修改开始----------------'
    
        
    fm_Range_all = "C1" & ":C" & row


'------排序-------'
    Range("B1").Select
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
 range_all = "$A$1:$C$" & row
 
 'Range(Range_all).Select
 
    ActiveSheet.Range(range_all).AutoFilter Field:=3, Criteria1:="#N/A"
    ActiveSheet.Range(range_all).AutoFilter Field:=2, Criteria1:="=dl*", _
        Operator:=xlOr, Criteria2:="=p101*"
 '------筛选-------'
   '--------------------DL--------------------'
    ActiveSheet.UsedRange.Columns(3).SpecialCells(xlCellTypeVisible).Cells = "DL"
        
   '--------------------IB--------------------'
   '------取消筛选-------'
    ActiveSheet.Range(range_all).AutoFilter Field:=2
    
    
    ActiveSheet.UsedRange.Columns(3).SpecialCells(xlCellTypeVisible).Cells = "IB"
    
    
   '------排序-------'
    Range("A1").Select
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
    ActiveSheet.Cells(1, 3) = "渠道"
 
 '--------------初级修改结束----------------'
 
    Columns("A:A").Select
    ActiveWorkbook.Names.Add Name:="CF_Name", RefersToR1C1:="=临时数据!C1"
    Columns("C:C").Select
    ActiveWorkbook.Names.Add Name:="CF_QD", RefersToR1C1:="=临时数据!C3"
 
 
 '-------------复制整理表 开始--------------------------------'


    Sheets("临时数据").Select
    Range("A:A,C:C").Select
    Selection.Copy
    Sheets("整理后数据").Select
    Range("A1").Select
    ActiveSheet.Paste
 '-------------复制整理表 结束--------------------------------'
 
 
 
  '-------------------------------修改+,*开始------------------------------'
 
 
    Dim col, cel_len, i, t, times, it As Integer
    Dim cel_tmp, cel_old, cel_new, cel_qd As String
    
    row = ActiveSheet.UsedRange.Rows.Count
    
    'col = ActiveSheet.UsedRange.Rows.Count

    
    
    i = 2
    Do
        If i > row Then Exit Do

            cel_len = Len(ActiveSheet.Cells(i, 1).Text)
            cel_old = ActiveSheet.Cells(i, 1)
'
            For t = 2 To cel_len Step 1
'
'
                cel_tmp = Mid(ActiveSheet.Cells(i, 1).Text, t - 1, 1)
'    '<<<<<<<<<<<<<<<<<<<<6.3更新*数字重复情况>>>>>>>>>>>>>>>>>>>>>>>>>'
                    If cel_tmp = "*" Then _

                        ActiveSheet.Cells(i, 1) = Mid(cel_old, 1, t - 2)
                        cel_new = Mid(cel_old, t + 2, cel_len)
                        
                        cel_qd = ActiveSheet.Cells(i, 2)
                        times = Mid(cel_old, t, 1)
                        cel_old = ActiveSheet.Cells(i, 1)

                        For it = 1 To times - 1 Step 1

                            Rows(i + it).Select
                            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

                            ActiveSheet.Cells(i + it, 1) = cel_old
                            ActiveSheet.Cells(i + it, 2) = cel_qd
                            row = row + 1
                        Next
                        
                        If cel_new = "" Then Exit For
                        Rows(i + times).Select
                        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                        
                        ActiveSheet.Cells(i + times, 1) = cel_new
                        ActiveSheet.Cells(i + times, 2) = cel_qd
                        row = row + 1
                        
                        Exit For
                    End If
                    
                    
                    If cel_tmp = "+" Then _

                        ActiveSheet.Cells(i, 1) = Mid(cel_old, 1, t - 2)
                
                        cel_new = Mid(cel_old, t, cel_len)
                        cel_qd = ActiveSheet.Cells(i, 2)
                
                        Rows(i + 1).Select
                        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                        ActiveSheet.Cells(i + 1, 1) = cel_new
                        ActiveSheet.Cells(i + 1, 2) = cel_qd
                        row = row + 1

                        Exit For
                    End If
            Next
            
          i = i + 1
    Loop
  '-------------------------------修改+,*结束------------------------------'
  
  
'    Sheets("临时数据").Select
'    Range("A:A,C:C").Select
'    Selection.Copy
'    Sheets("整理后数据").Select
'    Range("A1").Select
'    ActiveSheet.Paste
   
'
    Columns("A:A").Select
    ActiveWorkbook.Names.Add Name:="Name", RefersToR1C1:="=整理后数据!C1"
    Columns("B:B").Select
    ActiveWorkbook.Names.Add Name:="QD", RefersToR1C1:="=整理后数据!C2"

End Sub


Sub b清空命名空间()
'
' 清空命名空间Macro
'
'
    ActiveWorkbook.Names("Name").Delete
    ActiveWorkbook.Names("QD").Delete
    ActiveWorkbook.Names("CF_Name").Delete
    ActiveWorkbook.Names("CF_QD").Delete
    
        
    Sheets("整理后数据").Select
    Cells.Select
    Selection.ClearContents
    
    Sheets("临时数据").Select
    Cells.Select
    Selection.ClearContents
    
    Sheets("原数据").Select
    Cells.Select
    Selection.ClearContents
    
    
End Sub











