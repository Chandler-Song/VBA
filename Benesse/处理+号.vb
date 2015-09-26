Sub m_test()



     ActiveSheet.Select

'    MsgBox ActiveSheet.UsedRange.Rows.Count
'
'    MsgBox ActiveSheet.UsedRange.Columns.Count


    Dim row, col, cel_len, i, t, times, it As Integer
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
                        
                        cel_qd = ActiveSheet.Cells(i, 3)
                        times = Mid(cel_old, t, 1)
                        cel_old = ActiveSheet.Cells(i, 1)

                        For it = 1 To times - 1 Step 1

                            Rows(i + it).Select
                            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

                            ActiveSheet.Cells(i + it, 1) = cel_old
                            ActiveSheet.Cells(i + it, 3) = cel_qd
                            row = row + 1
                        Next
                        
                        If cel_new = "" Then Exit For
                        Rows(i + times).Select
                        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                        
                        ActiveSheet.Cells(i + times, 1) = cel_new
                        ActiveSheet.Cells(i + times, 3) = cel_qd
                        row = row + 1
                        
                        Exit For
                    End If
                    
                    
                    If cel_tmp = "+" Then _

                        ActiveSheet.Cells(i, 1) = Mid(cel_old, 1, t - 2)
                
                        cel_new = Mid(cel_old, t, cel_len)
                        cel_qd = ActiveSheet.Cells(i, 3)
                
                        Rows(i + 1).Select
                        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                
                        ActiveSheet.Cells(i + 1, 1) = cel_new
                        ActiveSheet.Cells(i + 1, 3) = cel_qd
                        row = row + 1

                        Exit For
                    End If
            Next
            
          i = i + 1
    Loop
    
    Sheets("临时数据").Select
    Range("A:A,C:C").Select
    Selection.Copy
    Sheets("整理后数据").Select
    Range("A1").Select
    ActiveSheet.Paste
    
    
End Sub
