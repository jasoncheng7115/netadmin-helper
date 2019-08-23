Attribute VB_Name = "Mod_vbalSGrid"

Public Function CheckGroupHeader(Src_Obj As vbalGrid)
'取出已設為群組之標題

    Dim sCol As Long
    
    For sCol = 1 To Src_Obj.Columns
        If (Src_Obj.ColumnIsGrouped(sCol)) = True Then
            Src_Obj.ColumnIsGrouped(sCol) = False
            Src_Obj.ColumnIsGrouped(sCol) = True
        End If
    Next

End Function

Public Function ExGroup(Src_Obj As vbalGrid)
'重新展開群組
    
With Src_Obj
    
    .Redraw = False
    
    Dim sRow As Long
    For sRow = 1 To .Rows
        
        If .RowIsGroup(sRow) Then
            
            .RowHeight(sRow) = .RowHeight(sRow) * 2
            .RowGroupingState(sRow) = ecgExpanded
        
        End If
    
    Next
    
    .Redraw = True
    
End With

End Function

Public Function SGrid_Serach_Row(Src_Obj As vbalGrid, Field As Integer, Serach_String As String, Optional Start_Row As Long = 1) As Long
'搜尋 ID 並傳回 Row_Index 資料

    Dim i As Long: i = 0
    
    For i = Start_Row To Src_Obj.Rows
        
        If CStr(Src_Obj.CellText(i, Field)) = CStr(Serach_String) Then
            SGrid_Serach_Row = i
            Exit For
        End If
        
        DoEvents
    Next

End Function

Public Function SGrid_Serach_Rows_Check(Src_Obj As vbalGrid, Get_Field As Integer, Field_Check As Integer) As String()
'功能 : 搜尋 SGrid 並傳回陣列
'用法 : 需再呼叫端先設一陣列並將本函數傳回陣列指定過去方可使用

'陣列元素起始大小
Dim Int_Arr As Long: Int_Arr = 0

'暫時存放陣列
Dim temp_arr() As String

With Src_Obj

    '開始列舉
    For i = 1 To .Rows
    
        If .RowIsGroup(i) = False Then
            
            '找到已勾選
            If .CellText(i, Field_Check) = "v" Then
            
            
                '擴展陣列
                ReDim Preserve temp_arr(Int_Arr)
                
                '判斷要指定取得欄位為第幾個欄位
                temp_arr(Int_Arr) = .CellText(i, Get_Field)
                
                '元素 + 1
                Int_Arr = Int_Arr + 1
                
            End If
            
        End If
        
        DoEvents
        
    Next i

End With

'將陣列回傳
SGrid_Serach_Rows_Check = temp_arr

End Function

Public Function GroupRows(Src_Form As Form) As Integer
'計算有幾個 Group 欄位

GroupRows = 0

Dim i As Long
For i = 1 To Src_Form.Sg1.Rows
    If Src_Form.Sg1.RowIsGroup(i) = True Then GroupRows = GroupRows + 1
Next i

End Function

Public Sub vbalGrid_Sort(Src_Obj As vbalGrid, lCol As Long, Optional Src_Sort_Order As Boolean)
'點選標題列排序
    
    'Show_StatusBar_Text "正在排序資料..."
    
    Dim sTag As String
    Dim iSortIndex As Long
          
        With Src_Obj.SortObject
          
          ' This demo allows grouping.  When a column is clicked
          ' for sorting, we only want to remove any grouped rows:
          .ClearNongrouped
          
          ' See if this column is already in the sort object:
          iSortIndex = .IndexOf(lCol)
          If (iSortIndex = 0) Then
             ' If not, we add it:
             iSortIndex = .Count + 1
             .SortColumn(iSortIndex) = lCol
          End If
       
          ' Determine which sort order to apply:
          sTag = Src_Obj.ColumnTag(lCol)
          
          '為 True 則代表照原本排序
          If Src_Sort_Order = True Then
          
            If (sTag = "") Then
               sTag = "ASC"
               .SortOrder(iSortIndex) = CCLOrderDescending
            Else
               sTag = "DESC"
               .SortOrder(iSortIndex) = CCLOrderAscending
            End If
                  
          '改變
          Else
        
            If (sTag = "") Then
               sTag = "DESC"
               .SortOrder(iSortIndex) = CCLOrderAscending
            Else
               sTag = ""
               .SortOrder(iSortIndex) = CCLOrderDescending
            End If
          
          End If
        
          Src_Obj.ColumnTag(lCol) = sTag
          
          ' Set the type of sorting:
          .SortType(iSortIndex) = Src_Obj.ColumnSortType(lCol)
          
        End With
        
        Src_Obj.Sort
    
    'Show_StatusBar_Text "資料排序完成"
    
End Sub


