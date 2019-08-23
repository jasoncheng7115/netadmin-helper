Attribute VB_Name = "Mod_vbalSGrid"

Public Function CheckGroupHeader(Src_Obj As vbalGrid)
'���X�w�]���s�դ����D

    Dim sCol As Long
    
    For sCol = 1 To Src_Obj.Columns
        If (Src_Obj.ColumnIsGrouped(sCol)) = True Then
            Src_Obj.ColumnIsGrouped(sCol) = False
            Src_Obj.ColumnIsGrouped(sCol) = True
        End If
    Next

End Function

Public Function ExGroup(Src_Obj As vbalGrid)
'���s�i�}�s��
    
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
'�j�M ID �öǦ^ Row_Index ���

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
'�\�� : �j�M SGrid �öǦ^�}�C
'�Ϊk : �ݦA�I�s�ݥ��]�@�}�C�ñN����ƶǦ^�}�C���w�L�h��i�ϥ�

'�}�C�����_�l�j�p
Dim Int_Arr As Long: Int_Arr = 0

'�Ȯɦs��}�C
Dim temp_arr() As String

With Src_Obj

    '�}�l�C�|
    For i = 1 To .Rows
    
        If .RowIsGroup(i) = False Then
            
            '���w�Ŀ�
            If .CellText(i, Field_Check) = "v" Then
            
            
                '�X�i�}�C
                ReDim Preserve temp_arr(Int_Arr)
                
                '�P�_�n���w���o��쬰�ĴX�����
                temp_arr(Int_Arr) = .CellText(i, Get_Field)
                
                '���� + 1
                Int_Arr = Int_Arr + 1
                
            End If
            
        End If
        
        DoEvents
        
    Next i

End With

'�N�}�C�^��
SGrid_Serach_Rows_Check = temp_arr

End Function

Public Function GroupRows(Src_Form As Form) As Integer
'�p�⦳�X�� Group ���

GroupRows = 0

Dim i As Long
For i = 1 To Src_Form.Sg1.Rows
    If Src_Form.Sg1.RowIsGroup(i) = True Then GroupRows = GroupRows + 1
Next i

End Function

Public Sub vbalGrid_Sort(Src_Obj As vbalGrid, lCol As Long, Optional Src_Sort_Order As Boolean)
'�I����D�C�Ƨ�
    
    'Show_StatusBar_Text "���b�ƧǸ��..."
    
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
          
          '�� True �h�N��ӭ쥻�Ƨ�
          If Src_Sort_Order = True Then
          
            If (sTag = "") Then
               sTag = "ASC"
               .SortOrder(iSortIndex) = CCLOrderDescending
            Else
               sTag = "DESC"
               .SortOrder(iSortIndex) = CCLOrderAscending
            End If
                  
          '����
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
    
    'Show_StatusBar_Text "��ƱƧǧ���"
    
End Sub


