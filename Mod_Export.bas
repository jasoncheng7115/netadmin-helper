Attribute VB_Name = "Mod_Export"
'Declarations
'===============================
Private Const xlAutomatic = -4105
Private Const xlUnderlineStyleNone = -4142
Private Const xlPrintNoComments = -4142
Private Const xlLandscape = 2
Private Const xlPaperA4 = 9
Private Const xlDownThenOver = 1
'===================================

Const xlNone = -4142
Const xlOn = 1

'Enum XlBorderWeight:
Const xlHairline = 1
Const xlThin = 2
Const xlMedium = -4138
Const xlThick = 4

'Enum XlBordersIndex:
Const xlDiagonalDown = 5
Const xlDiagonalUp = 6
Const xlEdgeLeft = 7
Const xlEdgeTop = 8
Const xlEdgeBottom = 9
Const xlEdgeRight = 10
Const xlInsideVertical = 11
Const xlInsideHorizontal = 12

'Enum XlLineStyle:
Const xlContinuous = 1
Const xlDash = -4115
Const xlDashDot = 4
Const xlDashDotDot = 5
Const xlDot = -4118
Const xlDouble = -4119
Const xlLineStyleNone = -4142
Const xlSlantDashDot = 13

'Enum XlColorIndex:
Const xlColorIndexAutomatic = -4105
Const xlColorIndexNone = -4142
  

Private Sub Excel_Style_Font_Set(Src_Obj As Object, FontSize As Integer)
'字型格式設定

With Src_Obj

    
    .Range("A1:IV65536").Select
    With .Selection.Font
        .Name = "Tahoma"
        .Size = FontSize
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With
    
    .Selection.NumberFormatLocal = "@"
    
End With

End Sub


Private Sub Excel_Style_Grid_Set(Src_Obj As Object, Row_Height As Integer)
'設定格列高、自動調整大小
   
With Src_Obj
   
    .Range("A1:IV65536").Select
    .Selection.RowHeight = Row_Height
    .Selection.Columns.AutoFit
    .Range("A1").Select

End With

End Sub

Private Sub Excel_Style_Set(Src_Obj As Object)
'字型格式設定

With Src_Obj

    '.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    '.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    'With .Selection.Borders(xlEdgeLeft)
    '    .LineStyle = xlContinuous
    '    .Weight = xlThin
    '    .ColorIndex = xlAutomatic
    'End With
    'With .Selection.Borders(xlEdgeTop)
    '    .LineStyle = xlContinuous
    '    .Weight = xlThin
    '    .ColorIndex = xlAutomatic
    'End With
    'With .Selection.Borders(xlEdgeBottom)
    '    .LineStyle = xlContinuous
    '    .Weight = xlThin
    '    .ColorIndex = xlAutomatic
    'End With
    'With .Selection.Borders(xlEdgeRight)
    '    .LineStyle = xlContinuous
    '    .Weight = xlThin
    '    .ColorIndex = xlAutomatic
    'End With
    'With .Selection.Borders(xlInsideVertical)
    '    .LineStyle = xlContinuous
    '    .Weight = xlThin
    '    .ColorIndex = xlAutomatic
    'End With
    'With .Selection.Borders(xlInsideHorizontal)
    '    .LineStyle = xlContinuous
    '    .Weight = xlThin
    '    .ColorIndex = xlAutomatic
    'End With
 
End With

End Sub

'CSEH: ErrReport
Public Function ExportToExcel(Src_SG As vbalGrid, Optional StartCol As Long = 1)

    '匯出到 Excel

    Dim actForm As Form
    Set actForm = Screen.ActiveForm

    With Src_SG

        If .Rows = 0 Then MsgBox "尚未統計", vbQuestion, "錯誤": Exit Function
       
        Dim sRow, sCol As Long
        Dim ExToXls As Object '編譯時需改成以此方式動態呼叫，並取消引用項目中的 Microsoft Excel 10.0
        'Dim ExToXls As Excel.Application
    
        Show_Msg "正在初始化 Excel 物件..."
    
            '先釋放物件然後把 ExToXls 設為實體物件
            Set ExToXls = Nothing
            Set ExToXls = CreateObject("Excel.Application")
        
        '增加 WorkBooks
        Show_Msg "加入一個新活頁簿..."
            ExToXls.Workbooks.Add
    
        Show_Msg "正在設定 Excel 字型樣式..."
            Excel_Style_Font_Set ExToXls, 9
    
        '寫入標題
        Show_Msg "正在寫入標題..."
            For sCol = StartCol To .Columns
                If .ColumnIsGrouped(sCol) <> True Then
                    ExToXls.Cells.Item(1, sCol - StartCol + 1) = .ColumnHeader(sCol)
                End If
            Next
    
        Show_Msg "正在寫入資料... "
            actForm.Enabled = False
    
            '縱向寫入資料
    
            For sRow = 1 To .Rows
            
                '若為群組列
    
                If .RowIsGroup(sRow) = True Then
                
                    '為該群組列加上 [ ] 號標示
    
                    For sCol = StartCol To .Columns
    
                        If .CellText(sRow, sCol) <> "" Then
    
                            ExToXls.Cells.Item(sRow + 1, .RowGroupingLevel(sRow)) = "[" & Trim(.CellText(sRow, sCol)) & "]"
                            Exit For
    
                        End If
    
                    Next
            
                Else
                
                    '橫向
    
                    For sCol = StartCol To .Columns
                        
                        If .ColumnIsGrouped(sCol) <> True Then
                            ExToXls.Cells.Item(sRow + 1, sCol - StartCol + 1) = Trim(.CellText(sRow, sCol))
                        End If
                        
                    Next
            
                End If
            
                '顯示進度表
                DoEvents
                Call Show_Msg("正在寫入資料... " & FormatPercent(sRow / .Rows))
            
            Next
        
            Show_Msg "資料寫入完成"
        
            Show_Msg "正在調整表格樣式..."
                Excel_Style_Grid_Set ExToXls, 12
                'ExToXls.Range("A1").Select
        
        Show_Msg "表格樣式調整完成"
    
        '顯示
        ExToXls.Visible = True
    
        '釋放物件
        Set ExToXls = Nothing

        actForm.Enabled = True

    End With
    
    Show_Msg "Excel 匯出完畢"
    
End Function

'CSEH: ErrReport
Public Function ExportToCSV(Src_SG As vbalGrid, Optional StartCol As Long = 1)
'匯出到 CSV
    

    '取得要匯出的檔名
    Dim sFile As String
    Dim cc As New cCommonDialog
    If cc.VBGetSaveFileName(sFile, , , "CSV 逗點分隔檔 (*.CSV)|*.CSV", , , "選擇要匯出的檔案", "CSV", Screen.ActiveForm.hwnd, OFN_HIDEREADONLY) Then
    End If

    If Trim(sFile) = "" Then Exit Function

    Dim actForm As Form
    Set actForm = Screen.ActiveForm

    With Src_SG

        If .Rows = 0 Then MsgBox "尚未統計", vbQuestion, "錯誤": Exit Function

        
        Show_Msg "正在建立檔案..."
        Dim sRow, sCol As Long
        Dim fs As Object, a As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(sFile, True)
    
        Dim tmp_title As String
    
        '寫入標題
        Show_Msg "正在寫入標題..."
        For sCol = StartCol To .Columns
            tmp_title = tmp_title & .ColumnHeader(sCol) & ","
        Next
        a.WriteLine tmp_title
            
        
        Dim tmp_Data
        Show_Msg "正在寫入資料... "
        actForm.Enabled = False

        '縱向寫入資料
        For sRow = 1 To .Rows
            
            If .RowIsGroup(sRow) <> True Then
                
                '橫向
                For sCol = StartCol To .Columns
                    tmp_Data = tmp_Data & Trim(.CellText(sRow, sCol)) & ","
                Next
                
                a.WriteLine tmp_Data
                tmp_Data = ""
                
                '顯示進度表
                DoEvents
                Call Show_Msg("正在寫入資料... " & FormatPercent(sRow / .Rows))
            
            End If
            
        Next
    
        Call Show_Msg("資料寫入完成")
    
        '釋放物件
        a.Close

    End With
    
    actForm.Enabled = True
    
    Call Show_Msg("Excel 匯出完畢")
    MsgBox "檔案已匯出至 " & sFile & "", vbInformation, "作業完成"
    
End Function


