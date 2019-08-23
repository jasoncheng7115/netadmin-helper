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
'�r���榡�]�w

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
'�]�w��C���B�۰ʽվ�j�p
   
With Src_Obj
   
    .Range("A1:IV65536").Select
    .Selection.RowHeight = Row_Height
    .Selection.Columns.AutoFit
    .Range("A1").Select

End With

End Sub

Private Sub Excel_Style_Set(Src_Obj As Object)
'�r���榡�]�w

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

    '�ץX�� Excel

    Dim actForm As Form
    Set actForm = Screen.ActiveForm

    With Src_SG

        If .Rows = 0 Then MsgBox "�|���έp", vbQuestion, "���~": Exit Function
       
        Dim sRow, sCol As Long
        Dim ExToXls As Object '�sĶ�ɻݧ令�H���覡�ʺA�I�s�A�è����ޥζ��ؤ��� Microsoft Excel 10.0
        'Dim ExToXls As Excel.Application
    
        Show_Msg "���b��l�� Excel ����..."
    
            '�����񪫥�M��� ExToXls �]�����骫��
            Set ExToXls = Nothing
            Set ExToXls = CreateObject("Excel.Application")
        
        '�W�[ WorkBooks
        Show_Msg "�[�J�@�ӷs����ï..."
            ExToXls.Workbooks.Add
    
        Show_Msg "���b�]�w Excel �r���˦�..."
            Excel_Style_Font_Set ExToXls, 9
    
        '�g�J���D
        Show_Msg "���b�g�J���D..."
            For sCol = StartCol To .Columns
                If .ColumnIsGrouped(sCol) <> True Then
                    ExToXls.Cells.Item(1, sCol - StartCol + 1) = .ColumnHeader(sCol)
                End If
            Next
    
        Show_Msg "���b�g�J���... "
            actForm.Enabled = False
    
            '�a�V�g�J���
    
            For sRow = 1 To .Rows
            
                '�Y���s�զC
    
                If .RowIsGroup(sRow) = True Then
                
                    '���Ӹs�զC�[�W [ ] ���Х�
    
                    For sCol = StartCol To .Columns
    
                        If .CellText(sRow, sCol) <> "" Then
    
                            ExToXls.Cells.Item(sRow + 1, .RowGroupingLevel(sRow)) = "[" & Trim(.CellText(sRow, sCol)) & "]"
                            Exit For
    
                        End If
    
                    Next
            
                Else
                
                    '��V
    
                    For sCol = StartCol To .Columns
                        
                        If .ColumnIsGrouped(sCol) <> True Then
                            ExToXls.Cells.Item(sRow + 1, sCol - StartCol + 1) = Trim(.CellText(sRow, sCol))
                        End If
                        
                    Next
            
                End If
            
                '��ܶi�ת�
                DoEvents
                Call Show_Msg("���b�g�J���... " & FormatPercent(sRow / .Rows))
            
            Next
        
            Show_Msg "��Ƽg�J����"
        
            Show_Msg "���b�վ���˦�..."
                Excel_Style_Grid_Set ExToXls, 12
                'ExToXls.Range("A1").Select
        
        Show_Msg "���˦��վ㧹��"
    
        '���
        ExToXls.Visible = True
    
        '���񪫥�
        Set ExToXls = Nothing

        actForm.Enabled = True

    End With
    
    Show_Msg "Excel �ץX����"
    
End Function

'CSEH: ErrReport
Public Function ExportToCSV(Src_SG As vbalGrid, Optional StartCol As Long = 1)
'�ץX�� CSV
    

    '���o�n�ץX���ɦW
    Dim sFile As String
    Dim cc As New cCommonDialog
    If cc.VBGetSaveFileName(sFile, , , "CSV �r�I���j�� (*.CSV)|*.CSV", , , "��ܭn�ץX���ɮ�", "CSV", Screen.ActiveForm.hwnd, OFN_HIDEREADONLY) Then
    End If

    If Trim(sFile) = "" Then Exit Function

    Dim actForm As Form
    Set actForm = Screen.ActiveForm

    With Src_SG

        If .Rows = 0 Then MsgBox "�|���έp", vbQuestion, "���~": Exit Function

        
        Show_Msg "���b�إ��ɮ�..."
        Dim sRow, sCol As Long
        Dim fs As Object, a As Object
        Set fs = CreateObject("Scripting.FileSystemObject")
        Set a = fs.CreateTextFile(sFile, True)
    
        Dim tmp_title As String
    
        '�g�J���D
        Show_Msg "���b�g�J���D..."
        For sCol = StartCol To .Columns
            tmp_title = tmp_title & .ColumnHeader(sCol) & ","
        Next
        a.WriteLine tmp_title
            
        
        Dim tmp_Data
        Show_Msg "���b�g�J���... "
        actForm.Enabled = False

        '�a�V�g�J���
        For sRow = 1 To .Rows
            
            If .RowIsGroup(sRow) <> True Then
                
                '��V
                For sCol = StartCol To .Columns
                    tmp_Data = tmp_Data & Trim(.CellText(sRow, sCol)) & ","
                Next
                
                a.WriteLine tmp_Data
                tmp_Data = ""
                
                '��ܶi�ת�
                DoEvents
                Call Show_Msg("���b�g�J���... " & FormatPercent(sRow / .Rows))
            
            End If
            
        Next
    
        Call Show_Msg("��Ƽg�J����")
    
        '���񪫥�
        a.Close

    End With
    
    actForm.Enabled = True
    
    Call Show_Msg("Excel �ץX����")
    MsgBox "�ɮפw�ץX�� " & sFile & "", vbInformation, "�@�~����"
    
End Function


