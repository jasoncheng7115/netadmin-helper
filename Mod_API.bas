Attribute VB_Name = "Mod_API"
'將表單設為最上層 api---------
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Const SWP_NOMOVE = &H2                          '不更動目前視窗位置
    Const SWP_NOSIZE = &H1                          '不更動目前視窗大小
    Const HWND_TOPMOST = -1                         '設定為最上層
    Const HWND_NOTOPMOST = -2                       '取消最上層設定
    Const flags = SWP_NOMOVE Or SWP_NOSIZE
'將表單設為最上層 api---------



Public Function Form_Set_Always_Top(Src_Form As Form, Src_Bool As Boolean)
''將表單設為最上層
    
    If Src_Bool = True Then
        SetWindowPos Src_Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags
    Else
        SetWindowPos Src_Form.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags
    End If

End Function


