Attribute VB_Name = "Mod_API"
'�N���]���̤W�h api---------
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Const SWP_NOMOVE = &H2                          '����ʥثe������m
    Const SWP_NOSIZE = &H1                          '����ʥثe�����j�p
    Const HWND_TOPMOST = -1                         '�]�w���̤W�h
    Const HWND_NOTOPMOST = -2                       '�����̤W�h�]�w
    Const flags = SWP_NOMOVE Or SWP_NOSIZE
'�N���]���̤W�h api---------



Public Function Form_Set_Always_Top(Src_Form As Form, Src_Bool As Boolean)
''�N���]���̤W�h
    
    If Src_Bool = True Then
        SetWindowPos Src_Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, flags
    Else
        SetWindowPos Src_Form.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags
    End If

End Function


