Attribute VB_Name = "Mod_Interface"
Public Declare Sub InitCommonControls Lib "COMCTL32" ()

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public m_hMod As Long

Public Sub Show_Msg(Src_String As String)
'��D�������A�C��ܤ�r

    If Src_String <> "" Then Frm_Main.vbalSBar1.PanelText(1) = Src_String
    
    Frm_Main.vbalSBar1.PanelText(2) = "�s�u�� : " & Total_Connected_Computers & "    "
 
End Sub

Public Sub Show_Batch_Log(Src_String As String, Src_Title As String)
'��ܧ妸�B�z���G
    
    Frm_Batch_Processed.Txt_Log.Text = Src_String
    Frm_Batch_Processed.Lbl_Title.Caption = Src_Title
    Frm_Batch_Processed.Show 1
    
End Sub


