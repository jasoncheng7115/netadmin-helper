VERSION 5.00
Begin VB.Form Frm_Show_Setup 
   Caption         =   "��ܶ��س]�w"
   ClientHeight    =   5580
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2655
   Icon            =   "Frm_Show_Setup.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5580
   ScaleWidth      =   2655
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.CommandButton Cmd_Save 
      Caption         =   "�x�s�]�w(&S)"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1995
   End
   Begin VB.ListBox Lst_Show_Setup 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   0
      Style           =   1  '���إ]�t�֨����
      TabIndex        =   0
      Top             =   360
      Width           =   1995
   End
End
Attribute VB_Name = "Frm_Show_Setup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cmd_Save_Click()


Call Save_Set
'Unload Me


End Sub

Private Sub Form_Load()

Form_Set_Always_Top Me, True

With Lst_Show_Setup

    '.AddItem "�@��"
    '.AddItem "�@�~�t��"
    .AddItem "�ѽX��"
    .AddItem "��ܾ�"
    .AddItem "��ܥd"
    .AddItem "��ܼҦ�"
    .AddItem "�����ϥΪ�"
    .AddItem "�����s��"
    .AddItem "�s����"
    .AddItem "����ʸ˸m"
    .AddItem "���Щʸ˸m"
    .AddItem "�H���Y�θ˸m"
    .AddItem "���ĸ˸m"
    .AddItem "�t��"
    .AddItem "�����ܼ�"
    .AddItem "�}������"
    .AddItem "�׺ݾ�"
    .AddItem "�n��w��"
    .AddItem "��s��"
    .AddItem "�޿�Ϻ�"
    .AddItem "�A��"
    .AddItem "�����"
    .AddItem "��������"
    .AddItem "�����˸m�]�w"
    .AddItem "�Ϻо�"
    .AddItem "�@��"
    .AddItem "�L���"
    .AddItem "�L��������s����"
    .AddItem "�L����X�ʵ{��"
    .AddItem "�ϺФ��ΰ�"

End With

Get_Set

End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    'Cmd_Save.Top = 0
    'Cmd_Save.Left = 0
    'Cmd_Save.Width = Me.ScaleWidth
    
    Lst_Show_Setup.TOp = 0
    Lst_Show_Setup.Left = 0
    Lst_Show_Setup.Width = Me.ScaleWidth
    Lst_Show_Setup.Height = Me.ScaleHeight

End Sub


Sub Get_Set()
'Ū���]�w�n����


On Error Resume Next

   Dim Show_Sets As Variant
   Dim i As Long
   Dim i_list As Long
   
   Show_Sets = GetAllSettings(App.Title, "Show_Setup")
   
   For i = 0 To UBound(Show_Sets, 1)
        For i_list = 0 To Lst_Show_Setup.ListCount - 1
            If Lst_Show_Setup.List(i_list) = Show_Sets(i, 0) Then
                 Lst_Show_Setup.Selected(i_list) = Show_Sets(i, 1)
                 Exit For
            End If
            DoEvents
        Next
        DoEvents
   Next

End Sub

Sub Save_Set()
'�g�J�]�w��n����
On Error Resume Next

Cmd_Save.Enabled = False
Lst_Show_Setup.Enabled = False

'���M��

    DeleteSetting App.Title, "Show_Setup"

'�A�g�J
    Dim i As Long
    For i = 0 To Lst_Show_Setup.ListCount - 1

        If Lst_Show_Setup.Selected(i) = True Then
            SaveSetting App.Title, "Show_Setup", Lst_Show_Setup.List(i), "1"
        Else
            SaveSetting App.Title, "Show_Setup", Lst_Show_Setup.List(i), "0"
        End If
    
    Next
    
Lst_Show_Setup.Enabled = True
Cmd_Save.Enabled = True
   
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form_Set_Always_Top Me, False
End Sub

Private Sub Lst_Show_Setup_Click()
Save_Set
End Sub
