VERSION 5.00
Begin VB.Form Frm_Manager_Printer_AddPrinter 
   BorderStyle     =   4  '��u�T�w�u�����
   Caption         =   "�s�W�L���"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '�t�ιw�]��
   Begin VB.TextBox Txt_Name 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   2
      Top             =   1080
      Width           =   3855
   End
   Begin VB.CommandButton Cmd_Cancel 
      Cancel          =   -1  'True
      Caption         =   "����(&C)"
      Height          =   435
      Left            =   3660
      TabIndex        =   11
      Top             =   3060
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Save 
      Caption         =   "�T�w(&S)"
      Default         =   -1  'True
      Height          =   435
      Left            =   2340
      TabIndex        =   10
      Top             =   3060
      Width           =   1215
   End
   Begin VB.TextBox Txt_Sharename 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   6
      Text            =   "�п�J�@�ΦW��"
      Top             =   2460
      Width           =   2955
   End
   Begin VB.CheckBox Chk_Shared 
      Caption         =   "�@��"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   2460
      Width           =   675
   End
   Begin VB.CheckBox Chk_Network 
      Caption         =   "�����L���"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   1980
      Value           =   1  '�֨�
      Width           =   1275
   End
   Begin VB.TextBox Txt_Location 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   1500
      Width           =   2295
   End
   Begin VB.ComboBox Cmb_PrinterTCPIPPort 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   1
      Top             =   660
      Width           =   3855
   End
   Begin VB.ComboBox Cmb_PrinterDriver 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1080
      Style           =   2  '��¤U�Ԧ�
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '�z��
      Caption         =   "�W�١G"
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '�z��
      Caption         =   "�ҡG15A �줽��"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3540
      TabIndex        =   12
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '�z��
      Caption         =   "��m�G"
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '�z��
      Caption         =   "�s����G"
      Height          =   315
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "�L����G"
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   300
      Width           =   975
   End
End
Attribute VB_Name = "Frm_Manager_Printer_AddPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Chk_Shared_Click()
Select Case Chk_Shared.Value
    Case 1
        'Lbl_Sharename.Enabled = False
        Txt_Sharename.Enabled = True
    Case 0
        'Lbl_Sharename.Enabled = True
        Txt_Sharename.Enabled = False
End Select
End Sub

Private Sub Cmb_PrinterDriver_Click()
Txt_Name.Text = Cmb_PrinterDriver.Text
End Sub

Private Sub Cmd_Cancel_Click()
Unload Me
End Sub

Private Sub Cmd_Save_Click()
        
        
        If Trim(Cmb_PrinterDriver.Text) = "" Then MsgBox "�S����ܦL����X�ʵ{��", vbQuestion, "���~": Exit Sub
        If Trim(Cmb_PrinterTCPIPPort.Text) = "" Then MsgBox "�S����ܦL����s����", vbQuestion, "���~": Exit Sub
        If Trim(Cmb_Protocol.Text) = "" Then MsgBox "�S����ܦL����q�T��w", vbQuestion, "���~": Exit Sub
        
        Set obj = objWMIService.Get("Win32_Printer").SpawnInstance_
        obj.DriverName = Cmb_PrinterDriver.Text
        obj.PortName = Cmb_PrinterTCPIPPort.Text
        obj.DeviceID = Txt_Name.Text
        obj.Location = Txt_Location.Text
        obj.Network = Chk_Network.Value
        obj.Shared = Chk_Shared.Value
        'obj.SNMPEnabled = True
        'obj.SNMPCommunity = "public"
        'obj.SNMPDevIndex = "1"
        
        'If Cmb_Protocol.Text = "RAW" Then
        '    obj.Protocol = 1
        '    obj.PortNumber = 515
        'Else
        '    obj.Protocol = 2
        '    obj.Queue = "PASSTHRU"
        'End If
        
        
        If Chk_Shared.Value = 1 Then
            obj.ShareName = Txt_Sharename.Text
        End If
        
        obj.Put_
        
        'Frm_Manager_Printer.Show_Printer
        
        Unload Me
        
End Sub

Private Sub Form_Load()

Dim cols, obj
Dim tmp_drvname() As String

'Ū�J�i���X�ʵ{��
Set cols = objWMIService.ExecQuery("Select * from Win32_PrinterDriver")
For Each obj In cols
    
    tmp_drvname = Split(obj.Name, ",")
    Cmb_PrinterDriver.AddItem tmp_drvname(0)
    DoEvents
Next

'Ū�J�i�� port
Set cols = objWMIService.ExecQuery("Select * from Win32_TCPIPPrinterPort")
For Each obj In cols
    Cmb_PrinterTCPIPPort.AddItem obj.Name
    DoEvents
Next

'�[�J�q�T��w
'Cmb_Protocol.AddItem "LPR"
'Cmb_Protocol.ItemData(Cmb_Protocol.ListCount - 1) = "2"
'Cmb_Protocol.AddItem "RAW"
'Cmb_Protocol.ItemData(Cmb_Protocol.ListCount - 1) = "1"


End Sub

Private Sub Option1_Click()

End Sub

Private Sub Txt_Name_GotFocus()
AutoSelStr Me.ActiveControl
End Sub

Private Sub Txt_Sharename_GotFocus()
AutoSelStr Me.ActiveControl
End Sub

Private Sub Txt_Location_GotFocus()
AutoSelStr Me.ActiveControl
End Sub
