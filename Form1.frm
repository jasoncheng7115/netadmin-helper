VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{2210EC79-A724-4033-AAF4-790E2467C0E8}#1.0#0"; "vbalCmdBar6.ocx"
Object = "{4F11FEBA-BBC2-4FB6-A3D3-AA5B5BA087F4}#1.0#0"; "vbalSbar6.ocx"
Begin VB.Form Frm_Main 
   Caption         =   "AD ����q���޲z���� ���ժ�"
   ClientHeight    =   8505
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9600
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   9600
   StartUpPosition =   3  '�t�ιw�]��
   Begin vbalCmdBar6.vbalCommandBar cmdBar 
      Align           =   1  '������W��
      Height          =   375
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MainMenu        =   -1  'True
   End
   Begin VB.TextBox Txt_Port 
      Alignment       =   2  '�m�����
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   1920
      TabIndex        =   5
      Text            =   "445"
      Top             =   3600
      Width           =   555
   End
   Begin VB.TextBox Txt_DomainName 
      Alignment       =   2  '�m�����
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Text            =   "test.com.tw"
      Top             =   3600
      Width           =   1755
   End
   Begin vbalSbar6.vbalStatusBar vbalSBar1 
      Align           =   2  '������U��
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   8130
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483633
      SimpleStyle     =   0
   End
   Begin vbalIml6.vbalImageList vbalImlMenu 
      Left            =   3360
      Top             =   4860
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   19516
      Images          =   "Form1.frx":57E2
      Version         =   131072
      KeyCount        =   17
      Keys            =   $"Form1.frx":A43E
   End
   Begin vbalIml6.vbalImageList vbalIml2 
      Left            =   2760
      Top             =   4860
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      ColourDepth     =   24
      Size            =   17648
      Images          =   "Form1.frx":A4C9
      Version         =   131072
      KeyCount        =   4
      Keys            =   "SHOW_SETUP?Check_Connect_Countinue?Check_Connect_Stop?Check_Connect"
   End
   Begin vbalIml6.vbalImageList vbalIml1 
      Left            =   2160
      Top             =   4860
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   4592
      Images          =   "Form1.frx":E9D9
      Version         =   131072
      KeyCount        =   4
      Keys            =   "?Disconnect?root?Connected"
   End
   Begin vbalCmdBar6.vbalCommandBar cmdBar 
      Align           =   1  '������W��
      Height          =   375
      Index           =   0
      Left            =   0
      Negotiate       =   -1  'True
      Top             =   375
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock Sck_Check_One 
      Left            =   300
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock Sck_List 
      Index           =   0
      Left            =   780
      Top             =   5040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin vbalTreeViewLib6.vbalTreeView Tvw1 
      Height          =   2235
      Left            =   0
      TabIndex        =   0
      Top             =   1140
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   3942
      NoCustomDraw    =   0   'False
      Indentation     =   12
      ItemHeight      =   18
      LineStyle       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbAcceleratorSGrid6.vbalGrid Sg1 
      Height          =   2235
      Left            =   3120
      TabIndex        =   1
      Top             =   1140
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   3942
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      GridLineMode    =   1
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
      DefaultRowHeight=   19
      AllowGrouping   =   -1  'True
      HideGroupingBox =   -1  'True
      GroupBoxHintText=   "�즲�s�������D�ܦ��A�즲��Ы��U���D�C�Ƨǥi�P�ɮi�}�Ҧ��s��"
      HotTrack        =   -1  'True
      SelectionOutline=   -1  'True
   End
   Begin vbalCmdBar6.vbalCommandBar cmdBar 
      Align           =   1  '������W��
      Height          =   225
      Index           =   2
      Left            =   0
      Negotiate       =   -1  'True
      Top             =   750
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   397
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin vbalIml6.vbalImageList vbalIml_Tools 
      Left            =   4260
      Top             =   4860
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   24
      Size            =   12628
      Images          =   "Form1.frx":FBE9
      Version         =   131072
      KeyCount        =   11
      Keys            =   $"Form1.frx":12D5D
   End
   Begin VB.Label Label1 
      Appearance      =   0  '����
      ForeColor       =   &H80000008&
      Height          =   2715
      Left            =   2940
      TabIndex        =   2
      Top             =   1020
      Width           =   45
   End
End
Attribute VB_Name = "Frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'�P�_�O�_���U�ƹ����� (�ǳƽվ�j�p)
Private mbResizing As Boolean

'����q���M��
Public tmp_Computers As Collection

'����C��X��
Public Stop_Computers_List As Boolean

'�s�u���դ��_�ɩҨ��o���q���s��
Public List_Last As Long

Private Sub cmdBar_ButtonClick(index As Integer, btn As vbalCmdBar6.cButton)
           
On Error Resume Next
           
    Select Case btn.Key
        '�s�u����
        Case "CONNECT_CHECK"
        
            With cmdBar(0).Toolbar.Buttons
                .Item("CONNECT_CHECK").Enabled = False
                .Item("CONNECT_CHECK_COUNTINUE").Enabled = False
                .Item("CONNECT_CHECK_STOP").Enabled = True
            End With
            
            '��l�ƦU�ܼ�
            Total_Connected_Computers = 0
            List_Last = 1
            Stop_Computers_List = False
            SGrid_Computers_List
            
        '�s�u�~�����
        Case "CONNECT_CHECK_COUNTINUE"
            
            With cmdBar(0).Toolbar.Buttons
                .Item("CONNECT_CHECK").Enabled = False
                .Item("CONNECT_CHECK_COUNTINUE").Enabled = False
                .Item("CONNECT_CHECK_STOP").Enabled = True
            End With
            
            '�q�W�����_�B�}�l
            List_Last = List_Last + 1
            Stop_Computers_List = False
            SGrid_Computers_List
                
        '�s�u���հ���
        Case "CONNECT_CHECK_STOP"
            
            With cmdBar(0).Toolbar.Buttons
                .Item("CONNECT_CHECK").Enabled = True
                .Item("CONNECT_CHECK_COUNTINUE").Enabled = True
                .Item("CONNECT_CHECK_STOP").Enabled = False
            End With
                       
            Stop_Computers_List = True
            Stop_Computer_List
        
        '�����Ҧ�
        Case "CONNECT_LOCAL": Get_Computers_From_Local
            
            
        
        '��ܳ]�w
        Case "SHOW_SETUP": Frm_Show_Setup.Show
        
        '�ܧ����
        Case "SETUP_DOMAINNAME_CHANGE"
        
            AD_Domain_Name = Txt_DomainName.Text
            Get_Computers_From_AD
        
        '�ܧ�s�u�T�{�� PORT
        Case "CONNECT_CHECK_PORT_SET_CHANGE"
            
            If IsNumeric(Txt_Port.Text) = False Then MsgBox "�𸹥��ݬO�Ʀr", vbInformation, "���~": Exit Sub
            If Txt_Port.Text < 1 Or Txt_Port.Text > 65536 Then MsgBox "�𸹥��ݤ��� 1~65536 ����", vbInformation, "���~": Exit Sub
            Port_Ping = Txt_Port.Text
            
        '�ץX�� Excel
        Case "FILE:EXPORT_TO_EXCEL": Call ExportToExcel(Me.Sg1)
        
        '�ץX�� CSV
        Case "FILE:EXPORT_TO_CSV": Call ExportToCSV(Me.Sg1)
        
        '���}�{��
        Case "FILE:EXIT":  Unload Me
        
        '�ԲӸ�T
        Case "MANAGER:INFORMATION": Show_Computer_Information Tvw1.SelectedItem
        
        '�޲z�A��
        Case "MANAGER:SERVICE":  Load Frm_Manager_Service
        
        '�޲z�����
        Case "MANAGER:PROCESS": Load Frm_Manager_Process
        
        '�޲z�L���
        Case "MANAGER:PRINTER": Load Frm_Manager_Printer
        
        '�޲z�n��
        Case "MANAGER:SOFTWARE": Load Frm_Manager_Software
        
        '���s�}��
        Case "MANAGER:SHUTDOWN:REBOOT": Action_Shutdown "", "ACTION:REBOOT"
        
        '����
        Case "MANAGER:SHUTDOWN:SHUTDOWN": Action_Shutdown "", "ACTION:SHUTDOWN"
                
        '�ܧ󥻾��޲z���K�X
        Case "MANAGER:LOCAL:CHANGE_ADMIN_PWD": Action_Local_Admin "", "ACTION:CHANGE_ADMIN_PASSWORD"
        
        '�妸�ܧ󥻾��޲z���K�X
        Case "BATCH:LOCAL:CHANGE_ADMIN_PWD": Batch_Action_Local_Admin "", "ACTION:CHANGE_ADMIN_PASSWORD", Get_Checked_Computers
        
        '�妸�޲z�}�l�A��
        Case "BATCH:SERVICE:START": Batch_Action_Service "ACTION:START", Get_Checked_Computers
        
        '�妸�޲z����A��
        Case "BATCH:SERVICE:STOP": Batch_Action_Service "ACTION:STOP", Get_Checked_Computers
        
        '�妸�}�Ұ����
        Case "BATCH:PROCESS:CREATE": Batch_Action_Process "ACTION:CREATE", Get_Checked_Computers
         
        '�妸���_�����
        Case "BATCH:PROCESS:TERMINATE": Batch_Action_Process "ACTION:TERMINATE", Get_Checked_Computers
        
        '����o�ӵ{��
        Case "ABOUT:PROGRAM": Frm_About.Show
        
        '����@��
        Case "ABOUT:MAKER": MsgBox "Jason Cheng", vbInformation, "Design By"
        
    
    End Select
        
End Sub

Private Function Get_Checked_Computers() As String
'���o�Ŀ諸�q���W��
    
    Dim tmp_checked As String
        
    Dim i As Long
    For i = 1 To Tvw1.NodeCount
                
        If Tvw1.Nodes(i).Key <> "root" Then
            If Tvw1.Nodes(i).Checked = True Then
                tmp_checked = tmp_checked & Tvw1.Nodes(i).Text
                If i < Tvw1.NodeCount Then tmp_checked = tmp_checked & ","
            End If
        End If
        
    Next
    
    Get_Checked_Computers = tmp_checked
    
End Function

Private Sub ConfigureButtons()
'�إ߳]�w�u��C�W�����s

'On Error Resume Next

    Dim btn As cButton
    
    cmdBar(0).ToolbarImageList = vbalIml2.hIml
    With cmdBar(0).Buttons
        
            '�إ� TOOLS ���s�s
            Set btn = .Add("TOOLS:CONNECT:SPLIT", , , eSeparator)
            
            Set btn = .Add("CONNECT_CHECK", vbalIml2.ItemIndex("Check_Connect") - 1, "�s�u�T�{", , "�T�{�Ҧ��q���O�_���s�u��")
            btn.ShowCaptionInToolbar = True
            btn.Enabled = True
            
            Set btn = .Add("CONNECT_CHECK_COUNTINUE", vbalIml2.ItemIndex("Check_Connect_Countinue") - 1, "�~��T�{", , "�q�W�����_�B���U�T�{�Ҧ��q���O�_���s�u��")
            btn.ShowCaptionInToolbar = True
            btn.Enabled = False
            
            Set btn = .Add("CONNECT_CHECK_STOP", vbalIml2.ItemIndex("Check_Connect_Stop") - 1, "�����˴�", , "���_�s�u���դu�@")
            btn.ShowCaptionInToolbar = True
            btn.Enabled = False
        
        
            '�إ� TOOLS ���s�s
            Set btn = .Add("TOOLS:SHOW_SETUP:SPLIT", , , eSeparator)
            Set btn = .Add("SHOW_SETUP", vbalIml2.ItemIndex("SHOW_SETUP") - 1, "��ܶ��س]�w", , "��ܹq����T�C��ɩҭn�[�ݪ���T")
            btn.ShowCaptionInToolbar = True
            btn.Enabled = True
        
    End With
    
    '����]�w
    'cmdBar(2).MenuImageList = vbalImlMenu.hIml
    With cmdBar(2).Buttons
            
            Set btn = .Add("TOOLS:SETUP_DOMAINNAME:SPLIT", , , eSeparator)
            
            Set btn = .Add("SETUP_DOMAINNAME_LABEL", , "����]�w�G")
            btn.ShowCaptionInToolbar = True: btn.Enabled = False
            
            Set btn = .Add("SETUP_DOMAINNAME", , "����", ePanel, "�]�w�z�n�޲z�� AD ����")
            btn.PanelWidth = 90: btn.PanelControl = Txt_DomainName
        
            Set btn = .Add("SETUP_DOMAINNAME_CHANGE", , "�C�X�q��", , "�C�X�o�Ӻ��쪺�Ҧ��q���M��")
            btn.ShowCaptionInToolbar = True
        
            '�����Ҧ�
            Set btn = .Add("TOOLS:SETUP_DOMAINNAME:SPLIT2", , , eSeparator)
            Set btn = .Add("SETUP_DOMAINNAME_LOCAL", , "�����Ҧ�", , "�����s�������q��")
            btn.ShowCaptionInToolbar = True
            btn.Enabled = True
        
            Set btn = .Add("TOOLS:CONNECT_CHECK_PORT_SET:SPLIT", , , eSeparator)
            
            Set btn = .Add("CONNECT_CHECK_PORT_SET_LABEL", , "�T�{�s�u�Ҩϥΰ𸹡G")
            btn.ShowCaptionInToolbar = True: btn.Enabled = False
            
            Set btn = .Add("CONNECT_CHECK_PORT_SET_TEXTBOX", , , ePanel, "�]�w�s�u�T�{�ɩҥΪ� PORT")
            btn.PanelWidth = 45: btn.PanelControl = Txt_Port
            
            Set btn = .Add("CONNECT_CHECK_PORT_SET_CHANGE", , "�ܧ�")
            btn.ShowCaptionInToolbar = True
            
 
    End With
    
    
    cmdBar(1).MenuImageList = vbalImlMenu.hIml
    With cmdBar(1).Buttons
               
        '�̤W�h
        Set btn = .Add("FILE", , "�ɮ�(&F)")
        btn.ShowCaptionInToolbar = True
        Set btn = .Add("VIEW", , "�˵�(&V)")
        btn.ShowCaptionInToolbar = True
        Set btn = .Add("OPTION", , "�ﶵ(&O)")
        btn.ShowCaptionInToolbar = True
        Set btn = .Add("MANAGER", , "�޲z(&M)")
        btn.ShowCaptionInToolbar = True
        Set btn = .Add("BATCH", , "�妸(&B)")
        btn.ShowCaptionInToolbar = True
        Set btn = .Add("SETUP", , "�]�w(&S)")
        btn.Enabled = False
        btn.ShowCaptionInToolbar = True
        Set btn = .Add("ABOUT", , "����(&A)")
        btn.ShowCaptionInToolbar = True
        
            '�إ� FILE ���s�s
            Set btn = .Add("FILE:EXPORT_TO_EXCEL", vbalImlMenu.ItemIndex("TOEXCEL") - 1, "�ץX��ƨ� Excel", , "�N Excel �}�ҨñN��ƶץX")
            Set btn = .Add("FILE:EXPORT_TO_CSV", vbalImlMenu.ItemIndex("TONOTEPAD") - 1, "�ץX��ƨ� CSV", , "�N��ƶץX�� CSV")
            Set btn = .Add("FILE:SPLIT1", , , eSeparator)
            btn.Visible = False
            Set btn = .Add("FILE:EXIT", , "���}", , "�����{��", vbKeyF12, 0)
            btn.Visible = False
        
            '�إ� MANAGER ���s�s
            Set btn = .Add("MANAGER:INFORMATION", vbalImlMenu.ItemIndex("INFORMATION") - 1, "�ԲӸ�T")
            Set btn = .Add("MANAGER:SPLIT1", , , eSeparator)
            Set btn = .Add("MANAGER:SERVICE", vbalImlMenu.ItemIndex("SERVICE") - 1, "�A��")
            Set btn = .Add("MANAGER:PROCESS", vbalImlMenu.ItemIndex("PROCESS") - 1, "�����")
            Set btn = .Add("MANAGER:PRINTER", vbalImlMenu.ItemIndex("PRINTER") - 1, "�L���")
            Set btn = .Add("MANAGER:SOFTWARE", vbalImlMenu.ItemIndex("SOFTWARE") - 1, "�n��")
            Set btn = .Add("MANAGER:SPLIT2", , , eSeparator)
            Set btn = .Add("MANAGER:SHUTDOWN", vbalImlMenu.ItemIndex("SHUTDOWN") - 1, "����")
                Set btn = .Add("MANAGER:SHUTDOWN:REBOOT", vbalImlMenu.ItemIndex("REBOOT") - 1, "���s�}��")
                Set btn = .Add("MANAGER:SHUTDOWN:SHUTDOWN", vbalImlMenu.ItemIndex("SHUTDOWN") - 1, "����")
            Set btn = .Add("MANAGER:LOCAL", vbalImlMenu.ItemIndex("LOCAL") - 1, "����")
                Set btn = .Add("MANAGER:LOCAL:CHANGE_ADMIN_PWD", vbalImlMenu.ItemIndex("CHANGE_PWD") - 1, "�ܧ󥻾� Administrator �K�X")
            
            '�إ� BATCH ���s�s
            Set btn = .Add("BATCH:SERVICE", vbalImlMenu.ItemIndex("SERVICE") - 1, "�A��")
                Set btn = .Add("BATCH:SERVICE:START", vbalImlMenu.ItemIndex("START") - 1, "�ҥ�")
                Set btn = .Add("BATCH:SERVICE:STOP", vbalImlMenu.ItemIndex("STOP") - 1, "����")
                
            Set btn = .Add("BATCH:PROCESS", vbalImlMenu.ItemIndex("PROCESS") - 1, "�����")
                Set btn = .Add("BATCH:PROCESS:CREATE", vbalImlMenu.ItemIndex("CREATE") - 1, "�إ�")
                Set btn = .Add("BATCH:PROCESS:TERMINATE", vbalImlMenu.ItemIndex("TERMINATE") - 1, "���_")
                
            Set btn = .Add("BATCH:SPLIT1", , , eSeparator)
            Set btn = .Add("BATCH:LOCAL", vbalImlMenu.ItemIndex("LOCAL") - 1, "����")
                Set btn = .Add("BATCH:LOCAL:CHANGE_ADMIN_PWD", vbalImlMenu.ItemIndex("CHANGE_PWD") - 1, "�ܧ󥻾� Administrator �K�X")
        
        
            '�إ� ABOUT ���s�s
            Set btn = .Add("ABOUT:PROGRAM", , "����", , "����o�ӵ{��...")
            Set btn = .Add("ABOUT:SPLIT1", , , eSeparator)
            Set btn = .Add("ABOUT:MAKER", vbalImlMenu.ItemIndex("INFORMATION") - 1, "�@��", , "�@�̸�T")
        
    End With
    
 
End Sub

Private Sub ConfigureBars()

'On Error Resume Next

    Dim bar, bar_1 As cCommandBar
    Dim Btns, Btns_1 As cCommandBarButtons


    '�u��C�s����
    With cmdBar(0)
            
        '�إߤ@�ӷs�u��C
        Set bar = .CommandBars.Add("STANDARD", "Standard Buttons")
        Set Btns = bar.Buttons
             
            '�إ�
            Btns.Add .Buttons.Item("TOOLS:CONNECT:SPLIT")
            Btns.Add .Buttons.Item("CONNECT_CHECK")
            Btns.Add .Buttons.Item("CONNECT_CHECK_COUNTINUE")
            Btns.Add .Buttons.Item("CONNECT_CHECK_STOP")
            Btns.Add .Buttons.Item("TOOLS:SHOW_SETUP:SPLIT")
            Btns.Add .Buttons.Item("SHOW_SETUP")
            
    End With

    '�ĤG�����]�w
    With cmdBar(2)
        
        Set bar = .CommandBars.Add("DOMAINNAME_SETUP", "Standard Button")
        Set Btns = bar.Buttons
            
            Btns.Add .Buttons.Item("TOOLS:SETUP_DOMAINNAME:SPLIT")
            Btns.Add .Buttons.Item("SETUP_DOMAINNAME_LABEL")
            Btns.Add .Buttons.Item("SETUP_DOMAINNAME")
            Btns.Add .Buttons.Item("SETUP_DOMAINNAME_CHANGE")
            Btns.Add .Buttons.Item("TOOLS:SETUP_DOMAINNAME:SPLIT2")
            Btns.Add .Buttons.Item("SETUP_DOMAINNAME_LOCAL")
            
            Btns.Add .Buttons.Item("TOOLS:CONNECT_CHECK_PORT_SET:SPLIT")
            Btns.Add .Buttons.Item("CONNECT_CHECK_PORT_SET_LABEL")
            Btns.Add .Buttons.Item("CONNECT_CHECK_PORT_SET_TEXTBOX")
            Btns.Add .Buttons.Item("CONNECT_CHECK_PORT_SET_CHANGE")
            
    End With
    
    '�\���s����
    With cmdBar(1)
    
        '�إ߳��h�\���
        Set bar = .CommandBars.Add("TOPMENU", "Menu")
        Set Btns = bar.Buttons
        Btns.Add .Buttons.Item("FILE")
        Btns.Add .Buttons.Item("VIEW")
        Btns.Add .Buttons.Item("OPTION")
        Btns.Add .Buttons.Item("MANAGER")
        Btns.Add .Buttons.Item("BATCH")
        Btns.Add .Buttons.Item("SETUP")
        Btns.Add .Buttons.Item("ABOUT")
        
            '�إߤ@�Ӥl�\��� FILE
            Set bar = .CommandBars.Add("FILEMENU", "FILE")
            Set Btns = bar.Buttons
            Btns.Add .Buttons.Item("FILE:EXPORT_TO_EXCEL")
            Btns.Add .Buttons.Item("FILE:EXPORT_TO_CSV")
            Btns.Add .Buttons.Item("FILE:SPLIT1")
            Btns.Add .Buttons.Item("FILE:EXIT")
            .Buttons.Item("FILE").bar = bar
      
            '�إߤ@�Ӥl�\��� MANAGER
            Set bar = .CommandBars.Add("MANAGERMENU", "MANAGER")
                  Set Btns = bar.Buttons
                  Btns.Add .Buttons.Item("MANAGER:INFORMATION")
                  Btns.Add .Buttons.Item("MANAGER:SPLIT1")
                  Btns.Add .Buttons.Item("MANAGER:SERVICE")
                  Btns.Add .Buttons.Item("MANAGER:PROCESS")
                  Btns.Add .Buttons.Item("MANAGER:PRINTER")
                  Btns.Add .Buttons.Item("MANAGER:SOFTWARE")
                  Btns.Add .Buttons.Item("MANAGER:SPLIT2")
                  Btns.Add .Buttons.Item("MANAGER:SHUTDOWN")
                  Btns.Add .Buttons.Item("MANAGER:LOCAL")
                  .Buttons.Item("MANAGER").bar = bar
                  
                      '�إߤ@�Ӥl2�\��� MANAGER:SHUTDOWN
                      Set bar = .CommandBars.Add("SHUTDOWN:MANAGERMENU", "SHUTDOWN")
                      Set Btns = bar.Buttons
                      Btns.Add .Buttons.Item("MANAGER:SHUTDOWN:REBOOT")
                      Btns.Add .Buttons.Item("MANAGER:SHUTDOWN:SHUTDOWN")
                      .Buttons.Item("MANAGER:SHUTDOWN").bar = bar
            
                      '�إߤ@�Ӥl2�\��� MANAGER:SHUTDOWN
                      Set bar = .CommandBars.Add("SHUTDOWN:LOCAL", "LOCAL")
                      Set Btns = bar.Buttons
                      Btns.Add .Buttons.Item("MANAGER:LOCAL:CHANGE_ADMIN_PWD")
                      .Buttons.Item("MANAGER:LOCAL").bar = bar
      
            
            '�إߤ@�Ӥl�\��� BATCH
            Set bar = .CommandBars.Add("BATCHMENU", "BACTH")
                  
                  '�إߤl��� SERVICE
                  Set Btns = bar.Buttons
                  Btns.Add .Buttons.Item("BATCH:SERVICE")
                  .Buttons.Item("BATCH").bar = bar
                  
                  '�إߤl��� PROCESS
                  Set Btns = bar.Buttons
                  Btns.Add .Buttons.Item("BATCH:PROCESS")
                  .Buttons.Item("BATCH").bar = bar
                  
                  Btns.Add .Buttons.Item("BATCH:SPLIT1")
                  
                  '�إߤ@�Ӥl��� LOCAL
                  Set Btns = bar.Buttons
                  Btns.Add .Buttons.Item("BATCH:LOCAL")
                  .Buttons.Item("BATCH").bar = bar
                  
                          '�إߤ@�Ӥl2�\��� BATCH:SERVICE
                          Set bar = .CommandBars.Add("BATCH:SERVICE", "SERVICE")
                          Set Btns = bar.Buttons
                          Btns.Add .Buttons.Item("BATCH:SERVICE:START")
                          Btns.Add .Buttons.Item("BATCH:SERVICE:STOP")
                          .Buttons.Item("BATCH:SERVICE").bar = bar
                
                          '�إߤ@�Ӥl2�\��� BATCH:PROCESS
                          Set bar = .CommandBars.Add("BATCH:PROCESS", "PROCESS")
                          Set Btns = bar.Buttons
                          Btns.Add .Buttons.Item("BATCH:PROCESS:CREATE")
                          Btns.Add .Buttons.Item("BATCH:PROCESS:TERMINATE")
                          .Buttons.Item("BATCH:PROCESS").bar = bar
                
                          '�إߤ@�Ӥl2�\��� BATCH:SHUTDOWN
                          Set bar = .CommandBars.Add("BATCH:LOCAL", "LOCAL")
                          Set Btns = bar.Buttons
                          Btns.Add .Buttons.Item("BATCH:LOCAL:CHANGE_ADMIN_PWD")
                          .Buttons.Item("BATCH:LOCAL").bar = bar
            
      
            '�إߤ@�Ӥl�\��� ABOUT
            Set bar = .CommandBars.Add("ABOUTMENU", "ABOUT")
                Set Btns = bar.Buttons
                Btns.Add .Buttons.Item("ABOUT:PROGRAM")
                Btns.Add .Buttons.Item("ABOUT:SPLIT1")
                Btns.Add .Buttons.Item("ABOUT:MAKER")
                .Buttons.Item("ABOUT").bar = bar
      
    End With
End Sub

Private Sub cmdBar_RequestNewInstance(index As Integer, ctl As Object)
   
   '��ܳo�ӥ\����l�ﶵ
   Dim lNewIndex As Long
   lNewIndex = cmdBar.UBound + 1
   Load cmdBar(lNewIndex)
   cmdBar(lNewIndex).Align = 0
   Set ctl = cmdBar(lNewIndex)

End Sub

Private Sub Form_Initialize()
m_hMod = LoadLibrary("shell32.dll")
Call InitCommonControls
End Sub

Private Sub Form_Load()

Me.Caption = Me.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision & "  By Jason "

'�w�]����
AD_Domain_Name = "test.com.tw"

'�w�] PORT
Port_Ping = "445"

'�إߥ\���B�u��C�P���s�s
ConfigureButtons
ConfigureBars

    '�D�\���
    cmdBar(1).MainMenu = True
    cmdBar(1).Toolbar = cmdBar(1).CommandBars("TOPMENU")
    
    '�D�n�u��C
    cmdBar(0).Toolbar = cmdBar(0).CommandBars("STANDARD")
    
    '����]�w�u��C
    cmdBar(2).Toolbar = cmdBar(2).CommandBars("DOMAINNAME_SETUP")


'�W�[���A�C����
vbalSBar1.AddPanel , , , , , True
vbalSBar1.AddPanel , "�s�u�� : 0  ", , , 80, , True

'�]�w���ʤ��ε����ɴ�м˦�
Label1.MousePointer = vbSizeWE

'�C�X�q��
Get_Computers_From_AD

End Sub

Private Sub Get_Computers_From_AD()
'�q AD �����Ҧ��q���M��

    Show_Msg "���b�q AD ���o�q���M��.."
    
    '���o�Ҧ��q��
    Set tmp_Computers = ADSI_Computer_List()


    If canUseAD = True Then
        Add_Computer_To_TreeView
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK").Enabled = True
        Show_Msg "�q AD ���o�q���M�槹��"
    Else
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK").Enabled = False
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK_COUNTINUE").Enabled = False
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK_STOP").Enabled = False
        Show_Msg "�L�k�q AD ���o�q���M��"
        
        If MsgBox("�L�k�q AD ���o�q���M��A�O�_�n�Ұʥ����Ҧ��H", vbQuestion + vbYesNo, "����") = vbYes Then
            Get_Computers_From_Local
        End If
    End If
    

End Sub

Private Sub Add_Computer_To_TreeView()
'�[�J�q���� TreeView

    Show_Msg "���b�N�q���[�J��𪬲M��.."
    
    Tvw1.ImageList = vbalIml1
    Tvw1.Nodes.Clear
    Tvw1.CheckBoxes = True

    Dim NodeRoot As cTreeViewNode
    Dim NodeChildren As cTreeViewNodes
    Dim NodeSub  As cTreeViewNode
    Dim Icon_Root As Long
    Dim Icon_Sub As Long
    
    Icon_Root = vbalIml1.ItemIndex("root") - 1
    
    '�[�J�ڥؿ�
    Set NodeRoot = Tvw1.Nodes.Add(, etvwFirst, "root", UCase(AD_Domain_Name), Icon_Root)
    Set NodeChildren = NodeRoot.Children
    
    Icon_Sub = vbalIml1.ItemIndex("Disconnect") - 1
    
    '�[�J
    Dim i As Long
    For i = 1 To tmp_Computers.Count
        Set NodeSub = NodeChildren.Add(, etvwChild, Format(i, "000#") & CStr(tmp_Computers(i)), CStr(tmp_Computers(i)), Icon_Sub)
        NodeSub.Tag = "0"
    Next

    NodeRoot.Expanded = True

    Show_Msg "�q���[�J�𪬲M�槹��"

End Sub

Private Sub Get_Computers_From_Local()
'���������q��

    Show_Msg "���b�q Local ���o�q���M��.."
    
    '�C�|�����ܼ�
    'Dim aa, i
    'Do
    '    i = i + 1: aa = aa & Environ(i) & vbCrLf
    '    DoEvents
    'Loop Until Environ(i) = ""
    
    '���o�q���W��
    Tvw1.Nodes.Clear
    Tvw1.CheckBoxes = False
    Tvw1.Nodes.Add , , "0001" & Environ("COMPUTERNAME"), Environ("COMPUTERNAME")

    cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK").Enabled = False
    cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK_COUNTINUE").Enabled = False
    cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK_STOP").Enabled = False
    
    Show_Msg "�q���M��w���o"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If MsgBox("�z�T�w�n���}�ܡH", vbQuestion + vbYesNo, "����") = vbNo Then
    Cancel = 1
End If

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    If Me.ScaleHeight < 3000 Then Me.ScaleHeight = 3000
    If Me.ScaleWidth < 3000 Then Me.ScaleWidth = 3000
    
    
    Dim Mid_Top As Long: Mid_Top = cmdBar(0).Height + cmdBar(1).Height + cmdBar(2).Height
    Dim Mid_Height As Long: Mid_Height = Me.ScaleHeight - Mid_Top - vbalSBar1.Height
    
    Tvw1.TOp = Mid_Top
    Tvw1.Height = Mid_Height
    
    Label1.TOp = Mid_Top
    Label1.Left = Tvw1.Left + Tvw1.Width
    Label1.Height = Mid_Height
    
    Sg1.TOp = Mid_Top
    Sg1.Left = Label1.Left + Label1.Width
    Sg1.Width = Me.ScaleWidth - Sg1.Left
    Sg1.Height = Mid_Height
    

End Sub

Private Sub Form_Unload(Cancel As Integer)
        
On Error Resume Next

    '����Ҧ� Socket
    Dim i
    For i = Sck_List.LBound + 1 To Sck_List.UBound
        Unload Sck_List(i)
    Next

    '���񱼩Ҧ��u��C�P���s
    Dim cCombar As Integer, cComdBars As Integer
    For cCombar = cmdBar.LBound To cmdBar.UBound
        For cComdBars = 1 To cmdBar(cCombar).CommandBars.Count
            cmdBar(cCombar).CommandBars(cComdBars).Buttons.Clear
        Next
    Next
        
    FreeLibrary m_hMod
    
    '����Ҧ����
    Dim the_Frms As Form
    For Each the_Frms In Forms
        Unload the_Frms
    Next

If Err.Number <> 0 Then
    'MsgBox "�o�Ϳ��~ : " & Err.Description & "(" & Err.Number & ")", vbCritical, "���~"
End If


End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '�ǳƽվ�j�p
        '<EhHeader>
        On Error GoTo Label1_MouseDown_Err
        '</EhHeader>

100     If Button = vbLeftButton Then mbResizing = True

        '<EhFooter>
        Exit Sub

Label1_MouseDown_Err:
        MsgBox Err.Description & vbCrLf & _
               "in NetAdmin.Frm_Main.Label1_MouseDown " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

  '���U�ƹ�����ò��ʮ�, �۰ʽվ�U����j�p
    If mbResizing Then
        
        Dim nX As Single: nX = Label1.Left + x
        
        If nX < 50 Then Exit Sub
        If nX > Me.ScaleWidth - 300 Then Exit Sub
        
        
        Sg1.Width = Me.ScaleWidth - nX - Label1.Width
        Sg1.Left = nX + Label1.Width
        
        Tvw1.Width = nX
        Label1.Left = nX
        
    End If

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
  '����վ�j�p
    mbResizing = False

End Sub


Private Sub Sck_Check_One_Connect()

If Sck_Check_One.State = sckConnected Then
    SGrid_Computer_List "���`�s�u"
Else
    SGrid_Computer_List "����"
End If

End Sub

Private Sub Sck_Check_One_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

On Error Resume Next
    
    Select Case Number
       
        '�S�����D��
        Case sckHostNotFound, sckHostNotFoundTryAgain
            
            Call SGrid_Computer_List("�S�����D��")
        
        '�O��
        Case sckTimedout
        
            Call SGrid_Computer_List("�s�u�O��")
        
        
        Case Else
        
            Call SGrid_Computer_List("�s�u����")
        
    End Select

End Sub

Function Check_Show_Setup(Src_Item As String) As Integer

Check_Show_Setup = CInt(GetSetting(App.Title, "Show_Setup", Src_Item, "0"))

End Function

Function SGrid_Computer_List(Src_State As String)
'�C�X��������e

On Error Resume Next
    
    Dim tmp_BackColor As Long
    
    Tvw1.Enabled = False
    tmp_BackColor = Tvw1.BackColor
    Tvw1.BackColor = &H8000000F
    
    '�q���W�ٻP IP
    Dim strComputer As String:   strComputer = Tvw1.SelectedItem.Text
    Dim strComputer_IP As String: strComputer_IP = Sck_Check_One.RemoteHostIP
    
    '�q���W�٫e���s������
    Dim strIndex As Integer: strIndex = CInt(Left(Tvw1.SelectedItem.Key, 4))
                       
                       
    Sck_Check_One.Close
    DoEvents
    
    '�Ȧs�}�l�B�z�ɶ�
    Dim tmr_Start As Single: tmr_Start = Timer
    
    
    With cmdBar(0).Toolbar.Buttons
        .Item("CONNECT_CHECK").Enabled = True
        .Item("CONNECT_CHECK_COUNTINUE").Enabled = False
        .Item("CONNECT_CHECK_STOP").Enabled = False
    End With
    
    
    
    With Sg1
    
        '.Redraw = False
        
        .Clear True
        
        .AddColumn , "����"
        .AddColumn , "���e"
        .AddColumn , "���O"
        
     
         If Src_State = "���`�s�u" Then
            
            Show_Msg "���b�P�ؼйq���إ߳s�u���A�еy��..."
            
            '�L�k�s�u
            If WMI_Service_Create(strComputer) = False Then
              
                SGrid_Computer_List_AddRow "�i��S���v���άO�����o�Ͱ��D", Src_State, "���~"
                  Show_Msg "�L�k�s�u��" & strComputer & "�A�z�i��S���v���άO�����o�Ͱ��D"
            
                'Call Show_Tvw_Computer_State(strIndex, "���u")
            
            '�i�s�u
            Else
              
                'Call Show_Tvw_Computer_State(strIndex, "���`�s�u")
              
                Show_Msg "Ū���q���W�٤��A�еy��..."
                    SGrid_Computer_List_AddRow "�q���W��", strComputer, "�@��"
                
                Show_Msg "Ū��������}���A�еy��..."
                    SGrid_Computer_List_AddRow "�q��������}", strComputer_IP, "�@��"
                
                Show_Msg "Ū���q���n�J�ϥΪ̤��A�еy��..."
                    WMI_Computer_Login_UserName strComputer
                
                Show_Msg "Ū���q���@�~�t�θ�T���A�еy��..."
                    WMI_Computer_OperatingSystem strComputer
    
                If Check_Show_Setup("�����s��") = 1 Then
                    Show_Msg "Ū���q�������s�դ��A�еy��..."
                    ADSI_Computer_Group strComputer
                End If
                
                If Check_Show_Setup("�����ϥΪ�") = 1 Then
                    Show_Msg "Ū���q�������ϥΪ̤��A�еy��..."
                    ADSI_Computer_User strComputer
                End If
    
                If Check_Show_Setup("��ܼҦ�") = 1 Then
                    Show_Msg "Ū���q���i����ܼҦ���T���A�еy��..."
                    CMI_Computer_VideoControllerResolution strComputer
                End If
                
                If Check_Show_Setup("�ѽX��") = 1 Then
                    Show_Msg "Ū���q���ѽX����T���A�еy��..."
                    WMI_Computer_CodecFile strComputer
                End If
                
                If Check_Show_Setup("��ܾ�") = 1 Then
                    Show_Msg "Ū���q����ܾ���T���A�еy��..."
                    WMI_Computer_DesktopMonitor strComputer
                End If
                
                If Check_Show_Setup("��ܥd") = 1 Then
                    Show_Msg "Ū���q����ܥd��T���A�еy��..."
                    WMI_Computer_Displays strComputer
                End If
                    
                If Check_Show_Setup("�s����") = 1 Then
                    Show_Msg "Ū���q���s�����T���A�еy��..."
                    WMI_Computer_SerialPort strComputer
                End If
                
                If Check_Show_Setup("����ʸ˸m") = 1 Then
                    Show_Msg "Ū���q������ʸ˸m��T���A�еy��..."
                    WMI_Computer_Keyboard strComputer
                End If
                
                If Check_Show_Setup("�H���Y�θ˸m") = 1 Then
                    Show_Msg "Ū���q���H���Y�θ˸m��T���A�еy��..."
                    WMI_Computer_PnPEntity strComputer
                End If
                
                If Check_Show_Setup("����ʸ˸m") = 1 Then
                    Show_Msg "Ū���q�����Щʸ˸m��T���A�еy��..."
                    WMI_Computer_PointingDevice strComputer
                End If
                   
                If Check_Show_Setup("���ĸ˸m") = 1 Then
                    Show_Msg "Ū���q�����ĸ˸m��T���A�еy��..."
                    WMI_Computer_SoundDevice strComputer
                End If
                
                If Check_Show_Setup("�t��") = 1 Then
                    Show_Msg "Ū���q���t�θ�T���A�еy��..."
                    WMI_Computer_System_Information strComputer
                End If
                    
                If Check_Show_Setup("�����ܼ�") = 1 Then
                    Show_Msg "Ū���q�������ܼƸ�T���A�еy��..."
                    WMI_Computer_Environment strComputer
                End If
                    
                If Check_Show_Setup("�}������") = 1 Then
                    Show_Msg "Ū���q���}�������T���A�еy��..."
                    WMI_Computer_StartupCommand strComputer
                End If
                
                If Check_Show_Setup("�׺ݾ�") = 1 Then
                    Show_Msg "Ū���q���׺ݾ���T���A�еy��..."
                    WMI_Computer_Terminal strComputer
                End If
                
                If Check_Show_Setup("�n��w��") = 1 Then
                    Show_Msg "Ū���q���n��w�˸�T���A�еy��..."
                    WMI_Computer_Product_Installed strComputer
                End If
                
                If Check_Show_Setup("��s��") = 1 Then
                    Show_Msg "Ū���q����s�ɸ�T���A�еy��..."
                    WMI_Computer_QuickFixEngineering strComputer
                End If
                
                If Check_Show_Setup("�޿�Ϻ�") = 1 Then
                    Show_Msg "Ū���q���޿�Ϻи�T���A�еy��..."
                    WMI_Computer_LogicalDisk strComputer
                End If
                
                If Check_Show_Setup("�A��") = 1 Then
                    Show_Msg "Ū���q���A�ȸ�T���A�еy��..."
                    WMI_Computer_Service strComputer
                End If
                
                If Check_Show_Setup("�����") = 1 Then
                    Show_Msg "Ū���q���������T���A�еy��..."
                    WMI_Computer_Process strComputer
                End If
                
                If Check_Show_Setup("��������") = 1 Then
                    Show_Msg "Ū���q������������T���A�еy��..."
                    WMI_Computer_NetworkAdapter strComputer
                End If
                
                If Check_Show_Setup("�����˸m�]�w") = 1 Then
                    Show_Msg "Ū���q�������˸m�]�w��T���A�еy��..."
                    WMI_Computer_NetworkAdapterConfiguration strComputer
                End If
                
                If Check_Show_Setup("�Ϻо�") = 1 Then
                    Show_Msg "Ū���q���Ϻо��˸m��T���A�еy��..."
                    WMI_Computer_DiskDrive strComputer
                End If
                
                If Check_Show_Setup("�@��") = 1 Then
                    Show_Msg "Ū���q���@�θ�T���A�еy��..."
                    WMI_Computer_Share strComputer
                End If
                
                If Check_Show_Setup("�L���") = 1 Then
                    Show_Msg "Ū���q���L�����T���A�еy��..."
                    WMI_Computer_Printer strComputer
                End If
                
                If Check_Show_Setup("�L��������s����") = 1 Then
                    Show_Msg "Ū���q���L��������s�����T���A�еy��..."
                    WMI_Computer_TCPIPPrinterPort strComputer
                End If
                
                If Check_Show_Setup("�L����X�ʵ{��") = 1 Then
                    Show_Msg "Ū���q���L����X�ʵ{����T���A�еy��..."
                    WMI_Computer_PrinterDriver strComputer
                End If
                
                
                
                If Check_Show_Setup("�ϺФ��ΰ�") = 1 Then
                    Show_Msg "Ū���q���ϺФ��ΰϸ�T���A�еy��..."
                    WMI_Computer_DiskPartition strComputer
                End If
                
                Show_Msg "�q�����Ū�����\"
                
            End If
        Else
        
            SGrid_Computer_List_AddRow "���~", Src_State, "���~"
            Show_Msg "�L�k�s�u�� " & strComputer
            
            
        End If
        
        
        Show_Msg "���b�N��Ƹs�դ����..."
        .ColumnIsGrouped(3) = False
        .ColumnIsGrouped(3) = True
        Show_Msg "��Ƹs�դ���ܧ���"
        
        .AutoWidthColumn 1
        .AutoWidthColumn 2
        
        .Redraw = True
        
    End With
    
    vbalGrid_Sort Me.Sg1, 1
    ExGroup Me.Sg1
    
    
    Show_Msg "�����ܧ����A�@�ӶO : " & Timer - tmr_Start & " ��"
    
    Tvw1.BackColor = tmp_BackColor
    Tvw1.Enabled = True

End Function

Sub SGrid_Computer_List_AddRow(Field1 As String, Field2 As String, Field3 As String)
'�N�q����T�[�J SGrid

On Error Resume Next

With Sg1
    
    .AddRow
    .CellDetails .Rows, 1, Field1
    .CellDetails .Rows, 2, Field2
    .CellDetails .Rows, 3, Field3
    .AutoWidthColumn 1
    .AutoWidthColumn 2

    '.EnsureVisible .Rows, 1

End With

End Sub

Function SGrid_Computers_List()
'�C�X����q���M��H�γs�u���A

On Error Resume Next

    '����C��
    'Call Stop_Computer_List
    
    If Stop_Computers_List = True Then
        Stop_Computer_List
        Exit Function
    End If
    
    '���o�Ҧ��q��
    Show_Msg "���b�C�X�q��..."
    
    Set tmp_Computers = ADSI_Computer_List()
    
    With Sg1
        
        '.Redraw = False
        
        '�Y���s�}�l�~�M���e��
        If List_Last = 1 Then
            .Clear True
        End If
                    
        .AddColumn "Computer_Name", "�q���W��"
        .AddColumn "Computer_Online", "�s�u"
        .AddColumn "Computer_IP", "IP Address"
            
        Dim i As Long
        
        If Sck_List.Count > 1 Then
            For i = 1 To Sck_List.UBound - 1
                Unload Sck_List(i)
            Next
        
        End If
        
        Show_Msg "���b�}�ҳs����..."
        
        For i = List_Last To tmp_Computers.Count
            Load Sck_List(i)
            Frm_Main.Sck_List(i).Connect CStr(tmp_Computers(i)), Port_Ping
            DoEvents
        Next
    
        Show_Msg "�s�u��w�}��"
        
        .AutoWidthColumn 1
        .Redraw = True
    End With



End Function

Sub Stop_Computer_List()
'����C��

On Error Resume Next

    Show_Msg "���b����C��..."
    Stop_Computers_List = True
    
    If Sck_List.Count > 1 Then
        
        Dim i As Long
        For i = 1 To Sck_List.UBound - 1
            Sck_List(i).Close: Unload Sck_List(i)
        Next
    
    End If
    
    Delay 5

    Show_Msg "�C��w����"

End Sub


Sub SGrid_Computers_List_Addrow(Src_Index As Integer, Src_Exist As String, Optional Src_IP As String = "")
'��ܹq�����p

If Stop_Computers_List = True Then Exit Sub
    

   '���\�s�u
    With Sg1
        
        .AddRow
        .CellDetails .Rows, .ColumnIndex("Computer_Name"), CStr(tmp_Computers(Src_Index))
        .CellDetails .Rows, .ColumnIndex("Computer_Online"), Src_Exist
        .CellDetails .Rows, .ColumnIndex("Computer_IP"), Src_IP
        
        .AutoWidthColumn 1
        .AutoWidthColumn 2
        .AutoWidthColumn 3
        
        .EnsureVisible .Rows, 1
        
    End With
        
    '���ܾ�t���q�����A
    Call Show_Tvw_Computers_State(Sck_List(Src_Index), Src_Exist)

    '�����ثe�̫�@������
    List_Last = Src_Index
    
    '�����o�� Socket Port
    Sck_List(Src_Index).Close
    
    Show_Msg "���bŪ���ϥΤ��q�� : " & Sg1.Rows & "/" & tmp_Computers.Count
    
    '���o����
    If Sg1.Rows = tmp_Computers.Count Then
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK").Enabled = True
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK_COUNTINUE").Enabled = True
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK_STOP").Enabled = False
        Show_Msg "Ū������ " & Sg1.Rows & "/" & tmp_Computers.Count
    End If
    

End Sub

Sub Show_Tvw_Computers_State(Src_Winsock As Winsock, Src_Exist As String)

    Dim Src_Index As Integer: Src_Index = Src_Winsock.index

    '�b��t��W�ХܥX���A
    Dim tmp_key As String: tmp_key = Format(Src_Index, "0000") & CStr(tmp_Computers(Src_Index))
    Dim SubNode As vbalTreeViewLib6.cTreeViewNode
      
    Set SubNode = Tvw1.Nodes.Item(tmp_key)

    If Src_Exist = "���`�s�u" Then
        
        '�]�w�i�s�u���ϥܻP��r
        SubNode.Bold = True
        
        'SubNode.Tag = Sck_List(Src_Index).RemoteHostIP
        SubNode.Tag = Src_Winsock.RemoteHostIP
        
        SubNode.Image = vbalIml1.ItemIndex("Connected") - 1
        SubNode.SelectedImage = vbalIml1.ItemIndex("Connected") - 1
        
        '�i�s�u�q���`�� +1
        Total_Connected_Computers = Total_Connected_Computers + 1
        Show_Msg ""
    Else
    
        '�]�w�L�k�s�u���ϥܻP��r
        SubNode.Bold = False
        SubNode.Tag = "0"
        SubNode.Image = vbalIml1.ItemIndex("Disconnect") - 1
        SubNode.SelectedImage = vbalIml1.ItemIndex("Disconnect") - 1
    End If

End Sub

Private Sub Sck_List_Connect(index As Integer)

If Sck_List(index).State = sckConnected Then
    Call SGrid_Computers_List_Addrow(index, "���`�s�u", Sck_List(index).RemoteHostIP)
End If
 
End Sub

Private Sub Sck_List_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
    
    Select Case Number
       
        '�S�����D��
        Case sckHostNotFound, sckHostNotFoundTryAgain
            
            Call SGrid_Computers_List_Addrow(index, "�S�����D��")
        
        '�O��
        Case sckTimedout
        
            Call SGrid_Computers_List_Addrow(index, "�s�u�O��")
        
        '�䥦���~ (�H�������~���h)
        Case Else
        
            Call SGrid_Computers_List_Addrow(index, "�s�u����")
        
    End Select
        

'MsgBox Number
End Sub

Private Sub Sg1_ColumnClick(ByVal lCol As Long)

    '�Ƨ�
    vbalGrid_Sort Me.Sg1, lCol
    ExGroup Me.Sg1

End Sub

Private Sub Sg1_Click(ByVal lRow As Long, ByVal lCol As Long)

End Sub

Private Sub Tvw1_nodeCheck(node As vbalTreeViewLib6.cTreeViewNode)

'�I��ڥؿ��A�i�����Υ�������
If node.Key = "root" Then

    'Dim tmpa As String
        
    Dim i As Long
    For i = 1 To Tvw1.NodeCount
        Tvw1.Nodes(i).Checked = node.Checked
        'tmpa = tmpa & Tvw1.Nodes(i).Text & " , "
    Next

End If

End Sub

Private Sub Tvw1_NodeClick(node As vbalTreeViewLib6.cTreeViewNode)
'�p�G�I��ڥؿ����B�z
If node.Key = "root" Then
    Can_Process_Computer = ""
    Exit Sub
End If

'���o��T
Can_Process_Computer = node.Text

End Sub

Private Sub Tvw1_NodeDblClick(node As vbalTreeViewLib6.cTreeViewNode)

'�p�G�I��ڥؿ����B�z
If node.Key = "root" Then Exit Sub

'���o��T
Can_Process_Computer = node.Text

Show_Computer_Information node

End Sub

Sub Show_Computer_Information(node As vbalTreeViewLib6.cTreeViewNode)
'�C�X�q���ԲӸ�T

    '���u���A�̤��i�B�z
    If node.Tag = "0" Then
        If MsgBox("�B�z���u���A�q���t�ױN�|�ܺC�A�z�T�w�n���ջP�ӹq���s�u�ܡH", vbYesNo, "����") = vbNo Then Exit Sub
    End If
    
    Show_Msg "�q�����Ū�����A�еy��..."
    Sck_Check_One.Close
    Sck_Check_One.Connect node.Text, Port_Ping


End Sub


Private Sub Tvw1_NodeRightClick(node As vbalTreeViewLib6.cTreeViewNode)

'�p�G�I��ڥؿ����B�z
If node.Key = "root" Then
    Can_Process_Computer = ""
    Exit Sub
End If

    '���ݥ��������A�o�˥\���\��i�J��~���q���W�٥i��
    node.Selected = True
    
    '���o��T
    Can_Process_Computer = node.Text
    
    
    '���o��Цb�ù��W����m
    Call GethWnd
    
    '�ƹ�W�o��i�H���n
    cmdBar(1).ClientCoordinatesToScreen Me.Left, Me.TOp, Me.hwnd
    
    '�q�X��ܧ���\���
    cmdBar(1).ShowPopupMenu tmp_x, tmp_y, cmdBar(1).CommandBars("MANAGERMENU")
    
End Sub

Private Sub Txt_DomainName_GotFocus()
AutoSelStr Me.ActiveControl
End Sub

Private Sub Txt_Port_GotFocus()
AutoSelStr Me.ActiveControl
End Sub
