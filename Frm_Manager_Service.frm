VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{2210EC79-A724-4033-AAF4-790E2467C0E8}#1.0#0"; "vbalCmdBar6.ocx"
Begin VB.Form Frm_Manager_Service 
   Caption         =   "�A�Ⱥ޲z"
   ClientHeight    =   4950
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5805
   Icon            =   "Frm_Manager_Service.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4950
   ScaleWidth      =   5805
   StartUpPosition =   3  '�t�ιw�]��
   Begin vbalCmdBar6.vbalCommandBar CmdBar1 
      Align           =   1  '������W��
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   556
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
   Begin vbAcceleratorSGrid6.vbalGrid Sg1 
      Height          =   4035
      Left            =   420
      TabIndex        =   0
      Top             =   540
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   7117
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
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
      DefaultRowHeight=   19
      AllowGrouping   =   -1  'True
      HideGroupingBox =   -1  'True
      GroupBoxHintText=   "�즲�s�������D�ܦ��A�즲��Ы��U���D�C�Ƨǥi�P�ɮi�}�Ҧ��s��"
      HotTrack        =   -1  'True
      SelectionOutline=   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '�z��
      Caption         =   "�B�z��..."
      Height          =   555
      Left            =   300
      TabIndex        =   1
      Top             =   660
      Width           =   2355
   End
End
Attribute VB_Name = "Frm_Manager_Service"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ConfigureButtons()
'�إ߳]�w�u��C�W�����s

On Error Resume Next

    Dim btn As cButton
    
    CmdBar1.ToolbarImageList = Frm_Main.vbalIml_Tools.hIml
    With CmdBar1.Buttons
    
        '�إ� TOOLS ���s�s
        Set btn = .Add("TOOLS:SERVICE:SPLIT", , , eSeparator)
        
        Set btn = .Add("ACTION:START", Frm_Main.vbalIml_Tools.ItemIndex("ACTION:START") - 1, "�Ұ�", , "�ҰʪA��")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = False
        
        Set btn = .Add("ACTION:STOP", Frm_Main.vbalIml_Tools.ItemIndex("ACTION:STOP") - 1, "����", , "����A��")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = False
        
        Set btn = .Add("ACTION:DELETE", Frm_Main.vbalIml_Tools.ItemIndex("ACTION:DELETE") - 1, "�R��", , "�R���A��")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = False
        
        Set btn = .Add("TOOLS:SERVICE:SPLIT2", , , eSeparator)
    
        'Frm_Main.vbalIml_Tools.ItemIndex("STARTMODE_AUTO") - 1
        Set btn = .Add("STARTMODE:AUTO", Frm_Main.vbalIml_Tools.ItemIndex("STARTMODE:AUTO") - 1, "�۰ʱҰ�", , "�N�A�ȳ]�w���۰ʱҰ�")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = False
        
        Set btn = .Add("STARTMODE:MANUAL", Frm_Main.vbalIml_Tools.ItemIndex("STARTMODE:MANUAL") - 1, "���", , "�N�A�ȳ]�w����ʱҰ�")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = False
        
        Set btn = .Add("STARTMODE:DISABLED", Frm_Main.vbalIml_Tools.ItemIndex("STARTMODE:DISABLED") - 1, "����", , "�N�A�ȳ]���w����")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = False
        
    
        Set btn = .Add("TOOLS:SERVICE:SPLIT3", , , eSeparator)
        
        Set btn = .Add("REFRESH", Frm_Main.vbalIml_Tools.ItemIndex("REFRESH") - 1, "���s��z", , "���s�C��̷s���A")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = True
        
    
    End With
End Sub

Private Sub ConfigureBars()

On Error Resume Next

    Dim bar As cCommandBar
    Dim Btns As cCommandBarButtons


    '�u��C�s����
    With CmdBar1
            
        '�إߤ@�ӷs�u��C
        Set bar = .CommandBars.Add("SERVICE", "Standard Buttons")
        Set Btns = bar.Buttons
             
            '�إ�
            Btns.Add .Buttons.Item("TOOLS:SERVICE:SPLIT")
            Btns.Add .Buttons.Item("ACTION:START")
            Btns.Add .Buttons.Item("ACTION:STOP")
            Btns.Add .Buttons.Item("ACTION:DELETE")
            Btns.Add .Buttons.Item("TOOLS:SERVICE:SPLIT2")
            Btns.Add .Buttons.Item("STARTMODE:AUTO")
            Btns.Add .Buttons.Item("STARTMODE:MANUAL")
            Btns.Add .Buttons.Item("STARTMODE:DISABLED")
            Btns.Add .Buttons.Item("TOOLS:SERVICE:SPLIT3")
            Btns.Add .Buttons.Item("REFRESH")
            
    End With

End Sub

Private Sub CmdBar1_ButtonClick(btn As vbalCmdBar6.cButton)

    If Sg1.Rows < 1 Then Exit Sub

    Dim sRow As Long: sRow = Sg1.SelectedRow
    Dim sCol As Long: sCol = Sg1.ColumnIndex("Name")

    Sg1.Enabled = False

    Select Case btn.Key
    
        Case "ACTION:START":  Service_Action Sg1.CellText(sRow, sCol), "ACTION:START", sRow
        Case "ACTION:STOP":   Service_Action Sg1.CellText(sRow, sCol), "ACTION:STOP", sRow
        Case "ACTION:DELETE":
            If MsgBox("�T�w�n�R���o�ӪA�ȡH", vbQuestion + vbYesNo, "����") = vbYes Then
                Service_Action Sg1.CellText(sRow, sCol), "ACTION:DELETE", sRow
            End If
        
        Case "STARTMODE:AUTO": Service_Action Sg1.CellText(sRow, sCol), "STARTMODE:AUTO", sRow
        Case "STARTMODE:MANUAL": Service_Action Sg1.CellText(sRow, sCol), "STARTMODE:MANUAL", sRow
        Case "STARTMODE:DISABLED": Service_Action Sg1.CellText(sRow, sCol), "STARTMODE:DISABLED", sRow
                
        Case "REFRESH": Show_Service
        
    End Select
    
    Sg1.Enabled = True

End Sub

Private Sub Form_Load()

'On Error Resume Next
'If CmdBar1.Buttons.Count < 1 Then

    ConfigureButtons
    ConfigureBars
    CmdBar1.Toolbar = CmdBar1.CommandBars("SERVICE")


'End If

If Trim(Can_Process_Computer) = "" Then
    MsgBox "�S���q�𪬲M�椤��ܭn�޲z���q��", vbQuestion, "���~"
    Unload Me
    Exit Sub
End If

If WMI_Service_Create(Can_Process_Computer) = False Then
    MsgBox "�L�k�s�u���q��", vbQuestion, "���~"
    Unload Me
    Exit Sub
End If
    
    


Me.Caption = Can_Process_Computer & " �W���A�ȲM��"

Show_Service

Me.Show

End Sub


Sub Show_Service()
'��ܪA��
    
'On Error Resume Next


    
    Dim cols, obj
    Set cols = objWMIService.ExecQuery("Select * from Win32_Service")
    
    Sg1.Clear True
    
    Sg1.Redraw = False
    Sg1.Visible = False

    Sg1.AddColumn "Name", "�W��"
    Sg1.AddColumn "State", "���A"
    Sg1.AddColumn "Startmode", "�ҰʼҦ�"
    Sg1.AddColumn "Path", "���|"
    Sg1.AddColumn "Detail", "����"
    
    If cols.Count > 0 Then
        
        For Each obj In cols
            Add_Row _
                StrNullToSpace(obj.Name), _
                StrNullToSpace(obj.State), _
                StrNullToSpace(obj.StartMode), _
                StrNullToSpace(obj.Pathname), _
                StrNullToSpace(obj.Description)
            DoEvents
        Next
    
    End If
    
    Sg1.Redraw = True
    Sg1.Visible = True
    
End Sub

Sub Add_Row(Field1 As String, Optional Field2 As String = "", Optional Field3 As String = "", Optional field4 As String = "", Optional field5 As String = "")

With Sg1

    .AddRow
    .CellDetails .Rows, .ColumnIndex("Name"), Field1
    .CellDetails .Rows, .ColumnIndex("State"), Field2
    .CellDetails .Rows, .ColumnIndex("Startmode"), Field3
    .CellDetails .Rows, .ColumnIndex("Path"), field4
    .CellDetails .Rows, .ColumnIndex("Detail"), field5
    

    .AutoWidthColumn .ColumnIndex("Name")
    .AutoWidthColumn .ColumnIndex("State")
    .AutoWidthColumn .ColumnIndex("Startmode")
    '.AutoWidthColumn .ColumnIndex("Path")
    .AutoWidthColumn .ColumnIndex("Detail")

End With
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    Sg1.TOp = CmdBar1.Height
    Sg1.Left = 0
    Sg1.Width = Me.ScaleWidth
    Sg1.Height = Me.ScaleHeight - CmdBar1.Height
    
End Sub

Private Sub Sg1_ColumnClick(ByVal lCol As Long)
    '�Ƨ�
    vbalGrid_Sort Me.Sg1, lCol
    ExGroup Me.Sg1

End Sub

Function Service_Action(Src_Service_Name As String, Src_State As String, Src_Row As Long)
'�B�z�A��

Dim tmp_flag As Boolean: tmp_flag = False
Dim tmp_return As String: tmp_return = Action_Service(Src_Service_Name, Src_State)
    
    Select Case Src_State
        
        'action:Start
        Case "ACTION:START"
            If tmp_return = "0" Then
                Sg1.CellDetails Src_Row, Sg1.ColumnIndex("State"), "Running": tmp_flag = True
            End If
        
        'action:Stop
        Case "ACTION:STOP"
            If tmp_return = "0" Then
                Sg1.CellDetails Src_Row, Sg1.ColumnIndex("State"), "Stopped": tmp_flag = True
            End If
        
        'action:Delete
        Case "ACTION:DELETE"
            If tmp_return = "0" Then
                Sg1.RemoveRow Src_Row: tmp_flag = True
            End If
        
        
        'startmode:Auto
        Case "STARTMODE:AUTO"
            If tmp_return = "0" Then
                Sg1.CellDetails Src_Row, Sg1.ColumnIndex("Startmode"), "Auto": tmp_flag = True
            End If
        
        'startmode:Manual
        Case "STARTMODE:MANUAL"
            If tmp_return = "0" Then
                Sg1.CellDetails Src_Row, Sg1.ColumnIndex("Startmode"), "Manual": tmp_flag = True
            End If
        
        'startmode:Stopped
        Case "STARTMODE:DISABLED"
            If tmp_return = "0" Then
                Sg1.CellDetails Src_Row, Sg1.ColumnIndex("Startmode"), "Disabled": tmp_flag = True
            End If
                
        
    End Select

    
If tmp_flag = False Then MsgBox "����", vbInformation, "���~"

End Function

Private Sub Sg1_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)

Select Case Sg1.CellText(lRow, Sg1.ColumnIndex("State"))
    Case "Running"
        CmdBar1.Toolbar.Buttons.Item("ACTION:START").Enabled = False
        CmdBar1.Toolbar.Buttons.Item("ACTION:STOP").Enabled = True
        CmdBar1.Toolbar.Buttons.Item("ACTION:DELETE").Enabled = True
    Case "Stopped"
        CmdBar1.Toolbar.Buttons.Item("ACTION:START").Enabled = True
        CmdBar1.Toolbar.Buttons.Item("ACTION:STOP").Enabled = False
        CmdBar1.Toolbar.Buttons.Item("ACTION:DELETE").Enabled = False
End Select

Select Case Sg1.CellText(lRow, Sg1.ColumnIndex("Startmode"))
    Case "Auto"
        CmdBar1.Toolbar.Buttons.Item("STARTMODE:AUTO").Enabled = False
        CmdBar1.Toolbar.Buttons.Item("STARTMODE:MANUAL").Enabled = True
        CmdBar1.Toolbar.Buttons.Item("STARTMODE:DISABLED").Enabled = True
    Case "Manual"
        CmdBar1.Toolbar.Buttons.Item("STARTMODE:AUTO").Enabled = True
        CmdBar1.Toolbar.Buttons.Item("STARTMODE:MANUAL").Enabled = False
        CmdBar1.Toolbar.Buttons.Item("STARTMODE:DISABLED").Enabled = True
    Case "Disabled"
        CmdBar1.Toolbar.Buttons.Item("STARTMODE:AUTO").Enabled = True
        CmdBar1.Toolbar.Buttons.Item("STARTMODE:MANUAL").Enabled = True
        CmdBar1.Toolbar.Buttons.Item("STARTMODE:DISABLED").Enabled = False
End Select


End Sub
