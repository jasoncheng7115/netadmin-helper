VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{2210EC79-A724-4033-AAF4-790E2467C0E8}#1.0#0"; "vbalCmdBar6.ocx"
Begin VB.Form Frm_Manager_Printer 
   Caption         =   "�L����޲z"
   ClientHeight    =   7035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10155
   Icon            =   "Frm_Manager_Printer.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7035
   ScaleWidth      =   10155
   StartUpPosition =   3  '�t�ιw�]��
   Begin vbalCmdBar6.vbalCommandBar CmdBar1 
      Align           =   1  '������W��
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   10155
      _ExtentX        =   17912
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
   Begin vbAcceleratorSGrid6.vbalGrid Sg_Printer 
      Height          =   1155
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2037
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
   Begin vbAcceleratorSGrid6.vbalGrid Sg_PrinterDriver 
      Height          =   1515
      Left            =   0
      TabIndex        =   2
      Top             =   1560
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2672
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
   Begin vbAcceleratorSGrid6.vbalGrid Sg_PrinterTCPIPPort 
      Height          =   1515
      Left            =   0
      TabIndex        =   3
      Top             =   3120
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   2672
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
Attribute VB_Name = "Frm_Manager_Printer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ConfigureButtons()
'�إ߳]�w�u��C�W�����s

'On Error Resume Next

    Dim btn As cButton
    
    'CmdBar1.ToolbarImageList = Frm_Main.vbalIml_Tools.hIml
    With CmdBar1.Buttons
    
        '�إ� TOOLS ���s�s
        Set btn = .Add("TOOLS:PRINTER:SPLIT", , , eSeparator)
        
        Set btn = .Add("PRINTER:ADD", , "�s�W�L���", , "�b�ؼйq���إߤ@�ӷs�������L���")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = True
        
        Set btn = .Add("PRINTER:DELETE", , "�R���L���", , "�R���ؼйq�����L���")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = False
        
        Set btn = .Add("TOOLS:PRINTERDRIVER:SPLIT", , , eSeparator)
        
        Set btn = .Add("PRINTERDRIVER:ADD", , "�s�W�X�ʵ{��", , "�b�ؼйq���إߤ@�ӷs���L����X�ʵ{��")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = True
        
        Set btn = .Add("PRINTERDRIVER:DELETE", , "�R���X�ʵ{���L���", , "�R���ؼйq�����L����X�ʵ{��")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = False
        btn.Visible = False
        
        Set btn = .Add("TOOLS:PRINTERTCPIPPORT:SPLIT", , , eSeparator)
        
        Set btn = .Add("PRINTERTCPIPPORT:ADD", , "�s�W�s����", , "�b�ؼйq���إߤ@�ӷs���L����s����")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = True
        
        Set btn = .Add("PRINTERTCPIPPORT:DELETE", , "�R���s����", , "�R���ؼйq�����L����s����")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = False
        
        Set btn = .Add("TOOLS:PRINTER:SPLIT2", , , eSeparator)
        
        Set btn = .Add("PRINTER:REFRESH", , "���s��z", , "���s�C��̷s���A")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = True
        
    
    End With
End Sub

Private Sub ConfigureBars()

    Dim bar As cCommandBar
    Dim Btns As cCommandBarButtons

    '�u��C�s����
    With CmdBar1
            
        '�إߤ@�ӷs�u��C
        Set bar = .CommandBars.Add("PRINTER", "Standard Buttons")
        Set Btns = bar.Buttons
             
            '�إ�
            Btns.Add .Buttons.Item("TOOLS:PRINTER:SPLIT")
            Btns.Add .Buttons.Item("PRINTER:ADD")
            Btns.Add .Buttons.Item("PRINTER:DELETE")
            Btns.Add .Buttons.Item("TOOLS:PRINTERDRIVER:SPLIT")
            Btns.Add .Buttons.Item("PRINTERDRIVER:ADD")
            Btns.Add .Buttons.Item("PRINTERDRIVER:DELETE")
            Btns.Add .Buttons.Item("TOOLS:PRINTERTCPIPPORT:SPLIT")
            Btns.Add .Buttons.Item("PRINTERTCPIPPORT:ADD")
            Btns.Add .Buttons.Item("PRINTERTCPIPPORT:DELETE")
            Btns.Add .Buttons.Item("TOOLS:PRINTER:SPLIT2")
            Btns.Add .Buttons.Item("PRINTER:REFRESH")
            
    End With

End Sub

Private Sub CmdBar1_ButtonClick(btn As vbalCmdBar6.cButton)

Dim tmp_Ret As String

Select Case btn.Key

    '�إߦL���
    Case "PRINTER:ADD"
        
        Call Action_Printer("", "ACTION:ADD")
        Show_Printer
        
    '�R���L���
    Case "PRINTER:DELETE"
        
        tmp_Ret = Sg_Printer.CellText(Sg_Printer.SelectedRow, Sg_Printer.ColumnIndex("Name"))
        If Sg_Printer.SelectedRow > 0 Then
            Call Action_Printer(tmp_Ret, "ACTION:DELETE")
            Show_Printer
        End If
        
            
    '�s�W�L����X�ʵ{��
    Case "PRINTERDRIVER:ADD"
        
        tmp_Ret = Action_PrinterDriver("", "ACTION:ADD")
        If tmp_Ret = "0" Then
            Show_PrinterDriver
            MsgBox "�s�W���\", vbInformation, "����"
        ElseIf tmp_Ret = "-1" Then
            
        Else
            MsgBox "����", vbInformation, "���~"
        End If
    
    '�إ߷s�L����s����
    Case "PRINTERTCPIPPORT:ADD"
        
        Call Action_PrinterTCPIPPort("", "ACTION:ADD")
        'If Action_PrinterTCPIPPort("", "ACTION:ADD") <> "0" Then
        '    MsgBox "����", vbInformation, "���~"
        'Elses
            Show_PrinterTCPIPPrinterPort
        'End If
    
    '�R���L����s����
    Case "PRINTERTCPIPPORT:DELETE"
        
        With Sg_PrinterTCPIPPort
            
            If .SelectedRow > 0 Then
                If .CellText(.SelectedRow, 2) <> "���䴩" Then
                    
                     If Action_PrinterTCPIPPort(.CellText(.SelectedRow, .ColumnIndex("Name")), "ACTION:DELETE") <> "0" Then
                        MsgBox "���ѡA��]�i��O���L����Q�]�w�ϥθӳs����", vbInformation, "���~"
                     Else
                        Show_PrinterTCPIPPrinterPort
                     End If
                     
                End If
            End If
            
        End With
        
    '���s��z
    Case "PRINTER:REFRESH"
    
            Show_Printer
            Show_PrinterDriver
            Show_PrinterTCPIPPrinterPort
    
End Select

End Sub

Private Sub Form_Load()

On Error Resume Next
'If CmdBar1.Buttons.Count < 1 Then

    ConfigureButtons
    ConfigureBars
    CmdBar1.Toolbar = CmdBar1.CommandBars("PRINTER")

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
    
    
Me.Caption = Can_Process_Computer & " �W���L����޲z"
   

Show_Printer
Show_PrinterDriver
Show_PrinterTCPIPPrinterPort
Me.Show

End Sub

Sub Show_Printer()
'��ܦL���

    '��ܪA��
    
On Error Resume Next

   
    Dim cols ', 'objWMIService
    Set cols = objWMIService.ExecQuery("Select * from Win32_Printer")
    
    With Sg_Printer
        
        .Clear True
        .Redraw = False
        .Visible = False

        .AddColumn "DriverName", "�L�������"
        .AddColumn "Name", "�W��"
        .AddColumn "Port", "�s����"
        .AddColumn "DataType", "��Ʈ榡"
    
        If cols.Count > 0 Then
            
            For Each obj In cols
            
                .AddRow
                .CellDetails .Rows, .ColumnIndex("DriverName"), obj.DriverName & " "
                .CellDetails .Rows, .ColumnIndex("Name"), obj.Name & " "
                .CellDetails .Rows, .ColumnIndex("Port"), obj.PortName & " "
                .CellDetails .Rows, .ColumnIndex("DataType"), obj.PrintJobDataType & " "
                DoEvents
            
            Next
        
        End If
    
        .AutoWidthColumn "DriverName"
        .AutoWidthColumn "Name"
        .AutoWidthColumn "Port"
        .AutoWidthColumn "DataType"
    
        .Redraw = True
        .Visible = True
    
    End With
    
End Sub


Sub Show_PrinterDriver()
'��ܦL����X�ʵ{��

    '��ܪA��
    
On Error GoTo ErrMsg
'On Error Resume Next

    Dim cols ', 'objWMIService
    Set cols = objWMIService.ExecQuery("Select * from Win32_PrinterDriver")
    
    With Sg_PrinterDriver
        
        .Clear True
        .Redraw = False
        '.Visible = False

        .AddColumn "Name", "�X�ʵ{���W��"
        .AddColumn "Version", "����"
        .AddColumn "Description", "����"
        .AddColumn "DriverPath", "���|"
    
        If cols.Count > 0 Then
            
            For Each obj In cols
            
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Name"), obj.Name & " "
                .CellDetails .Rows, .ColumnIndex("Version"), Format(obj.Version, "#0.00") & " "
                .CellDetails .Rows, .ColumnIndex("Description"), obj.Description & " "
                .CellDetails .Rows, .ColumnIndex("DriverPath"), obj.DriverPath & " "
                DoEvents
            
            Next
        
        End If
    
        .AutoWidthColumn "Name"
        .AutoWidthColumn "Version"
        .AutoWidthColumn "Description"
        .AutoWidthColumn "DriverPath"
        
        .Redraw = True
        .Visible = True
    
    End With
    
    CmdBar1.Toolbar.Buttons.Item("PRINTERDRIVER:ADD").Enabled = True
    
Exit Sub

ErrMsg:

    With Sg_PrinterDriver
    
        .Clear True
        .Redraw = False
    
        .AddColumn "Name", "�X�ʵ{���W��"
        .AddColumn "Version", "����"
        .AddColumn "Description", "����"
        .AddColumn "DriverPath", "���|"
        
        .AddRow
        .CellDetails .Rows, .ColumnIndex("Name"), "�o�ӥ\��ݭn�t�Φb XP �~�䴩���"
        .CellDetails .Rows, .ColumnIndex("Version"), "���䴩"
        .CellDetails .Rows, .ColumnIndex("Description"), "���䴩"
        .CellDetails .Rows, .ColumnIndex("DriverPath"), "���䴩"
        .AutoWidthColumn "Name"
        .AutoWidthColumn "Version"
        .AutoWidthColumn "Description"
        .AutoWidthColumn "DriverPath"

        .Redraw = True
        .Visible = True
        
    End With
    
    CmdBar1.Toolbar.Buttons.Item("PRINTERDRIVER:ADD").Enabled = False
    CmdBar1.Toolbar.Buttons.Item("PRINTERDRIVER:DELETE").Enabled = False

End Sub

Sub Show_PrinterTCPIPPrinterPort()
'��ܦL��������s����

    '��ܪA��
On Error GoTo ErrMsg
'On Error Resume Next

    Dim cols ', 'objWMIService
    Set cols = objWMIService.ExecQuery("Select * from Win32_TCPIPPrinterPort")
    
'If Err.Number > 0 Then GoTo ErrMsg
    
    With Sg_PrinterTCPIPPort
        
        .Clear True
        .Redraw = False
        .Visible = False

        .AddColumn "Name", "�s����W��"
        .AddColumn "Address", "��}"
        .AddColumn "Protocol", "�q�T��w"
        .AddColumn "SNMP", "SNMP"

        If cols.Count > 0 Then
            
            For Each obj In cols
            
                .AddRow
                .CellDetails .Rows, .ColumnIndex("Name"), obj.Name & " "
                .CellDetails .Rows, .ColumnIndex("Address"), obj.HostAddress & " "
                .CellDetails .Rows, .ColumnIndex("Protocol"), IIf(obj.Protocol = 1, "RAW", "LPR") & " "
                .CellDetails .Rows, .ColumnIndex("SNMP"), IIf(obj.SNMPEnabled = True, "�ҥ�", "����")
                DoEvents
            
            Next
        
        End If
    
        .AutoWidthColumn "Name"
        .AutoWidthColumn "Address"
        .AutoWidthColumn "Protocol"
        .AutoWidthColumn "SNMP"
        
        .Redraw = True
        .Visible = True
    
    End With

    CmdBar1.Toolbar.Buttons.Item("PRINTERTCPIPPORT:ADD").Enabled = True

Exit Sub

ErrMsg:

    With Sg_PrinterTCPIPPort
    
        .Clear True
        .Redraw = False
    
        .AddColumn "Name", "�s����W��"
        .AddColumn "Address", "��}"
        .AddColumn "Protocol", "�q�T��w"
        .AddColumn "SNMP", "SNMP"
        
        .AddRow
        .CellDetails .Rows, .ColumnIndex("Name"), "�o�ӥ\��ݭn�t�Φb XP �~�䴩���"
        .CellDetails .Rows, .ColumnIndex("Address"), "���䴩"
        .CellDetails .Rows, .ColumnIndex("Protocol"), "���䴩"
        .CellDetails .Rows, .ColumnIndex("SNMP"), "���䴩"
        
        .AutoWidthColumn "Name"
        .AutoWidthColumn "Address"
        .AutoWidthColumn "Protocol"
        .AutoWidthColumn "SNMP"

        .Redraw = True
        .Visible = True
        
    End With
        
    CmdBar1.Toolbar.Buttons.Item("PRINTERTCPIPPORT:ADD").Enabled = False
    CmdBar1.Toolbar.Buttons.Item("PRINTERTCPIPPORT:DELETE").Enabled = False

End Sub



Private Sub Form_Resize()
On Error Resume Next
    
    Dim hTop As Long: hTop = CmdBar1.Height
    Dim totalTop As Long: totalTop = Me.ScaleHeight - hTop
    
    Sg_Printer.TOp = hTop
    Sg_Printer.Left = 0
    Sg_Printer.Width = Me.ScaleWidth
    Sg_Printer.Height = totalTop / 3
    
    Sg_PrinterDriver.TOp = hTop + totalTop * (1 / 3)
    Sg_PrinterDriver.Left = 0
    Sg_PrinterDriver.Width = Me.ScaleWidth
    Sg_PrinterDriver.Height = totalTop / 3
    
    Sg_PrinterTCPIPPort.TOp = hTop + totalTop * (2 / 3)
    Sg_PrinterTCPIPPort.Left = 0
    Sg_PrinterTCPIPPort.Width = Me.ScaleWidth
    Sg_PrinterTCPIPPort.Height = totalTop / 3
    
End Sub

Private Sub Sg_Printer_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)

With Sg_Printer
    
    If lRow > 0 Then
        CmdBar1.Toolbar.Buttons.Item("PRINTER:DELETE").Enabled = True
    Else
        CmdBar1.Toolbar.Buttons.Item("PRINTER:DELETE").Enabled = False
    End If
    
End With

End Sub

Private Sub Sg_PrinterDriver_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)

With Sg_PrinterDriver
    
    If .CellText(lRow, 2) <> "���䴩" Then
        CmdBar1.Toolbar.Buttons.Item("PRINTERDRIVER:DELETE").Enabled = True
    Else
        CmdBar1.Toolbar.Buttons.Item("PRINTERDRIVER:DELETE").Enabled = False
    End If
    
End With

End Sub

Private Sub Sg_PrinterTCPIPPort_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)

With Sg_PrinterTCPIPPort
    
    If .CellText(lRow, 2) <> "���䴩" Then
        CmdBar1.Toolbar.Buttons.Item("PRINTERTCPIPPORT:DELETE").Enabled = True
    Else
        CmdBar1.Toolbar.Buttons.Item("PRINTERTCPIPPORT:DELETE").Enabled = False
    End If
    
End With

End Sub
