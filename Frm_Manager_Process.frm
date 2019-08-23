VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{2210EC79-A724-4033-AAF4-790E2467C0E8}#1.0#0"; "vbalCmdBar6.ocx"
Begin VB.Form Frm_Manager_Process 
   Caption         =   "執行緒管理"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5670
   Icon            =   "Frm_Manager_Process.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4875
   ScaleWidth      =   5670
   StartUpPosition =   3  '系統預設值
   Begin vbalCmdBar6.vbalCommandBar CmdBar1 
      Align           =   1  '對齊表單上方
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   5670
      _ExtentX        =   10001
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
      Left            =   240
      TabIndex        =   0
      Top             =   420
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
      GroupBoxHintText=   "拖曳群組欄位標題至此，拖曳後請按下標題列排序可同時展開所有群組"
      HotTrack        =   -1  'True
      SelectionOutline=   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "處理中..."
      Height          =   555
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   2355
   End
End
Attribute VB_Name = "Frm_Manager_Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ConfigureButtons()
'建立設定工具列上的按鈕

On Error Resume Next

    Dim btn As cButton
    
    CmdBar1.ToolbarImageList = Frm_Main.vbalIml_Tools.hIml
    With CmdBar1.Buttons
    
        '建立 TOOLS 按鈕群
        Set btn = .Add("TOOLS:PROCESS:SPLIT", , , eSeparator)
        
        Set btn = .Add("PROCESS:CREATE", Frm_Main.vbalIml_Tools.ItemIndex("PROCESS:CREATE") - 1, "建立", , "在目標電腦建立一個背景執行緒啟動服務")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = True
        
        Set btn = .Add("PROCESS:TERMINATE", Frm_Main.vbalIml_Tools.ItemIndex("PROCESS:TERMINATE") - 1, "中斷", , "砍掉目標電腦的執行緒")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = True
        
        Set btn = .Add("TOOLS:PROCESS:SPLIT2", , , eSeparator)
    
        Set btn = .Add("PROCESS:REFRESH", Frm_Main.vbalIml_Tools.ItemIndex("REFRESH") - 1, "重新整理", , "重新列表最新狀態")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = True
        
    
    End With
End Sub

Private Sub ConfigureBars()


On Error Resume Next

    Dim bar As cCommandBar
    Dim Btns As cCommandBarButtons

    '工具列新面版
    With CmdBar1
            
        '建立一個新工具列
        Set bar = .CommandBars.Add("PROCESS", "Standard Buttons")
        Set Btns = bar.Buttons
             
            '建立
            Btns.Add .Buttons.Item("TOOLS:PROCESS:SPLIT")
            Btns.Add .Buttons.Item("PROCESS:CREATE")
            Btns.Add .Buttons.Item("PROCESS:TERMINATE")
            Btns.Add .Buttons.Item("TOOLS:PROCESS:SPLIT2")
            Btns.Add .Buttons.Item("PROCESS:REFRESH")
            
    End With

End Sub

Private Sub CmdBar1_ButtonClick(btn As vbalCmdBar6.cButton)

    Dim sRow As Long: sRow = Sg1.SelectedRow
    Dim sCol As Long: sCol = Sg1.ColumnIndex("Name")
    
    Dim tmp_Ret As String
    
    Sg1.Enabled = False

    Select Case btn.Key
    
        Case "PROCESS:CREATE"
            
            tmp_Ret = Action_Process("", "ACTION:CREATE")
            If tmp_Ret = "0" Then
                MsgBox "已在目標電腦上開啟該執行緒", vbInformation, "完成"
                Call Show_Process
            ElseIf tmp_Ret = "-1" Then
                MsgBox "動作取消", vbInformation, "訊息"
            Else
            End If
        
        Case "PROCESS:TERMINATE"
            
            If MsgBox("您確定要中止這個執行緒？", vbQuestion + vbYesNo, "提示") = vbYes Then
                If Sg1.SelectedRow > 0 Then Call Action_Process(Sg1.CellText(sRow, sCol), "ACTION:TERMINATE")
                 Call Show_Process
            End If
            
        Case "PROCESS:REFRESH": Show_Process
    
    End Select

    Sg1.Enabled = True
    
End Sub

Private Sub Form_Load()

On Error Resume Next
'If CmdBar1.Buttons.Count < 1 Then

    ConfigureButtons
    ConfigureBars
    CmdBar1.Toolbar = CmdBar1.CommandBars("PROCESS")

'End If

If Trim(Can_Process_Computer) = "" Then
    MsgBox "沒有從樹狀清單中選擇要管理的電腦", vbQuestion, "錯誤"
    Unload Me
    Exit Sub
End If

If WMI_Service_Create(Can_Process_Computer) = False Then
    MsgBox "無法連線此電腦", vbQuestion, "錯誤"
    Unload Me
    Exit Sub
End If
    
    
Me.Caption = Can_Process_Computer & " 上的執行緒清單"

Show_Process

Me.Show
End Sub


Sub Show_Process()
'顯示服務
    
On Error Resume Next


    Dim cols ', 'objWMIService
    Set cols = objWMIService.ExecQuery("Select * from Win32_Process")
    
    Sg1.Clear True
    Sg1.Redraw = False
    Sg1.Visible = False
    

    Sg1.AddColumn "Name", "名稱"
    Sg1.AddColumn "PID", "PID"
    Sg1.AddColumn "Path", "路徑"
    'Sg1.AddColumn "Detail", "說明"
    
    If cols.Count > 0 Then
        
        For Each obj In cols
            
            Add_Row _
                StrNullToSpace(obj.Name), _
                StrNullToSpace(obj.ProcessID), _
                StrNullToSpace(obj.ExecutablePath) ', _
                StrNullToSpace(obj.Description)
            DoEvents
        Next
    
    End If
    
    Sg1.Visible = True
    Sg1.Redraw = True

    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    Sg1.TOp = CmdBar1.Height
    Sg1.Left = 0
    Sg1.Width = Me.ScaleWidth
    Sg1.Height = Me.ScaleHeight - CmdBar1.Height
    
End Sub

Sub Add_Row(Field1 As String, Optional Field2 As String = "", Optional Field3 As String = "", Optional field4 As String = "")

With Sg1

    .AddRow
    .CellDetails .Rows, .ColumnIndex("Name"), Field1
    .CellDetails .Rows, .ColumnIndex("PID"), Field2
    .CellDetails .Rows, .ColumnIndex("Path"), Field3
    '.CellDetails .Rows, .ColumnIndex("Detail"), field4

    .AutoWidthColumn .ColumnIndex("Name")
    .AutoWidthColumn .ColumnIndex("PID")
    .AutoWidthColumn .ColumnIndex("Path")
    '.AutoWidthColumn .ColumnIndex("Detail")

End With
End Sub

Private Sub Sg1_ColumnClick(ByVal lCol As Long)
    
    '排序
    vbalGrid_Sort Me.Sg1, lCol
    ExGroup Me.Sg1


End Sub

