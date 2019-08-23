VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{2210EC79-A724-4033-AAF4-790E2467C0E8}#1.0#0"; "vbalCmdBar6.ocx"
Begin VB.Form Frm_Manager_Software 
   Caption         =   "軟體管理"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   Icon            =   "Frm_Manager_Software.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4860
   ScaleWidth      =   5295
   StartUpPosition =   3  '系統預設值
   Begin vbalCmdBar6.vbalCommandBar CmdBar1 
      Align           =   1  '對齊表單上方
      Height          =   315
      Left            =   0
      Top             =   0
      Width           =   5295
      _ExtentX        =   9340
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
      GroupBoxHintText=   "拖曳群組欄位標題至此，拖曳後請按下標題列排序可同時展開所有群組"
      HotTrack        =   -1  'True
      SelectionOutline=   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "處理中..."
      Height          =   555
      Left            =   300
      TabIndex        =   1
      Top             =   660
      Width           =   2355
   End
End
Attribute VB_Name = "Frm_Manager_Software"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub ConfigureButtons()
'建立設定工具列上的按鈕

On Error Resume Next

    Dim btn As cButton
    
    CmdBar1.ToolbarImageList = Frm_Main.vbalIml_Tools.hIml
    With CmdBar1.Buttons
    
        '建立 TOOLS 按鈕群
        Set btn = .Add("SOFTWARE:SPLIT1", , , eSeparator)
        
        Set btn = .Add("SOFTWARE:ACTION:ADD", Frm_Main.vbalIml_Tools.ItemIndex("SOFTWARE:ADD") - 1, "安裝", , "安裝一個軟體到目標電腦上")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = False
        
        Set btn = .Add("SOFTWARE:ACTION:DELETE", Frm_Main.vbalIml_Tools.ItemIndex("SOFTWARE:DELETE") - 1, "移除", , "從目標電腦上移除這套軟體")
        btn.ShowCaptionInToolbar = True
        btn.Enabled = False
        
        Set btn = .Add("REFRESH", Frm_Main.vbalIml_Tools.ItemIndex("REFRESH") - 1, "重新整理", , "重新列表最新狀態")
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
        Set bar = .CommandBars.Add("SOFTWARE", "Standard Buttons")
        Set Btns = bar.Buttons
             
            '建立
            Btns.Add .Buttons.Item("SOFTWARE:SPLIT1")
            Btns.Add .Buttons.Item("SOFTWARE:ACTION:ADD")
            Btns.Add .Buttons.Item("SOFTWARE:ACTION:DELETE")
            Btns.Add .Buttons.Item("REFRESH")
            
    End With

End Sub

Private Sub CmdBar1_ButtonClick(btn As vbalCmdBar6.cButton)

    If Sg1.Rows < 1 Then Exit Sub

    Dim sRow As Long: sRow = Sg1.SelectedRow
    Dim sCol As Long: sCol = Sg1.ColumnIndex("Name")

    Sg1.Enabled = False

    Select Case btn.Key
    
        Case "SOFTWARE:ACTION:ADD":  MsgBox Action_Software("", "ACTION:ADD")
        
        Case "SOFTWARE:ACTION:DELETE":
            
                If Action_Software(Sg1.CellText(sRow, sCol), "ACTION:DELETE") = "0" Then
                    Sg1.RemoveRow sRow
                    Show_Software
                Else
                    MsgBox "移除失敗", vbInformation, "錯誤"
                End If
            
        
        Case "REFRESH": Show_Software
        
    End Select
    
    Sg1.Enabled = True

End Sub

Private Sub Form_Load()


'On Error Resume Next
'If CmdBar1.Buttons.Count < 1 Then

    ConfigureButtons
    ConfigureBars
    CmdBar1.Toolbar = CmdBar1.CommandBars("SOFTWARE")


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
    
    


Me.Caption = Can_Process_Computer & " 上的軟體清單"
Show_Software

Me.Show

End Sub

Sub Show_Software()
'顯示服務
    
On Error Resume Next

    
    Dim cols, obj
    Set cols = objWMIService.ExecQuery("Select * from Win32_Product")
    
    With Sg1
        
        .Clear True
        
        .Redraw = False
        .Visible = False
    
        .AddColumn "Name", "名稱"
        .AddColumn "Version", "版本"
        .AddColumn "Path", "路徑"
        .AddColumn "InstallDate", "安裝日期"
        
        If cols.Count > 0 Then
            
            For Each obj In cols
                Add_Row _
                    StrNullToSpace(obj.Name), _
                    StrNullToSpace(obj.Version), _
                    StrNullToSpace(obj.InstallLocation), _
                    Format(StrNullToSpace(obj.InstallDate), "####/##/##")
                        
                DoEvents
            Next
        
        End If
        
        vbalGrid_Sort Me.Sg1, 1
        ExGroup Me.Sg1
    
        .Redraw = True
        .Visible = True
    
    End With
    
End Sub

Sub Add_Row(Field1 As String, Optional Field2 As String = "", Optional Field3 As String = "", Optional field4 As String = "", Optional field5 As String = "")

With Sg1

    .AddRow
    .CellDetails .Rows, .ColumnIndex("Name"), Field1
    .CellDetails .Rows, .ColumnIndex("Version"), Field2
    .CellDetails .Rows, .ColumnIndex("Path"), Field3
    .CellDetails .Rows, .ColumnIndex("InstallDate"), field4

    .AutoWidthColumn .ColumnIndex("Name")
    .AutoWidthColumn .ColumnIndex("Version")
    .AutoWidthColumn .ColumnIndex("Path")
    .AutoWidthColumn .ColumnIndex("InstallDate")
    
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
    '排序
    vbalGrid_Sort Me.Sg1, lCol
    ExGroup Me.Sg1
End Sub

Private Sub Sg1_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)

If lRow > 0 Then
    CmdBar1.Toolbar.Buttons.Item("SOFTWARE:ACTION:DELETE").Enabled = True
Else
    CmdBar1.Toolbar.Buttons.Item("SOFTWARE:ACTION:DELETE").Enabled = False
End If

End Sub
