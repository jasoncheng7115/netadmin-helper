VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Object = "{CA5A8E1E-C861-4345-8FF8-EF0A27CD4236}#1.1#0"; "vbalTreeView6.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "Mswinsck.ocx"
Object = "{2210EC79-A724-4033-AAF4-790E2467C0E8}#1.0#0"; "vbalCmdBar6.ocx"
Object = "{4F11FEBA-BBC2-4FB6-A3D3-AA5B5BA087F4}#1.0#0"; "vbalSbar6.ocx"
Begin VB.Form Frm_Main 
   Caption         =   "AD 網域電腦管理幫手 測試版"
   ClientHeight    =   8505
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9600
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   9600
   StartUpPosition =   3  '系統預設值
   Begin vbalCmdBar6.vbalCommandBar cmdBar 
      Align           =   1  '對齊表單上方
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
      Alignment       =   2  '置中對齊
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
      Alignment       =   2  '置中對齊
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
      Align           =   2  '對齊表單下方
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
      Align           =   1  '對齊表單上方
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
      GroupBoxHintText=   "拖曳群組欄位標題至此，拖曳後請按下標題列排序可同時展開所有群組"
      HotTrack        =   -1  'True
      SelectionOutline=   -1  'True
   End
   Begin vbalCmdBar6.vbalCommandBar cmdBar 
      Align           =   1  '對齊表單上方
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
      Appearance      =   0  '平面
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

'判斷是否按下滑鼠左鍵 (準備調整大小)
Private mbResizing As Boolean

'網域電腦清單
Public tmp_Computers As Collection

'停止列表旗標
Public Stop_Computers_List As Boolean

'連線測試中斷時所取得的電腦編號
Public List_Last As Long

Private Sub cmdBar_ButtonClick(index As Integer, btn As vbalCmdBar6.cButton)
           
On Error Resume Next
           
    Select Case btn.Key
        '連線測試
        Case "CONNECT_CHECK"
        
            With cmdBar(0).Toolbar.Buttons
                .Item("CONNECT_CHECK").Enabled = False
                .Item("CONNECT_CHECK_COUNTINUE").Enabled = False
                .Item("CONNECT_CHECK_STOP").Enabled = True
            End With
            
            '初始化各變數
            Total_Connected_Computers = 0
            List_Last = 1
            Stop_Computers_List = False
            SGrid_Computers_List
            
        '連線繼續測試
        Case "CONNECT_CHECK_COUNTINUE"
            
            With cmdBar(0).Toolbar.Buttons
                .Item("CONNECT_CHECK").Enabled = False
                .Item("CONNECT_CHECK_COUNTINUE").Enabled = False
                .Item("CONNECT_CHECK_STOP").Enabled = True
            End With
            
            '從上次中斷處開始
            List_Last = List_Last + 1
            Stop_Computers_List = False
            SGrid_Computers_List
                
        '連線測試停止
        Case "CONNECT_CHECK_STOP"
            
            With cmdBar(0).Toolbar.Buttons
                .Item("CONNECT_CHECK").Enabled = True
                .Item("CONNECT_CHECK_COUNTINUE").Enabled = True
                .Item("CONNECT_CHECK_STOP").Enabled = False
            End With
                       
            Stop_Computers_List = True
            Stop_Computer_List
        
        '本機模式
        Case "CONNECT_LOCAL": Get_Computers_From_Local
            
            
        
        '顯示設定
        Case "SHOW_SETUP": Frm_Show_Setup.Show
        
        '變更網域
        Case "SETUP_DOMAINNAME_CHANGE"
        
            AD_Domain_Name = Txt_DomainName.Text
            Get_Computers_From_AD
        
        '變更連線確認用 PORT
        Case "CONNECT_CHECK_PORT_SET_CHANGE"
            
            If IsNumeric(Txt_Port.Text) = False Then MsgBox "埠號必需是數字", vbInformation, "錯誤": Exit Sub
            If Txt_Port.Text < 1 Or Txt_Port.Text > 65536 Then MsgBox "埠號必需介於 1~65536 之間", vbInformation, "錯誤": Exit Sub
            Port_Ping = Txt_Port.Text
            
        '匯出到 Excel
        Case "FILE:EXPORT_TO_EXCEL": Call ExportToExcel(Me.Sg1)
        
        '匯出到 CSV
        Case "FILE:EXPORT_TO_CSV": Call ExportToCSV(Me.Sg1)
        
        '離開程式
        Case "FILE:EXIT":  Unload Me
        
        '詳細資訊
        Case "MANAGER:INFORMATION": Show_Computer_Information Tvw1.SelectedItem
        
        '管理服務
        Case "MANAGER:SERVICE":  Load Frm_Manager_Service
        
        '管理執行緒
        Case "MANAGER:PROCESS": Load Frm_Manager_Process
        
        '管理印表機
        Case "MANAGER:PRINTER": Load Frm_Manager_Printer
        
        '管理軟體
        Case "MANAGER:SOFTWARE": Load Frm_Manager_Software
        
        '重新開機
        Case "MANAGER:SHUTDOWN:REBOOT": Action_Shutdown "", "ACTION:REBOOT"
        
        '關機
        Case "MANAGER:SHUTDOWN:SHUTDOWN": Action_Shutdown "", "ACTION:SHUTDOWN"
                
        '變更本機管理員密碼
        Case "MANAGER:LOCAL:CHANGE_ADMIN_PWD": Action_Local_Admin "", "ACTION:CHANGE_ADMIN_PASSWORD"
        
        '批次變更本機管理員密碼
        Case "BATCH:LOCAL:CHANGE_ADMIN_PWD": Batch_Action_Local_Admin "", "ACTION:CHANGE_ADMIN_PASSWORD", Get_Checked_Computers
        
        '批次管理開始服務
        Case "BATCH:SERVICE:START": Batch_Action_Service "ACTION:START", Get_Checked_Computers
        
        '批次管理停止服務
        Case "BATCH:SERVICE:STOP": Batch_Action_Service "ACTION:STOP", Get_Checked_Computers
        
        '批次開啟執行緒
        Case "BATCH:PROCESS:CREATE": Batch_Action_Process "ACTION:CREATE", Get_Checked_Computers
         
        '批次中斷執行緒
        Case "BATCH:PROCESS:TERMINATE": Batch_Action_Process "ACTION:TERMINATE", Get_Checked_Computers
        
        '關於這個程式
        Case "ABOUT:PROGRAM": Frm_About.Show
        
        '關於作者
        Case "ABOUT:MAKER": MsgBox "Jason Cheng", vbInformation, "Design By"
        
    
    End Select
        
End Sub

Private Function Get_Checked_Computers() As String
'取得勾選的電腦名稱
    
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
'建立設定工具列上的按鈕

'On Error Resume Next

    Dim btn As cButton
    
    cmdBar(0).ToolbarImageList = vbalIml2.hIml
    With cmdBar(0).Buttons
        
            '建立 TOOLS 按鈕群
            Set btn = .Add("TOOLS:CONNECT:SPLIT", , , eSeparator)
            
            Set btn = .Add("CONNECT_CHECK", vbalIml2.ItemIndex("Check_Connect") - 1, "連線確認", , "確認所有電腦是否有連線中")
            btn.ShowCaptionInToolbar = True
            btn.Enabled = True
            
            Set btn = .Add("CONNECT_CHECK_COUNTINUE", vbalIml2.ItemIndex("Check_Connect_Countinue") - 1, "繼續確認", , "從上次中斷處往下確認所有電腦是否有連線中")
            btn.ShowCaptionInToolbar = True
            btn.Enabled = False
            
            Set btn = .Add("CONNECT_CHECK_STOP", vbalIml2.ItemIndex("Check_Connect_Stop") - 1, "停止檢測", , "中斷連線測試工作")
            btn.ShowCaptionInToolbar = True
            btn.Enabled = False
        
        
            '建立 TOOLS 按鈕群
            Set btn = .Add("TOOLS:SHOW_SETUP:SPLIT", , , eSeparator)
            Set btn = .Add("SHOW_SETUP", vbalIml2.ItemIndex("SHOW_SETUP") - 1, "顯示項目設定", , "選擇電腦資訊列表時所要觀看的資訊")
            btn.ShowCaptionInToolbar = True
            btn.Enabled = True
        
    End With
    
    '網域設定
    'cmdBar(2).MenuImageList = vbalImlMenu.hIml
    With cmdBar(2).Buttons
            
            Set btn = .Add("TOOLS:SETUP_DOMAINNAME:SPLIT", , , eSeparator)
            
            Set btn = .Add("SETUP_DOMAINNAME_LABEL", , "網域設定：")
            btn.ShowCaptionInToolbar = True: btn.Enabled = False
            
            Set btn = .Add("SETUP_DOMAINNAME", , "網域", ePanel, "設定您要管理的 AD 網域")
            btn.PanelWidth = 90: btn.PanelControl = Txt_DomainName
        
            Set btn = .Add("SETUP_DOMAINNAME_CHANGE", , "列出電腦", , "列出這個網域的所有電腦清單")
            btn.ShowCaptionInToolbar = True
        
            '本機模式
            Set btn = .Add("TOOLS:SETUP_DOMAINNAME:SPLIT2", , , eSeparator)
            Set btn = .Add("SETUP_DOMAINNAME_LOCAL", , "本機模式", , "直接瀏覽本機電腦")
            btn.ShowCaptionInToolbar = True
            btn.Enabled = True
        
            Set btn = .Add("TOOLS:CONNECT_CHECK_PORT_SET:SPLIT", , , eSeparator)
            
            Set btn = .Add("CONNECT_CHECK_PORT_SET_LABEL", , "確認連線所使用埠號：")
            btn.ShowCaptionInToolbar = True: btn.Enabled = False
            
            Set btn = .Add("CONNECT_CHECK_PORT_SET_TEXTBOX", , , ePanel, "設定連線確認時所用的 PORT")
            btn.PanelWidth = 45: btn.PanelControl = Txt_Port
            
            Set btn = .Add("CONNECT_CHECK_PORT_SET_CHANGE", , "變更")
            btn.ShowCaptionInToolbar = True
            
 
    End With
    
    
    cmdBar(1).MenuImageList = vbalImlMenu.hIml
    With cmdBar(1).Buttons
               
        '最上層
        Set btn = .Add("FILE", , "檔案(&F)")
        btn.ShowCaptionInToolbar = True
        Set btn = .Add("VIEW", , "檢視(&V)")
        btn.ShowCaptionInToolbar = True
        Set btn = .Add("OPTION", , "選項(&O)")
        btn.ShowCaptionInToolbar = True
        Set btn = .Add("MANAGER", , "管理(&M)")
        btn.ShowCaptionInToolbar = True
        Set btn = .Add("BATCH", , "批次(&B)")
        btn.ShowCaptionInToolbar = True
        Set btn = .Add("SETUP", , "設定(&S)")
        btn.Enabled = False
        btn.ShowCaptionInToolbar = True
        Set btn = .Add("ABOUT", , "關於(&A)")
        btn.ShowCaptionInToolbar = True
        
            '建立 FILE 按鈕群
            Set btn = .Add("FILE:EXPORT_TO_EXCEL", vbalImlMenu.ItemIndex("TOEXCEL") - 1, "匯出資料到 Excel", , "將 Excel 開啟並將資料匯出")
            Set btn = .Add("FILE:EXPORT_TO_CSV", vbalImlMenu.ItemIndex("TONOTEPAD") - 1, "匯出資料到 CSV", , "將資料匯出到 CSV")
            Set btn = .Add("FILE:SPLIT1", , , eSeparator)
            btn.Visible = False
            Set btn = .Add("FILE:EXIT", , "離開", , "結束程式", vbKeyF12, 0)
            btn.Visible = False
        
            '建立 MANAGER 按鈕群
            Set btn = .Add("MANAGER:INFORMATION", vbalImlMenu.ItemIndex("INFORMATION") - 1, "詳細資訊")
            Set btn = .Add("MANAGER:SPLIT1", , , eSeparator)
            Set btn = .Add("MANAGER:SERVICE", vbalImlMenu.ItemIndex("SERVICE") - 1, "服務")
            Set btn = .Add("MANAGER:PROCESS", vbalImlMenu.ItemIndex("PROCESS") - 1, "執行緒")
            Set btn = .Add("MANAGER:PRINTER", vbalImlMenu.ItemIndex("PRINTER") - 1, "印表機")
            Set btn = .Add("MANAGER:SOFTWARE", vbalImlMenu.ItemIndex("SOFTWARE") - 1, "軟體")
            Set btn = .Add("MANAGER:SPLIT2", , , eSeparator)
            Set btn = .Add("MANAGER:SHUTDOWN", vbalImlMenu.ItemIndex("SHUTDOWN") - 1, "關機")
                Set btn = .Add("MANAGER:SHUTDOWN:REBOOT", vbalImlMenu.ItemIndex("REBOOT") - 1, "重新開機")
                Set btn = .Add("MANAGER:SHUTDOWN:SHUTDOWN", vbalImlMenu.ItemIndex("SHUTDOWN") - 1, "關機")
            Set btn = .Add("MANAGER:LOCAL", vbalImlMenu.ItemIndex("LOCAL") - 1, "本機")
                Set btn = .Add("MANAGER:LOCAL:CHANGE_ADMIN_PWD", vbalImlMenu.ItemIndex("CHANGE_PWD") - 1, "變更本機 Administrator 密碼")
            
            '建立 BATCH 按鈕群
            Set btn = .Add("BATCH:SERVICE", vbalImlMenu.ItemIndex("SERVICE") - 1, "服務")
                Set btn = .Add("BATCH:SERVICE:START", vbalImlMenu.ItemIndex("START") - 1, "啟用")
                Set btn = .Add("BATCH:SERVICE:STOP", vbalImlMenu.ItemIndex("STOP") - 1, "停止")
                
            Set btn = .Add("BATCH:PROCESS", vbalImlMenu.ItemIndex("PROCESS") - 1, "執行緒")
                Set btn = .Add("BATCH:PROCESS:CREATE", vbalImlMenu.ItemIndex("CREATE") - 1, "建立")
                Set btn = .Add("BATCH:PROCESS:TERMINATE", vbalImlMenu.ItemIndex("TERMINATE") - 1, "中斷")
                
            Set btn = .Add("BATCH:SPLIT1", , , eSeparator)
            Set btn = .Add("BATCH:LOCAL", vbalImlMenu.ItemIndex("LOCAL") - 1, "本機")
                Set btn = .Add("BATCH:LOCAL:CHANGE_ADMIN_PWD", vbalImlMenu.ItemIndex("CHANGE_PWD") - 1, "變更本機 Administrator 密碼")
        
        
            '建立 ABOUT 按鈕群
            Set btn = .Add("ABOUT:PROGRAM", , "關於", , "關於這個程式...")
            Set btn = .Add("ABOUT:SPLIT1", , , eSeparator)
            Set btn = .Add("ABOUT:MAKER", vbalImlMenu.ItemIndex("INFORMATION") - 1, "作者", , "作者資訊")
        
    End With
    
 
End Sub

Private Sub ConfigureBars()

'On Error Resume Next

    Dim bar, bar_1 As cCommandBar
    Dim Btns, Btns_1 As cCommandBarButtons


    '工具列新面版
    With cmdBar(0)
            
        '建立一個新工具列
        Set bar = .CommandBars.Add("STANDARD", "Standard Buttons")
        Set Btns = bar.Buttons
             
            '建立
            Btns.Add .Buttons.Item("TOOLS:CONNECT:SPLIT")
            Btns.Add .Buttons.Item("CONNECT_CHECK")
            Btns.Add .Buttons.Item("CONNECT_CHECK_COUNTINUE")
            Btns.Add .Buttons.Item("CONNECT_CHECK_STOP")
            Btns.Add .Buttons.Item("TOOLS:SHOW_SETUP:SPLIT")
            Btns.Add .Buttons.Item("SHOW_SETUP")
            
    End With

    '第二行網域設定
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
    
    '功能表新面版
    With cmdBar(1)
    
        '建立頂層功能表
        Set bar = .CommandBars.Add("TOPMENU", "Menu")
        Set Btns = bar.Buttons
        Btns.Add .Buttons.Item("FILE")
        Btns.Add .Buttons.Item("VIEW")
        Btns.Add .Buttons.Item("OPTION")
        Btns.Add .Buttons.Item("MANAGER")
        Btns.Add .Buttons.Item("BATCH")
        Btns.Add .Buttons.Item("SETUP")
        Btns.Add .Buttons.Item("ABOUT")
        
            '建立一個子功能表 FILE
            Set bar = .CommandBars.Add("FILEMENU", "FILE")
            Set Btns = bar.Buttons
            Btns.Add .Buttons.Item("FILE:EXPORT_TO_EXCEL")
            Btns.Add .Buttons.Item("FILE:EXPORT_TO_CSV")
            Btns.Add .Buttons.Item("FILE:SPLIT1")
            Btns.Add .Buttons.Item("FILE:EXIT")
            .Buttons.Item("FILE").bar = bar
      
            '建立一個子功能表 MANAGER
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
                  
                      '建立一個子2功能表 MANAGER:SHUTDOWN
                      Set bar = .CommandBars.Add("SHUTDOWN:MANAGERMENU", "SHUTDOWN")
                      Set Btns = bar.Buttons
                      Btns.Add .Buttons.Item("MANAGER:SHUTDOWN:REBOOT")
                      Btns.Add .Buttons.Item("MANAGER:SHUTDOWN:SHUTDOWN")
                      .Buttons.Item("MANAGER:SHUTDOWN").bar = bar
            
                      '建立一個子2功能表 MANAGER:SHUTDOWN
                      Set bar = .CommandBars.Add("SHUTDOWN:LOCAL", "LOCAL")
                      Set Btns = bar.Buttons
                      Btns.Add .Buttons.Item("MANAGER:LOCAL:CHANGE_ADMIN_PWD")
                      .Buttons.Item("MANAGER:LOCAL").bar = bar
      
            
            '建立一個子功能表 BATCH
            Set bar = .CommandBars.Add("BATCHMENU", "BACTH")
                  
                  '建立子選單 SERVICE
                  Set Btns = bar.Buttons
                  Btns.Add .Buttons.Item("BATCH:SERVICE")
                  .Buttons.Item("BATCH").bar = bar
                  
                  '建立子選單 PROCESS
                  Set Btns = bar.Buttons
                  Btns.Add .Buttons.Item("BATCH:PROCESS")
                  .Buttons.Item("BATCH").bar = bar
                  
                  Btns.Add .Buttons.Item("BATCH:SPLIT1")
                  
                  '建立一個子選單 LOCAL
                  Set Btns = bar.Buttons
                  Btns.Add .Buttons.Item("BATCH:LOCAL")
                  .Buttons.Item("BATCH").bar = bar
                  
                          '建立一個子2功能表 BATCH:SERVICE
                          Set bar = .CommandBars.Add("BATCH:SERVICE", "SERVICE")
                          Set Btns = bar.Buttons
                          Btns.Add .Buttons.Item("BATCH:SERVICE:START")
                          Btns.Add .Buttons.Item("BATCH:SERVICE:STOP")
                          .Buttons.Item("BATCH:SERVICE").bar = bar
                
                          '建立一個子2功能表 BATCH:PROCESS
                          Set bar = .CommandBars.Add("BATCH:PROCESS", "PROCESS")
                          Set Btns = bar.Buttons
                          Btns.Add .Buttons.Item("BATCH:PROCESS:CREATE")
                          Btns.Add .Buttons.Item("BATCH:PROCESS:TERMINATE")
                          .Buttons.Item("BATCH:PROCESS").bar = bar
                
                          '建立一個子2功能表 BATCH:SHUTDOWN
                          Set bar = .CommandBars.Add("BATCH:LOCAL", "LOCAL")
                          Set Btns = bar.Buttons
                          Btns.Add .Buttons.Item("BATCH:LOCAL:CHANGE_ADMIN_PWD")
                          .Buttons.Item("BATCH:LOCAL").bar = bar
            
      
            '建立一個子功能表 ABOUT
            Set bar = .CommandBars.Add("ABOUTMENU", "ABOUT")
                Set Btns = bar.Buttons
                Btns.Add .Buttons.Item("ABOUT:PROGRAM")
                Btns.Add .Buttons.Item("ABOUT:SPLIT1")
                Btns.Add .Buttons.Item("ABOUT:MAKER")
                .Buttons.Item("ABOUT").bar = bar
      
    End With
End Sub

Private Sub cmdBar_RequestNewInstance(index As Integer, ctl As Object)
   
   '顯示這個功能表的子選項
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

'預設網域
AD_Domain_Name = "test.com.tw"

'預設 PORT
Port_Ping = "445"

'建立功能表、工具列與按鈕群
ConfigureButtons
ConfigureBars

    '主功能表
    cmdBar(1).MainMenu = True
    cmdBar(1).Toolbar = cmdBar(1).CommandBars("TOPMENU")
    
    '主要工具列
    cmdBar(0).Toolbar = cmdBar(0).CommandBars("STANDARD")
    
    '網域設定工具列
    cmdBar(2).Toolbar = cmdBar(2).CommandBars("DOMAINNAME_SETUP")


'增加狀態列面版
vbalSBar1.AddPanel , , , , , True
vbalSBar1.AddPanel , "連線中 : 0  ", , , 80, , True

'設定移動分割視窗時游標樣式
Label1.MousePointer = vbSizeWE

'列出電腦
Get_Computers_From_AD

End Sub

Private Sub Get_Computers_From_AD()
'從 AD 提取所有電腦清單

    Show_Msg "正在從 AD 取得電腦清單.."
    
    '取得所有電腦
    Set tmp_Computers = ADSI_Computer_List()


    If canUseAD = True Then
        Add_Computer_To_TreeView
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK").Enabled = True
        Show_Msg "從 AD 取得電腦清單完畢"
    Else
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK").Enabled = False
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK_COUNTINUE").Enabled = False
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK_STOP").Enabled = False
        Show_Msg "無法從 AD 取得電腦清單"
        
        If MsgBox("無法從 AD 取得電腦清單，是否要啟動本機模式？", vbQuestion + vbYesNo, "提示") = vbYes Then
            Get_Computers_From_Local
        End If
    End If
    

End Sub

Private Sub Add_Computer_To_TreeView()
'加入電腦到 TreeView

    Show_Msg "正在將電腦加入到樹狀清單.."
    
    Tvw1.ImageList = vbalIml1
    Tvw1.Nodes.Clear
    Tvw1.CheckBoxes = True

    Dim NodeRoot As cTreeViewNode
    Dim NodeChildren As cTreeViewNodes
    Dim NodeSub  As cTreeViewNode
    Dim Icon_Root As Long
    Dim Icon_Sub As Long
    
    Icon_Root = vbalIml1.ItemIndex("root") - 1
    
    '加入根目錄
    Set NodeRoot = Tvw1.Nodes.Add(, etvwFirst, "root", UCase(AD_Domain_Name), Icon_Root)
    Set NodeChildren = NodeRoot.Children
    
    Icon_Sub = vbalIml1.ItemIndex("Disconnect") - 1
    
    '加入
    Dim i As Long
    For i = 1 To tmp_Computers.Count
        Set NodeSub = NodeChildren.Add(, etvwChild, Format(i, "000#") & CStr(tmp_Computers(i)), CStr(tmp_Computers(i)), Icon_Sub)
        NodeSub.Tag = "0"
    Next

    NodeRoot.Expanded = True

    Show_Msg "電腦加入樹狀清單完畢"

End Sub

Private Sub Get_Computers_From_Local()
'提取本機電腦

    Show_Msg "正在從 Local 取得電腦清單.."
    
    '列舉環境變數
    'Dim aa, i
    'Do
    '    i = i + 1: aa = aa & Environ(i) & vbCrLf
    '    DoEvents
    'Loop Until Environ(i) = ""
    
    '取得電腦名稱
    Tvw1.Nodes.Clear
    Tvw1.CheckBoxes = False
    Tvw1.Nodes.Add , , "0001" & Environ("COMPUTERNAME"), Environ("COMPUTERNAME")

    cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK").Enabled = False
    cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK_COUNTINUE").Enabled = False
    cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK_STOP").Enabled = False
    
    Show_Msg "電腦清單已取得"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If MsgBox("您確定要離開嗎？", vbQuestion + vbYesNo, "提示") = vbNo Then
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

    '釋放所有 Socket
    Dim i
    For i = Sck_List.LBound + 1 To Sck_List.UBound
        Unload Sck_List(i)
    Next

    '釋放掉所有工具列與按鈕
    Dim cCombar As Integer, cComdBars As Integer
    For cCombar = cmdBar.LBound To cmdBar.UBound
        For cComdBars = 1 To cmdBar(cCombar).CommandBars.Count
            cmdBar(cCombar).CommandBars(cComdBars).Buttons.Clear
        Next
    Next
        
    FreeLibrary m_hMod
    
    '釋放所有表單
    Dim the_Frms As Form
    For Each the_Frms In Forms
        Unload the_Frms
    Next

If Err.Number <> 0 Then
    'MsgBox "發生錯誤 : " & Err.Description & "(" & Err.Number & ")", vbCritical, "錯誤"
End If


End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    '準備調整大小
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

  '按下滑鼠左鍵並移動時, 自動調整各控制項大小
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
  '停止調整大小
    mbResizing = False

End Sub


Private Sub Sck_Check_One_Connect()

If Sck_Check_One.State = sckConnected Then
    SGrid_Computer_List "正常連線"
Else
    SGrid_Computer_List "失敗"
End If

End Sub

Private Sub Sck_Check_One_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

On Error Resume Next
    
    Select Case Number
       
        '沒有找到主機
        Case sckHostNotFound, sckHostNotFoundTryAgain
            
            Call SGrid_Computer_List("沒有找到主機")
        
        '逾時
        Case sckTimedout
        
            Call SGrid_Computer_List("連線逾時")
        
        
        Case Else
        
            Call SGrid_Computer_List("連線失敗")
        
    End Select

End Sub

Function Check_Show_Setup(Src_Item As String) As Integer

Check_Show_Setup = CInt(GetSetting(App.Title, "Show_Setup", Src_Item, "0"))

End Function

Function SGrid_Computer_List(Src_State As String)
'列出單機的內容

On Error Resume Next
    
    Dim tmp_BackColor As Long
    
    Tvw1.Enabled = False
    tmp_BackColor = Tvw1.BackColor
    Tvw1.BackColor = &H8000000F
    
    '電腦名稱與 IP
    Dim strComputer As String:   strComputer = Tvw1.SelectedItem.Text
    Dim strComputer_IP As String: strComputer_IP = Sck_Check_One.RemoteHostIP
    
    '電腦名稱前的編號索引
    Dim strIndex As Integer: strIndex = CInt(Left(Tvw1.SelectedItem.Key, 4))
                       
                       
    Sck_Check_One.Close
    DoEvents
    
    '暫存開始處理時間
    Dim tmr_Start As Single: tmr_Start = Timer
    
    
    With cmdBar(0).Toolbar.Buttons
        .Item("CONNECT_CHECK").Enabled = True
        .Item("CONNECT_CHECK_COUNTINUE").Enabled = False
        .Item("CONNECT_CHECK_STOP").Enabled = False
    End With
    
    
    
    With Sg1
    
        '.Redraw = False
        
        .Clear True
        
        .AddColumn , "項目"
        .AddColumn , "內容"
        .AddColumn , "類別"
        
     
         If Src_State = "正常連線" Then
            
            Show_Msg "正在與目標電腦建立連線中，請稍候..."
            
            '無法連線
            If WMI_Service_Create(strComputer) = False Then
              
                SGrid_Computer_List_AddRow "可能沒有權限或是網路發生問題", Src_State, "錯誤"
                  Show_Msg "無法連線到" & strComputer & "，您可能沒有權限或是網路發生問題"
            
                'Call Show_Tvw_Computer_State(strIndex, "離線")
            
            '可連線
            Else
              
                'Call Show_Tvw_Computer_State(strIndex, "正常連線")
              
                Show_Msg "讀取電腦名稱中，請稍候..."
                    SGrid_Computer_List_AddRow "電腦名稱", strComputer, "一般"
                
                Show_Msg "讀取網路位址中，請稍候..."
                    SGrid_Computer_List_AddRow "電腦網路位址", strComputer_IP, "一般"
                
                Show_Msg "讀取電腦登入使用者中，請稍候..."
                    WMI_Computer_Login_UserName strComputer
                
                Show_Msg "讀取電腦作業系統資訊中，請稍候..."
                    WMI_Computer_OperatingSystem strComputer
    
                If Check_Show_Setup("本機群組") = 1 Then
                    Show_Msg "讀取電腦本機群組中，請稍候..."
                    ADSI_Computer_Group strComputer
                End If
                
                If Check_Show_Setup("本機使用者") = 1 Then
                    Show_Msg "讀取電腦本機使用者中，請稍候..."
                    ADSI_Computer_User strComputer
                End If
    
                If Check_Show_Setup("顯示模式") = 1 Then
                    Show_Msg "讀取電腦可用顯示模式資訊中，請稍候..."
                    CMI_Computer_VideoControllerResolution strComputer
                End If
                
                If Check_Show_Setup("解碼器") = 1 Then
                    Show_Msg "讀取電腦解碼器資訊中，請稍候..."
                    WMI_Computer_CodecFile strComputer
                End If
                
                If Check_Show_Setup("顯示器") = 1 Then
                    Show_Msg "讀取電腦顯示器資訊中，請稍候..."
                    WMI_Computer_DesktopMonitor strComputer
                End If
                
                If Check_Show_Setup("顯示卡") = 1 Then
                    Show_Msg "讀取電腦顯示卡資訊中，請稍候..."
                    WMI_Computer_Displays strComputer
                End If
                    
                If Check_Show_Setup("連接埠") = 1 Then
                    Show_Msg "讀取電腦連接埠資訊中，請稍候..."
                    WMI_Computer_SerialPort strComputer
                End If
                
                If Check_Show_Setup("按鍵性裝置") = 1 Then
                    Show_Msg "讀取電腦按鍵性裝置資訊中，請稍候..."
                    WMI_Computer_Keyboard strComputer
                End If
                
                If Check_Show_Setup("隨插即用裝置") = 1 Then
                    Show_Msg "讀取電腦隨插即用裝置資訊中，請稍候..."
                    WMI_Computer_PnPEntity strComputer
                End If
                
                If Check_Show_Setup("按鍵性裝置") = 1 Then
                    Show_Msg "讀取電腦指標性裝置資訊中，請稍候..."
                    WMI_Computer_PointingDevice strComputer
                End If
                   
                If Check_Show_Setup("音效裝置") = 1 Then
                    Show_Msg "讀取電腦音效裝置資訊中，請稍候..."
                    WMI_Computer_SoundDevice strComputer
                End If
                
                If Check_Show_Setup("系統") = 1 Then
                    Show_Msg "讀取電腦系統資訊中，請稍候..."
                    WMI_Computer_System_Information strComputer
                End If
                    
                If Check_Show_Setup("環境變數") = 1 Then
                    Show_Msg "讀取電腦環境變數資訊中，請稍候..."
                    WMI_Computer_Environment strComputer
                End If
                    
                If Check_Show_Setup("開機執行") = 1 Then
                    Show_Msg "讀取電腦開機執行資訊中，請稍候..."
                    WMI_Computer_StartupCommand strComputer
                End If
                
                If Check_Show_Setup("終端機") = 1 Then
                    Show_Msg "讀取電腦終端機資訊中，請稍候..."
                    WMI_Computer_Terminal strComputer
                End If
                
                If Check_Show_Setup("軟體安裝") = 1 Then
                    Show_Msg "讀取電腦軟體安裝資訊中，請稍候..."
                    WMI_Computer_Product_Installed strComputer
                End If
                
                If Check_Show_Setup("更新檔") = 1 Then
                    Show_Msg "讀取電腦更新檔資訊中，請稍候..."
                    WMI_Computer_QuickFixEngineering strComputer
                End If
                
                If Check_Show_Setup("邏輯磁碟") = 1 Then
                    Show_Msg "讀取電腦邏輯磁碟資訊中，請稍候..."
                    WMI_Computer_LogicalDisk strComputer
                End If
                
                If Check_Show_Setup("服務") = 1 Then
                    Show_Msg "讀取電腦服務資訊中，請稍候..."
                    WMI_Computer_Service strComputer
                End If
                
                If Check_Show_Setup("執行緒") = 1 Then
                    Show_Msg "讀取電腦執行緒資訊中，請稍候..."
                    WMI_Computer_Process strComputer
                End If
                
                If Check_Show_Setup("網路介面") = 1 Then
                    Show_Msg "讀取電腦網路介面資訊中，請稍候..."
                    WMI_Computer_NetworkAdapter strComputer
                End If
                
                If Check_Show_Setup("網路裝置設定") = 1 Then
                    Show_Msg "讀取電腦網路裝置設定資訊中，請稍候..."
                    WMI_Computer_NetworkAdapterConfiguration strComputer
                End If
                
                If Check_Show_Setup("磁碟機") = 1 Then
                    Show_Msg "讀取電腦磁碟機裝置資訊中，請稍候..."
                    WMI_Computer_DiskDrive strComputer
                End If
                
                If Check_Show_Setup("共用") = 1 Then
                    Show_Msg "讀取電腦共用資訊中，請稍候..."
                    WMI_Computer_Share strComputer
                End If
                
                If Check_Show_Setup("印表機") = 1 Then
                    Show_Msg "讀取電腦印表機資訊中，請稍候..."
                    WMI_Computer_Printer strComputer
                End If
                
                If Check_Show_Setup("印表機網路連接埠") = 1 Then
                    Show_Msg "讀取電腦印表機網路連接埠資訊中，請稍候..."
                    WMI_Computer_TCPIPPrinterPort strComputer
                End If
                
                If Check_Show_Setup("印表機驅動程式") = 1 Then
                    Show_Msg "讀取電腦印表機驅動程式資訊中，請稍候..."
                    WMI_Computer_PrinterDriver strComputer
                End If
                
                
                
                If Check_Show_Setup("磁碟分割區") = 1 Then
                    Show_Msg "讀取電腦磁碟分割區資訊中，請稍候..."
                    WMI_Computer_DiskPartition strComputer
                End If
                
                Show_Msg "電腦資料讀取成功"
                
            End If
        Else
        
            SGrid_Computer_List_AddRow "錯誤", Src_State, "錯誤"
            Show_Msg "無法連線到 " & strComputer
            
            
        End If
        
        
        Show_Msg "正在將資料群組化顯示..."
        .ColumnIsGrouped(3) = False
        .ColumnIsGrouped(3) = True
        Show_Msg "資料群組化顯示完畢"
        
        .AutoWidthColumn 1
        .AutoWidthColumn 2
        
        .Redraw = True
        
    End With
    
    vbalGrid_Sort Me.Sg1, 1
    ExGroup Me.Sg1
    
    
    Show_Msg "資料顯示完畢，共耗費 : " & Timer - tmr_Start & " 秒"
    
    Tvw1.BackColor = tmp_BackColor
    Tvw1.Enabled = True

End Function

Sub SGrid_Computer_List_AddRow(Field1 As String, Field2 As String, Field3 As String)
'將電腦資訊加入 SGrid

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
'列出網域電腦清單以及連線狀態

On Error Resume Next

    '停止列表
    'Call Stop_Computer_List
    
    If Stop_Computers_List = True Then
        Stop_Computer_List
        Exit Function
    End If
    
    '取得所有電腦
    Show_Msg "正在列出電腦..."
    
    Set tmp_Computers = ADSI_Computer_List()
    
    With Sg1
        
        '.Redraw = False
        
        '若重新開始才清除畫面
        If List_Last = 1 Then
            .Clear True
        End If
                    
        .AddColumn "Computer_Name", "電腦名稱"
        .AddColumn "Computer_Online", "連線"
        .AddColumn "Computer_IP", "IP Address"
            
        Dim i As Long
        
        If Sck_List.Count > 1 Then
            For i = 1 To Sck_List.UBound - 1
                Unload Sck_List(i)
            Next
        
        End If
        
        Show_Msg "正在開啟連接埠..."
        
        For i = List_Last To tmp_Computers.Count
            Load Sck_List(i)
            Frm_Main.Sck_List(i).Connect CStr(tmp_Computers(i)), Port_Ping
            DoEvents
        Next
    
        Show_Msg "連線埠已開啟"
        
        .AutoWidthColumn 1
        .Redraw = True
    End With



End Function

Sub Stop_Computer_List()
'停止列表

On Error Resume Next

    Show_Msg "正在停止列表..."
    Stop_Computers_List = True
    
    If Sck_List.Count > 1 Then
        
        Dim i As Long
        For i = 1 To Sck_List.UBound - 1
            Sck_List(i).Close: Unload Sck_List(i)
        Next
    
    End If
    
    Delay 5

    Show_Msg "列表已停止"

End Sub


Sub SGrid_Computers_List_Addrow(Src_Index As Integer, Src_Exist As String, Optional Src_IP As String = "")
'顯示電腦狀況

If Stop_Computers_List = True Then Exit Sub
    

   '成功連線
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
        
    '改變樹系中電腦狀態
    Call Show_Tvw_Computers_State(Sck_List(Src_Index), Src_Exist)

    '紀錄目前最後一筆紀錄
    List_Last = Src_Index
    
    '關閉這個 Socket Port
    Sck_List(Src_Index).Close
    
    Show_Msg "正在讀取使用中電腦 : " & Sg1.Rows & "/" & tmp_Computers.Count
    
    '取得完畢
    If Sg1.Rows = tmp_Computers.Count Then
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK").Enabled = True
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK_COUNTINUE").Enabled = True
        cmdBar(0).Toolbar.Buttons.Item("CONNECT_CHECK_STOP").Enabled = False
        Show_Msg "讀取完畢 " & Sg1.Rows & "/" & tmp_Computers.Count
    End If
    

End Sub

Sub Show_Tvw_Computers_State(Src_Winsock As Winsock, Src_Exist As String)

    Dim Src_Index As Integer: Src_Index = Src_Winsock.index

    '在樹系單上標示出狀態
    Dim tmp_key As String: tmp_key = Format(Src_Index, "0000") & CStr(tmp_Computers(Src_Index))
    Dim SubNode As vbalTreeViewLib6.cTreeViewNode
      
    Set SubNode = Tvw1.Nodes.Item(tmp_key)

    If Src_Exist = "正常連線" Then
        
        '設定可連線的圖示與文字
        SubNode.Bold = True
        
        'SubNode.Tag = Sck_List(Src_Index).RemoteHostIP
        SubNode.Tag = Src_Winsock.RemoteHostIP
        
        SubNode.Image = vbalIml1.ItemIndex("Connected") - 1
        SubNode.SelectedImage = vbalIml1.ItemIndex("Connected") - 1
        
        '可連線電腦總數 +1
        Total_Connected_Computers = Total_Connected_Computers + 1
        Show_Msg ""
    Else
    
        '設定無法連線的圖示與文字
        SubNode.Bold = False
        SubNode.Tag = "0"
        SubNode.Image = vbalIml1.ItemIndex("Disconnect") - 1
        SubNode.SelectedImage = vbalIml1.ItemIndex("Disconnect") - 1
    End If

End Sub

Private Sub Sck_List_Connect(index As Integer)

If Sck_List(index).State = sckConnected Then
    Call SGrid_Computers_List_Addrow(index, "正常連線", Sck_List(index).RemoteHostIP)
End If
 
End Sub

Private Sub Sck_List_Error(index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
    
    Select Case Number
       
        '沒有找到主機
        Case sckHostNotFound, sckHostNotFoundTryAgain
            
            Call SGrid_Computers_List_Addrow(index, "沒有找到主機")
        
        '逾時
        Case sckTimedout
        
            Call SGrid_Computers_List_Addrow(index, "連線逾時")
        
        '其它錯誤 (以本機錯誤較多)
        Case Else
        
            Call SGrid_Computers_List_Addrow(index, "連線失敗")
        
    End Select
        

'MsgBox Number
End Sub

Private Sub Sg1_ColumnClick(ByVal lCol As Long)

    '排序
    vbalGrid_Sort Me.Sg1, lCol
    ExGroup Me.Sg1

End Sub

Private Sub Sg1_Click(ByVal lRow As Long, ByVal lCol As Long)

End Sub

Private Sub Tvw1_nodeCheck(node As vbalTreeViewLib6.cTreeViewNode)

'點選根目錄，進行全選或全部取消
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
'如果點選根目錄不處理
If node.Key = "root" Then
    Can_Process_Computer = ""
    Exit Sub
End If

'取得資訊
Can_Process_Computer = node.Text

End Sub

Private Sub Tvw1_NodeDblClick(node As vbalTreeViewLib6.cTreeViewNode)

'如果點選根目錄不處理
If node.Key = "root" Then Exit Sub

'取得資訊
Can_Process_Computer = node.Text

Show_Computer_Information node

End Sub

Sub Show_Computer_Information(node As vbalTreeViewLib6.cTreeViewNode)
'列出電腦詳細資訊

    '離線狀態者不可處理
    If node.Tag = "0" Then
        If MsgBox("處理離線狀態電腦速度將會很慢，您確定要嘗試與該電腦連線嗎？", vbYesNo, "提示") = vbNo Then Exit Sub
    End If
    
    Show_Msg "電腦資料讀取中，請稍候..."
    Sck_Check_One.Close
    Sck_Check_One.Connect node.Text, Port_Ping


End Sub


Private Sub Tvw1_NodeRightClick(node As vbalTreeViewLib6.cTreeViewNode)

'如果點選根目錄不處理
If node.Key = "root" Then
    Can_Process_Computer = ""
    Exit Sub
End If

    '必需先讓其選取，這樣功能表功能進入後才有電腦名稱可用
    node.Selected = True
    
    '取得資訊
    Can_Process_Computer = node.Text
    
    
    '取得游標在螢幕上的位置
    Call GethWnd
    
    '事實上這行可以不要
    cmdBar(1).ClientCoordinatesToScreen Me.Left, Me.TOp, Me.hwnd
    
    '秀出顯示快顯功能表
    cmdBar(1).ShowPopupMenu tmp_x, tmp_y, cmdBar(1).CommandBars("MANAGERMENU")
    
End Sub

Private Sub Txt_DomainName_GotFocus()
AutoSelStr Me.ActiveControl
End Sub

Private Sub Txt_Port_GotFocus()
AutoSelStr Me.ActiveControl
End Sub
