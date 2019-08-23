Attribute VB_Name = "Mod_WMI_LDAP"
'許多功能必需在 XP 及 .NET Server 以上方可執行
'可由系統資訊判斷，5.1 以上才可
'
Public AD_Domain_Name  As String

Public Can_Process_Computer As String

Public canUseAD As Boolean
Public objWMIService

'暫存作業系統版本
Public OS_Version As Single

Public Sub Delay(ByVal N As Single)
    
    '延遲副程式
    Dim tm1, tm2 As Single
    tm1 = Timer
    Do
        tm2 = Timer
        If tm2 < tm1 Then tm2 = tm2 + 86400
        If tm2 - tm1 > N Then Exit Do
        DoEvents
    Loop
   
End Sub

Function ADSI_Computer_List() As Collection
'列舉網域中電腦
    
On Error GoTo ErrMsg

    '解析網域名稱
    Dim i As Long: i = 0
    Dim tmp_str() As String
    Dim tmp_Domain_Name As String: tmp_Domain_Name = ""
    tmp_str = Split(Trim(AD_Domain_Name), ".")
    For i = LBound(tmp_str) To UBound(tmp_str)
        tmp_Domain_Name = tmp_Domain_Name & "DC=" & tmp_str(i)
        If i < UBound(tmp_str) Then tmp_Domain_Name = tmp_Domain_Name & ","
    Next
    
    '建立 ADO 連線
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    '開啟查詢
    Set objCommand = CreateObject("ADODB.Command")
    Set objCommand.ActiveConnection = objConnection
    
    Const ADS_SCOPE_SUBTREE = 2
    
    '定義 ADO 參數
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Timeout") = 30
    objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
    objCommand.Properties("Cache Results") = False
    
    '送出向 AD 要求電腦的清單
    objCommand.CommandText = _
        "SELECT Name, Location FROM 'LDAP://" & tmp_Domain_Name & "' " & _
        "WHERE objectClass='computer' ORDER BY Name"
    Set objRecordSet = objCommand.Execute
    
    '暫時存放電腦清單
    Dim tmp_Computers As New Collection
    
    '開始取出資料
    objRecordSet.MoveFirst
    Do Until objRecordSet.EOF
        i = i + 1
        
        '將電腦加入集合
        tmp_Computers.Add objRecordSet.Fields("Name").Value, "i" & i
        DoEvents
        objRecordSet.MoveNext
    Loop
    Set objConnection = Nothing
    
    '將集合傳回
    Set ADSI_Computer_List = tmp_Computers
    
    canUseAD = True
    
Exit Function

ErrMsg:

    If Err.Number = -2147217865 Then
        'MsgBox "無法存取網域", vbInformation, "錯誤"
    Else
        MsgBox "發生其它錯誤 : " & Err.Description & "(" & Err.Number & ")", vbInformation, "錯誤"
    End If
    
    canUseAD = False

End Function

Function ADSI_User_List() As Collection
'列舉網域中 User

On Error Resume Next
    
    '解析網域名稱
    Dim i As Long: i = 0
    Dim tmp_str() As String
    Dim tmp_Domain_Name As String: tmp_Domain_Name = ""
    tmp_str = Split(Trim(AD_Domain_Name), ".")
    For i = LBound(tmp_str) To UBound(tmp_str)
        tmp_Domain_Name = tmp_Domain_Name & "DC=" & tmp_str(i)
        If i < UBound(tmp_str) Then tmp_Domain_Name = tmp_Domain_Name & ","
    Next
    
    
    Const ADS_SCOPE_SUBTREE = 2
    
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    Set objCommand.ActiveConnection = objConnection
    
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
    
    objCommand.CommandText = "SELECT * FROM 'LDAP://" & tmp_Domain_Name & "' WHERE objectCategory='user'"
    Set objRecordSet = objCommand.Execute
    
    objRecordSet.MoveFirst
    Do Until objRecordSet.EOF
        
        For i = 0 To objRecordSet.Fields.Count
    
            aa = aa & objRecordSet.Fields(i).Value & " , "
        Next
        
        aa = aa & vbCrLf
        DoEvents
        objRecordSet.MoveNext
        
    Loop

End Function
Function WMI_Service_Create(strComputer As String) As Boolean

On Error GoTo ErrMsg
    
    Set objWMIService = GetObject("winmgmts:" _
        & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

    WMI_Service_Create = True

Exit Function

ErrMsg:
    
    WMI_Service_Create = False

End Function

Function WMI_Computer_Ping(strComputer As String) As Boolean
'檢視電腦是否可用

'On Error GoTo ErrMsg
On Error Resume Next
    'strComputer = "."
    
    Set colPingedComputers = objWMIService.ExecQuery _
        ("Select * from Win32_PingStatus Where Address = '" & strComputer & "'")
    
    For Each objComputer In colPingedComputers
        
        If objComputer.StatusCode = 0 Then
            'Frm_Main.SGrid_Computer_List_AddRow "Remote computer responded."
            WMI_Computer_Ping = True
            'Exit For
        Else
            'Frm_Main.SGrid_Computer_List_AddRow "Remote computer did not respond."
            WMI_Computer_Ping = False
            'Exit For
        End If
        
        DoEvents
    Next
    
Exit Function

ErrMsg:

WMI_Computer_Ping = False

End Function

Function WMI_Computer_OperatingSystem(strComputer As String)
'查看作業系統

On Error Resume Next
  
    Set cols = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    
    For Each obj In cols

        Frm_Main.SGrid_Computer_List_AddRow "作業系統", obj.Caption & " " & obj.Version, "作業系統"
        OS_Version = CSng(Left(obj.Version, 3))
        
    Next


End Function

Function WMI_Computer_Login_UserName(strComputer As String)
'取得登入者名稱

On Error GoTo ErrMsg

    Set colServices = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    
    If colServices.Count > 0 Then
    
       For Each objService In colServices
           Frm_Main.SGrid_Computer_List_AddRow "登入帳號", objService.UserName, "一般"
           DoEvents
       Next
    
    Else
       
       Frm_Main.SGrid_Computer_List_AddRow "登入帳號", "無", "一般"
    
    End If

Exit Function

ErrMsg:

    Frm_Main.SGrid_Computer_List_AddRow "登入帳號", "無", "一般"
    
End Function



Function WMI_Computer_System_Information(strComputer As String) As String()
'系統資訊

On Error Resume Next
    
    Dim i As Long
    
    Set colSettings = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    
    For Each objOperatingSystem In colSettings
        
        
        Frm_Main.SGrid_Computer_List_AddRow "作業系統版本", objOperatingSystem.Version, "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "Service Pack", objOperatingSystem.ServicePackMajorVersion _
                                                        & "." & objOperatingSystem.ServicePackMinorVersion, "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "作業系統製造商", objOperatingSystem.Manufacturer, "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "Windows 目錄", objOperatingSystem.WindowsDirectory, "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "國別碼", strLocale(objOperatingSystem.Locale) & " - (" & objOperatingSystem.Locale & ")", "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "國碼", objOperatingSystem.CountryCode, "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "字碼頁", strCodepage(objOperatingSystem.CodeSet) & " - (" & objOperatingSystem.CodeSet & ")", "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "資料執行防止 (DEP)", IIf(objOperatingSystem.DataExecutionPrevention_Available = True, "啟用", "關閉"), "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "資料執行防止 (DEP) - 驅動程式", IIf(objOperatingSystem.DataExecutionPrevention_Drivers = True, "啟用", "關閉"), "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "加密等級", objOperatingSystem.EncryptionLevel & " 位元", "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "系統安裝日期", Change_GMT(objOperatingSystem.InstallDate), "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "最近一次開機", Change_GMT(objOperatingSystem.LastBootUpTime), "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "允許最大執行緒數目", objOperatingSystem.MaxNumberOfProcesses, "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "允許最大執行緒記憶體大小", Format_MB_By_K(objOperatingSystem.MaxProcessMemorySize), "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "目前執行緒數目", objOperatingSystem.NumberOfProcesses, "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "目前使用者數目", objOperatingSystem.NumberOfUsers, "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "組織", objOperatingSystem.NumberOfUsers, "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "作業系統語言", strOsLang(objOperatingSystem.OSLanguage) & " - (" & objOperatingSystem.OSLanguage & ")", "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "註冊使用者", objOperatingSystem.RegisteredUser, "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "所有磁碟分頁檔", Format_MB_By_K(objOperatingSystem.SizeStoredInPagingFiles), "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "實體記憶體剩餘", Format_MB_By_K(objOperatingSystem.FreePhysicalMemory), "記憶體"
        Frm_Main.SGrid_Computer_List_AddRow "虛擬記憶體大小", Format_MB_By_K(objOperatingSystem.TotalVirtualMemorySize), "記憶體"
        Frm_Main.SGrid_Computer_List_AddRow "虛擬記憶體剩餘", Format_MB_By_K(objOperatingSystem.FreeVirtualMemory), "記憶體"
        
        DoEvents
    Next
    
    Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    
    For Each objComputer In colSettings
        
        'Frm_Main.SGrid_Computer_List_AddRow "電腦名稱", objComputer.Name, "一般"
        Frm_Main.SGrid_Computer_List_AddRow "製造商", objComputer.Manufacturer, "硬體"
        Frm_Main.SGrid_Computer_List_AddRow "規格", objComputer.Model, "硬體"
        Frm_Main.SGrid_Computer_List_AddRow "BootROM", IIf(objComputer.BootROMSupported = True, "支援", "不支援"), "硬體"
        Frm_Main.SGrid_Computer_List_AddRow "時區", "GMT " & IIf(objComputer.CurrentTimeZone > 0, "+", "-") & objComputer.CurrentTimeZone / 60, "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "日光節約時間", IIf(objComputer.EnableDaylightSavingsTime = True, "啟用", "關閉"), "作業系統"
        Frm_Main.SGrid_Computer_List_AddRow "中央處理器數目", objComputer.NumberOfProcessors, "硬體"
        Frm_Main.SGrid_Computer_List_AddRow "實體記憶體大小", Format_MB_By_B(objComputer.TotalPhysicalMemory), "記憶體"
        DoEvents
        
    Next
    
    i = 0
    Set colSettings = objWMIService.ExecQuery("Select * from Win32_Processor")
    For Each objprocessor In colSettings
        i = i + 1
        Frm_Main.SGrid_Computer_List_AddRow "名稱", Trim(objprocessor.Name), "硬體 - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "Socket 型式", objprocessor.SocketDesignation, "硬體 - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "最高時脈", objprocessor.MaxClockSpeed & " MHz", "硬體 - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "標準時脈", objprocessor.CurrentClockSpeed & " MHz", "硬體 - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "外頻", objprocessor.ExtClock & " MHz", "硬體 - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "資料頻寬", objprocessor.DataWidth & " Bits", "硬體 - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "目前負載", objprocessor.LoadPercentage & " %", "硬體 - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "L2 Cache", objprocessor.L2CacheSize & " K", "硬體 - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "L2 Cache Speed", IIf(IsNull(objprocessor.L2CacheSpeed), "0", objprocessor.L2CacheSpeed) & " MHz", "硬體 - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "電力管理", IIf(objprocessor.PowerManagementSupported, "支援", "不支援"), "硬體 - CPU " & i
        
        DoEvents
    Next
    i = 0
    
    Set colSettings = objWMIService.ExecQuery("Select * from Win32_BIOS")
    For Each objBIOS In colSettings
        Frm_Main.SGrid_Computer_List_AddRow "BIOS 版本", objBIOS.Version, "硬體"
        DoEvents
    Next
    

End Function


Function WMI_Computer_Product_Installed(strComputer As String)
On Error Resume Next
    
    Set colSoftware = objWMIService.ExecQuery("Select * from Win32_Product")
    
    For Each objSoftware In colSoftware
        
        Frm_Main.SGrid_Computer_List_AddRow objSoftware.Caption & " " & objSoftware.Version, objSoftware.InstallLocation, "軟體安裝"
        DoEvents
    Next


End Function


Function WMI_Computer_LogicalDisk(srcComputer As String)
'列出邏輯磁碟機

On Error Resume Next
    
    Set colLogicalDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
    For Each objdisk In colLogicalDisks
        
        Frm_Main.SGrid_Computer_List_AddRow _
            objdisk.DeviceID & " " & objdisk.VolumeName & " - (" & objdisk.Description & ")", _
            "總共: " & Format_MB_By_B(objdisk.Size) & " / 可用: " & Format_MB_By_B(objdisk.FreeSpace) & IIf(IsNull(objdisk.FileSystem), "", " (" & objdisk.FileSystem & ")"), _
            "邏輯磁碟機"
        DoEvents
    Next

End Function


Function WMI_Computer_Service(strComputer As String)
'列舉服務
On Error Resume Next

    Set colServices = objWMIService.ExecQuery("Select * from Win32_Service")
    
    If colServices.Count > 0 Then
        
        For Each objService In colServices
            Frm_Main.SGrid_Computer_List_AddRow _
                objService.Name & " (" & objService.Pathname & ")", _
                objService.State, "服務"
            DoEvents
        Next
    
    End If

End Function

Function WMI_Computer_Process(strComputer As String)
'檢視執行緒狀態

On Error Resume Next
    
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")
    
    For Each objProcess In colProcesses
        'objDictionary.Add objProcess.ProcessID, objProcess.Name
        Frm_Main.SGrid_Computer_List_AddRow _
            objProcess.Name & " (" & objProcess.ExecutablePath & ")", _
            "PID (" & objProcess.ProcessID & ") ", "執行緒"
        DoEvents
    Next
    
    'Set colThreads = objWMIService.ExecQuery("Select * from Win32_Thread")
    'aa = ""
    'For Each objThread In colThreads
    '    intProcessID = CInt(objThread.ProcessHandle)
    '    strProcessName = objDictionary.Item(intProcessID)
    '    aa = aa & strProcessName & vbTab & objThread.ProcessHandle vbTab & objThread.Handle & vbTab & objThread.ThreadState & vbCrLf
    'Next
    
End Function


Function WMI_Computer_NetworkAdapter(strComputer As String)
'檢視網路裝置狀態

On Error Resume Next

    Set colNA = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter")
    
    For Each objNA In colNA

        Frm_Main.SGrid_Computer_List_AddRow _
            objNA.Name & " (" & objNA.Manufacturer & ")", _
            "MAC - (" & objNA.MACAddress & ") ", "網路介面"
        DoEvents
    Next

End Function


Function WMI_Computer_DiskDrive(strComputer As String)
'列舉實體磁碟內容

On Error Resume Next

    Set colDiskDrives = objWMIService.ExecQuery("Select * from Win32_DiskDrive")
    
    For Each objDiskDrive In colDiskDrives
    
        Frm_Main.SGrid_Computer_List_AddRow _
            objDiskDrive.Model & " / " & _
            objDiskDrive.InterfaceType & _
            " / " & objDiskDrive.Manufacturer, _
            Format_MB_By_B(objDiskDrive.Size), _
            "磁碟機"
        DoEvents
    Next

End Function

Function WMI_Computer_Share(strComputer As String)
'列舉共用資料夾內容
On Error Resume Next

    Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")
    
    For Each objShare In colShares
    
        Frm_Main.SGrid_Computer_List_AddRow _
            objShare.Name & " - ", _
            objShare.Path, "共用資源"
        DoEvents
    Next

End Function


Function WMI_Computer_Printer(strComputer As String)
'列舉印表機內容 僅適用 XP 以上
On Error Resume Next

'    If OS_Version > 5.1 Then  '(大於 XP 以上才可)
        
        Set colPrinters = objWMIService.ExecQuery("Select * from Win32_Printer")
        For Each objPrinter In colPrinters
        
            Frm_Main.SGrid_Computer_List_AddRow _
                objPrinter.DriverName & " - " & objPrinter.Name & " (狀況:" & objPrinter.Status & ")", _
                objPrinter.PortName, "印表機"
            DoEvents
        Next
        

End Function

Function WMI_Computer_PrinterDriver(strComputer As String)
'列舉印表機驅動程式

On Error Resume Next

 If OS_Version < 5.1 Then  '(小於 XP 以不支援)
            
            Frm_Main.SGrid_Computer_List_AddRow _
                "此功能只有目標電腦作業系統在 XP (5.1) 以上才支援顯示", _
                "未支援", _
                "印表機驅動程式"
    Else
        
        Set cols = objWMIService.ExecQuery("Select * from Win32_PrinterDriver")
        For Each obj In cols
        
            Frm_Main.SGrid_Computer_List_AddRow _
                obj.Name & " " & obj.Version & " - " & obj.Description, _
                obj.DriverPath, _
                "印表機驅動程式"
            DoEvents
        Next
        
  End If
  
End Function



Function WMI_Computer_TCPIPPrinterPort(strComputer As String)
'列舉印表機 Port 內容 僅適用 XP 以上

On Error Resume Next

    If OS_Version < 5.1 Then  '(小於 XP 以不支援)
            
            Frm_Main.SGrid_Computer_List_AddRow _
                "此功能只有目標電腦作業系統在 XP (5.1) 以上才支援顯示", _
                "未支援", _
                "印表機網路連接埠"
    Else
            
        Set colTPs = objWMIService.ExecQuery("Select * from Win32_TCPIPPrinterPort")
        For Each objtp In colTPs
        
            Frm_Main.SGrid_Computer_List_AddRow _
                objtp.Name & " " & " PortNumber:" & objtp.PortNumber & " SNMP: " & IIf(objtp.SNMPEnabled = True, "是", "否"), _
                objtp.HostAddress, "印表機網路連接埠"
            DoEvents
        Next
    
    End If

End Function

Function WMI_Computer_DiskPartition(strComputer As String)
'列舉磁碟分割區

On Error Resume Next

    Set colDiskPartitions = objWMIService.ExecQuery("Select * from Win32_DiskPartition")
    
    For Each objDP In colDiskPartitions
            
        Frm_Main.SGrid_Computer_List_AddRow _
            objDP.Name & IIf(objDP.BootPartition = True, " [Boot]", ""), _
            Format_MB_By_B(objDP.Size), _
            "磁碟分割區"
        DoEvents
    Next

End Function


Function WMI_Computer_Terminal(strComputer As String)

'列舉終端機資訊

On Error Resume Next
    
    If OS_Version < 5.1 Then  '(小於 XP 不支援)
    
            Frm_Main.SGrid_Computer_List_AddRow _
                "此功能只有目標電腦作業系統在 XP (5.1) 以上才支援顯示", _
                "未支援", _
                "終端機"
    Else
            
        Set colTerminals = objWMIService.ExecQuery("Select * from Win32_Terminal")
        For Each objTerminal In colTerminals
        
            Frm_Main.SGrid_Computer_List_AddRow _
                objTerminal.TerminalName, _
                IIf(objTerminal.fEnableTerminal = 1, "啟用", "停用"), _
                "終端機"
        
            DoEvents
        Next
    
    End If

End Function


Function WMI_Computer_StartupCommand(strComputer As String)
'列舉開機執行

On Error Resume Next
    
    Set colSCs = objWMIService.ExecQuery("Select * from Win32_StartupCommand")
    For Each objsc In colSCs
    
        Frm_Main.SGrid_Computer_List_AddRow _
            objsc.Name & " (User:" & objsc.User & " / Loc.:" & objsc.Location & ")", _
            objsc.Command, _
            "開機執行"
    
        DoEvents
    Next

End Function


Function WMI_Computer_SoundDevice(strComputer As String)
'列舉音效裝置

On Error Resume Next
    Set cols = objWMIService.ExecQuery("Select * from Win32_SoundDevice")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name, _
            obj.Manufacturer, _
            "音效裝置"
        DoEvents
    Next

End Function

Function WMI_Computer_SerialPort(strComputer As String)
'列舉連接埠裝置

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_SerialPort")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name, _
            IIf(IsNull(obj.Status), "不可用", "可用"), _
            "連接埠"
        DoEvents
    Next


    Set cols = objWMIService.ExecQuery("Select * from Win32_ParallelPort")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name, _
            IIf(IsNull(obj.Status), "不可用", "可用"), _
            "連接埠"
        DoEvents
    Next

End Function



Function WMI_Computer_ScheduledJob(strComputer As String)
'列舉工作排程 停用

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_ScheduledJob")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name, _
            IIf(IsNull(obj.Status), "不可用", "可用"), _
            "連接埠"
        DoEvents
    Next

End Function

Function WMI_Computer_PointingDevice(strComputer As String)
'列舉指標性裝置

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_Keyboard")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name & " (" & obj.Description & ")", _
            IIf(IsNull(obj.Status), "不可用", "可用"), _
            "按鍵性裝置"
        DoEvents
    Next

End Function

Function WMI_Computer_Keyboard(strComputer As String)
'列舉按鍵性裝置

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_PointingDevice")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name & " (" & obj.Manufacturer & ")", _
            "按鈕數: " & obj.NumberOfButtons, _
            "指標性裝置"
        DoEvents
    Next

End Function

Function WMI_Computer_PnPEntity(strComputer As String)
'列舉隨插即用裝置

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_PnPEntity")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name & " - " & obj.Description & " " & obj.Manufacturer & "", _
            IIf(IsNull(obj.Status), "未連接", "已連接"), _
            "隨插即用裝置"
        DoEvents
    Next

End Function

Function WMI_Computer_Group(strComputer As String)
'列舉本機群組

'已改用 ADSI 寫法

'On Error Resume Next
'    'MsgBox strComputer
'    Set cols = objWMIService.ExecQuery("Select * from Win32_Group Domain = '" & strComputer & "'")
'    For Each obj In cols
'
'        Frm_Main.SGrid_Computer_List_AddRow _
'            obj.Name, _
'            obj.Description, _
'            "本機群組"
'        DoEvents
'    Next

End Function


Function WMI_Computer_Displays(strComputer As String)
'列舉顯示卡資訊

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_DisplayConfiguration")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow "顯示卡型號", StrNullToSpace(obj.Name), "顯示卡"
        Frm_Main.SGrid_Computer_List_AddRow "顯示卡型號說明", StrNullToSpace(obj.Description), "顯示卡"
        Frm_Main.SGrid_Computer_List_AddRow "驅動程式版本", StrNullToSpace(obj.DriverVersion), "顯示卡"
        Frm_Main.SGrid_Computer_List_AddRow "資料頻寬", StrNullToSpace(obj.BitsPerPel), "顯示卡"
        Frm_Main.SGrid_Computer_List_AddRow "垂直更新頻率", StrNullToSpace(obj.DisplayFrequency), "顯示卡"
        DoEvents
        
    Next

    Set cols = objWMIService.ExecQuery("Select * from Win32_DisplayControllerConfiguration")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow "晶片型號", StrNullToSpace(obj.Name), "顯示卡"
        Frm_Main.SGrid_Computer_List_AddRow "晶片型號說明", StrNullToSpace(obj.Description), "顯示卡"
        Frm_Main.SGrid_Computer_List_AddRow "解析度", StrNullToSpace(obj.HorizontalResolution) & "x" & StrNullToSpace(obj.VerticalResolution), "顯示卡"
        Frm_Main.SGrid_Computer_List_AddRow "色彩品質", StrNullToSpace(obj.BitsPerPixel), "顯示卡"
        Frm_Main.SGrid_Computer_List_AddRow "顯示模式", StrNullToSpace(obj.VideoMode), "顯示卡"
        DoEvents
        
    Next


End Function

Function WMI_Computer_DesktopMonitor(strComputer As String)
'列舉顯示器可用顯示模式

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name & " " & obj.Screenwidth & "x" & obj.ScreenHeight, _
            obj.DeviceID, _
            "顯示器"
        DoEvents
    Next

End Function


Function WMI_Computer_Environment(strComputer As String)
'列舉環境變數變數

    On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_Environment")

    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.UserName & " | " & obj.Name & " | 系統變數:" & IIf(obj.SystemVariable = True, "是", "否") & " (" & obj.Description & ") ", _
            obj.VariableValue, _
            "環境變數"
        
        DoEvents
    Next
    

End Function

Function WMI_Computer_CodecFile(strComputer As String)
'列舉解碼器資訊

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_CodecFile")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            StrNullToSpace(obj.Group) & " | " & StrNullToSpace(obj.Description) & " " & StrNullToSpace(obj.Version) & " (" & StrNullToSpace(obj.Manufacturer) & ")", _
            StrNullToSpace(obj.Name), _
            "解碼器"
        DoEvents
    Next

End Function



Function WMI_Computer_NetworkAdapterConfiguration(strComputer As String)
'列舉網路設定

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    Dim i As Integer
    
    For Each obj In cols
    
        i = i + 1
        
        Frm_Main.SGrid_Computer_List_AddRow "裝置", StrNullToSpace(obj.index) & " - " & StrNullToSpace(obj.ServiceName), "網路設定 " & i
        
        
        Frm_Main.SGrid_Computer_List_AddRow "裝置名稱", StrNullToSpace(obj.Description), "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "啟用 DHCP", StrNullToSpace(obj.DHCPEnabled), "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "DHCP 伺服器", StrNullToSpace(obj.dhcpserver), "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "DNS 領域", obj.DNSDomain, "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "IP 位址", IIf(IsNull(obj.IPAddress) = True, "", Join(obj.IPAddress)), "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "子網路遮罩", IIf(IsNull(obj.IPSubnet) = True, "", Join(obj.IPSubnet)), "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "預設閘道", IIf(IsNull(obj.DefaultIPGateway) = True, "", Join(obj.DefaultIPGateway)), "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "DNS 伺服器", IIf(IsNull(obj.DNSServerSearchOrder) = True, "", Join(obj.DNSServerSearchOrder)), "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "最少閘道", IIf(IsNull(obj.GatewayCostMetric) = True, "", Join(obj.GatewayCostMetric)), "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "最少連線", StrNullToSpace(obj.IPConnectionMetric), "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "啟用 IPX", StrNullToSpace(obj.IPXEnabled), "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "IPX 位址", IIf(IsNull(obj.IPXAddress) = True, "", Join(obj.IPXAddress)), "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "在 DNS 上註冊電腦名稱", StrNullToSpace(obj.DomainDNSRegistrationEnabled), "網路設定 " & i
        Frm_Main.SGrid_Computer_List_AddRow "MAC 卡號", StrNullToSpace(obj.MACAddress), "網路設定 " & i
        
        DoEvents
    Next
    
    
End Function


Function WMI_Computer_QuickFixEngineering(strComputer As String)
'列出已安裝的 Hot Fixes
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_QuickFixEngineering")
    
    For Each obj In cols
        
        If obj.HotFixID <> "File 1" Then
            Frm_Main.SGrid_Computer_List_AddRow obj.HotFixID, obj.Description, "更新檔"
        End If
        
        DoEvents
    Next
        
End Function

Function ADSI_Computer_Group(strComputer As String)
'列舉目標電腦的群組
    Set cols = GetObject("WinNT://" & strComputer & ",computer")
    cols.Filter = Array("group")
    
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow obj.Name, obj.Description, "本機群組"
    Next

End Function

Function ADSI_Computer_User(strComputer As String)
'列舉目標電腦的帳號
    
    'Dim cmp
    
    'Dim usr As IADsUser
    Dim usr_grp As IADsGroup
    Dim tmp_grp As String
    
    Set usr = GetObject("WinNT://" & strComputer & ",computer")
    usr.Filter = Array("user")
        
    
    '列舉 User
    For Each obj In usr
        
        '取出 User 所隸屬的群組
        For Each usr_grp In obj.Groups
            tmp_grp = tmp_grp & usr_grp.Name & ";"
        Next
        tmp_grp = Left(tmp_grp, Len(tmp_grp) - 1)
        
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name & " (" & tmp_grp & ") " & IIf(obj.AccountDisabled = True, "(啟用)", "(停用)"), _
            obj.Description, _
            "本機使用者"
        
        tmp_grp = ""
        
    Next

End Function


