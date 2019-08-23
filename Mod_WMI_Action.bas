Attribute VB_Name = "Mod_WMI_Action"
Function Action_Service(Src_Service_Name As String, Src_State As String) As String
'對服務做處理

    Dim cols
    Set cols = objWMIService.ExecQuery("Select * from Win32_Service Where Name = '" & Src_Service_Name & "'")
    
    For Each obj In cols
        
        Select Case Src_State
            
            'action:Start
            Case "ACTION:START": Action_Service = CStr(obj.StartService)
            
            'action:Stop
            Case "ACTION:STOP": Action_Service = CStr(obj.StopService)
                
            'action:Delete
            Case "ACTION:DELETE": obj.StopService: Action_Service = CStr(obj.Delete)
            
            
            'startmode:Auto
            Case "STARTMODE:AUTO": Action_Service = CStr(obj.ChangeStartmode("Automatic"))
            
            'startmode:Manual
            Case "STARTMODE:MANUAL": Action_Service = CStr(obj.ChangeStartmode("Manual"))
            
            'startmode:Stopped
            Case "STARTMODE:DISABLED": Action_Service = CStr(obj.ChangeStartmode("Disabled"))
            
        End Select
    Next

End Function

Function Action_Process(Src_Process_Name As String, Src_State As String) As String
'對執行緒做處理

    
Select Case Src_State
    
            
    'action:CREATE
    Case "ACTION:CREATE"
    
        Dim tmp_cmd As String
        Dim tmp_Exec As String: tmp_Exec = InputBox("請輸入目標執行檔的路徑與檔名", "建立執行緒", "notepad.exe")
        Dim tmp_Active As String: tmp_Active = InputBox("要開啟於前台或是隱藏此執行緒？1=開啟,0=隱藏", "顯示", "1")
        Dim tmp_Runat As String: tmp_Runat = InputBox("於什麼時候執行？格式 HH:MM (會以你輸入的時間再延後一分鐘)", "時間點", Format(Time, "hh:mm"))
        
        If tmp_Exec = "" Then Action_Process = "-1": Exit Function
        If tmp_Active = "" Then Action_Process = "-1": Exit Function
        If tmp_Runat = "" Then Action_Process = "-1": Exit Function
        
        tmp_Runat = Format(DateAdd("n", "1", tmp_Runat), "hh:mm")
        tmp_cmd = "Cmd.exe /c " & Chr(34) & "AT " & tmp_Runat & IIf(tmp_Active = "1", " /interactive ", " ") & tmp_Exec & Chr(34)
        
        Set obj = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & Can_Process_Computer & "\root\cimv2:Win32_Process")
        Action_Process = CStr(obj.Create(tmp_cmd, Null, Null, intProcessID))
        
    
    'action:TERMINATE
    Case "ACTION:TERMINATE"
    
        Dim cols
        Set cols = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & Src_Process_Name & "'")
    
        For Each obj In cols
            obj.Terminate
        Next

End Select
End Function


Function Action_Shutdown(Src_Null As String, Src_State As String) As String
'對服務做開關機處理

    If MsgBox("您確定要將 " & Can_Process_Computer & " 重新開機 / 關機 嗎？", vbQuestion + vbYesNo, "提示") = vbYes Then
            
        Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\" & Can_Process_Computer & "\root\cimv2")
        
        Set cols = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
            
        Select Case Src_State
                    
            'action:重開
            Case "ACTION:REBOOT"
            
                For Each obj In cols
                    obj.Reboot
                Next
            
            'action:關機
            Case "ACTION:SHUTDOWN"
            
                For Each obj In cols
                    obj.Shutdown
                Next
         
        End Select

    End If
    
End Function


Function Action_Printer(Src_PrinterDriver_Name As String, Src_State As String) As String
'對印表機處理
    
Select Case Src_State
    
            
    Case "ACTION:ADD"
        
        Frm_Manager_Printer_AddPrinter.Show 1

    Case "ACTION:DELETE"
        
        If MsgBox("您確定刪除 " & Can_Process_Computer & " 上的 " & Src_PrinterDriver_Name & "？", vbQuestion + vbYesNo, "提示") = vbYes Then

            Dim cols
            Set cols = objWMIService.ExecQuery("Select * from Win32_Printer where Name = '" & Trim(Src_PrinterDriver_Name) & "'")
            
            For Each obj In cols
                obj.Delete_
            Next

        End If
        
End Select
End Function

Function Action_PrinterDriver(Src_PrinterDriver_Name As String, Src_State As String) As String
'對印表機驅動程式處理
    
Select Case Src_State
    
            
    Case "ACTION:ADD"
        
        '顯示選驅動程式資料夾視窗
        Dim sFile As String
        Dim cc As New cCommonDialog
        If cc.VBGetOpenFileName(sFile, , , , , , "驅動程式資訊檔 (*.INF)|*.INF", , , "選擇印表機驅動程式", "INF", Screen.ActiveForm.hwnd, OFN_HIDEREADONLY) Then
            
        End If

        '有選到檔案才動作
        If Len(sFile) > 0 Then
        
            Dim sPath As String:  sPath = Left(sFile, InStrRev(sFile, "\"))
            
            '取得驅動程式中的 DRVNAME
            Dim fs As New FileSystemObject
            Dim a As TextStream
            Set a = fs.OpenTextFile(sFile, ForReading, False)
            Dim tmp_File As String: tmp_File = a.ReadAll
            
            '設定開始找的區段
            Dim tmp_Serach_Start As Long: tmp_Serach_Start = InStr(1, tmp_File, "[Strings]")
                        
            
            '設定 DRVNAME 資訊該行的位置
            Dim tmp_Drvname_Start As Long: tmp_Drvname_Start = InStr(tmp_Serach_Start, tmp_File, "DRVNAME")
            Dim tmp_Drvname_End As Long: tmp_Drvname_End = InStr(tmp_Drvname_Start, tmp_File, vbCrLf)
            
            '取出整段
            Dim tmp_Drvname_Str As String: tmp_Drvname_Str = Mid(tmp_File, tmp_Drvname_Start, tmp_Drvname_End - tmp_Drvname_Start)
             
            '剔除掉不必要的字
            tmp_Drvname_Str = Replace(tmp_Drvname_Str, "DRVNAME", "")
            tmp_Drvname_Str = Replace(tmp_Drvname_Str, "=", "")
            tmp_Drvname_Str = Trim(Replace(tmp_Drvname_Str, """", ""))
            
            a.Close
            Set fs = Nothing
                        
            
            If MsgBox("是否要安裝 " & tmp_Drvname_Str & " 於 " & Can_Process_Computer & " 上？", vbQuestion + vbYesNo, "確認") = vbYes Then
            
                '寫入
                
                MsgBox "p:" & Left(sPath, Len(sPath) - 1)
                MsgBox "f:" & sFile
                Set obj = objWMIService.Get("Win32_PrinterDriver")
                obj.Name = tmp_Drvname_Str
                obj.SupportedPlatform = "Windows NT x86"
                obj.Version = "3"
                obj.FilePath = Left(sPath, Len(sPath) - 1)
                obj.InfName = sFile
                Action_PrinterDriver = CStr(obj.AddPrinterDriver(obj))
            
            Else
            
                Action_PrinterDriver = "-1"
            
            End If

        Else
        
            
        End If
        
    Case "ACTION:DELETE"
    

End Select
End Function


Function Action_PrinterTCPIPPort(Src_PrinterTCPIPPort_Name As String, Src_State As String) As String
'對印表機網路連接埠處理
    
Select Case Src_State
            
    Case "ACTION:ADD"
        
        Dim tmp_ip As String:  tmp_ip = InputBox("請輸入 IP:", "新增 TCPIP 連接埠")
        Dim tmp_protocol As String: tmp_protocol = InputBox("請選擇通訊協定: " & vbCrLf & "RAW 或 LPR (預設)", "新增 TCPIP 連接埠", "LPR")
        
        Set obj = objWMIService.Get("Win32_TCPIPPrinterPort").SpawnInstance_
        obj.Name = "IP_" & tmp_ip
        obj.HostAddress = tmp_ip
        
        Select Case tmp_protocol
            Case "RAW"
                obj.Protocol = 1
                obj.PortNumber = "515"
            Case "LPR"
                obj.Protocol = 2
                obj.Queue = "PASSTHRU"
            Case Else
                obj.Protocol = 2
                obj.Queue = "PASSTHRU"
        End Select
        
        obj.SNMPEnabled = False
        obj.SNMPCommunity = "public"
        obj.SNMPDevIndex = 1
        obj.Put_
    
    Case "ACTION:DELETE"

        If MsgBox("是否要刪除 " & Src_PrinterTCPIPPort_Name & " 於 " & Can_Process_Computer & " 上？", vbQuestion + vbYesNo, "確認") = vbYes Then

            Set cols = objWMIService.ExecQuery("Select * from Win32_TCPIPPrinterPort Where Name = '" & Src_PrinterTCPIPPort_Name & "'")
            For Each obj In cols
                Action_PrinterTCPIPPort = CStr(obj.Delete)
            Next

        End If
        
End Select

End Function


Function Action_Software(Src_Software_Name As String, Src_State As String) As String
'對軟體做處理
    
Dim cols
            
Select Case Src_State
    
    Case "ACTION:ADD"
        
        '顯示選軟體封包資料夾視窗
        Dim sFile As String
        Dim cc As New cCommonDialog
        If cc.VBGetOpenFileName(sFile, , , , , , "軟體安裝檔 (*.MSI)|*.MSI|軟體安裝檔 (*.EXE)|*.EXE", , , "選擇要安裝的軟體", "MSI;EXE", Screen.ActiveForm.hwnd, OFN_HIDEREADONLY) Then
        End If
        
        '有選到檔案才動作
        If Len(sFile) > 0 Then
        
            'Set cols = objWMIService.ExecQuery("Win32_Product")
            'Action_Software = CStr(cols.Install(sFile, , True))
            
            Dim tmp_SID As String: tmp_SID = InputBox("請輸入擁有管理者權限的帳號：", "提示", "test\Administrator")
            Dim tmp_Pwd As String: tmp_Pwd = InputBox("請輸入密碼：", "提示")
            
            Const wbemImpersonationLevelDelegate = 4
            Set objwbemLocator = CreateObject("WbemScripting.SWbemLocator")
            Set objConnection = objwbemLocator.ConnectServer _
                (Can_Process_Computer, "root\cimv2", tmp_SID, tmp_Pwd, , "kerberos:" & Can_Process_Computer)
            
            objConnection.Security_.ImpersonationLevel = wbemImpersonationLevelDelegate
            
            Set objSoftware = objConnection.Get("Win32_Product")
            Action_Software = CStr(objSoftware.Install(sFile, , True))

        
        End If
        
    
    Case "ACTION:DELETE"
    
        If MsgBox("是否要移除 " & Src_Software_Name & " 於 " & Can_Process_Computer & " 上？", vbQuestion + vbYesNo, "確認") = vbYes Then
            
            Set cols = objWMIService.ExecQuery("Select * from Win32_Product Where Name = '" & Src_Software_Name & "'")
            For Each obj In cols
                Action_Software = CStr(obj.Uninstall)
            Next

        End If
    
End Select

End Function


Function Action_Local_Admin(Src_Name As String, Src_State As String) As String
'變更本機管理原密碼

Select Case Src_State
    
    '變更管理員帳號密碼
    Case "ACTION:CHANGE_ADMIN_PASSWORD"
        
        If Can_Process_Computer = "" Then MsgBox "沒有選擇要處理的電腦", vbInformation, "錯誤": Exit Function
        
        Dim tmp_Pwd As String: tmp_Pwd = InputBox("請輸入要變更的密碼 : ", "變更管理員密碼")
        
        If Len(tmp_Pwd) = 0 Then
            MsgBox "變更作業取消", vbInformation, "訊息"
        Else
            On Error Resume Next
            Dim obj: Set obj = GetObject("WinNT://" & Can_Process_Computer & "/Administrator")
            obj.SetPassword (tmp_Pwd)
            If Err.Number <> 0 Then
                MsgBox "變更失敗，請檢查是否有權限或是該電腦上帳號正確性。", vbInformation, "錯誤"
                Err.Clear
            Else
                MsgBox "變更成功", vbInformation, "完成"
            End If
        End If
        
End Select
End Function


Function Batch_Action_Local_Admin(Src_Name As String, Src_State As String, Src_Computers As String) As String
'批次變更本機管理原密碼

If Trim(Src_Computers) = "" Then
    MsgBox "沒有選擇要處理的電腦", vbInformation, "錯誤"
    Exit Function
End If

'取得電腦勾選的清單
Dim tmp_com() As String: tmp_com = Split(Src_Computers, ",")

'處理失敗的電腦清單
Dim tmp_com_Fail As String

'處理成功的電腦清單
Dim tmp_com_Success As String

Select Case Src_State

    '變更管理員帳號密碼
    Case "ACTION:CHANGE_ADMIN_PASSWORD"

        Dim tmp_Pwd As String: tmp_Pwd = InputBox("請輸入要變更的密碼 : ", "變更管理員密碼")
        If Len(tmp_Pwd) = 0 Then
        
            MsgBox "變更作業取消", vbInformation, "訊息"
        
        Else
        
            On Error Resume Next
            
            Dim i As Long
            '開始跑各台電腦
            For i = 0 To UBound(tmp_com) - 1
            
                Show_Msg "正在變更 " & tmp_com(i) & " 的本機管理員密碼... " & i + 1 & "/" & UBound(tmp_com)
                
                Dim obj: Set obj = GetObject("WinNT://" & tmp_com(i) & "/Administrator")
                obj.SetPassword tmp_Pwd
                
                If Err.Number <> 0 Then
                    tmp_com_Fail = tmp_com_Fail & tmp_com(i) & vbCrLf
                    Err.Clear
                Else
                    tmp_com_Success = tmp_com_Success & tmp_com(i) & vbCrLf
                End If
                
                DoEvents
                
            Next
            
            Show_Msg "處理結束"
            
            Show_Batch_Log _
                "成功：" & vbCrLf & tmp_com_Success & vbCrLf & _
                "失敗：" & vbCrLf & tmp_com_Fail, _
                "批次變更本機管理員密碼作業結果如下"

        End If

End Select
End Function


Function Batch_Action_Service(Src_State As String, Src_Computers As String) As String
'批次變更本機管理原密碼

If Trim(Src_Computers) = "" Then
    MsgBox "沒有選擇要處理的電腦", vbInformation, "錯誤"
    Exit Function
End If

'取得電腦勾選的清單
Dim tmp_com() As String: tmp_com = Split(Src_Computers, ",")

'處理失敗的電腦清單
Dim tmp_com_Fail As String

'處理成功的電腦清單
Dim tmp_com_Success As String

Dim cols

Dim tmp_over As Boolean
Dim tmp_ret As String


        Dim tmp_Service As String: tmp_Service = InputBox("請輸入要處理服務名稱 : ", "服務")
        If Len(tmp_Service) = 0 Then
        
            MsgBox "變更作業取消", vbInformation, "訊息"
        
        Else
        
            '開始跑各台電腦
            Dim i As Long
            For i = 0 To UBound(tmp_com) - 1
            
                Show_Msg "正在處理 " & tmp_com(i) & " 的 " & tmp_Service & " 服務... " & i + 1 & "/" & UBound(tmp_com)
                
                '無法跟目標電腦建立連線
                If WMI_Service_Create(tmp_com(i)) = False Then
                    tmp_com_Fail = tmp_com_Fail & tmp_com(i) & vbCrLf
                Else
                    
                    '電腦可連線，開始停止服務
                    Set cols = objWMIService.ExecQuery("Select * from Win32_Service Where Name = '" & tmp_Service & "'")
                    For Each obj In cols
                        
                        Select Case Src_State

                            '開始服務
                            Case "ACTION:START": tmp_ret = CStr(obj.StartService)
                        
                            '停止服務
                            Case "ACTION:STOP": tmp_ret = CStr(obj.StopService)
                        
                        End Select

                        
                        If tmp_ret = "0" Then
                            tmp_over = True
                        Else
                            tmp_over = False
                        End If
                    Next
                
                    '判定服務處理結果
                    If tmp_over = True Then tmp_com_Success = tmp_com_Success & tmp_com(i) & vbCrLf
                    If tmp_over = False Then tmp_com_Fail = tmp_com_Fail & tmp_com(i) & vbCrLf
                    
                End If
                DoEvents
            Next
            
            Show_Msg "處理結束"
            
            Show_Batch_Log _
                "成功：" & vbCrLf & tmp_com_Success & vbCrLf & _
                "失敗：" & vbCrLf & tmp_com_Fail, _
                "批次處理服務作業結果如下"

        End If


End Function

Function Batch_Action_Process(Src_State As String, Src_Computers As String) As String
'批次變更執行緒

If Trim(Src_Computers) = "" Then
    MsgBox "沒有選擇要處理的電腦", vbInformation, "錯誤"
    Exit Function
End If

'取得電腦勾選的清單
Dim tmp_com() As String: tmp_com = Split(Src_Computers, ",")

'處理失敗的電腦清單
Dim tmp_com_Fail As String

'處理成功的電腦清單
Dim tmp_com_Success As String
Dim cols
Dim tmp_over As Boolean
Dim tmp_ret As String


        Dim tmp_Process As String: tmp_Process = InputBox("請輸入要處理的執行緒完整檔名 (含副檔名) : ", "執行緒")
        
        '建立
        If Src_State = "ACTION:CREATE" Then
            Dim tmp_cmd As String
            Dim tmp_Exec As String: tmp_Exec = tmp_Process
            Dim tmp_Active As String: tmp_Active = InputBox("要開啟於前台或是隱藏此執行緒？1=開啟,0=隱藏", "顯示", "1")
            Dim tmp_Runat As String: tmp_Runat = InputBox("於什麼時候執行？格式 HH:MM (會以你輸入的時間再延後一分鐘)", "時間點", Format(Time, "hh:mm"))
            
            If tmp_Exec = "" Then Batch_Action_Process = "-1": Exit Function
            If tmp_Active = "" Then Batch_Action_Process = "-1": Exit Function
            If tmp_Runat = "" Then Batch_Action_Process = "-1": Exit Function
            
            tmp_Runat = Format(DateAdd("n", "1", tmp_Runat), "hh:mm")
            tmp_cmd = "Cmd.exe /c " & Chr(34) & "AT " & tmp_Runat & IIf(tmp_Active = "1", " /interactive ", " ") & tmp_Exec & Chr(34)
    
        End If
        
        
        
        If Len(tmp_Process) = 0 Then
        
            MsgBox "變更作業取消", vbInformation, "訊息"
        
        Else
        
            '開始跑各台電腦
            Dim i As Long
            For i = 0 To UBound(tmp_com) - 1
            
                Show_Msg "正在處理 " & tmp_com(i) & " 的 " & tmp_Process & " 執行緒... " & i + 1 & "/" & UBound(tmp_com)
                
                '無法跟目標電腦建立連線
                If WMI_Service_Create(tmp_com(i)) = False Then
                    tmp_com_Fail = tmp_com_Fail & tmp_com(i) & vbCrLf
                Else
                 
                        Select Case Src_State

                            '建立執行緒
                            Case "ACTION:CREATE"
                            
                            
                                   Set obj = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & tmp_com(i) & "\root\cimv2:Win32_Process")
                                   tmp_ret = CStr(obj.Create(tmp_cmd, Null, Null, intProcessID))
                       
                            '停止服務
                            Case "ACTION:TERMINATE"
                               
                                '電腦可連線，開始停止服務
                                Set cols = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & tmp_Process & "'")
                                For Each obj In cols
                                    tmp_ret = CStr(obj.Terminate)
                                Next
                                'If cols.Count = 0 Then tmp_Ret = "-1"
                        
                        End Select

                        If tmp_ret = "0" Then tmp_over = True
                        If tmp_ret <> "0" Then tmp_over = False
                
                    '判定服務處理結果
                    If tmp_over = True Then tmp_com_Success = tmp_com_Success & tmp_com(i) & vbCrLf
                    If tmp_over = False Then tmp_com_Fail = tmp_com_Fail & tmp_com(i) & vbCrLf
                    
                End If
                DoEvents
            Next
         
            Show_Msg "處理結束"
            
            Show_Batch_Log _
                "成功：" & vbCrLf & tmp_com_Success & vbCrLf & _
                "失敗：" & vbCrLf & tmp_com_Fail, _
                "批次處理執行緒作業結果如下"

        End If


End Function
