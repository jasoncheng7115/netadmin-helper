Attribute VB_Name = "Mod_WMI_Action"
Function Action_Service(Src_Service_Name As String, Src_State As String) As String
'��A�Ȱ��B�z

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
'���������B�z

    
Select Case Src_State
    
            
    'action:CREATE
    Case "ACTION:CREATE"
    
        Dim tmp_cmd As String
        Dim tmp_Exec As String: tmp_Exec = InputBox("�п�J�ؼа����ɪ����|�P�ɦW", "�إ߰����", "notepad.exe")
        Dim tmp_Active As String: tmp_Active = InputBox("�n�}�ҩ�e�x�άO���æ�������H1=�}��,0=����", "���", "1")
        Dim tmp_Runat As String: tmp_Runat = InputBox("�󤰻�ɭ԰���H�榡 HH:MM (�|�H�A��J���ɶ��A����@����)", "�ɶ��I", Format(Time, "hh:mm"))
        
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
'��A�Ȱ��}�����B�z

    If MsgBox("�z�T�w�n�N " & Can_Process_Computer & " ���s�}�� / ���� �ܡH", vbQuestion + vbYesNo, "����") = vbYes Then
            
        Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\" & Can_Process_Computer & "\root\cimv2")
        
        Set cols = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
            
        Select Case Src_State
                    
            'action:���}
            Case "ACTION:REBOOT"
            
                For Each obj In cols
                    obj.Reboot
                Next
            
            'action:����
            Case "ACTION:SHUTDOWN"
            
                For Each obj In cols
                    obj.Shutdown
                Next
         
        End Select

    End If
    
End Function


Function Action_Printer(Src_PrinterDriver_Name As String, Src_State As String) As String
'��L����B�z
    
Select Case Src_State
    
            
    Case "ACTION:ADD"
        
        Frm_Manager_Printer_AddPrinter.Show 1

    Case "ACTION:DELETE"
        
        If MsgBox("�z�T�w�R�� " & Can_Process_Computer & " �W�� " & Src_PrinterDriver_Name & "�H", vbQuestion + vbYesNo, "����") = vbYes Then

            Dim cols
            Set cols = objWMIService.ExecQuery("Select * from Win32_Printer where Name = '" & Trim(Src_PrinterDriver_Name) & "'")
            
            For Each obj In cols
                obj.Delete_
            Next

        End If
        
End Select
End Function

Function Action_PrinterDriver(Src_PrinterDriver_Name As String, Src_State As String) As String
'��L����X�ʵ{���B�z
    
Select Case Src_State
    
            
    Case "ACTION:ADD"
        
        '��ܿ��X�ʵ{����Ƨ�����
        Dim sFile As String
        Dim cc As New cCommonDialog
        If cc.VBGetOpenFileName(sFile, , , , , , "�X�ʵ{����T�� (*.INF)|*.INF", , , "��ܦL����X�ʵ{��", "INF", Screen.ActiveForm.hwnd, OFN_HIDEREADONLY) Then
            
        End If

        '������ɮפ~�ʧ@
        If Len(sFile) > 0 Then
        
            Dim sPath As String:  sPath = Left(sFile, InStrRev(sFile, "\"))
            
            '���o�X�ʵ{������ DRVNAME
            Dim fs As New FileSystemObject
            Dim a As TextStream
            Set a = fs.OpenTextFile(sFile, ForReading, False)
            Dim tmp_File As String: tmp_File = a.ReadAll
            
            '�]�w�}�l�䪺�Ϭq
            Dim tmp_Serach_Start As Long: tmp_Serach_Start = InStr(1, tmp_File, "[Strings]")
                        
            
            '�]�w DRVNAME ��T�Ӧ檺��m
            Dim tmp_Drvname_Start As Long: tmp_Drvname_Start = InStr(tmp_Serach_Start, tmp_File, "DRVNAME")
            Dim tmp_Drvname_End As Long: tmp_Drvname_End = InStr(tmp_Drvname_Start, tmp_File, vbCrLf)
            
            '���X��q
            Dim tmp_Drvname_Str As String: tmp_Drvname_Str = Mid(tmp_File, tmp_Drvname_Start, tmp_Drvname_End - tmp_Drvname_Start)
             
            '�簣�������n���r
            tmp_Drvname_Str = Replace(tmp_Drvname_Str, "DRVNAME", "")
            tmp_Drvname_Str = Replace(tmp_Drvname_Str, "=", "")
            tmp_Drvname_Str = Trim(Replace(tmp_Drvname_Str, """", ""))
            
            a.Close
            Set fs = Nothing
                        
            
            If MsgBox("�O�_�n�w�� " & tmp_Drvname_Str & " �� " & Can_Process_Computer & " �W�H", vbQuestion + vbYesNo, "�T�{") = vbYes Then
            
                '�g�J
                
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
'��L��������s����B�z
    
Select Case Src_State
            
    Case "ACTION:ADD"
        
        Dim tmp_ip As String:  tmp_ip = InputBox("�п�J IP:", "�s�W TCPIP �s����")
        Dim tmp_protocol As String: tmp_protocol = InputBox("�п�ܳq�T��w: " & vbCrLf & "RAW �� LPR (�w�])", "�s�W TCPIP �s����", "LPR")
        
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

        If MsgBox("�O�_�n�R�� " & Src_PrinterTCPIPPort_Name & " �� " & Can_Process_Computer & " �W�H", vbQuestion + vbYesNo, "�T�{") = vbYes Then

            Set cols = objWMIService.ExecQuery("Select * from Win32_TCPIPPrinterPort Where Name = '" & Src_PrinterTCPIPPort_Name & "'")
            For Each obj In cols
                Action_PrinterTCPIPPort = CStr(obj.Delete)
            Next

        End If
        
End Select

End Function


Function Action_Software(Src_Software_Name As String, Src_State As String) As String
'��n�鰵�B�z
    
Dim cols
            
Select Case Src_State
    
    Case "ACTION:ADD"
        
        '��ܿ�n��ʥ]��Ƨ�����
        Dim sFile As String
        Dim cc As New cCommonDialog
        If cc.VBGetOpenFileName(sFile, , , , , , "�n��w���� (*.MSI)|*.MSI|�n��w���� (*.EXE)|*.EXE", , , "��ܭn�w�˪��n��", "MSI;EXE", Screen.ActiveForm.hwnd, OFN_HIDEREADONLY) Then
        End If
        
        '������ɮפ~�ʧ@
        If Len(sFile) > 0 Then
        
            'Set cols = objWMIService.ExecQuery("Win32_Product")
            'Action_Software = CStr(cols.Install(sFile, , True))
            
            Dim tmp_SID As String: tmp_SID = InputBox("�п�J�֦��޲z���v�����b���G", "����", "test\Administrator")
            Dim tmp_Pwd As String: tmp_Pwd = InputBox("�п�J�K�X�G", "����")
            
            Const wbemImpersonationLevelDelegate = 4
            Set objwbemLocator = CreateObject("WbemScripting.SWbemLocator")
            Set objConnection = objwbemLocator.ConnectServer _
                (Can_Process_Computer, "root\cimv2", tmp_SID, tmp_Pwd, , "kerberos:" & Can_Process_Computer)
            
            objConnection.Security_.ImpersonationLevel = wbemImpersonationLevelDelegate
            
            Set objSoftware = objConnection.Get("Win32_Product")
            Action_Software = CStr(objSoftware.Install(sFile, , True))

        
        End If
        
    
    Case "ACTION:DELETE"
    
        If MsgBox("�O�_�n���� " & Src_Software_Name & " �� " & Can_Process_Computer & " �W�H", vbQuestion + vbYesNo, "�T�{") = vbYes Then
            
            Set cols = objWMIService.ExecQuery("Select * from Win32_Product Where Name = '" & Src_Software_Name & "'")
            For Each obj In cols
                Action_Software = CStr(obj.Uninstall)
            Next

        End If
    
End Select

End Function


Function Action_Local_Admin(Src_Name As String, Src_State As String) As String
'�ܧ󥻾��޲z��K�X

Select Case Src_State
    
    '�ܧ�޲z���b���K�X
    Case "ACTION:CHANGE_ADMIN_PASSWORD"
        
        If Can_Process_Computer = "" Then MsgBox "�S����ܭn�B�z���q��", vbInformation, "���~": Exit Function
        
        Dim tmp_Pwd As String: tmp_Pwd = InputBox("�п�J�n�ܧ󪺱K�X : ", "�ܧ�޲z���K�X")
        
        If Len(tmp_Pwd) = 0 Then
            MsgBox "�ܧ�@�~����", vbInformation, "�T��"
        Else
            On Error Resume Next
            Dim obj: Set obj = GetObject("WinNT://" & Can_Process_Computer & "/Administrator")
            obj.SetPassword (tmp_Pwd)
            If Err.Number <> 0 Then
                MsgBox "�ܧ󥢱ѡA���ˬd�O�_���v���άO�ӹq���W�b�����T�ʡC", vbInformation, "���~"
                Err.Clear
            Else
                MsgBox "�ܧ󦨥\", vbInformation, "����"
            End If
        End If
        
End Select
End Function


Function Batch_Action_Local_Admin(Src_Name As String, Src_State As String, Src_Computers As String) As String
'�妸�ܧ󥻾��޲z��K�X

If Trim(Src_Computers) = "" Then
    MsgBox "�S����ܭn�B�z���q��", vbInformation, "���~"
    Exit Function
End If

'���o�q���Ŀ諸�M��
Dim tmp_com() As String: tmp_com = Split(Src_Computers, ",")

'�B�z���Ѫ��q���M��
Dim tmp_com_Fail As String

'�B�z���\���q���M��
Dim tmp_com_Success As String

Select Case Src_State

    '�ܧ�޲z���b���K�X
    Case "ACTION:CHANGE_ADMIN_PASSWORD"

        Dim tmp_Pwd As String: tmp_Pwd = InputBox("�п�J�n�ܧ󪺱K�X : ", "�ܧ�޲z���K�X")
        If Len(tmp_Pwd) = 0 Then
        
            MsgBox "�ܧ�@�~����", vbInformation, "�T��"
        
        Else
        
            On Error Resume Next
            
            Dim i As Long
            '�}�l�]�U�x�q��
            For i = 0 To UBound(tmp_com) - 1
            
                Show_Msg "���b�ܧ� " & tmp_com(i) & " �������޲z���K�X... " & i + 1 & "/" & UBound(tmp_com)
                
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
            
            Show_Msg "�B�z����"
            
            Show_Batch_Log _
                "���\�G" & vbCrLf & tmp_com_Success & vbCrLf & _
                "���ѡG" & vbCrLf & tmp_com_Fail, _
                "�妸�ܧ󥻾��޲z���K�X�@�~���G�p�U"

        End If

End Select
End Function


Function Batch_Action_Service(Src_State As String, Src_Computers As String) As String
'�妸�ܧ󥻾��޲z��K�X

If Trim(Src_Computers) = "" Then
    MsgBox "�S����ܭn�B�z���q��", vbInformation, "���~"
    Exit Function
End If

'���o�q���Ŀ諸�M��
Dim tmp_com() As String: tmp_com = Split(Src_Computers, ",")

'�B�z���Ѫ��q���M��
Dim tmp_com_Fail As String

'�B�z���\���q���M��
Dim tmp_com_Success As String

Dim cols

Dim tmp_over As Boolean
Dim tmp_ret As String


        Dim tmp_Service As String: tmp_Service = InputBox("�п�J�n�B�z�A�ȦW�� : ", "�A��")
        If Len(tmp_Service) = 0 Then
        
            MsgBox "�ܧ�@�~����", vbInformation, "�T��"
        
        Else
        
            '�}�l�]�U�x�q��
            Dim i As Long
            For i = 0 To UBound(tmp_com) - 1
            
                Show_Msg "���b�B�z " & tmp_com(i) & " �� " & tmp_Service & " �A��... " & i + 1 & "/" & UBound(tmp_com)
                
                '�L�k��ؼйq���إ߳s�u
                If WMI_Service_Create(tmp_com(i)) = False Then
                    tmp_com_Fail = tmp_com_Fail & tmp_com(i) & vbCrLf
                Else
                    
                    '�q���i�s�u�A�}�l����A��
                    Set cols = objWMIService.ExecQuery("Select * from Win32_Service Where Name = '" & tmp_Service & "'")
                    For Each obj In cols
                        
                        Select Case Src_State

                            '�}�l�A��
                            Case "ACTION:START": tmp_ret = CStr(obj.StartService)
                        
                            '����A��
                            Case "ACTION:STOP": tmp_ret = CStr(obj.StopService)
                        
                        End Select

                        
                        If tmp_ret = "0" Then
                            tmp_over = True
                        Else
                            tmp_over = False
                        End If
                    Next
                
                    '�P�w�A�ȳB�z���G
                    If tmp_over = True Then tmp_com_Success = tmp_com_Success & tmp_com(i) & vbCrLf
                    If tmp_over = False Then tmp_com_Fail = tmp_com_Fail & tmp_com(i) & vbCrLf
                    
                End If
                DoEvents
            Next
            
            Show_Msg "�B�z����"
            
            Show_Batch_Log _
                "���\�G" & vbCrLf & tmp_com_Success & vbCrLf & _
                "���ѡG" & vbCrLf & tmp_com_Fail, _
                "�妸�B�z�A�ȧ@�~���G�p�U"

        End If


End Function

Function Batch_Action_Process(Src_State As String, Src_Computers As String) As String
'�妸�ܧ�����

If Trim(Src_Computers) = "" Then
    MsgBox "�S����ܭn�B�z���q��", vbInformation, "���~"
    Exit Function
End If

'���o�q���Ŀ諸�M��
Dim tmp_com() As String: tmp_com = Split(Src_Computers, ",")

'�B�z���Ѫ��q���M��
Dim tmp_com_Fail As String

'�B�z���\���q���M��
Dim tmp_com_Success As String
Dim cols
Dim tmp_over As Boolean
Dim tmp_ret As String


        Dim tmp_Process As String: tmp_Process = InputBox("�п�J�n�B�z������������ɦW (�t���ɦW) : ", "�����")
        
        '�إ�
        If Src_State = "ACTION:CREATE" Then
            Dim tmp_cmd As String
            Dim tmp_Exec As String: tmp_Exec = tmp_Process
            Dim tmp_Active As String: tmp_Active = InputBox("�n�}�ҩ�e�x�άO���æ�������H1=�}��,0=����", "���", "1")
            Dim tmp_Runat As String: tmp_Runat = InputBox("�󤰻�ɭ԰���H�榡 HH:MM (�|�H�A��J���ɶ��A����@����)", "�ɶ��I", Format(Time, "hh:mm"))
            
            If tmp_Exec = "" Then Batch_Action_Process = "-1": Exit Function
            If tmp_Active = "" Then Batch_Action_Process = "-1": Exit Function
            If tmp_Runat = "" Then Batch_Action_Process = "-1": Exit Function
            
            tmp_Runat = Format(DateAdd("n", "1", tmp_Runat), "hh:mm")
            tmp_cmd = "Cmd.exe /c " & Chr(34) & "AT " & tmp_Runat & IIf(tmp_Active = "1", " /interactive ", " ") & tmp_Exec & Chr(34)
    
        End If
        
        
        
        If Len(tmp_Process) = 0 Then
        
            MsgBox "�ܧ�@�~����", vbInformation, "�T��"
        
        Else
        
            '�}�l�]�U�x�q��
            Dim i As Long
            For i = 0 To UBound(tmp_com) - 1
            
                Show_Msg "���b�B�z " & tmp_com(i) & " �� " & tmp_Process & " �����... " & i + 1 & "/" & UBound(tmp_com)
                
                '�L�k��ؼйq���إ߳s�u
                If WMI_Service_Create(tmp_com(i)) = False Then
                    tmp_com_Fail = tmp_com_Fail & tmp_com(i) & vbCrLf
                Else
                 
                        Select Case Src_State

                            '�إ߰����
                            Case "ACTION:CREATE"
                            
                            
                                   Set obj = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & tmp_com(i) & "\root\cimv2:Win32_Process")
                                   tmp_ret = CStr(obj.Create(tmp_cmd, Null, Null, intProcessID))
                       
                            '����A��
                            Case "ACTION:TERMINATE"
                               
                                '�q���i�s�u�A�}�l����A��
                                Set cols = objWMIService.ExecQuery("Select * from Win32_Process Where Name = '" & tmp_Process & "'")
                                For Each obj In cols
                                    tmp_ret = CStr(obj.Terminate)
                                Next
                                'If cols.Count = 0 Then tmp_Ret = "-1"
                        
                        End Select

                        If tmp_ret = "0" Then tmp_over = True
                        If tmp_ret <> "0" Then tmp_over = False
                
                    '�P�w�A�ȳB�z���G
                    If tmp_over = True Then tmp_com_Success = tmp_com_Success & tmp_com(i) & vbCrLf
                    If tmp_over = False Then tmp_com_Fail = tmp_com_Fail & tmp_com(i) & vbCrLf
                    
                End If
                DoEvents
            Next
         
            Show_Msg "�B�z����"
            
            Show_Batch_Log _
                "���\�G" & vbCrLf & tmp_com_Success & vbCrLf & _
                "���ѡG" & vbCrLf & tmp_com_Fail, _
                "�妸�B�z������@�~���G�p�U"

        End If


End Function
