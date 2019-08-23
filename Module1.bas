Attribute VB_Name = "Mod_WMI_LDAP"
'�\�h�\�ॲ�ݦb XP �� .NET Server �H�W��i����
'�i�Ѩt�θ�T�P�_�A5.1 �H�W�~�i
'
Public AD_Domain_Name  As String

Public Can_Process_Computer As String

Public canUseAD As Boolean
Public objWMIService

'�Ȧs�@�~�t�Ϊ���
Public OS_Version As Single

Public Sub Delay(ByVal N As Single)
    
    '����Ƶ{��
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
'�C�|���줤�q��
    
On Error GoTo ErrMsg

    '�ѪR����W��
    Dim i As Long: i = 0
    Dim tmp_str() As String
    Dim tmp_Domain_Name As String: tmp_Domain_Name = ""
    tmp_str = Split(Trim(AD_Domain_Name), ".")
    For i = LBound(tmp_str) To UBound(tmp_str)
        tmp_Domain_Name = tmp_Domain_Name & "DC=" & tmp_str(i)
        If i < UBound(tmp_str) Then tmp_Domain_Name = tmp_Domain_Name & ","
    Next
    
    '�إ� ADO �s�u
    Set objConnection = CreateObject("ADODB.Connection")
    objConnection.Provider = "ADsDSOObject"
    objConnection.Open "Active Directory Provider"
    
    '�}�Ҭd��
    Set objCommand = CreateObject("ADODB.Command")
    Set objCommand.ActiveConnection = objConnection
    
    Const ADS_SCOPE_SUBTREE = 2
    
    '�w�q ADO �Ѽ�
    objCommand.Properties("Page Size") = 1000
    objCommand.Properties("Timeout") = 30
    objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE
    objCommand.Properties("Cache Results") = False
    
    '�e�X�V AD �n�D�q�����M��
    objCommand.CommandText = _
        "SELECT Name, Location FROM 'LDAP://" & tmp_Domain_Name & "' " & _
        "WHERE objectClass='computer' ORDER BY Name"
    Set objRecordSet = objCommand.Execute
    
    '�Ȯɦs��q���M��
    Dim tmp_Computers As New Collection
    
    '�}�l���X���
    objRecordSet.MoveFirst
    Do Until objRecordSet.EOF
        i = i + 1
        
        '�N�q���[�J���X
        tmp_Computers.Add objRecordSet.Fields("Name").Value, "i" & i
        DoEvents
        objRecordSet.MoveNext
    Loop
    Set objConnection = Nothing
    
    '�N���X�Ǧ^
    Set ADSI_Computer_List = tmp_Computers
    
    canUseAD = True
    
Exit Function

ErrMsg:

    If Err.Number = -2147217865 Then
        'MsgBox "�L�k�s������", vbInformation, "���~"
    Else
        MsgBox "�o�ͨ䥦���~ : " & Err.Description & "(" & Err.Number & ")", vbInformation, "���~"
    End If
    
    canUseAD = False

End Function

Function ADSI_User_List() As Collection
'�C�|���줤 User

On Error Resume Next
    
    '�ѪR����W��
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
'�˵��q���O�_�i��

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
'�d�ݧ@�~�t��

On Error Resume Next
  
    Set cols = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    
    For Each obj In cols

        Frm_Main.SGrid_Computer_List_AddRow "�@�~�t��", obj.Caption & " " & obj.Version, "�@�~�t��"
        OS_Version = CSng(Left(obj.Version, 3))
        
    Next


End Function

Function WMI_Computer_Login_UserName(strComputer As String)
'���o�n�J�̦W��

On Error GoTo ErrMsg

    Set colServices = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    
    If colServices.Count > 0 Then
    
       For Each objService In colServices
           Frm_Main.SGrid_Computer_List_AddRow "�n�J�b��", objService.UserName, "�@��"
           DoEvents
       Next
    
    Else
       
       Frm_Main.SGrid_Computer_List_AddRow "�n�J�b��", "�L", "�@��"
    
    End If

Exit Function

ErrMsg:

    Frm_Main.SGrid_Computer_List_AddRow "�n�J�b��", "�L", "�@��"
    
End Function



Function WMI_Computer_System_Information(strComputer As String) As String()
'�t�θ�T

On Error Resume Next
    
    Dim i As Long
    
    Set colSettings = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    
    For Each objOperatingSystem In colSettings
        
        
        Frm_Main.SGrid_Computer_List_AddRow "�@�~�t�Ϊ���", objOperatingSystem.Version, "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "Service Pack", objOperatingSystem.ServicePackMajorVersion _
                                                        & "." & objOperatingSystem.ServicePackMinorVersion, "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "�@�~�t�λs�y��", objOperatingSystem.Manufacturer, "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "Windows �ؿ�", objOperatingSystem.WindowsDirectory, "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "��O�X", strLocale(objOperatingSystem.Locale) & " - (" & objOperatingSystem.Locale & ")", "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "��X", objOperatingSystem.CountryCode, "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "�r�X��", strCodepage(objOperatingSystem.CodeSet) & " - (" & objOperatingSystem.CodeSet & ")", "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "��ư��樾�� (DEP)", IIf(objOperatingSystem.DataExecutionPrevention_Available = True, "�ҥ�", "����"), "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "��ư��樾�� (DEP) - �X�ʵ{��", IIf(objOperatingSystem.DataExecutionPrevention_Drivers = True, "�ҥ�", "����"), "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "�[�K����", objOperatingSystem.EncryptionLevel & " �줸", "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "�t�Φw�ˤ��", Change_GMT(objOperatingSystem.InstallDate), "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "�̪�@���}��", Change_GMT(objOperatingSystem.LastBootUpTime), "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "���\�̤j������ƥ�", objOperatingSystem.MaxNumberOfProcesses, "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "���\�̤j������O����j�p", Format_MB_By_K(objOperatingSystem.MaxProcessMemorySize), "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "�ثe������ƥ�", objOperatingSystem.NumberOfProcesses, "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "�ثe�ϥΪ̼ƥ�", objOperatingSystem.NumberOfUsers, "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "��´", objOperatingSystem.NumberOfUsers, "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "�@�~�t�λy��", strOsLang(objOperatingSystem.OSLanguage) & " - (" & objOperatingSystem.OSLanguage & ")", "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "���U�ϥΪ�", objOperatingSystem.RegisteredUser, "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "�Ҧ��ϺФ�����", Format_MB_By_K(objOperatingSystem.SizeStoredInPagingFiles), "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "����O����Ѿl", Format_MB_By_K(objOperatingSystem.FreePhysicalMemory), "�O����"
        Frm_Main.SGrid_Computer_List_AddRow "�����O����j�p", Format_MB_By_K(objOperatingSystem.TotalVirtualMemorySize), "�O����"
        Frm_Main.SGrid_Computer_List_AddRow "�����O����Ѿl", Format_MB_By_K(objOperatingSystem.FreeVirtualMemory), "�O����"
        
        DoEvents
    Next
    
    Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    
    For Each objComputer In colSettings
        
        'Frm_Main.SGrid_Computer_List_AddRow "�q���W��", objComputer.Name, "�@��"
        Frm_Main.SGrid_Computer_List_AddRow "�s�y��", objComputer.Manufacturer, "�w��"
        Frm_Main.SGrid_Computer_List_AddRow "�W��", objComputer.Model, "�w��"
        Frm_Main.SGrid_Computer_List_AddRow "BootROM", IIf(objComputer.BootROMSupported = True, "�䴩", "���䴩"), "�w��"
        Frm_Main.SGrid_Computer_List_AddRow "�ɰ�", "GMT " & IIf(objComputer.CurrentTimeZone > 0, "+", "-") & objComputer.CurrentTimeZone / 60, "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "����`���ɶ�", IIf(objComputer.EnableDaylightSavingsTime = True, "�ҥ�", "����"), "�@�~�t��"
        Frm_Main.SGrid_Computer_List_AddRow "�����B�z���ƥ�", objComputer.NumberOfProcessors, "�w��"
        Frm_Main.SGrid_Computer_List_AddRow "����O����j�p", Format_MB_By_B(objComputer.TotalPhysicalMemory), "�O����"
        DoEvents
        
    Next
    
    i = 0
    Set colSettings = objWMIService.ExecQuery("Select * from Win32_Processor")
    For Each objprocessor In colSettings
        i = i + 1
        Frm_Main.SGrid_Computer_List_AddRow "�W��", Trim(objprocessor.Name), "�w�� - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "Socket ����", objprocessor.SocketDesignation, "�w�� - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "�̰��ɯ�", objprocessor.MaxClockSpeed & " MHz", "�w�� - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "�зǮɯ�", objprocessor.CurrentClockSpeed & " MHz", "�w�� - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "�~�W", objprocessor.ExtClock & " MHz", "�w�� - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "����W�e", objprocessor.DataWidth & " Bits", "�w�� - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "�ثe�t��", objprocessor.LoadPercentage & " %", "�w�� - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "L2 Cache", objprocessor.L2CacheSize & " K", "�w�� - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "L2 Cache Speed", IIf(IsNull(objprocessor.L2CacheSpeed), "0", objprocessor.L2CacheSpeed) & " MHz", "�w�� - CPU " & i
        Frm_Main.SGrid_Computer_List_AddRow "�q�O�޲z", IIf(objprocessor.PowerManagementSupported, "�䴩", "���䴩"), "�w�� - CPU " & i
        
        DoEvents
    Next
    i = 0
    
    Set colSettings = objWMIService.ExecQuery("Select * from Win32_BIOS")
    For Each objBIOS In colSettings
        Frm_Main.SGrid_Computer_List_AddRow "BIOS ����", objBIOS.Version, "�w��"
        DoEvents
    Next
    

End Function


Function WMI_Computer_Product_Installed(strComputer As String)
On Error Resume Next
    
    Set colSoftware = objWMIService.ExecQuery("Select * from Win32_Product")
    
    For Each objSoftware In colSoftware
        
        Frm_Main.SGrid_Computer_List_AddRow objSoftware.Caption & " " & objSoftware.Version, objSoftware.InstallLocation, "�n��w��"
        DoEvents
    Next


End Function


Function WMI_Computer_LogicalDisk(srcComputer As String)
'�C�X�޿�Ϻо�

On Error Resume Next
    
    Set colLogicalDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk")
    For Each objdisk In colLogicalDisks
        
        Frm_Main.SGrid_Computer_List_AddRow _
            objdisk.DeviceID & " " & objdisk.VolumeName & " - (" & objdisk.Description & ")", _
            "�`�@: " & Format_MB_By_B(objdisk.Size) & " / �i��: " & Format_MB_By_B(objdisk.FreeSpace) & IIf(IsNull(objdisk.FileSystem), "", " (" & objdisk.FileSystem & ")"), _
            "�޿�Ϻо�"
        DoEvents
    Next

End Function


Function WMI_Computer_Service(strComputer As String)
'�C�|�A��
On Error Resume Next

    Set colServices = objWMIService.ExecQuery("Select * from Win32_Service")
    
    If colServices.Count > 0 Then
        
        For Each objService In colServices
            Frm_Main.SGrid_Computer_List_AddRow _
                objService.Name & " (" & objService.Pathname & ")", _
                objService.State, "�A��"
            DoEvents
        Next
    
    End If

End Function

Function WMI_Computer_Process(strComputer As String)
'�˵���������A

On Error Resume Next
    
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process")
    
    For Each objProcess In colProcesses
        'objDictionary.Add objProcess.ProcessID, objProcess.Name
        Frm_Main.SGrid_Computer_List_AddRow _
            objProcess.Name & " (" & objProcess.ExecutablePath & ")", _
            "PID (" & objProcess.ProcessID & ") ", "�����"
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
'�˵������˸m���A

On Error Resume Next

    Set colNA = objWMIService.ExecQuery("Select * from Win32_NetworkAdapter")
    
    For Each objNA In colNA

        Frm_Main.SGrid_Computer_List_AddRow _
            objNA.Name & " (" & objNA.Manufacturer & ")", _
            "MAC - (" & objNA.MACAddress & ") ", "��������"
        DoEvents
    Next

End Function


Function WMI_Computer_DiskDrive(strComputer As String)
'�C�|����ϺФ��e

On Error Resume Next

    Set colDiskDrives = objWMIService.ExecQuery("Select * from Win32_DiskDrive")
    
    For Each objDiskDrive In colDiskDrives
    
        Frm_Main.SGrid_Computer_List_AddRow _
            objDiskDrive.Model & " / " & _
            objDiskDrive.InterfaceType & _
            " / " & objDiskDrive.Manufacturer, _
            Format_MB_By_B(objDiskDrive.Size), _
            "�Ϻо�"
        DoEvents
    Next

End Function

Function WMI_Computer_Share(strComputer As String)
'�C�|�@�θ�Ƨ����e
On Error Resume Next

    Set colShares = objWMIService.ExecQuery("Select * from Win32_Share")
    
    For Each objShare In colShares
    
        Frm_Main.SGrid_Computer_List_AddRow _
            objShare.Name & " - ", _
            objShare.Path, "�@�θ귽"
        DoEvents
    Next

End Function


Function WMI_Computer_Printer(strComputer As String)
'�C�|�L������e �ȾA�� XP �H�W
On Error Resume Next

'    If OS_Version > 5.1 Then  '(�j�� XP �H�W�~�i)
        
        Set colPrinters = objWMIService.ExecQuery("Select * from Win32_Printer")
        For Each objPrinter In colPrinters
        
            Frm_Main.SGrid_Computer_List_AddRow _
                objPrinter.DriverName & " - " & objPrinter.Name & " (���p:" & objPrinter.Status & ")", _
                objPrinter.PortName, "�L���"
            DoEvents
        Next
        

End Function

Function WMI_Computer_PrinterDriver(strComputer As String)
'�C�|�L����X�ʵ{��

On Error Resume Next

 If OS_Version < 5.1 Then  '(�p�� XP �H���䴩)
            
            Frm_Main.SGrid_Computer_List_AddRow _
                "���\��u���ؼйq���@�~�t�Φb XP (5.1) �H�W�~�䴩���", _
                "���䴩", _
                "�L����X�ʵ{��"
    Else
        
        Set cols = objWMIService.ExecQuery("Select * from Win32_PrinterDriver")
        For Each obj In cols
        
            Frm_Main.SGrid_Computer_List_AddRow _
                obj.Name & " " & obj.Version & " - " & obj.Description, _
                obj.DriverPath, _
                "�L����X�ʵ{��"
            DoEvents
        Next
        
  End If
  
End Function



Function WMI_Computer_TCPIPPrinterPort(strComputer As String)
'�C�|�L��� Port ���e �ȾA�� XP �H�W

On Error Resume Next

    If OS_Version < 5.1 Then  '(�p�� XP �H���䴩)
            
            Frm_Main.SGrid_Computer_List_AddRow _
                "���\��u���ؼйq���@�~�t�Φb XP (5.1) �H�W�~�䴩���", _
                "���䴩", _
                "�L��������s����"
    Else
            
        Set colTPs = objWMIService.ExecQuery("Select * from Win32_TCPIPPrinterPort")
        For Each objtp In colTPs
        
            Frm_Main.SGrid_Computer_List_AddRow _
                objtp.Name & " " & " PortNumber:" & objtp.PortNumber & " SNMP: " & IIf(objtp.SNMPEnabled = True, "�O", "�_"), _
                objtp.HostAddress, "�L��������s����"
            DoEvents
        Next
    
    End If

End Function

Function WMI_Computer_DiskPartition(strComputer As String)
'�C�|�ϺФ��ΰ�

On Error Resume Next

    Set colDiskPartitions = objWMIService.ExecQuery("Select * from Win32_DiskPartition")
    
    For Each objDP In colDiskPartitions
            
        Frm_Main.SGrid_Computer_List_AddRow _
            objDP.Name & IIf(objDP.BootPartition = True, " [Boot]", ""), _
            Format_MB_By_B(objDP.Size), _
            "�ϺФ��ΰ�"
        DoEvents
    Next

End Function


Function WMI_Computer_Terminal(strComputer As String)

'�C�|�׺ݾ���T

On Error Resume Next
    
    If OS_Version < 5.1 Then  '(�p�� XP ���䴩)
    
            Frm_Main.SGrid_Computer_List_AddRow _
                "���\��u���ؼйq���@�~�t�Φb XP (5.1) �H�W�~�䴩���", _
                "���䴩", _
                "�׺ݾ�"
    Else
            
        Set colTerminals = objWMIService.ExecQuery("Select * from Win32_Terminal")
        For Each objTerminal In colTerminals
        
            Frm_Main.SGrid_Computer_List_AddRow _
                objTerminal.TerminalName, _
                IIf(objTerminal.fEnableTerminal = 1, "�ҥ�", "����"), _
                "�׺ݾ�"
        
            DoEvents
        Next
    
    End If

End Function


Function WMI_Computer_StartupCommand(strComputer As String)
'�C�|�}������

On Error Resume Next
    
    Set colSCs = objWMIService.ExecQuery("Select * from Win32_StartupCommand")
    For Each objsc In colSCs
    
        Frm_Main.SGrid_Computer_List_AddRow _
            objsc.Name & " (User:" & objsc.User & " / Loc.:" & objsc.Location & ")", _
            objsc.Command, _
            "�}������"
    
        DoEvents
    Next

End Function


Function WMI_Computer_SoundDevice(strComputer As String)
'�C�|���ĸ˸m

On Error Resume Next
    Set cols = objWMIService.ExecQuery("Select * from Win32_SoundDevice")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name, _
            obj.Manufacturer, _
            "���ĸ˸m"
        DoEvents
    Next

End Function

Function WMI_Computer_SerialPort(strComputer As String)
'�C�|�s����˸m

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_SerialPort")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name, _
            IIf(IsNull(obj.Status), "���i��", "�i��"), _
            "�s����"
        DoEvents
    Next


    Set cols = objWMIService.ExecQuery("Select * from Win32_ParallelPort")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name, _
            IIf(IsNull(obj.Status), "���i��", "�i��"), _
            "�s����"
        DoEvents
    Next

End Function



Function WMI_Computer_ScheduledJob(strComputer As String)
'�C�|�u�@�Ƶ{ ����

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_ScheduledJob")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name, _
            IIf(IsNull(obj.Status), "���i��", "�i��"), _
            "�s����"
        DoEvents
    Next

End Function

Function WMI_Computer_PointingDevice(strComputer As String)
'�C�|���Щʸ˸m

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_Keyboard")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name & " (" & obj.Description & ")", _
            IIf(IsNull(obj.Status), "���i��", "�i��"), _
            "����ʸ˸m"
        DoEvents
    Next

End Function

Function WMI_Computer_Keyboard(strComputer As String)
'�C�|����ʸ˸m

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_PointingDevice")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name & " (" & obj.Manufacturer & ")", _
            "���s��: " & obj.NumberOfButtons, _
            "���Щʸ˸m"
        DoEvents
    Next

End Function

Function WMI_Computer_PnPEntity(strComputer As String)
'�C�|�H���Y�θ˸m

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_PnPEntity")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name & " - " & obj.Description & " " & obj.Manufacturer & "", _
            IIf(IsNull(obj.Status), "���s��", "�w�s��"), _
            "�H���Y�θ˸m"
        DoEvents
    Next

End Function

Function WMI_Computer_Group(strComputer As String)
'�C�|�����s��

'�w��� ADSI �g�k

'On Error Resume Next
'    'MsgBox strComputer
'    Set cols = objWMIService.ExecQuery("Select * from Win32_Group Domain = '" & strComputer & "'")
'    For Each obj In cols
'
'        Frm_Main.SGrid_Computer_List_AddRow _
'            obj.Name, _
'            obj.Description, _
'            "�����s��"
'        DoEvents
'    Next

End Function


Function WMI_Computer_Displays(strComputer As String)
'�C�|��ܥd��T

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_DisplayConfiguration")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow "��ܥd����", StrNullToSpace(obj.Name), "��ܥd"
        Frm_Main.SGrid_Computer_List_AddRow "��ܥd��������", StrNullToSpace(obj.Description), "��ܥd"
        Frm_Main.SGrid_Computer_List_AddRow "�X�ʵ{������", StrNullToSpace(obj.DriverVersion), "��ܥd"
        Frm_Main.SGrid_Computer_List_AddRow "����W�e", StrNullToSpace(obj.BitsPerPel), "��ܥd"
        Frm_Main.SGrid_Computer_List_AddRow "������s�W�v", StrNullToSpace(obj.DisplayFrequency), "��ܥd"
        DoEvents
        
    Next

    Set cols = objWMIService.ExecQuery("Select * from Win32_DisplayControllerConfiguration")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow "��������", StrNullToSpace(obj.Name), "��ܥd"
        Frm_Main.SGrid_Computer_List_AddRow "������������", StrNullToSpace(obj.Description), "��ܥd"
        Frm_Main.SGrid_Computer_List_AddRow "�ѪR��", StrNullToSpace(obj.HorizontalResolution) & "x" & StrNullToSpace(obj.VerticalResolution), "��ܥd"
        Frm_Main.SGrid_Computer_List_AddRow "��m�~��", StrNullToSpace(obj.BitsPerPixel), "��ܥd"
        Frm_Main.SGrid_Computer_List_AddRow "��ܼҦ�", StrNullToSpace(obj.VideoMode), "��ܥd"
        DoEvents
        
    Next


End Function

Function WMI_Computer_DesktopMonitor(strComputer As String)
'�C�|��ܾ��i����ܼҦ�

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name & " " & obj.Screenwidth & "x" & obj.ScreenHeight, _
            obj.DeviceID, _
            "��ܾ�"
        DoEvents
    Next

End Function


Function WMI_Computer_Environment(strComputer As String)
'�C�|�����ܼ��ܼ�

    On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_Environment")

    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.UserName & " | " & obj.Name & " | �t���ܼ�:" & IIf(obj.SystemVariable = True, "�O", "�_") & " (" & obj.Description & ") ", _
            obj.VariableValue, _
            "�����ܼ�"
        
        DoEvents
    Next
    

End Function

Function WMI_Computer_CodecFile(strComputer As String)
'�C�|�ѽX����T

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_CodecFile")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            StrNullToSpace(obj.Group) & " | " & StrNullToSpace(obj.Description) & " " & StrNullToSpace(obj.Version) & " (" & StrNullToSpace(obj.Manufacturer) & ")", _
            StrNullToSpace(obj.Name), _
            "�ѽX��"
        DoEvents
    Next

End Function



Function WMI_Computer_NetworkAdapterConfiguration(strComputer As String)
'�C�|�����]�w

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    Dim i As Integer
    
    For Each obj In cols
    
        i = i + 1
        
        Frm_Main.SGrid_Computer_List_AddRow "�˸m", StrNullToSpace(obj.index) & " - " & StrNullToSpace(obj.ServiceName), "�����]�w " & i
        
        
        Frm_Main.SGrid_Computer_List_AddRow "�˸m�W��", StrNullToSpace(obj.Description), "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "�ҥ� DHCP", StrNullToSpace(obj.DHCPEnabled), "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "DHCP ���A��", StrNullToSpace(obj.dhcpserver), "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "DNS ���", obj.DNSDomain, "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "IP ��}", IIf(IsNull(obj.IPAddress) = True, "", Join(obj.IPAddress)), "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "�l�����B�n", IIf(IsNull(obj.IPSubnet) = True, "", Join(obj.IPSubnet)), "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "�w�]�h�D", IIf(IsNull(obj.DefaultIPGateway) = True, "", Join(obj.DefaultIPGateway)), "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "DNS ���A��", IIf(IsNull(obj.DNSServerSearchOrder) = True, "", Join(obj.DNSServerSearchOrder)), "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "�ֹ̤h�D", IIf(IsNull(obj.GatewayCostMetric) = True, "", Join(obj.GatewayCostMetric)), "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "�ֳ̤s�u", StrNullToSpace(obj.IPConnectionMetric), "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "�ҥ� IPX", StrNullToSpace(obj.IPXEnabled), "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "IPX ��}", IIf(IsNull(obj.IPXAddress) = True, "", Join(obj.IPXAddress)), "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "�b DNS �W���U�q���W��", StrNullToSpace(obj.DomainDNSRegistrationEnabled), "�����]�w " & i
        Frm_Main.SGrid_Computer_List_AddRow "MAC �d��", StrNullToSpace(obj.MACAddress), "�����]�w " & i
        
        DoEvents
    Next
    
    
End Function


Function WMI_Computer_QuickFixEngineering(strComputer As String)
'�C�X�w�w�˪� Hot Fixes
    
    Set cols = objWMIService.ExecQuery("Select * from Win32_QuickFixEngineering")
    
    For Each obj In cols
        
        If obj.HotFixID <> "File 1" Then
            Frm_Main.SGrid_Computer_List_AddRow obj.HotFixID, obj.Description, "��s��"
        End If
        
        DoEvents
    Next
        
End Function

Function ADSI_Computer_Group(strComputer As String)
'�C�|�ؼйq�����s��
    Set cols = GetObject("WinNT://" & strComputer & ",computer")
    cols.Filter = Array("group")
    
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow obj.Name, obj.Description, "�����s��"
    Next

End Function

Function ADSI_Computer_User(strComputer As String)
'�C�|�ؼйq�����b��
    
    'Dim cmp
    
    'Dim usr As IADsUser
    Dim usr_grp As IADsGroup
    Dim tmp_grp As String
    
    Set usr = GetObject("WinNT://" & strComputer & ",computer")
    usr.Filter = Array("user")
        
    
    '�C�| User
    For Each obj In usr
        
        '���X User �����ݪ��s��
        For Each usr_grp In obj.Groups
            tmp_grp = tmp_grp & usr_grp.Name & ";"
        Next
        tmp_grp = Left(tmp_grp, Len(tmp_grp) - 1)
        
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.Name & " (" & tmp_grp & ") " & IIf(obj.AccountDisabled = True, "(�ҥ�)", "(����)"), _
            obj.Description, _
            "�����ϥΪ�"
        
        tmp_grp = ""
        
    Next

End Function


