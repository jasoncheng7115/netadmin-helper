'***********************************************
'���g : Jason Cheng
'�Ѧ� : MSDN
'�Ϊk : ���p�b�s�խ�h�����n�J Script
'
'�p����K�ЫO�d������T�A����
'�p�� Bug�A�гq���A����
'***********************************************

On Error Resume Next

    '���o���� windows �ؿ��P����
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set cols = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    For Each obj In cols
        tmp_windir = obj.WindowsDirectory
        tmp_ver = Left(obj.Version, 3)
    Next

    '�Y�b Windows �����b XP(5.1) �H�W�~�i�� --> (�~�����ب�����A���M�F���ק�?)
    If CSng(tmp_ver) >= CSng(5.1) Then

      '�}�l�ק慨����]�w�� netfw.inf -----------------------------------------------------

          '�w�q�n�g�J inf �ɪ��r��
          Dim Write_Command
          Write_Command = "HKLM," & Chr(34)
          Write_Command = "SYSTEM\CurrentControlSet\Services\SharedAccess\"
          Write_Command = "Parameters\FirewallPolicy\DomainProfile\RemoteAdminSettings"
          Write_Command = Chr(34) & "," & Chr(34) & "Enabled" & Chr(34) 
          Write_Command = ",0x00010001,1" & Chr(34)

          '�w�q�ɦW
          Dim sFile
          sfile = tmp_windir & "\inf\netfw.inf"

          '�ŧi�ɮ׾ާ@����
          Set objFSO = CreateObject("Scripting.FileSystemObject")

          'Ū�g�ݩ�
          Const ForReading = 1, ForWriting = 2, ForAppending = 8
          
          '�r���� (-1=Unicode,0=ASCII)
          Const TristateFalse = 0, TristateTrue = -1, TristateUseDefault = -2 

          'Ū�J�ɮ�
          Set objFile = objFSO.OpenTextFile(sfile, ForReading, False, -1)
          tmp_File = objFile.Readall
          objFile.Close

          '�j�M�O�_�w�s�b�����O�A�S���~�g�J
          If InStr(1, tmp_File, Write_Command) = 0 Then

              '�j�M�n�g�J���Ϭq Section
              tmp_Section = InStr(1, tmp_File, "[ICF.AddReg.DomainProfile]")

              '���o�j�M�ӰϬq�U�@�� (�Ѧ���m�}�l���J�A+2 �O���F�׶}�浲���P����Ÿ�)
              tmp_Section_Nextline = InStr(tmp_Section, tmp_File, vbCrLf) + 2

              '�զX�s�ɮ׭n�g�J�����e
              target_file = ""
              target_file = target_file & Left(tmp_File, tmp_Section_Nextline - 1)
              target_file = target_file & Write_Command & vbCrLf
              target_file = target_file & Mid(tmp_File, tmp_Section_Nextline)

              '���N�쥻�ɦW�ק�ƥ�
              IF objFSO.FileExists(sfile & ".bak") then objFSO.DeleteFile sfile & ".bak"
              objFSO.MoveFile sfile, sfile & ".bak"

              '�g�^
              Set objFile = objFSO.OpenTextFile(sfile, ForAppending, True, -1)
              objFile.Write target_file
              objFile.Close

          End If

      '�����ק慨����]�w�� netfw.inf -------------------------------------------------



      '�}�l�b������}�һ��ݺ޲z -------------------------------------------------------

        Set objFirewall = CreateObject("HNetCfg.FwMgr")
        Set objPolicy = objFirewall.LocalPolicy.CurrentProfile
        Set objAdminSettings = objPolicy.RemoteAdminSettings
        objAdminSettings.Enabled = True

      '�����b������}�һ��ݺ޲z -------------------------------------------------------


        '�}�l�b������}���ɮפΦC�L�@�ΪA�� ---------------------------------------

            '�ŧi�`��
            Const NET_FW_SERVICE_FILE_AND_PRINT = 0
            Const NET_FW_SERVICE_UPNP = 1
            Const NET_FW_SERVICE_REMOTE_DESKTOP = 2

            '���
            Const NET_FW_SCOPE_ALL = 0             '����
            Const NET_FW_SCOPE_LOCAL_SUBNET = 1    '���a�l����

            '�ŧi�ܼ�
            Dim service
            Dim port

            '�إߤ@�Ө��������
            Dim fwMgr
            Set fwMgr = CreateObject("HNetCfg.FwMgr")

            '���o����������]�w��
            Dim profile
            Set profile = fwMgr.LocalPolicy.CurrentProfile

            '���X�ɮ׻P�C�L�@��
            Set service = profile.Services.Item(NET_FW_SERVICE_FILE_AND_PRINT)

            '�}�ҪA��
            service.Enabled = TRUE

            '�]�w�i�Ϊ����ݦ�}�A�b�o�̤�������
            service.RemoteAddresses = "*"

            '�]�w�ӷ�����
            service.Scope = NET_FW_SCOPE_ALL

            '�����䤤�� 445 (�L�n���ؿ��A�� Port�A�p�b AD ���Ҥ��A���i����)
            '�ñN 139 ��W�]���Ȧ����a�l�����i�ϥ�
            For Each port In service.GloballyOpenPorts
                'if port.Port = 139 Then port.Scope = NET_FW_SCOPE_LOCAL_SUBNET
            Next

        '�����b������}���ɮפΦC�L�@�ΪA�� --------------------------------------

    End If

