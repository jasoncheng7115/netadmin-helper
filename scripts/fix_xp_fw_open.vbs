'***********************************************
'撰寫 : Jason Cheng
'參考 : MSDN
'用法 : 部署在群組原則中的登入 Script
'
'如有轉貼請保留相關資訊，謝謝
'如有 Bug，請通知，謝謝
'***********************************************

On Error Resume Next

    '取得本機 windows 目錄與版本
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set cols = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
    For Each obj In cols
        tmp_windir = obj.WindowsDirectory
        tmp_ver = Left(obj.Version, 3)
    Next

    '若在 Windows 版本在 XP(5.1) 以上才進行 --> (才有內建防火牆，不然幹嘛修改?)
    If CSng(tmp_ver) >= CSng(5.1) Then

      '開始修改防火牆設定檔 netfw.inf -----------------------------------------------------

          '定義要寫入 inf 檔的字串
          Dim Write_Command
          Write_Command = "HKLM," & Chr(34)
          Write_Command = "SYSTEM\CurrentControlSet\Services\SharedAccess\"
          Write_Command = "Parameters\FirewallPolicy\DomainProfile\RemoteAdminSettings"
          Write_Command = Chr(34) & "," & Chr(34) & "Enabled" & Chr(34) 
          Write_Command = ",0x00010001,1" & Chr(34)

          '定義檔名
          Dim sFile
          sfile = tmp_windir & "\inf\netfw.inf"

          '宣告檔案操作物件
          Set objFSO = CreateObject("Scripting.FileSystemObject")

          '讀寫屬性
          Const ForReading = 1, ForWriting = 2, ForAppending = 8
          
          '字元集 (-1=Unicode,0=ASCII)
          Const TristateFalse = 0, TristateTrue = -1, TristateUseDefault = -2 

          '讀入檔案
          Set objFile = objFSO.OpenTextFile(sfile, ForReading, False, -1)
          tmp_File = objFile.Readall
          objFile.Close

          '搜尋是否已存在此指令，沒有才寫入
          If InStr(1, tmp_File, Write_Command) = 0 Then

              '搜尋要寫入的區段 Section
              tmp_Section = InStr(1, tmp_File, "[ICF.AddReg.DomainProfile]")

              '取得搜尋該區段下一行 (由此位置開始插入，+2 是為了避開行結束與換行符號)
              tmp_Section_Nextline = InStr(tmp_Section, tmp_File, vbCrLf) + 2

              '組合新檔案要寫入的內容
              target_file = ""
              target_file = target_file & Left(tmp_File, tmp_Section_Nextline - 1)
              target_file = target_file & Write_Command & vbCrLf
              target_file = target_file & Mid(tmp_File, tmp_Section_Nextline)

              '先將原本檔名修改備份
              IF objFSO.FileExists(sfile & ".bak") then objFSO.DeleteFile sfile & ".bak"
              objFSO.MoveFile sfile, sfile & ".bak"

              '寫回
              Set objFile = objFSO.OpenTextFile(sfile, ForAppending, True, -1)
              objFile.Write target_file
              objFile.Close

          End If

      '完成修改防火牆設定檔 netfw.inf -------------------------------------------------



      '開始在防火牆開啟遠端管理 -------------------------------------------------------

        Set objFirewall = CreateObject("HNetCfg.FwMgr")
        Set objPolicy = objFirewall.LocalPolicy.CurrentProfile
        Set objAdminSettings = objPolicy.RemoteAdminSettings
        objAdminSettings.Enabled = True

      '結束在防火牆開啟遠端管理 -------------------------------------------------------


        '開始在防火牆開啟檔案及列印共用服務 ---------------------------------------

            '宣告常數
            Const NET_FW_SERVICE_FILE_AND_PRINT = 0
            Const NET_FW_SERVICE_UPNP = 1
            Const NET_FW_SERVICE_REMOTE_DESKTOP = 2

            '領域
            Const NET_FW_SCOPE_ALL = 0             '任何
            Const NET_FW_SCOPE_LOCAL_SUBNET = 1    '本地子網路

            '宣告變數
            Dim service
            Dim port

            '建立一個防火牆實體
            Dim fwMgr
            Set fwMgr = CreateObject("HNetCfg.FwMgr")

            '取得本機防火牆設定檔
            Dim profile
            Set profile = fwMgr.LocalPolicy.CurrentProfile

            '取出檔案與列印共用
            Set service = profile.Services.Item(NET_FW_SERVICE_FILE_AND_PRINT)

            '開啟服務
            service.Enabled = TRUE

            '設定可用的遠端位址，在這裡不做限制
            service.RemoteAddresses = "*"

            '設定來源網域
            service.Scope = NET_FW_SCOPE_ALL

            '取消其中的 445 (微軟的目錄服務 Port，如在 AD 環境中，不可關閉)
            '並將 139 單獨設成僅有本地子網路可使用
            For Each port In service.GloballyOpenPorts
                'if port.Port = 139 Then port.Scope = NET_FW_SCOPE_LOCAL_SUBNET
            Next

        '完成在防火牆開啟檔案及列印共用服務 --------------------------------------

    End If

