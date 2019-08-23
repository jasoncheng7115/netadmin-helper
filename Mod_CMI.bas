Attribute VB_Name = "Mod_CMI"

Function CMI_Computer_VideoControllerResolution(strComputer As String)
'列舉支援的解析度資訊

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from CIM_VideoControllerResolution")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.settingid, _
            "", _
            "顯示模式"
        DoEvents
    Next

End Function





