Attribute VB_Name = "Mod_CMI"

Function CMI_Computer_VideoControllerResolution(strComputer As String)
'�C�|�䴩���ѪR�׸�T

On Error Resume Next
    
    Set cols = objWMIService.ExecQuery("Select * from CIM_VideoControllerResolution")
    For Each obj In cols
    
        Frm_Main.SGrid_Computer_List_AddRow _
            obj.settingid, _
            "", _
            "��ܼҦ�"
        DoEvents
    Next

End Function





