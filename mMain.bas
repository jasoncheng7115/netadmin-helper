Attribute VB_Name = "mMain"
Option Explicit

Private Type tagInitCommonControlsEx
    dwSize As Long
    dwICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "Comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean

Private Const ICC_BAR_CLASSES = &H4
Private Const ICC_COOL_CLASSES = &H400


Public Sub Main()
Dim tIccex As tagInitCommonControlsEx
   On Error Resume Next
   With tIccex
       .dwSize = LenB(tIccex)
       .dwICC = ICC_BAR_CLASSES
   End With
   InitCommonControlsEx tIccex
   On Error GoTo 0

   'Frm_Main.Show
   
End Sub
