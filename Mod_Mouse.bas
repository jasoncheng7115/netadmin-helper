Attribute VB_Name = "Mod_Mouse"
Option Explicit

Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Public tmp_x As Long
Public tmp_y As Long


Public Function GethWnd() As Long
Dim lpPoint As POINTAPI
GetCursorPos lpPoint

Debug.Print lpPoint.x & ", " & lpPoint.y
tmp_x = lpPoint.x
tmp_y = lpPoint.y
'GethWnd = WindowFromPoint(lpPoint.x, lpPoint.y)
End Function

