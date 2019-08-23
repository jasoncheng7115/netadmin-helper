VERSION 5.00
Begin VB.Form Frm_Batch_Processed 
   BorderStyle     =   5  '可調整工具視窗
   Caption         =   "批次作業處理結果"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '系統預設值
   Begin VB.TextBox Txt_Log 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   0
      Top             =   480
      Width           =   4815
   End
   Begin VB.Label Lbl_Title 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   60
      TabIndex        =   1
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "Frm_Batch_Processed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()

On Error Resume Next

Txt_Log.Width = Me.ScaleWidth - Txt_Log.Left - Txt_Log.Left
Txt_Log.Height = Me.ScaleHeight - Txt_Log.TOp - Txt_Log.Left

End Sub

Private Sub Txt_Log_Click()
AutoSelStr Me.ActiveControl
End Sub
