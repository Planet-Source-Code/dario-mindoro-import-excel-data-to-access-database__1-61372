VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Access Database Password"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3105
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1155
   ScaleWidth      =   3105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   1560
      TabIndex        =   2
      Top             =   720
      Width           =   1125
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "Done"
      Height          =   345
      Left            =   330
      TabIndex        =   1
      Top             =   720
      Width           =   1125
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   180
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   180
      Width           =   2775
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    IMPORTINFO.accesspassword = ""
    Unload Me
End Sub

Private Sub cmdDone_Click()
    IMPORTINFO.accesspassword = txtPassword.Text
    Unload Me
End Sub


Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdDone_Click
End Sub
