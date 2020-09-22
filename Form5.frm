VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "Form5.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5175
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   1950
      TabIndex        =   4
      Top             =   3360
      Width           =   1305
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3315
      Left            =   -30
      ScaleHeight     =   3255
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   -30
      Width           =   5235
      Begin VB.Image Image3 
         Height          =   630
         Left            =   720
         Picture         =   "Form5.frx":0442
         Stretch         =   -1  'True
         Top             =   660
         Width           =   630
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   600
         Picture         =   "Form5.frx":074C
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   315
         Left            =   4110
         Picture         =   "Form5.frx":0D39
         Top             =   630
         Width           =   330
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "For bugs/suggestions please contact me using info below."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   210
         TabIndex        =   10
         Top             =   1860
         Width           =   4725
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile: +639207874658"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   9
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "dards@mindworksoft.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   2670
         TabIndex        =   8
         Top             =   2670
         Width           =   2385
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "email:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   7
         Top             =   2670
         Width           =   915
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "http://www.mindworksoft.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   2670
         MouseIcon       =   "Form5.frx":1143
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   2460
         Width           =   2385
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Homepage:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   5
         Top             =   2460
         Width           =   915
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Programmer: Dario Mindoro "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1680
         TabIndex        =   3
         Top             =   2250
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form5.frx":1585
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   180
         TabIndex        =   2
         Top             =   1200
         Width           =   4755
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Import Excel to Access Database v.1.0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   885
         Left            =   510
         TabIndex        =   1
         Top             =   90
         Width           =   4335
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Label5_Click()
ShellExecute Me.hwnd, "open", "http://www.mindworksoft.com", vbNullString, vbNullString, SW_SHOWDEFAULT
End Sub
