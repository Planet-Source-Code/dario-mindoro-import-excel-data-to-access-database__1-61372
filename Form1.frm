VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MindWorkSoft.Com"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cdialog 
      Left            =   6510
      Top             =   660
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   5910
      TabIndex        =   11
      Top             =   4380
      Width           =   1125
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >"
      Height          =   345
      Left            =   4740
      TabIndex        =   10
      Top             =   4380
      Width           =   1125
   End
   Begin VB.Frame Frame2 
      Caption         =   "MS Access Database file (*.mdb)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2580
      TabIndex        =   3
      Top             =   2280
      Width           =   4485
      Begin VB.CommandButton cmdBrowseAccess 
         Caption         =   "..."
         Height          =   405
         Left            =   4020
         TabIndex        =   7
         Top             =   330
         Width           =   375
      End
      Begin VB.TextBox txtAccessFile 
         Height          =   405
         Left            =   90
         TabIndex        =   5
         Top             =   330
         Width           =   3915
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Excel File (*.xls)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2580
      TabIndex        =   2
      Top             =   1230
      Width           =   4485
      Begin VB.CommandButton cmdBrowseExcel 
         Caption         =   "..."
         Height          =   405
         Left            =   4020
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox txtExcelFile 
         Height          =   405
         Left            =   90
         TabIndex        =   4
         Top             =   360
         Width           =   3915
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   5115
      Left            =   -30
      ScaleHeight     =   5055
      ScaleWidth      =   2385
      TabIndex        =   0
      Top             =   -30
      Width           =   2445
      Begin VB.Line Line1 
         X1              =   150
         X2              =   2220
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Programmer: Dario Mindoro, site: www.mindworksoft.com, email: dards@mindworksoft.com, mobile: +639207874658"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Import Excel to Access Database"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   1245
         Left            =   150
         TabIndex        =   1
         Top             =   180
         Width           =   2055
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Find the source Excel file and the distination MS Access Database file and then click next button below"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2610
      TabIndex        =   8
      Top             =   300
      Width           =   4425
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowseAccess_Click()
  ' Set CancelError is True
  cdialog.CancelError = True
  On Error GoTo ErrHandlerAccess
  ' Set flags
  cdialog.Flags = cdlOFNHideReadOnly
  ' Set filters
  cdialog.Filter = "MS Access file|*.mdb"
  ' Display the Open dialog box
  cdialog.ShowOpen
  ' Display name of selected file
  txtAccessFile.Text = cdialog.FileName
  Exit Sub
ErrHandlerAccess:
  'User pressed the Cancel button
  Exit Sub
End Sub

Private Sub cmdBrowseExcel_Click()
  ' Set CancelError is True
  cdialog.CancelError = True
  On Error GoTo ErrHandlerExcel
  ' Set flags
  cdialog.Flags = cdlOFNHideReadOnly
  ' Set filters
  cdialog.Filter = "Excel file|*.xls"
  ' Display the Open dialog box
  cdialog.ShowOpen
  ' Display name of selected file
  txtExcelFile.Text = cdialog.FileName
  Exit Sub
ErrHandlerExcel:
  'User pressed the Cancel button
  Exit Sub
End Sub

Private Sub cmdCancel_Click()
    EndProgram
End Sub

Private Sub cmdNext_Click()
    Dim retval As Long
    Dim retvalxls As Long
    'check empty box
    If Trim(txtExcelFile.Text) = "" Then
        MsgBox "Please specify Excel source file", vbOKOnly, "Error occured"
        Exit Sub
    End If
    
    If Trim(txtAccessFile.Text) = "" Then
        MsgBox "Please specify MS Access file", vbOKOnly, "Error occured"
        Exit Sub
    End If
    
    IMPORTINFO.excelfile = txtExcelFile.Text
    IMPORTINFO.accessfile = txtAccessFile.Text
    IMPORTINFO.accesspassword = ""

ConnectMDB:
    'open access database
    retval = OpenMDB
    If retval = -2147217843 Then GoTo ErrNoPassword
    If retval <> 0 Then GoTo ErrUnknown

ConnectXLS:
    
    'open excel file
    retvalxls = OpenXLS(IMPORTINFO.excelfile)
    If retvalxls <> 0 Then GoTo ErrUnknown

    'no probs proceed
    Form2.Show
    Unload Me
    
    
    Exit Sub
ErrNoPassword:
    Form3.Show 1
    If Trim(IMPORTINFO.accesspassword) = "" Then
        MsgBox "The database requires password, you must supply password!", vbOKOnly, "Password Error"
        Exit Sub
    End If
    GoTo ConnectMDB
    
ErrUnknown:
    MsgBox "Unexptected Error occured, exiting...", vbCritical, "Error occred"
    End
End Sub

Private Sub Form_Load()
    txtExcelFile.Text = IMPORTINFO.excelfile
    txtAccessFile.Text = IMPORTINFO.accessfile
End Sub
