VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Table matching Setup"
   ClientHeight    =   7320
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   11550
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   345
      Left            =   9120
      TabIndex        =   19
      Top             =   6900
      Width           =   1155
   End
   Begin VB.CommandButton cmdBackup 
      Caption         =   "Back-up Database"
      Height          =   345
      Left            =   2730
      TabIndex        =   18
      Top             =   6900
      Width           =   1635
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "< Back"
      Height          =   345
      Left            =   120
      TabIndex        =   8
      Top             =   6900
      Width           =   1155
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   6165
      Left            =   60
      ScaleHeight     =   6105
      ScaleWidth      =   11385
      TabIndex        =   5
      Top             =   660
      Width           =   11445
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import now"
         Height          =   345
         Left            =   5340
         TabIndex        =   14
         Top             =   5700
         Width           =   1155
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1350
         Top             =   4500
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form2.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form2.frx":089A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form2.frx":0CF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form2.frx":0E4E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form2.frx":12A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form2.frx":13FE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   3885
         Left            =   5310
         TabIndex        =   12
         Top             =   270
         Width           =   2145
         _ExtentX        =   3784
         _ExtentY        =   6853
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fields"
            Object.Width           =   3528
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   5775
         Left            =   30
         TabIndex        =   11
         Top             =   270
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   10186
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         OLEDragMode     =   1
         OLEDropMode     =   1
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OLEDragMode     =   1
         OLEDropMode     =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fields"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Type"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Excel Fields"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Initial Values"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "field index"
            Object.Width           =   1764
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   3885
         Left            =   7500
         TabIndex        =   7
         Top             =   270
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   6853
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   6
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblStatus 
         BackStyle       =   0  'Transparent
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7260
         TabIndex        =   16
         Top             =   5760
         Width           =   3975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   6630
         TabIndex        =   15
         Top             =   5790
         Width           =   675
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"Form2.frx":155A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5310
         TabIndex        =   13
         Top             =   4200
         Width           =   6015
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Excel Fields"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   255
         Left            =   5340
         TabIndex        =   10
         Top             =   30
         Width           =   1035
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Excel Data preview"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7530
         TabIndex        =   9
         Top             =   30
         Width           =   1755
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Access Fields"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   255
         Left            =   90
         TabIndex        =   6
         Top             =   30
         Width           =   1185
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   345
      Left            =   10320
      TabIndex        =   4
      Top             =   6900
      Width           =   1155
   End
   Begin VB.ComboBox cmbSheets 
      Height          =   315
      Left            =   7230
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   2235
   End
   Begin VB.ComboBox cmbTables 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   2235
   End
   Begin MSComctlLib.ProgressBar progbar 
      Height          =   285
      Left            =   5400
      TabIndex        =   17
      Top             =   6930
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      Caption         =   "Select Sheet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6150
      TabIndex        =   3
      Top             =   270
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "Select Table"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   270
      Width           =   945
   End
   Begin VB.Menu pMenu 
      Caption         =   "pMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuresetexcel 
         Caption         =   "Reset excel fields"
      End
      Begin VB.Menu mnusetdefaultvalue 
         Caption         =   "Set default value"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear fields"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tblList As ADODB.Recordset
Dim tblListxls As ADODB.Recordset
Dim fldlst As ADODB.Recordset
Dim fldlstxls As ADODB.Recordset
Dim rs_source As ADODB.Recordset




Private Sub cmbSheets_Click()
    ShowExcelFields
End Sub

Private Sub cmbTables_Click()
    ShowFields
End Sub

Private Sub cmdAbout_Click()
    Form5.Show 1
End Sub

Private Sub cmdBack_Click()
    Form1.Show
    Unload Me
End Sub

Private Sub cmdBackup_Click()
    Dim bakcnt As Integer
    Dim bakfilename As String
    
AGAIN1:
    bakcnt = bakcnt + 1
    bakfilename = Left$(IMPORTINFO.accessfile, Len(IMPORTINFO.accessfile) - 4)
    bakfilename = bakfilename & "_BAK" & bakcnt & ".mdb"
        
    'ops file already exist increment back count again
    If FileExists(bakfilename) = True Then GoTo AGAIN1

    'now copy anyway
    CopyFileWindowsWay IMPORTINFO.accessfile, bakfilename
    
    MsgBox "Backing up succeded!", vbOKOnly, "Backup"
End Sub

Private Sub cmdClose_Click()
    EndProgram
End Sub

Private Sub ShowTables()
    Set tblList = New ADODB.Recordset
    Set tblList = Cn.OpenSchema(adSchemaTables)

    Me.MousePointer = vbHourglass
    'list all the table from the database
    Do While Not tblList.EOF
        If tblList.Fields(3) = "TABLE" Then
        cmbTables.AddItem tblList.Fields(2)
        End If
        tblList.MoveNext
    Loop
    Me.MousePointer = vbNormal
End Sub

Private Sub ShowExcelSheets()
    Set tblListxls = New ADODB.Recordset
    Set tblListxls = CnXls.OpenSchema(adSchemaTables)

    Me.MousePointer = vbHourglass
    'list all the table from the database
    Do While Not tblListxls.EOF
        'If tblListxls.Fields(3) = "TABLE" Then
        cmbSheets.AddItem tblListxls.Fields(2)
        'End If
        tblListxls.MoveNext
    Loop
    Me.MousePointer = vbNormal
End Sub

Private Sub cmdImport_Click()
    Dim destinationtble As String
    Dim sourcetble As String
    Dim selflds As String
    Dim i As Integer
    Dim defval As String
    Dim qSTRUCT() As QUERYSTRUCT
    Dim fldCount As Integer
    Dim sqltxt1 As String
    Dim sqltxt2 As String
    Dim retval As String
    Dim notnull As Boolean
    Dim cnt As Integer
    
    
    If MsgBox(" W   A   R   N   I   N   G  !" & vbCrLf & "You are about to make changes to Access Database file, be sure to BACKUP first the database before doing this!. Continue anyway?", vbYesNo, "Import Data Warning!") = vbNo Then Exit Sub
    
    If ListView1.ListItems.Count = 0 Then
        MsgBox "There are no table selected from Access Database", vbOKOnly, "Import Error!"
        Exit Sub
    End If
    
    On Error GoTo ErrHandlerImport
    
    'build query string for excel data
    sourcetble = " SELECT * FROM [" & cmbSheets.Text & "]"
    
    'find how many fields needed for the operation
    For i = 1 To ListView1.ListItems.Count
        selflds = ListView1.ListItems(i).SubItems(2)
        defval = ListView1.ListItems(i).SubItems(3)
        If Trim(selflds) <> "" Or Trim(defval) <> "" Then
                fldCount = fldCount + 1
        End If
    Next i
                
    'build query structor
    ReDim qSTRUCT(fldCount) As QUERYSTRUCT
    For i = 1 To ListView1.ListItems.Count
        selflds = ListView1.ListItems(i).SubItems(2)
        defval = ListView1.ListItems(i).SubItems(3)
        If Trim(selflds) <> "" Or Trim(defval) <> "" Then
                cnt = cnt + 1
                qSTRUCT(cnt).destinationfields = ListView1.ListItems(i).Text
                qSTRUCT(cnt).sourcefields = selflds
                qSTRUCT(cnt).defaultvalues = defval
        End If
    Next i
                
    Set rs_source = OpenRSXLS(sourcetble)
    If rs_source.RecordCount <> 0 Then
        progbar.Max = rs_source.RecordCount
        Do While Not rs_source.EOF
            DoEvents
            
            lblStatus.Caption = "Processing...."
            fldlst.AddNew
            For i = 1 To UBound(qSTRUCT)
                sqltxt1 = qSTRUCT(i).destinationfields
                sqltxt2 = qSTRUCT(i).sourcefields
                
                If Trim(sqltxt2) = "" Then
                    fldlst.Fields("" & sqltxt1 & "").Value = qSTRUCT(i).defaultvalues
                Else
                    fldlst.Fields("" & sqltxt1 & "").Value = rs_source.Fields("" & sqltxt2 & "").Value
                End If
                
            Next i
            fldlst.Update
            
            
            
            progbar.Value = rs_source.AbsolutePosition - 1
            rs_source.MoveNext
        Loop
    End If

    progbar.Value = 0
    lblStatus.Caption = ""
    
    MsgBox "Done!"
    Exit Sub
ErrHandlerImport:
    MsgBox Err.Description, vbOKOnly, "Error Occured"
End Sub

Private Sub Form_Load()
    'display all tables from the access file
    ShowTables
    'display all sheets from excel file
    ShowExcelSheets
    
End Sub

Private Sub ShowFields()
    Dim i As Integer
    Dim lst As ListItem
If Trim(cmbTables.Text) = "" Then Exit Sub

    Me.MousePointer = vbHourglass

    On Error GoTo ErrHandlerShowFields
    Set fldlst = OpenRS("SELECT * FROM `" & cmbTables.Text & "`")
    ListView1.ListItems.Clear
    
        For i = 0 To fldlst.Fields.Count - 1
        Set lst = ListView1.ListItems.Add
        lst.Text = fldlst.Fields(i).Name
        lst.SubItems(1) = GetType(fldlst.Fields(i).Type)
        lst.SmallIcon = 3
        Next i
    
    Me.MousePointer = vbNormal
    Exit Sub
ErrHandlerShowFields:
    Me.MousePointer = vbNormal
    MsgBox Err.Number & " : " & Err.Description, vbOKOnly, "Error occured!"
End Sub

Private Function GetType(typenumber As Integer) As String
    Select Case typenumber
    Case Is = 130
        GetType = "Char"
    Case Is = 202
        GetType = "Text"
    Case Is = 205
        GetType = "OLE Object"
    Case Is = 3
        GetType = "Long Integer"
    Case Is = 7
        GetType = "Date/Time"
    Case Is = 6
        GetType = "Currency"
    Case Is = 5
        GetType = "Double"
    Case Is = 11
        GetType = "Yes/No"
    Case Else
        GetType = typenumber
    End Select
End Function

Private Sub ShowExcelFields()
    Dim lst As ListItem
    Dim i As Integer
    If Trim(cmbSheets.Text) = "" Then Exit Sub
    Me.MousePointer = vbHourglass
    Set fldlstxls = OpenRSXLS("SELECT * FROM [" & cmbSheets.Text & "]")
    Set DataGrid1.DataSource = fldlstxls
    ListView2.ListItems.Clear
    For i = 0 To fldlstxls.Fields.Count - 1
        Set lst = ListView2.ListItems.Add
        lst.Text = fldlstxls.Fields(i).Name
        lst.SmallIcon = 6
    Next i
    Me.MousePointer = vbNormal
End Sub



Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        If ListView1.ListItems.Count = 0 Then Exit Sub
        PopupMenu pMenu
    End If
End Sub

Private Sub ListView1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim selectedField As String
    Dim objDrop As ListItem
    
    If ListView1.ListItems.Count = 0 Then
        Set ListView1.DropHighlight = Nothing
        Exit Sub
    End If
    
    Set objDrop = ListView1.HitTest(x, y)
    selectedField = ListView2.SelectedItem.Text
    
    If InStr(selectedField, "#") Then
        MsgBox "The field name contains invalid character like [#,double qoutes,single quotes]" & vbCrLf & "To solve this problem open the excel file and rename the field heading name", vbCritical, "Error"
        Set objDrop = Nothing
        Set ListView1.DropHighlight = Nothing
        Exit Sub
    End If
        
    ListView1.ListItems(objDrop.Index).SubItems(2) = selectedField
    Set objDrop = Nothing
    Set ListView1.DropHighlight = Nothing
End Sub

Private Sub ListView1_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    ListView1.DropHighlight = ListView1.HitTest(x, y)
End Sub

Private Sub mnuClear_Click()
    ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(2) = ""
    ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(3) = ""
End Sub

Private Sub mnuresetexcel_Click()
    Dim i As Integer
    For i = 1 To ListView1.ListItems.Count
        ListView1.ListItems(i).SubItems(2) = ""
    Next i
End Sub

Private Sub mnusetdefaultvalue_Click()
    fldsValue = ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(3)
    Form4.Show 1
    If Trim(fldsValue) <> "" Then
        ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(3) = fldsValue
    End If
End Sub

