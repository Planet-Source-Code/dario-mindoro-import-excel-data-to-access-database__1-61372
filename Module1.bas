Attribute VB_Name = "Module1"
'This is for opening a file
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Public Type SHFILEOPSTRUCT
     hwnd As Long
     wFunc As Long
     pFrom As String
     pTo As String
     fFlags As Integer
     fAnyOperationsAborted As Boolean
     hNameMappings As Long
     lpszProgressTitle As String
End Type

Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Const FO_COPY = &H2
Public Const FOF_ALLOWUNDO = &H40
Public Counts As Long


Type OPERATIONSTRUCT
    excelfile As String
    accessfile As String
    accesspassword As String
End Type

Public Type QUERYSTRUCT
    destinationfields As String
    sourcefields As String
    defaultvalues As String
End Type

Public IMPORTINFO As OPERATIONSTRUCT
Public fldsValue As String


Public Function FileExists(ByVal sPathName As String) As Boolean
    On Error Resume Next
    FileExists = (GetAttr(sPathName) And vbNormal) = vbNormal
    On Error GoTo 0
End Function

Public Sub EndProgram()
    Set Cn = Nothing
    Set CnXls = Nothing
    End
End Sub


Public Sub CopyFileWindowsWay(SourceFile As String, DestinationFile As String)

     Dim lngReturn As Long
     Dim typFileOperation As SHFILEOPSTRUCT

     With typFileOperation
        .hwnd = 0
        .wFunc = FO_COPY
        .pFrom = SourceFile & vbNullChar & vbNullChar 'source file
        .pTo = DestinationFile & vbNullChar & vbNullChar 'destination file
        .fFlags = FOF_ALLOWUNDO
     End With
     lngReturn = SHFileOperation(typFileOperation)
End Sub

'    CopyFileWindowsWay Source, Destination



