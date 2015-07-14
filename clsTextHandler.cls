VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit


Private mstrFileName As String
Private mboolUseFirstLineForColumnNames As Boolean
Private mstrLastError As String
Private mstrDelimiter As String

Private mstrSchemaFile As String
Private mboolTreatAllDataAsString As Boolean

Public Property Get TreatAllDataAsString() As Boolean
   TreatAllDataAsString = mboolTreatAllDataAsString
End Property

Public Property Let TreatAllDataAsString(ByVal boolTreatAllDataAsString As Boolean)
   mboolTreatAllDataAsString = boolTreatAllDataAsString
End Property
Public Property Get Delimiter() As String
   Delimiter = mstrDelimiter
End Property

Public Property Let Delimiter(ByVal strDelimiter As String)
   mstrDelimiter = strDelimiter
End Property

Private Sub CreateSchemaFile(szFileFolder As String, szFileTitle As String)
    Dim handle As Integer
    Dim szTextType As String
    Dim sColumnDescription As String
        
'   Open up schema ini file for use with ODBC connections to text files
    handle = FreeFile
    Open szFileFolder & "schema.ini" For Output As handle     'Open file for output.
    Print #handle, "[" & Trim(szFileTitle) & "]" 'Required; ODBC; file; Name
    
    If mboolUseFirstLineForColumnNames Then
        Print #handle, "ColNameHeader=True"
    Else
        Print #handle, "ColNameHeader=False"
    End If
    
    Select Case mstrDelimiter
        Case ","
            szTextType = "CSVDelimited"
        Case vbTab
            szTextType = "TabDelimited"
        Case Else
            szTextType = "Delimited(" & mstrDelimiter & ")"
    End Select
    
    Print #handle, "Format=" & szTextType
    
    Print #handle, "MaxScanRows=25"
    Print #handle, "CharacterSet=OEM"
    
    If mboolTreatAllDataAsString Then
        sColumnDescription = GetColumnDescription()
        If Len(sColumnDescription) > 0 Then
            Print #handle, sColumnDescription
        End If
    End If
    
    Close #handle     ' Close file.
    
    mstrSchemaFile = szFileFolder & "schema.ini"
    
End Sub





Private Function GetColumnDescription() As String
    On Error GoTo errHandle
    Dim iHandle As Integer
    Dim sFirstLine As String
    Dim arrayColumns() As String
    Dim sColumnDescription As String
    Dim x As Long
    
    iHandle = FreeFile
    Open mstrFileName For Input As iHandle
    Line Input #iHandle, sFirstLine
    
    arrayColumns = Split(sFirstLine, mstrDelimiter)
    
    For x = 0 To UBound(arrayColumns)
        'validate
        arrayColumns(x) = Trim$(arrayColumns(x))
        If Len(arrayColumns(x)) = 0 Then arrayColumns(x) = "F" & x + 1
            
        sColumnDescription = sColumnDescription & "Col" & x + 1
        sColumnDescription = sColumnDescription & "=" & Chr$(34) & arrayColumns(x) & Chr$(34) & " Text" & vbCrLf
    Next
    GetColumnDescription = sColumnDescription
    
exitPoint:
    'clean up
    Close #iHandle
    Exit Function
errHandle:
    GoTo exitPoint

End Function


Private Sub EraseSchemaFile()
    On Error Resume Next
    Kill mstrSchemaFile
End Sub

Private Function GetFileRecordSet() As ADODB.Recordset
    On Error GoTo errGetFileRecordSet
    Dim sConnectionString As String
    Dim rs As ADODB.Recordset
    Dim szFileTitle As String
    Dim szFileFolder As String
    
    BreakUpFilename mstrFileName, szFileFolder, szFileTitle
    CreateSchemaFile szFileFolder, szFileTitle
    
    sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;"
    sConnectionString = sConnectionString + "Data Source=" & szFileFolder & ";"
    sConnectionString = sConnectionString + "Extended Properties=Text"
    
    Set rs = New ADODB.Recordset
    rs.Open "select * from [" & szFileTitle & "]", sConnectionString
    
    Set GetFileRecordSet = rs
    Set rs = Nothing
    
    Exit Function
errGetFileRecordSet:
    mstrLastError = "Error Number: " & Err.Number & vbCrLf & Err.Description
    Set GetFileRecordSet = Nothing
End Function

Private Sub BreakUpFilename(fullfile As String, CurPath As String, File As String)
    Dim Found As Boolean
    Dim textPos As Integer

    textPos = Len(fullfile)
    Do
        If Mid$(fullfile, textPos, 1) = "\" Then
            CurPath = Left$(fullfile, textPos - 1)
            'If Len(CurPath) = 2 Then    'i.e. c: d:
            '    CurPath = CurPath + "\"
            'End If
            
            File = Mid$(fullfile, textPos + 1)
            Found = True
        Else
            textPos = textPos - 1
            Found = False
        End If
    Loop While Not Found Or textPos < 1

    If Not Found Then File = fullfile
    CurPath = CurPath & "\"
End Sub


Public Property Get LastError() As String
   LastError = mstrLastError
End Property

Public Function Execute(rs As Recordset) As Boolean
    Set rs = GetFileRecordSet()
    If Not rs Is Nothing Then
        Execute = True
    Else
        Execute = False
    End If
End Function

Public Property Get UseFirstLineForColumnNames() As Boolean
   UseFirstLineForColumnNames = mboolUseFirstLineForColumnNames
End Property

Public Property Let UseFirstLineForColumnNames(ByVal boolUseFirstLineForColumnNames As Boolean)
   mboolUseFirstLineForColumnNames = boolUseFirstLineForColumnNames
End Property
Public Property Get FileName() As String
   FileName = mstrFileName
End Property

Public Property Let FileName(ByVal strFileName As String)
   mstrFileName = strFileName
End Property


Private Sub Class_Terminate()
'    EraseSchemaFile
End Sub


