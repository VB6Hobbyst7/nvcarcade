VERSION 5.00
Begin VB.Form frmLoader 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   Picture         =   "frmLoader.frx":0000
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblList5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game5"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   285
      TabIndex        =   11
      Top             =   3480
      Width           =   4725
   End
   Begin VB.Image imgHighlight 
      Height          =   615
      Left            =   240
      Picture         =   "frmLoader.frx":E1044
      Top             =   3450
      Width           =   4815
   End
   Begin VB.Image imgControls 
      Height          =   1095
      Left            =   5520
      Picture         =   "frmLoader.frx":EAAEC
      Top             =   5880
      Width           =   3735
   End
   Begin VB.Label lblInfoYear 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Release Year"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5520
      TabIndex        =   10
      Top             =   4920
      Width           =   3735
   End
   Begin VB.Label lblInfoCo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturer"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   9
      Top             =   4440
      Width           =   3735
   End
   Begin VB.Label lblInfoName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Title"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      TabIndex        =   8
      Top             =   3960
      Width           =   3735
   End
   Begin VB.Image imgScreenshot 
      Appearance      =   0  'Flat
      Height          =   2505
      Left            =   5730
      Stretch         =   -1  'True
      Top             =   660
      Width           =   3330
   End
   Begin VB.Label lblList9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game9"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   285
      TabIndex        =   7
      Top             =   6240
      Width           =   4725
   End
   Begin VB.Label lblList8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game8"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   285
      TabIndex        =   6
      Top             =   5550
      Width           =   4725
   End
   Begin VB.Label lblList7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game7"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   285
      TabIndex        =   5
      Top             =   4860
      Width           =   4725
   End
   Begin VB.Label lblList6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game6"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   285
      TabIndex        =   4
      Top             =   4170
      Width           =   4725
   End
   Begin VB.Label lblList4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game4"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   285
      TabIndex        =   3
      Top             =   2790
      Width           =   4725
   End
   Begin VB.Label lblList3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game3"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   285
      TabIndex        =   2
      Top             =   2100
      Width           =   4725
   End
   Begin VB.Label lblList2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game2"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   285
      TabIndex        =   1
      Top             =   1410
      Width           =   4725
   End
   Begin VB.Label lblList1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game1"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   285
      TabIndex        =   0
      Top             =   720
      Width           =   4725
   End
End
Attribute VB_Name = "frmLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbGameList() As String
Dim index As Integer
Dim current As Integer
Dim ProgramPath As String
Dim ExecString As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim gay As String
    If KeyCode = vbKey1 Then          '1P Start
      gay = Shell(ExecString, vbNormalFocus)
    ElseIf KeyCode = vbKey2 Then      '2P Start
      gay = Shell(ExecString, vbNormalFocus)
    ElseIf KeyCode = vbKeyLeft Then   '1P Left Arrow
      gay = ScrollListUp(9)
    ElseIf KeyCode = vbKeyD Then      '2P Left Arrow
      gay = ScrollListUp(9)
    ElseIf KeyCode = vbKeyUp Then     '1P Up Arrow
      gay = ScrollListUp(1)
    ElseIf KeyCode = vbKeyR Then      '2P Up Arrow
      gay = ScrollListUp(1)
    ElseIf KeyCode = vbKeyRight Then  '1P Right Arrow
      gay = ScrollListDown(9)
    ElseIf KeyCode = vbKeyG Then      '2P Right Arrow
      gay = ScrollListDown(9)
    ElseIf KeyCode = vbKeyDown Then   '1P Down Arrow
      gay = ScrollListDown(1)
    ElseIf KeyCode = vbKeyF Then      '2P Down Arrow
      gay = ScrollListDown(1)
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim gay As String
    If KeyAscii = vbKeyTab Then        'Tab
      Unload Me
    End If
End Sub

Private Sub Form_Load()
    'Snap the form to the top left, this thing is meant for 640x480
    Me.Left = 0
    Me.Top = 0

    'Read in the game datbase to an array
    Dim FilePath As String
    
    'Get the path to the game database
    ProgramPath = App.Path
    If Right$(ProgramPath, 1) <> "\" Then ProgramPath = ProgramPath & "\"
    FilePath = ProgramPath & "gamelist.txt"
        
    Dim handle As Integer
    Dim sCurrentLine As String
    ReDim dbGameList(9000)
    index = 0
    
    'Read each line of the file to an array, don't bother with splitting yet
    'because multi-dimensional arrays in VB are ass slow
    handle = FreeFile
    Open FilePath For Input As handle
    Do While Not EOF(handle)
        Line Input #handle, sCurrentLine
        dbGameList(index) = sCurrentLine
        index = index + 1
    Loop
    
    'Size the game list back down to how many entries there really are
    ReDim Preserve dbGameList(index)
    
    'Set our game to the first one in the list
    current = 0
    Dim gay As String
    gay = ScrollListDown(0)
    
End Sub

Private Function ScrollListUp(x As Integer)
    'Scroll list up by x
    
    'Set the new current index
    current = current - x
    
    'The current item is the one highlighted in red, our actual start
    'needs to be earlier
    Dim start As Integer
    Dim tokens() As String
    start = current - 4
    'If our start is less than 0, we need to go to the end of
    'the list so it loops
    If start < 0 Then start = index + start
    
    Dim i As Integer
    
    For i = 0 To 9
      If start + i + 1 > index Then start = -1 * i
      If i = 0 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList1.Caption = tokens(0)
      ElseIf i = 1 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList2.Caption = tokens(0)
      ElseIf i = 2 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList3.Caption = tokens(0)
      ElseIf i = 3 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList4.Caption = tokens(0)
      ElseIf i = 4 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList5.Caption = tokens(0)
        lblInfoName.Caption = tokens(0)
        lblInfoCo.Caption = tokens(2)
        lblInfoYear.Caption = tokens(3)
        imgScreenshot = LoadPicture(ProgramPath & "snap\" & tokens(1) & ".gif")
        imgControls = LoadPicture(ProgramPath & "keys\" & tokens(4) & ".gif")
        ExecString = """" & ProgramPath & "emu\" & tokens(5) & "\" & tokens(5) & ".exe"" " & tokens(1)
      ElseIf i = 5 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList6.Caption = tokens(0)
      ElseIf i = 6 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList7.Caption = tokens(0)
      ElseIf i = 7 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList8.Caption = tokens(0)
      ElseIf i = 8 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList9.Caption = tokens(0)
      End If
    Next
      
End Function

Private Function ScrollListDown(x As Integer)
    'Scroll list down by x
    
    'Set the new current index
    current = current + x
    
    'The current item is the one highlighted in red, our actual start
    'needs to be earlier
    Dim start As Integer
    Dim tokens() As String
    start = current - 4
    'If our start is less than 0, we need to go to the end of
    'the list so it loops
    If start < 0 Then start = index + start
    
    Dim i As Integer
    
    For i = 0 To 9
      If start + i + 1 > index Then start = -1 * i
      If i = 0 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList1.Caption = tokens(0)
      ElseIf i = 1 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList2.Caption = tokens(0)
      ElseIf i = 2 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList3.Caption = tokens(0)
      ElseIf i = 3 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList4.Caption = tokens(0)
      ElseIf i = 4 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList5.Caption = tokens(0)
        lblInfoName.Caption = tokens(0)
        lblInfoCo.Caption = tokens(2)
        lblInfoYear.Caption = tokens(3)
        imgScreenshot = LoadPicture(ProgramPath & "snap\" & tokens(1) & ".gif")
        imgControls = LoadPicture(ProgramPath & "keys\" & tokens(4) & ".gif")
        ExecString = """" & ProgramPath & "emu\" & tokens(5) & "\" & tokens(5) & ".exe"" " & tokens(1)
      ElseIf i = 5 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList6.Caption = tokens(0)
      ElseIf i = 6 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList7.Caption = tokens(0)
      ElseIf i = 7 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList8.Caption = tokens(0)
      ElseIf i = 8 Then
        tokens = Split(dbGameList(start + i), vbTab)
        lblList9.Caption = tokens(0)
      End If
    Next
      
End Function
