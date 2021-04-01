VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUpdate 
   Caption         =   "Automatic Update"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   5895
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2880
      TabIndex        =   20
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2880
      TabIndex        =   19
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox TxtUpdateFileName 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3120
      TabIndex        =   23
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1560
      TabIndex        =   18
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1560
      TabIndex        =   17
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Close"
      Height          =   375
      Left            =   600
      TabIndex        =   13
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Update"
      Enabled         =   0   'False
      Height          =   375
      Left            =   600
      TabIndex        =   12
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Check For Update"
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox TxtUpdateVersion 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Top             =   360
      Width           =   2655
   End
   Begin VB.TextBox TxtUpdateSize 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3120
      TabIndex        =   5
      Top             =   2160
      Width           =   2655
   End
   Begin VB.TextBox TxtUpdateInfo 
      Height          =   1095
      Left            =   3120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2760
      Width           =   2655
   End
   Begin VB.TextBox TxtUpdateDate 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3120
      TabIndex        =   3
      Top             =   960
      Width           =   2655
   End
   Begin VB.TextBox TxtCurVersion 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   2655
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4335
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Caption         =   "Time Left: 00:00:00"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   3600
      Width           =   3015
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4680
      TabIndex        =   25
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   4680
      TabIndex        =   24
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Update File Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   22
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   3960
      Width           =   5895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   2880
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Update Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "File Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   9
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "Update Info."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   2520
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Update Release Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   7
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Current Version"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "FrmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mFTP As cFTP
Attribute mFTP.VB_VarHelpID = -1
Private BeginTransfer                   As Single
Private TransferRate                    As Single
Private Declare Function ClipCursor Lib "user32" _
    (lpRect As Any) As Long

Private FilePathName As String
Private Filename As String
Private FormName As String

Private Declare Function OSGetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function OSGetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function OSWritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function OSWritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Declare Function OSGetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function OSGetProfileSection Lib "kernel32" Alias "GetProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function OSGetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long

Private Declare Function OSWriteProfileSection Lib "kernel32" Alias "WriteProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String) As Long
Private Declare Function OSWriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long

Private Const nBUFSIZEINI = 1024
Private Const nBUFSIZEINIALL = 4096
Private NewVersion As String
Private OldVersion As String
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Function ConvertTime(ByVal TheTime As Single) As String
    Dim NewTime                         As String
    Dim Sec                             As Single
    Dim Min                             As Single
    Dim H                               As Single
    If TheTime > 60 Then
        Sec = TheTime
        Min = Sec / 60
        Min = Int(Min)
        Sec = Sec - Min * 60
        H = Int(Min / 60)
        Min = Min - H * 60
        NewTime = H & ":" & Min & ":" & Sec
        If H < 0 Then H = 0
        If Min < 0 Then Min = 0
        If Sec < 0 Then Sec = 0
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
    If TheTime < 60 Then
        NewTime = "00:00:" & TheTime
        NewTime = Format(NewTime, "HH:MM:SS")
        ConvertTime = NewTime
    End If
End Function

Public Function RunUpdate(UpdateURL As String)
HyperJump UpdateURL
End Function
Private Function HyperJump(ByVal url As String) As Long
    HyperJump = ShellExecute(0&, vbNullString, url, vbNullString, vbNullString, vbNormalFocus)
End Function
Private Sub Command2_Click()
    Dim lTimer As Long
    Dim strRemote As String
    Dim strLocal As String
BeginTransfer = Timer
strRemote = Text4.Text & "/" & Text5.Text
strLocal = App.Path & "\Updates\" & Text5.Text
lTimer = Timer
Label3.Caption = "Status: Downloading Updates..."
If Text6.Text = "Yes" Then
MsgBox "This update requires that the program not be running. When the Update is downloaded and started this program will close."
End If


If mFTP.OpenConnection(Text3.Text, "anonymous", "anonymous") Then
mFTP.SetFTPDirectory "/"
If Not mFTP.FTPDownloadFile(strLocal, strRemote) Then
Label3.Caption = "Status: Error"
MsgBox mFTP.GetLastErrorMessage
Else
Label3.Caption = "Status: Download Complete"
DoEvents

RunUpdate App.Path & "\Updates\" & Text5.Text
DoEvents
If Text6.Text = "Yes" Then
End
End If
End If
DoEvents
mFTP.CloseConnection
End If


End Sub

Private Sub Command3_Click()
End
End Sub

Public Sub mFTP_FileTransferProgress(lCurrentBytes As Long, lTotalBytes As Long)
On Error Resume Next
Dim j As Long
Dim j2 As Long
TransferRate = Format(Int(lCurrentBytes / (Timer - BeginTransfer)) / 1000, "####.00")
    PB.Max = lTotalBytes
    PB.Min = 0
  j = PB.Value
  j2 = PB.Value \ 1024
  DoEvents
        PB.Value = lCurrentBytes
        DoEvents
        PB.ToolTipText = PB.Value & " Bytes of " & PB.Max & " Bytes Transfered"
        DoEvents
        Label7.Caption = PB.Value \ 1024 & " KB of " & PB.Max \ 1024 & " KB Transfered"
        DoEvents
        Label9.Caption = Format$(CLng((j / PB.Max) * 100)) + "%"
        DoEvents
        Label10.Caption = Format(TransferRate, "##.#0#") & " Kbps"
        Label11.Caption = "Time Left: " & ConvertTime(Int(((PB.Max - PB.Value) / 1024) / TransferRate))
        If PB.Value = PB.Max Then
        Label9.Caption = "100%"
        End If
End Sub
Private Function GetPrivateProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String, ByVal szFileName As String) As String
   ' *** Get an entry in the inifile ***

   Dim szTmp                     As String
   Dim nRet                      As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetPrivateProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL, szFileName)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetPrivateProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI, szFileName)
   End If
   GetPrivateProfileString = Left$(szTmp, nRet)

End Function
Private Function GetProfileString(ByVal szSection As String, ByVal szEntry As Variant, ByVal szDefault As String) As String
   ' *** Get an entry in the WIN inifile ***

   Dim szTmp                    As String
   Dim nRet                     As Long

   If (IsNull(szEntry)) Then
      ' *** Get names of all entries in the named Section ***
      szTmp = String$(nBUFSIZEINIALL, 0)
      nRet = OSGetProfileString(szSection, 0&, szDefault, szTmp, nBUFSIZEINIALL)
   Else
      ' *** Get the value of the named Entry ***
      szTmp = String$(nBUFSIZEINI, 0)
      nRet = OSGetProfileString(szSection, CStr(szEntry), szDefault, szTmp, nBUFSIZEINI)
   End If
   GetProfileString = Left$(szTmp, nRet)

End Function
Private Sub Command1_Click()
    Dim lTimer As Long
    Dim strRemote As String
    Dim strLocal As String
Dim NewVer As String
Dim Oldver As String
Dim url As String
Dim AppDir As String
Dim YourVersion As String
Dim DOR As String
Dim FileSize As String
Dim WhatNew As String
Dim AppTitle As String
AppTitle = App_Title + ": Auto Update"
NewVer = "none"
Oldver = "none"
AppDir = App.Path
YourVersion = TxtCurVersion.Text
    
strRemote = Text2.Text & "Update.inf"
strLocal = App.Path & "\Updates\Update.inf"
lTimer = Timer


If mFTP.OpenConnection(Text1.Text, "anonymous", "anonymous") Then
mFTP.SetFTPDirectory "/"
If Not mFTP.FTPDownloadFile(strLocal, strRemote) Then
Label3.Caption = "Status: Error"
MsgBox mFTP.GetLastErrorMessage
Else
Label3.Caption = "Status: Download Complete"
End If
DoEvents
mFTP.CloseConnection
End If


  
'Gets your Version
Oldver = YourVersion

'State & Access 'Version.inf' file
FilePathName = AppDir + "\Updates\Update.inf"
NewVer = GetPrivateProfileString("Version", "Version", "", FilePathName)
NewVersion = NewVer
DOR = GetPrivateProfileString("Version", "DOR", "", FilePathName)
FileSize1 = GetPrivateProfileString("Version", "Filesize1", "", FilePathName)
FileSize2 = GetPrivateProfileString("Version", "Filesize2", "", FilePathName)
WhatNew = GetPrivateProfileString("Version", "Whatsnew", "", FilePathName)
Downloadsite = GetPrivateProfileString("Version", "DownloadSite", "", FilePathName)
DownloadPath = GetPrivateProfileString("Version", "DownloadPath", "", FilePathName)
DownloadFile = GetPrivateProfileString("Version", "DownloadFile", "", FilePathName)
CloseProgramBeforeUpdate = GetPrivateProfileString("Version", "CloseProgramBeforeUpdate", "", FilePathName)

'Compare for newer version
If Oldver > NewVer Then
Command1.Enabled = False
TxtUpdateVersion.Text = NewVer
TxtUpdateFileName = ""
TxtUpdateDate.Text = ""
TxtUpdateSize.Text = ""
TxtUpdateInfo.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Label3.Caption = "Status: You are Up to Date"
Else
TxtUpdateVersion.Text = NewVer
TxtUpdateDate.Text = DOR
TxtUpdateFileName = DownloadFile
TxtUpdateSize.Text = FileSize1 & " " & FileSize2
TxtUpdateInfo.Text = WhatNew
Text3.Text = Downloadsite
Text4.Text = DownloadPath
Text5.Text = DownloadFile
Text6.Text = CloseProgramBeforeUpdate
Command2.Enabled = True
Label3.Caption = "Status: Update Available"
End If
           
End Sub

Private Sub Form_Load()
Dim intfile As Integer
Dim pass As String
Dim pass2 As String
TxtCurVersion.Text = App.Major & "." & App.Minor & "." & App.Revision
intfile = FreeFile
  Open App.Path & "\Updates\UpdateSettings.ini" For Input As #intfile
  Input #intfile, pass
  Input #intfile, pass2
  Text1.Text = pass
  Text2.Text = pass2
  Close #intfile
  DoEvents

Set mFTP = New cFTP
   mFTP.SetModeActive
   mFTP.SetTransferBinary
End Sub

