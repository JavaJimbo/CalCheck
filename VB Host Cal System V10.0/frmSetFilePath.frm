VERSION 5.00
Begin VB.Form frmFilePath 
   Caption         =   "Set File Path"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   5
      Top             =   6480
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6840
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtLocalFolder 
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   960
      Width           =   5535
   End
   Begin VB.DirListBox localDirectoryBox 
      Height          =   3465
      Left            =   600
      TabIndex        =   2
      Top             =   2520
      Width           =   5535
   End
   Begin VB.DriveListBox localDriveBox 
      Height          =   315
      Left            =   600
      TabIndex        =   1
      Top             =   2040
      Width           =   5535
   End
   Begin VB.Label lblLocalPathName 
      Caption         =   "Local Path Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   7
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label lblLocalDriveName 
      Caption         =   "Local Drive Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6600
      TabIndex        =   6
      Top             =   240
      Width           =   5415
   End
   Begin VB.Label lblPathMessage 
      Caption         =   "Path for saving local spreadsheet:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "frmFilePath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' In general Declarations
   Dim tvn As Node
       Dim fs, d, s


Private Sub Command1_Click()

End Sub

Private Sub Dir1_Change()
    txtFilePath.Text = Dir1.Path
End Sub

Private Sub Drive1_Change()
    Dim DriveTestObj As Object
    Dim testDrive As String
    
    testDrive = Drive1.Drive
    Dir1.Path = testDrive
    txtFilePath.Text = testDrive
    
    Call CheckPath(testDrive)
    
End Sub

Sub CheckPath(drvPath)

    Dim DriveTestFlag As Boolean
    Dim PathTestFlag As Boolean
    Dim DriveString As String
    Dim BumString As String
    
    DriveTestFlag = False
    PathTestFlag = False
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set d = fs.GetDrive(fs.GetDriveName(drvPath))

    DriveString = fs.GetDriveName(drvPath)
    DriveTestFlag = fs.DriveExists(DriveString)
    If (DriveTestFlag <> True) Then
        result = MsgBox("Error - drive does not exist", vbOKCancel, "DRIVE ERROR")
    End If
           
    PathTestFlag = fs.FolderExists(drvPath)
    
'    s = "Drive " & d.DriveLetter & ": - "
'    s = s & d.VolumeName & vbCrLf
'    s = s & "Free Space: " & FormatNumber(d.FreeSpace / 1024, 0)
'    s = s & " Kbytes"
'    MsgBox s
End Sub

Private Sub cmdSave_Click()
    CheckPath (txtLocalFolder.Text)
End Sub

Private Sub Form_Load()
    frmFilePath.txtLocalFolder = strLocalFolder
End Sub




