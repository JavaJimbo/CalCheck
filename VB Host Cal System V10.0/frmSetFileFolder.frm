VERSION 5.00
Begin VB.Form frmFileFolder 
   Caption         =   "Set Spreadsheet Folders"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12450
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   12450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
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
      Left            =   5760
      TabIndex        =   13
      Top             =   6480
      Width           =   1215
   End
   Begin VB.CheckBox chkNetworkEnable 
      Caption         =   "Network enabled"
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
      Left            =   7200
      TabIndex        =   12
      Top             =   6000
      Width           =   2775
   End
   Begin VB.DirListBox networkFolderBox 
      BackColor       =   &H80000004&
      Height          =   3465
      Left            =   7200
      TabIndex        =   11
      Top             =   2400
      Width           =   4815
   End
   Begin VB.DriveListBox networkDriveBox 
      Height          =   315
      Left            =   7200
      TabIndex        =   10
      Top             =   1560
      Width           =   4815
   End
   Begin VB.TextBox txtNetworkFolder 
      Height          =   495
      Left            =   7200
      TabIndex        =   8
      Top             =   840
      Width           =   4815
   End
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
      Left            =   3960
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
      Left            =   2400
      TabIndex        =   4
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtLocalFolder 
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   840
      Width           =   4095
   End
   Begin VB.DirListBox localFolderBox 
      Height          =   3465
      Left            =   2400
      TabIndex        =   2
      Top             =   2400
      Width           =   4095
   End
   Begin VB.DriveListBox localDriveBox 
      Height          =   315
      Left            =   2400
      TabIndex        =   1
      Top             =   1560
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Network folder:"
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
      Left            =   7200
      TabIndex        =   9
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblFolder 
      Caption         =   "Select Folder:"
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
      Left            =   360
      TabIndex        =   7
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblDrive 
      Caption         =   "Select Drive:"
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
      Left            =   360
      TabIndex        =   6
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label lblFolderMessage 
      Caption         =   "Local folder:"
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
      Left            =   2400
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
End
Attribute VB_Name = "frmFileFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Function CheckPath(drvPath As String) As Boolean
    Dim DriveTestFlag As Boolean
    Dim PathTestFlag As Boolean
    Dim DriveString As String
    Dim BumString As String
    Dim errorCheck As Integer
    Dim fs As Variant
    Dim d As Variant
        
    DriveTestFlag = False
    PathTestFlag = False
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    Set d = fs.GetDrive(fs.GetDriveName(drvPath))
    errorCheck = Err.Number

    DriveString = fs.GetDriveName(drvPath)
    DriveTestFlag = fs.DriveExists(DriveString)
    
    If (DriveTestFlag = False) Then
        CheckPath = False
        result = MsgBox("Error - drive does not exist", vbOKOnly, "ERROR: Can't find drive " + DriveString)
    Else
        PathTestFlag = fs.FolderExists(drvPath)
        If (PathTestFlag = False) Then
            result = MsgBox("The folder you specified does not exist." + vbCrLf + "Save name anyway?", vbYesNo + vbDefaultButton2, "ERROR: Cannot find " + drvPath)
            If (result = vbYes) Then
                CheckPath = True
            Else
                CheckPath = False
            End If
        Else
            CheckPath = True
        End If
    End If
    
End Function

'This routine is not currently used here.
Function CheckFolder(strFolder As String) As Boolean
    Dim fs As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    CheckFolder = fs.FolderExists(strFolder)
End Function


Private Sub cmdCancel_Click()
    frmFileFolder.Hide
End Sub

Private Sub cmdReset_Click()
    strLocalFolder = "C:\Cal Data\"
    strNetworkFolder = "S:\SRH_HUMIDITY\Sensor Tip Calibration Run Data\"
    Call setupFileFolder
End Sub

'This routine checks the network and local folder paths
'to make sure they are valid, then it saves them.
'The network drive is optional. The local drive
'is mandatory. So the local drive must be valid,
'regardless of whether the network drive is valid.
'Otherwise this routine won't save the new settings.

Private Sub cmdSave_Click()

    If (chkNetworkEnable.value = Checked) Then
        'Check network drive and folder:
        If (CheckPath(txtNetworkFolder.Text) = True) Then
            strNetworkFolder = txtNetworkFolder.Text
        End If
    Else
        strNetworkFolder = txtNetworkFolder.Text
    End If
    
    'Check local drive and folder then save them if valid:
    If (CheckPath(txtLocalFolder.Text) = True) Then
        strLocalFolder = txtLocalFolder.Text
        Call SaveSettings
        frmFileFolder.Hide
    End If
End Sub

Public Sub setupFileFolder()
    localDriveBox.Drive = strLocalFolder
    localFolderBox.Path = strLocalFolder
    txtLocalFolder.Text = strLocalFolder

    networkDriveBox.Drive = strNetworkFolder
    networkFolderBox.Path = strNetworkFolder
    txtNetworkFolder.Text = strNetworkFolder
End Sub


Private Sub localFolderBox_Change()
    txtLocalFolder.Text = localFolderBox.Path
End Sub

Private Sub localDriveBox_Change()
    localFolderBox.Path = localDriveBox.Drive
    txtLocalFolder.Text = localFolderBox.Path
End Sub

Private Sub networkDriveBox_Change()
    networkFolderBox.Path = networkDriveBox.Drive
    txtNetworkFolder.Text = networkFolderBox.Path
End Sub

Private Sub networkFolderBox_Change()
    txtNetworkFolder.Text = networkFolderBox.Path
End Sub

