VERSION 5.00
Begin VB.Form frmPortSettings 
   Caption         =   "Com Port Assignments"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   7890
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraBitRate 
      Caption         =   "Baud Rate"
      Height          =   1095
      Left            =   960
      TabIndex        =   23
      Top             =   6720
      Visible         =   0   'False
      Width           =   2655
      Begin VB.ComboBox cboBitRate 
         Height          =   315
         Left            =   480
         TabIndex        =   24
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Frame fraVoltmeter 
      Caption         =   "Voltmeter"
      Height          =   1215
      Left            =   960
      TabIndex        =   9
      Top             =   2760
      Width           =   5655
      Begin VB.TextBox txtVoltmeterPort 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   20
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optVoltmeterClosed 
         Caption         =   "CLOSED"
         Height          =   435
         Left            =   4200
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optVoltmeterOpen 
         Caption         =   "OPEN"
         Height          =   435
         Left            =   4200
         TabIndex        =   13
         Top             =   120
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "COM Port for Voltmeter"
         Height          =   405
         Left            =   600
         TabIndex        =   10
         Top             =   480
         Width           =   2685
      End
   End
   Begin VB.Frame fraControl 
      Caption         =   "Thunder Chamber Control"
      Height          =   1215
      Left            =   960
      TabIndex        =   6
      Top             =   5400
      Width           =   5655
      Begin VB.TextBox txtControlPort 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   22
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optControlClosed 
         Caption         =   "CLOSED"
         Height          =   435
         Left            =   4320
         TabIndex        =   18
         Top             =   720
         Width           =   1155
      End
      Begin VB.OptionButton optControlOpen 
         Caption         =   "OPEN"
         Height          =   435
         Left            =   4320
         TabIndex        =   17
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.Label lblControlPort 
         Caption         =   "COM Port for Chamber Control"
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Frame fraReference 
      Caption         =   "Reference"
      Height          =   1215
      Left            =   960
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   5655
      Begin VB.TextBox txtRefPort 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   21
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optRefClosed 
         Caption         =   "CLOSED"
         Height          =   315
         Left            =   4320
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton optRefOpen 
         Caption         =   "OPEN"
         Height          =   495
         Left            =   4320
         TabIndex        =   15
         Top             =   120
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "COM Port for Chamber Reference"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   600
         Width           =   2655
      End
   End
   Begin VB.Frame fraMainPort 
      Caption         =   "Main Port (Interface Board)"
      Height          =   1215
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   5655
      Begin VB.TextBox txtMainPort 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   19
         Top             =   480
         Width           =   495
      End
      Begin VB.OptionButton optMainClosed 
         Caption         =   "CLOSED"
         Height          =   315
         Left            =   4200
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton optMainOpen 
         Caption         =   "OPEN"
         Height          =   400
         Left            =   4200
         TabIndex        =   11
         Top             =   240
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.Label Label1 
         Caption         =   "COM Port for Main Board:"
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.TextBox txtStatus 
      Height          =   975
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   240
      Width           =   5655
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   492
      Left            =   3840
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   492
      Left            =   5400
      TabIndex        =   0
      Top             =   6960
      Width           =   1215
   End
End
Attribute VB_Name = "frmPortSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Enables users to select a serial port and bit rate.
'Revision history:
' 1-27-2008 JBS - Multiple Port version with VoltmeterPort added
' 1-27-2008 JBS - Multiple four Port version with Chamber ports added
' 1-27-2008 JBS - Added CheckIfPortsExist(). Not sure about how this works. TODO: check this!
' 2-4-2008 JBS  -   Changed port routines to be more reliable. Got rid of Combo boxes
'                   and complicated COM port polling routines.
' 4-10-2008 JBS - Modified initializePortSettings() so that MSComm3 does not get opened
'                   if Thunder chamber is being used.=, since only four ports a needed.

Dim allowPortChanges As Boolean


Private Sub cboBitRate_Change()
Call VbSetCommTimeouts(BitRate)
End Sub

Private Sub cmdCancel_Click()
    frmPortSettings.Hide
End Sub

Private Sub cmdOK_Click()

    allowPortChanges = False
    PortNumber = CInt(Val(txtMainPort))
    VoltmeterPortNumber = CInt(Val(txtVoltmeterPort))
    ChamberRefPortNumber = CInt(Val(txtRefPort))
    ChamberControlPortNumber = CInt(Val(txtControlPort))
    
    txtStatus.Text = "Checking settings on each COM port. Please wait..."
    

'The application's main form reads the new settings.
    DoEvents
    
If checkNewPortNumbers = True Then
    'All four COM ports are opened here. Only baud rate of 1200 is allowed:
    'PortOpen = frmMain.openComPorts(1200, PortNumber, VoltmeterPortNumber, ChamberRefPortNumber, ChamberControlPortNumber)
    PortOpen = frmMain.openComPorts(1200, PortNumber, VoltmeterPortNumber, ChamberControlPortNumber)
    If (PortOpen = True) Then
        frmPortSettings.Hide
    Else
        txtStatus.Text = "There was an error opening one of the ports." + vbCrLf + "Try opening each port individually."
    End If
End If
End Sub



Private Sub InitializeBitRateComboBox()
cboBitRate.AddItem ("300")
cboBitRate.AddItem ("1200")
cboBitRate.AddItem ("2400")
cboBitRate.AddItem ("4800")
cboBitRate.AddItem ("9600")
cboBitRate.AddItem ("19200")
cboBitRate.AddItem ("57600")
cboBitRate.AddItem ("115200")
End Sub

Public Sub SetBitRateComboBox()
'Set the index of the BitRate combo box.
Do
    cboBitRate.ListIndex = cboBitRate.ListIndex + 1
Loop Until Val(cboBitRate.Text) = BitRate _
    Or cboBitRate.ListIndex = cboBitRate.ListCount - 1
End Sub


Public Sub initializePortSettings()
    txtStatus.Text = ""
    allowPortChanges = False

    txtMainPort.Text = Format$(PortNumber)
    txtVoltmeterPort.Text = Format$(VoltmeterPortNumber)
    txtRefPort.Text = Format$(ChamberRefPortNumber)
    txtControlPort.Text = Format$(ChamberControlPortNumber)

    If (frmMain.MSComm1.PortOpen = False) Then
        optMainClosed.value = True
        txtStatus.Text = "Main COM Port (Interface board) is not open" + vbCrLf
    Else
        optMainOpen.value = True
        'optMainOpen.Enabled = False
    End If
    
    If (frmMain.MSComm2.PortOpen = False) Then
        optVoltmeterClosed.value = True
        txtStatus.Text = txtStatus.Text + "Voltmeter COM port is not open" + vbCrLf
    Else
        optVoltmeterOpen.value = True
        'optVoltmeterOpen.Enabled = False
    End If
    
    If (ChamberType = HUMILAB) Then
        If (frmMain.MSComm3.PortOpen = False) Then
            optRefClosed.value = True
            txtStatus.Text = txtStatus.Text + "Chamber reference COM port not open" + vbCrLf
        Else
            optRefOpen.value = True
            'optRefOpen.Enabled = False
        End If
    End If
    
    If (frmMain.MSComm4.PortOpen = False) Then
        optControlClosed.value = True
        txtStatus.Text = txtStatus.Text + "Chamber control COM port not open" + vbCrLf
    Else
        optControlOpen.value = True
        'optControlOpen.Enabled = False
    End If
    
    allowPortChanges = True
    If (txtStatus.Text = "") Then txtStatus.Text = "All COM ports are open!" + vbCrLf + "If you wish to change any port number," + vbCrLf + "please close it first."
End Sub

'This routine opens just the Main (Interface board) port on MSComm1
Private Sub optMainOpen_Click()
    If (allowPortChanges = True) Then
        PortNumber = CInt(Val(txtMainPort))
        If (True = checkNewPortNumbers) Then
            If (frmMain.MSComm1.PortOpen = True) Then frmMain.MSComm1.PortOpen = False
            frmMain.MSComm1.CommPort = PortNumber
            On Error Resume Next
            frmMain.MSComm1.PortOpen = True
            If (frmMain.MSComm1.PortOpen = False) Then
                txtStatus.Text = "Can't open COM port #" + Format$(PortNumber)
                Delay (300)
                optMainClosed.value = True
            Else
                txtStatus.Text = "Port #" + Format$(PortNumber) + " opened successfully!"
            End If
        Else
            optMainClosed.value = True
        End If
    End If
End Sub


Private Sub optVoltmeterOpen_Click()
    If (allowPortChanges = True) Then
        VoltmeterPortNumber = CInt(Val(txtVoltmeterPort))
        If (True = checkNewPortNumbers) Then
            If (frmMain.MSComm2.PortOpen = True) Then frmMain.MSComm2.PortOpen = False
            frmMain.MSComm2.CommPort = VoltmeterPortNumber
            On Error Resume Next
            frmMain.MSComm2.PortOpen = True
            If (frmMain.MSComm2.PortOpen = False) Then
                txtStatus.Text = "Can't open COM port #" + Format$(VoltmeterPortNumber)
                Delay (300)
                optVoltmeterClosed.value = True
            Else
                txtStatus.Text = "Port #" + Format$(VoltmeterPortNumber) + " opened successfully!"
            End If
        End If
    End If
End Sub

Private Sub optRefOpen_Click()
    If (allowPortChanges = True) Then
        ChamberRefPortNumber = CInt(Val(txtRefPort))
        If (True = checkNewPortNumbers) Then
            If (frmMain.MSComm3.PortOpen = True) Then frmMain.MSComm3.PortOpen = False
            frmMain.MSComm3.CommPort = ChamberRefPortNumber
            On Error Resume Next
            frmMain.MSComm3.PortOpen = True
            If (frmMain.MSComm3.PortOpen = False) Then
                txtStatus.Text = "Can't open COM port #" + Format$(ChamberRefPortNumber)
                Delay (300)
                optRefClosed.value = True
            Else
                txtStatus.Text = "Port #" + Format$(ChamberRefPortNumber) + " opened successfully!"
            End If
        End If
    End If
End Sub

Private Sub optControlOpen_Click()
    If (allowPortChanges = True) Then
        ChamberControlPortNumber = CInt(Val(txtControlPort))
        If (True = checkNewPortNumbers) Then
            If (frmMain.MSComm4.PortOpen = True) Then frmMain.MSComm4.PortOpen = False
            frmMain.MSComm4.CommPort = ChamberControlPortNumber
            On Error Resume Next
            frmMain.MSComm4.PortOpen = True
            If (frmMain.MSComm4.PortOpen = False) Then
                txtStatus.Text = "Can't open COM port #" + Format$(ChamberControlPortNumber)
                Delay (300)
                optControlClosed.value = True
            Else
                txtStatus.Text = "Port #" + Format$(ChamberControlPortNumber) + " opened successfully!"
            End If
        End If
    End If
End Sub


Private Sub optMainClosed_Click()
    txtMainPort.Enabled = True
    If (frmMain.MSComm1.PortOpen = True) Then frmMain.MSComm1.PortOpen = False
End Sub


Private Sub optVoltmeterClosed_Click()
    txtVoltmeterPort.Enabled = True
    If (frmMain.MSComm2.PortOpen = True) Then frmMain.MSComm2.PortOpen = False
End Sub

Private Sub optRefClosed_Click()
    txtRefPort.Enabled = True
    If (frmMain.MSComm3.PortOpen = True) Then frmMain.MSComm3.PortOpen = False
End Sub

Private Sub optControlClosed_Click()
    txtControlPort.Enabled = True
    If (frmMain.MSComm4.PortOpen = True) Then frmMain.MSComm4.PortOpen = False
End Sub


