VERSION 5.00
Begin VB.Form frmPortSettings 
   Caption         =   "Serial Port Complete"
   ClientHeight    =   2256
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3180
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   2256
   ScaleWidth      =   3180
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtStatus 
      Height          =   612
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   2772
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   492
      Left            =   2040
      TabIndex        =   5
      Top             =   960
      Width           =   972
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   492
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   1092
   End
   Begin VB.Frame fraBitRate 
      Caption         =   "Bit Rate"
      Height          =   612
      Left            =   1680
      TabIndex        =   2
      Top             =   120
      Width           =   1332
      Begin VB.ComboBox cboBitRate 
         Height          =   288
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   1092
      End
   End
   Begin VB.Frame fraPort 
      Caption         =   "Port"
      Height          =   612
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1332
      Begin VB.ComboBox cboPort 
         Height          =   288
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1092
      End
   End
End
Attribute VB_Name = "frmPortSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Enables users to select a serial port and bit rate.

Private Sub cboBitRate_Change()
Call VbSetCommTimeouts(BitRate)
End Sub

Private Sub cmdCancel_Click()
Hide
End Sub

Private Sub cmdOK_Click()
'The application's main form reads the new settings.
Call GetNewSettings
ValidPort = fncCheckForValidPort
If ValidPort = True Then
    Hide
End If
End Sub

Private Sub Form_Load()
Dim Count As Integer
Call FindPorts
'Set default values if a retrieved setting is invalid.
'Be sure the selected port exists.
PortExists = False
For Count = 1 To UBound(CommPorts())
    'Compare the selected port number with the names in CommPorts.
    If "COM" & CStr(PortNumber) = CommPorts(Count) Then
        PortExists = True
    End If
Next Count
'Display the Setup window if the retrieved port number is invalid,
'or if the port is unavailable.
ValidPort = fncCheckForValidPort
Call InitializePortComboBox
Call InitializeBitRateComboBox
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

Private Sub InitializePortComboBox()
Dim Count As Integer
For Count = 1 To UBound(CommPorts())
    cboPort.AddItem CommPorts(Count)
Next Count
End Sub

Public Function fncCheckForValidPort()
'Find out if the selected port exists and is available.
'If not, display the Setup window to enable user to select another.
fncCheckForValidPort = True
If PortNumber < 1 Then
    Show
    cboPort.ListIndex = -1
    txtStatus.Text = "Please select a COM port."
    fncCheckForValidPort = False
    End If
If PortExists = False Then
    Show
    cboPort.ListIndex = -1
    txtStatus.Text = "COM" & PortNumber & " is unavailable. Please select a different port."
    fncCheckForValidPort = False
End If
End Function

Public Sub SetBitRateComboBox()
'Set the index of the BitRate combo box.
Do
    cboBitRate.ListIndex = cboBitRate.ListIndex + 1
Loop Until Val(cboBitRate.Text) = BitRate _
    Or cboBitRate.ListIndex = cboBitRate.ListCount - 1
End Sub

Public Sub SetPortComboBox()
'Set the index of the Port combo box.
'Read the numeric characters in the name of the selected port:
'"COM1", "COM2", etc.
Do
    cboPort.ListIndex = cboPort.ListIndex + 1
Loop Until _
    Val(Right(cboPort.Text, (Len(cboPort.Text) - 3))) = PortNumber _
    Or cboPort.ListIndex = cboPort.ListCount - 1
End Sub

