VERSION 5.00
Begin VB.Form frmPassword 
   Caption         =   "Diagnostics Password"
   ClientHeight    =   1035
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1035
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2040
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label lblPassword 
      Caption         =   "Enter password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtPassword_Change()
    If (txtPassword.Text = PASSWORD) Then
        Call frmMain.SetupDiagnostics
        frmPassword.Hide
    End If
End Sub
