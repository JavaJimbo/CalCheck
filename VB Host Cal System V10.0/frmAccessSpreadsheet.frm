VERSION 5.00
Begin VB.Form frmAccessSpreadsheet 
   Caption         =   "Form1"
   ClientHeight    =   1440
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   1440
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAccessFile 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   4095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Data"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtExcelFile 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Access Database"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Excel Spreadsheet"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "frmAccessSpreadsheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()
Dim excel_app As Object
Dim excel_sheet As Object
Dim max_row As Integer
Dim max_col As Integer
Dim row As Integer
Dim col As Integer

Dim statement As String
Dim new_value As String

    Screen.MousePointer = vbHourglass
    DoEvents

    ' Create the Excel application.
    Set excel_app = CreateObject("Excel.Application")

    ' Uncomment this line to make Excel visible.
'    excel_app.Visible = True

    ' Open the Excel spreadsheet.
    excel_app.Workbooks.Open filename:=txtExcelFile.Text

    ' Check for later versions.
    If Val(excel_app.Application.Version) >= 8 Then
        Set excel_sheet = excel_app.ActiveSheet
    Else
        Set excel_sheet = excel_app
    End If

    ' Get the last used row and column.
    max_row = excel_sheet.UsedRange.Rows.Count
    max_col = excel_sheet.UsedRange.Columns.Count

    ' Loop through the Excel spreadsheet rows,
    ' skipping the first row which contains
    ' the column headers.
    For row = 2 To max_row
        ' Compose an INSERT statement.
        statement = "INSERT INTO Books VALUES ("
        For col = 1 To max_col
            If col > 1 Then statement = statement & ","
            new_value = Trim$(excel_sheet.Cells(row, col).Value)
            If IsNumeric(new_value) Then
                statement = statement & _
                    new_value
            Else
                statement = statement & _
                    "'" & _
                    new_value & _
                    "'"
            End If
        Next col
        statement = statement & ")"

    Next row

    ' Comment the Close and Quit lines to keep
    ' Excel running so you can see it.

    ' Close the workbook saving changes.
    excel_app.ActiveWorkbook.Close True
    excel_app.Quit
    Set excel_sheet = Nothing
    Set excel_app = Nothing

    Screen.MousePointer = vbDefault
    MsgBox "Copied " & Format$(max_row - 1) & " values."
End Sub
' Note that this project contains a reference to
' Microsoft ADO Object Library 2.5.
Private Sub Form_Load()
Dim file_path As String

    file_path = App.Path
    If Right$(file_path, 1) <> "\" Then file_path = file_path & "\"
    txtExcelFile.Text = file_path & "Books.xls"
    txtAccessFile.Text = file_path & "Books.mdb"
End Sub


Private Sub Form_Resize()
Dim wid As Single

    wid = ScaleWidth - txtAccessFile.Left - 120
    If wid < 120 Then wid = 120
    txtAccessFile.Width = wid
    txtExcelFile.Width = wid
    cmdLoad.Left = (ScaleWidth - cmdLoad.Width) / 2
End Sub
