VERSION 5.00
Begin VB.Form frmExcel 
   Caption         =   "Save Spreadsheet "
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Data"
      Default         =   -1  'True
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtExcelFile 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Save As"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdLoad_Click()
Dim excel_app As Excel.Application
Dim row As Integer

    Screen.MousePointer = vbHourglass
    DoEvents

    ' Create the Excel application.
    Set excel_app = CreateObject("Excel.Application")

    ' Uncomment this line to make Excel visible.
    excel_app.Visible = True

    ' Create a new spreadsheet.
    excel_app.Workbooks.Add
     
    ' Insert data into Excel.
    With excel_app
        .Range("A1").Select
        .ActiveCell.FormulaR1C1 = "Title"
        .Columns("A:A").ColumnWidth = 35
        With .Selection.Font
            .Name = "Arial"
            .FontStyle = "Bold"
            .Size = 10
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = 5
        End With

        .Columns("B:B").ColumnWidth = 13
        .Range("B1").Select
        .ActiveCell.FormulaR1C1 = "ISBN"
        With .Selection.Font
            .Name = "Arial"
            .FontStyle = "Bold"
            .Size = 10
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = xlUnderlineStyleNone
            .ColorIndex = 5
        End With

        row = 2
        .Range("A" & Format$(row)).Select
        .ActiveCell.FormulaR1C1 = "'Advanced Visual Basic Techniques"
        .Range("B" & Format$(row)).Select
        .ActiveCell.FormulaR1C1 = "'0-471-18881-6"

        row = row + 1
        .Range("A" & Format$(row)).Select
        .ActiveCell.FormulaR1C1 = "'Ready-to-Run Visual Basic Algorithms"
        .Range("B" & Format$(row)).Select
        .ActiveCell.FormulaR1C1 = "'0-471-24268-3"

        row = row + 1
        .Range("A" & Format$(row)).Select
        .ActiveCell.FormulaR1C1 = "'Custom Controls Library"
        .Range("B" & Format$(row)).Select
        .ActiveCell.FormulaR1C1 = "'0-471-24267-5"

        row = row + 1
        .Range("A" & Format$(row)).Select
        .ActiveCell.FormulaR1C1 = "'Bug Proofing Visual Basic"
        .Range("B" & Format$(row)).Select
        .ActiveCell.FormulaR1C1 = "'0-471-32351-9"

        row = row + 1
        .Range("A" & Format$(row)).Select
        .ActiveCell.FormulaR1C1 = "'Ready-to-Run Visual Basic Code Library"
        .Range("B" & Format$(row)).Select
        .ActiveCell.FormulaR1C1 = "'0-471-33345-X"

        row = row + 1
        .Range("A" & Format$(row)).Select
        .ActiveCell.FormulaR1C1 = "'Visual Basic Graphics Programming"
        .Range("B" & Format$(row)).Select
        .ActiveCell.FormulaR1C1 = "'0-471-35599-2"

        ' Save the results.
        .ActiveWorkbook.SaveAs filename:=txtExcelFile, _
            FileFormat:=xlNormal, _
            Password:="", _
            WriteResPassword:="", _
            ReadOnlyRecommended:=False, _
            CreateBackup:=False
    End With

    ' Comment the rest of the lines to keep
    ' Excel running so you can see it.

    ' Close the workbook without saving.
    excel_app.ActiveWorkbook.Close False

    ' Close Excel.
    excel_app.Quit
    Set excel_app = Nothing

    Screen.MousePointer = vbDefault
    MsgBox "Ok"
End Sub

Private Sub Form_Load()
Dim file_path As String

    file_path = App.Path
    If Right$(file_path, 1) <> "\" Then file_path = file_path & "\"
    txtExcelFile.Text = file_path & "Books.xls"
End Sub


