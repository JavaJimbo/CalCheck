Attribute VB_Name = "ExcelRoutines"
'Routines for reading, writing, and creating Excel Spreadsheet
Public excel_sheet As Object
Public excel_app As Object
Public dataFilename As String

'These constants set the background colors for the STATUS column
'when the final PASS/FAIL data is written:
Global Const TWO_PERCENT_COLOR = 4     'This is GREEN
Global Const THREE_PERCENT_COLOR = 6   'This is YELLOW
Global Const FIVE_PERCENT_COLOR = 46   'Tis is ORANGE
Global Const FAIL_COLOR = 3            'This is RED
Global Const NO_COLOR = 2              'This is WHITE - the normal background color for a spreadsheet cell
Global Const BLACK = 1
Global Const BLUE = 5

Global Const ROW_OFFSET = 9
Global Const SETPOINT_ROW = 8
Global Const TITLE_ROW = SETPOINT_ROW + 1
Global Const OFFSET = TITLE_ROW

Global Const BLANK = "   "
Global Const UUT_WIDTH = 5
Global Const REF_WIDTH = 5
Global Const ERR_WIDTH = 5
Global Const COMMENT_WIDTH = 10

Global Const SENSOR_COLUMN = 1
Global Const COMMENT_COLUMN = 2
Global Const STATUS_COLUMN = 3

Global Const POT1_COLUMN = 4
Global Const CAL_REF1_COLUMN = 5
Global Const CAL_UUT1_COLUMN = 6
Global Const CAL_ERR1_COLUMN = 7
Global Const BLK0_COLUMN = 8

Global Const POT2_COLUMN = 9
Global Const CAL_REF2_COLUMN = 10
Global Const CAL_UUT2_COLUMN = 11
Global Const CAL_ERR2_COLUMN = 12

Global Const BLK1_COLUMN = 13
Global Const REF1_COLUMN = 14
Global Const UUT1_COLUMN = 15
Global Const ERR1_COLUMN = 16

Global Const BLK2_COLUMN = 17
Global Const REF2_COLUMN = 18
Global Const UUT2_COLUMN = 19
Global Const ERR2_COLUMN = 20
Global Const BLK3_COLUMN = 21
Global Const REF3_COLUMN = 22
Global Const UUT3_COLUMN = 23
Global Const ERR3_COLUMN = 24
Global Const BLK4_COLUMN = 25
Global Const REF4_COLUMN = 26
Global Const UUT4_COLUMN = 27
Global Const ERR4_COLUMN = 28
Global Const BLK5_COLUMN = 29
Global Const REF5_COLUMN = 30
Global Const UUT5_COLUMN = 31
Global Const ERR5_COLUMN = 32

 
Public Sub createNewSpreadsheet()
Dim ColumnTitle(1 To 31) As String
Dim ColumnWidth(1 To 31) As Integer
Dim sheet As Object
Dim currentDateTime As String
Dim SensorTipNumber As Integer
Dim setpointColumn As Integer
Dim setpoint(1 To 4) As Integer
Dim calpoint(1 To 2) As Integer
Dim i As Integer
Dim result As Integer

ColumnTitle(1) = "Sensor"
ColumnWidth(1) = Len("Sensor")

ColumnTitle(2) = "Comments"
ColumnWidth(2) = Len("Comments")


ColumnTitle(3) = "Status"
ColumnWidth(3) = Len("Status    ") 'Allow a little extra room for comments, ie: "FAIL POT ERROR", etc.
ColumnTitle(4) = "Pot 1"
ColumnWidth(4) = Len("Pot 1")

ColumnTitle(5) = "Ref"
ColumnWidth(5) = REF_WIDTH
ColumnTitle(6) = "UUT"
ColumnWidth(6) = UUT_WIDTH
ColumnTitle(7) = "Error"
ColumnWidth(7) = ERR_WIDTH
ColumnTitle(8) = BLANK
ColumnWidth(8) = Len(BLANK)


ColumnTitle(9) = "Pot 2"
ColumnWidth(9) = Len("Pot 2")
ColumnTitle(10) = "Ref"
ColumnWidth(10) = REF_WIDTH
ColumnTitle(11) = "UUT"
ColumnWidth(11) = UUT_WIDTH
ColumnTitle(12) = "Error"
ColumnWidth(12) = ERR_WIDTH

ColumnTitle(13) = BLANK
ColumnWidth(13) = Len(BLANK)
ColumnTitle(14) = "Ref"
ColumnWidth(14) = REF_WIDTH
ColumnTitle(15) = "UUT"
ColumnWidth(15) = UUT_WIDTH
ColumnTitle(16) = "Error"
ColumnWidth(16) = ERR_WIDTH
ColumnTitle(17) = BLANK
ColumnWidth(17) = Len(BLANK)

ColumnTitle(18) = "Ref"
ColumnWidth(18) = REF_WIDTH
ColumnTitle(19) = "UUT"
ColumnWidth(19) = UUT_WIDTH
ColumnTitle(20) = "Error"
ColumnWidth(20) = ERR_WIDTH
ColumnTitle(21) = BLANK
ColumnWidth(21) = Len(BLANK)
ColumnTitle(22) = "Ref"
ColumnWidth(22) = REF_WIDTH
ColumnTitle(23) = "UUT"
ColumnWidth(23) = UUT_WIDTH
ColumnTitle(24) = "Error"
ColumnWidth(24) = ERR_WIDTH
ColumnTitle(25) = BLANK
ColumnWidth(25) = Len(BLANK)
ColumnTitle(26) = "Ref"
ColumnWidth(26) = REF_WIDTH
ColumnTitle(27) = "UUT"
ColumnWidth(27) = UUT_WIDTH
ColumnTitle(28) = "Error"
ColumnWidth(28) = ERR_WIDTH

calpoint(1) = 20    'This the BALANCE setpoint
calpoint(2) = 80    'This is the SPAN setpoint
setpoint(1) = 90    'These are the four vaildation setpoints
setpoint(2) = 50
setpoint(3) = 10
setpoint(4) = 50

    'Set output filename to default and delete existing file by that name:
    dataFilename = DEFAULT_SHEETNAME
    On Error Resume Next
    Kill dataFilename
        
    'Make sure Excel isn't already running. If it is, we need to close it down:
    result = vbOK
    Do
        On Error Resume Next
        Set excel_app = GetObject(, "Excel.Application")
        errorCheck = Err.Number
        '429 means Excel is NOT running. 0 means it is already running.
        'If it is already running, then we don't need to start it up.
    
        'Close whatever is already running
        If (errorCheck = 0) Then
            result = MsgBox("Please save and close spreadsheet.", vbOKCancel + vbCritical + vbDefaultButton1, "EXCEL is already running.")
                       
            If (result = vbOK) Then
                excel_app.Visible = True
                result = MsgBox("Click OK to continue.", vbOKCancel + vbDefaultButton1, "Ready to go?")
                excel_app.Quit
                Set excel_sheet = Nothing
                Set excel_app = Nothing
            End If
            
            If (result = vbCancel) Then
                dataFilename = ""
                RunFlag = False
                GoTo Quit
            End If
            
        End If
    Loop While (errorCheck = 0)
    
    
    On Error Resume Next
    Kill dataFilename
         
    Set excel_app = CreateObject("Excel.Application")
    
    ' Uncomment this line to make Excel visible.
    If (mnuDiagnostics.Checked = True) Then excel_app.Visible = True
         
    ' Check for later versions.
     If Val(excel_app.Application.VERSION) >= 8 Then
         Set excel_sheet = excel_app.ActiveSheet
     Else
         Set excel_sheet = excel_app
     End If
     
    ' Create a new spreadsheet.
    excel_app.Workbooks.Add
    
    ' Insert data into Excel.
    With excel_app
    
        'First row is the title header with software version:
        row = 1
        .Cells(row, 1).Select
        .ActiveCell.FormulaR1C1 = VERSION
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
        
        'Second row is the calibration start time:
        row = 2
        currentDateTime = Now
        .Cells(row, 1).Select
        .ActiveCell.FormulaR1C1 = "Calibration start time: " + currentDateTime 'TODO: can this be fixed?
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
        
        'Third row is the calibration completion time:
        row = 3
        .Cells(row, 1).Select
        .ActiveCell.FormulaR1C1 = "Calibration completion time: "
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
        
        'Fourth row is last task completed:
        row = 4
        .Cells(row, 1).Select
        .ActiveCell.FormulaR1C1 = "Current task:"
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

        'Fifth row is additional notes:
        row = 5
        .Cells(row, 1).Select
        .ActiveCell.FormulaR1C1 = "Additional Notes:"
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
        
        'The TITLE row contains column headings for each
        'individual column, ie: "UUT", REF", "ERROR" etc.
        row = TITLE_ROW
        For i = 1 To 31
            .Cells(row, i).Select
            .ActiveCell.FormulaR1C1 = ColumnTitle(i)
            .Columns(ColumnRange).ColumnWidth = ColumnWidth(i)
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
        Next i
       
        'The SETPOINT_ROW contains column headings
        'for each setpoint ie: "Setpoint 1: 30%" etc.
        'The first setpoint column begins at the REFERENCE 1 COLUMN
        row = SETPOINT_ROW
        
        'These are the two calibration setpoints:
        setpointColumn = POT1_COLUMN
        For i = 1 To 2
            .Cells(row, setpointColumn).Select
            .ActiveCell.FormulaR1C1 = "Cal " + Format$(i) + ": " + Format$(calpoint(i)) + "%, "
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
            setpointColumn = setpointColumn + 5
        Next i
        
        'These are the four validation setpoints:
        setpointColumn = REF1_COLUMN
        For i = 1 To 4
            .Cells(row, setpointColumn).Select
            .ActiveCell.FormulaR1C1 = "Val " + Format$(i) + ": " + Format$(setpoint(i)) + "%, "
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
            setpointColumn = setpointColumn + 4
        Next i
        
        
        ' Save the results.
        .ActiveWorkbook.SaveAs FileName:=dataFilename
        frmMain.Caption = VERSION + "      " + dataFilename
                        
        'Initialize TASK cell with starting task = 0
        excel_app.Cells(4, 3) = 0
    End With


Quit:
End Sub

Public Sub StartExcelAndOpenFile(spreadsheetName As String)
Dim errorCheck As Integer

    DoEvents

    If (spreadsheetName <> "") Then
        result = True
        
        On Error Resume Next
        Set excel_app = GetObject(, "Excel.Application")
        errorCheck = Err.Number
        '429 means Excel is NOT running. 0 means it is already running.
        'If it is already running, then we don't need to start it up.
        
        ' Create the Excel application, if Excel isn't already running:
        If (errorCheck <> 0) Then Set excel_app = CreateObject("Excel.Application")
        
        ' Make Excel visible:
        excel_app.Visible = True
        ' Open the Excel spreadsheet.
        excel_app.Workbooks.Open FileName:=spreadsheetName

        ' Check for later versions.
        If Val(excel_app.Application.VERSION) >= 8 Then
            Set excel_sheet = excel_app.ActiveSheet
        Else
            Set excel_sheet = excel_app
        End If
    End If
End Sub

Function ExcelCheck() As Boolean
Dim SheetOpen As Boolean

Dim obExcelCheck As Object
Dim Testsheet As Object
Dim i, j, EndPath As Integer
Dim NumberOfOpenWorkBooks As Integer
Dim XLAppFx As Excel.Application
Dim TestName, NameOnly As String
Dim length As Integer
    
    length = Len(dataFilename)
    For i = 1 To length
        j = InStr(i, dataFilename, "\")
        If (j > 0) Then EndPath = j
    Next i
    NameOnly = Mid(dataFilename, EndPath + 1, length - EndPath)
    
    ExcelCheck = False
    
    On Error Resume Next
    'Is Excel Running?
    Set XLAppFx = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then Exit Function
    
    With excel_app
    
    'NumberOfOpenWorkBooks = XLAppFx.Workbooks.Count
        NumberOfOpenWorkBooks = .Workbooks.Count
    'Loop through all open workbooks in such instance
    For i = NumberOfOpenWorkBooks To 1 Step -1
        'TestName = XLAppFx.Workbooks(i).Name
            TestName = .Workbooks(i).Name
        If TestName = NameOnly Then Exit For
    Next i
    
    End With
    

    If (i <> 0) Then
        ExcelCheck = True
    Else
        excel_app.Quit
        Set excel_sheet = Nothing
        Set excel_app = Nothing
    End If
End Function


Public Sub addPassFailText()
Dim row As Integer
Dim i As Integer
Dim totalSensors As Integer
Dim totalTwoPercentSensors As Integer
Dim totalThreePercentSensors As Integer
Dim totalFivePercentSensors As Integer
Dim totalFailSensors As Integer
Dim dblTotalSensors As Double
Dim dblTwoPercentSensors As Double
Dim dblThreePercentSensors As Double
Dim dblFivePercentSensors As Double
Dim dblFailSensors As Double
Dim passPercent As String
Dim length As Integer

    
    totalSensors = 0
    totalTwoPercentSensors = 0
    totalThreePercentSensors = 0
    totalFivePercentSensors = 0
    totalFailSensors = 0
       
    For i = 1 To Cal.lastSensor
        If Cal.SensorTip(i).Used = True Then totalSensors = totalSensors + 1
    Next i
    
    totalSensors = Cal.lastSensor

    If (ExcelCheck() = True) Then
        With excel_app
        
            row = OFFSET
            For i = 1 To Cal.lastSensor
                row = i + OFFSET
                If Cal.SensorTip(i).Used = True Then
                    statusString = excel_app.Cells(row, STATUS_COLUMN).value
                    If (InStr(1, statusString, "PASS 2%") > 0) Then
                        totalTwoPercentSensors = totalTwoPercentSensors + 1
                    ElseIf (InStr(1, statusString, "PASS 3%") > 0) Then
                        totalThreePercentSensors = totalThreePercentSensors + 1
                    ElseIf (InStr(1, statusString, "PASS 5%") > 0) Then
                        totalFivePercentSensors = totalFivePercentSensors + 1
                    Else
                        totalFailSensors = totalFailSensors + 1
                    End If
                End If
            Next i
            
            If (totalSensors > 0) Then
                dblTotalSensors = CDbl(totalSensors)
                dblTwoPercentSensors = CDbl(totalTwoPercentSensors * 100) / dblTotalSensors
                dblThreePercentSensors = CDbl(totalThreePercentSensors * 100) / dblTotalSensors
                dblFivePercentSensors = CDbl(totalFivePercentSensors * 100) / dblTotalSensors
                dblFailSensors = CDbl(totalFailSensors * 100) / dblTotalSensors
            End If
        
            row = OFFSET + Cal.lastSensor + 2
            .Cells(row, STATUS_COLUMN).Select
            .ActiveCell.FormulaR1C1 = "FINAL RESULTS"
            .Cells(row, STATUS_COLUMN).Interior.ColorIndex = NO_COLOR
            With .Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .ColorIndex = BLUE
            End With
            
            row = row + 1
            .Cells(row, STATUS_COLUMN).Select
            .ActiveCell.FormulaR1C1 = "PASS 2%"
            .Cells(row, STATUS_COLUMN).Interior.ColorIndex = TWO_PERCENT_COLOR
            With .Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .ColorIndex = BLACK
            End With
            excel_app.Cells(row, STATUS_COLUMN - 1).value = totalTwoPercentSensors
            
            row = row + 1
            .Cells(row, STATUS_COLUMN).Select
            .ActiveCell.FormulaR1C1 = "PASS 3%"
            .Cells(row, STATUS_COLUMN).Interior.ColorIndex = THREE_PERCENT_COLOR
            With .Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .ColorIndex = BLACK
            End With
            excel_app.Cells(row, STATUS_COLUMN - 1).value = totalThreePercentSensors
            
            row = row + 1
            .Cells(row, STATUS_COLUMN).Select
            .ActiveCell.FormulaR1C1 = "PASS 5%"
            .Cells(row, STATUS_COLUMN).Interior.ColorIndex = FIVE_PERCENT_COLOR
            With .Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .ColorIndex = BLACK
            End With
            excel_app.Cells(row, STATUS_COLUMN - 1).value = totalFivePercentSensors
            
            row = row + 1
            .Cells(row, STATUS_COLUMN).Select
            .ActiveCell.FormulaR1C1 = "FAIL"
            .Cells(row, STATUS_COLUMN).Interior.ColorIndex = FAIL_COLOR
            With .Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .ColorIndex = BLACK
            End With
            excel_app.Cells(row, STATUS_COLUMN - 1).value = totalFailSensors
    
            row = row + 1
            .Cells(row, STATUS_COLUMN).Select
            .ActiveCell.FormulaR1C1 = "TOTAL"
            .Cells(row, STATUS_COLUMN).Interior.ColorIndex = NO_COLOR
            With .Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .ColorIndex = BLUE
            End With
            excel_app.Cells(row, STATUS_COLUMN - 1).value = totalSensors
                
            row = row + 2
            .Cells(row, STATUS_COLUMN).Select
            .ActiveCell.FormulaR1C1 = "YIELD"
            .Cells(row, STATUS_COLUMN).Interior.ColorIndex = NO_COLOR
            With .Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .ColorIndex = BLUE
            End With
            
            
            row = row + 1
            .Cells(row, STATUS_COLUMN).Select
            .ActiveCell.FormulaR1C1 = "PASS 2%"
            .Cells(row, STATUS_COLUMN).Interior.ColorIndex = TWO_PERCENT_COLOR
            With .Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .ColorIndex = BLACK
            End With
            passPercent = Format$(dblTwoPercentSensors, "##.##")
            If (Len(passPercent) = 1) Then passPercent = "0.0"
            excel_app.Cells(row, STATUS_COLUMN - 1).value = passPercent + "%"
            
            row = row + 1
            .Cells(row, STATUS_COLUMN).Select
            .ActiveCell.FormulaR1C1 = "PASS 3%"
            .Cells(row, STATUS_COLUMN).Interior.ColorIndex = THREE_PERCENT_COLOR
            With .Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .ColorIndex = BLACK
            End With
            passPercent = Format$(dblThreePercentSensors, "##.##")
            If (Len(passPercent) = 1) Then passPercent = "0.0"
            excel_app.Cells(row, STATUS_COLUMN - 1).value = passPercent + "%"
                        
            row = row + 1
            .Cells(row, STATUS_COLUMN).Select
            .ActiveCell.FormulaR1C1 = "PASS 5%"
            .Cells(row, STATUS_COLUMN).Interior.ColorIndex = FIVE_PERCENT_COLOR
            With .Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .ColorIndex = BLACK
            End With
            passPercent = Format$(dblFivePercentSensors, "##.##")
            If (Len(passPercent) = 1) Then passPercent = "0.0"
            excel_app.Cells(row, STATUS_COLUMN - 1).value = passPercent + "%"
            
            
            row = row + 1
            .Cells(row, STATUS_COLUMN).Select
            .ActiveCell.FormulaR1C1 = "FAIL"
            .Cells(row, STATUS_COLUMN).Interior.ColorIndex = FAIL_COLOR
            With .Selection.Font
                .Name = "Arial"
                .FontStyle = "Bold"
                .Size = 10
                .ColorIndex = BLACK
            End With
            passPercent = Format$(dblFailSensors, "##.##")
            If (Len(passPercent) = 1) Then passPercent = "0.0"
            excel_app.Cells(row, STATUS_COLUMN - 1).value = passPercent + "%"
            
        End With
    End If
End Sub


