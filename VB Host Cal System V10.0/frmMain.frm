VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Caption         =   "Setra Humidity Calibration"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   12435
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10710
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar barProgress 
      Height          =   375
      Left            =   7080
      TabIndex        =   34
      Top             =   9720
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CheckBox chkCompressorShutdown 
      Caption         =   "Shut off compressor when done."
      Height          =   495
      Left            =   3840
      TabIndex        =   32
      Top             =   1920
      Width           =   4575
   End
   Begin VB.TextBox txtTemperatureSetpoint 
      Enabled         =   0   'False
      Height          =   345
      Left            =   10200
      TabIndex        =   31
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox txtTemperatureWaitTime 
      Enabled         =   0   'False
      Height          =   345
      Left            =   11040
      TabIndex        =   30
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton cmdCompressor 
      Caption         =   "Turn On Compressor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   27
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CheckBox chkRack 
      Caption         =   "Rack D"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   4
      Left            =   240
      TabIndex        =   26
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CheckBox chkRack 
      Caption         =   "Rack C"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   3
      Left            =   240
      TabIndex        =   25
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CheckBox chkRack 
      Caption         =   "Rack B"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   240
      TabIndex        =   24
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CheckBox chkRack 
      Caption         =   "Rack A"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   240
      TabIndex        =   23
      Top             =   2400
      Width           =   1215
   End
   Begin VB.VScrollBar scrTasks 
      Enabled         =   0   'False
      Height          =   615
      Left            =   3000
      Max             =   32
      TabIndex        =   22
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton cmdLEDtest 
      Caption         =   "LED TEST OFF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      TabIndex        =   21
      Top             =   9000
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtChamberRhWaitTime 
      Enabled         =   0   'False
      Height          =   345
      Left            =   6360
      TabIndex        =   19
      Top             =   1440
      Width           =   735
   End
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   1080
      Top             =   9720
   End
   Begin VB.CommandButton cmdSetLEDS 
      Caption         =   "Calibration Complete. Click here to remove Failed units."
      Height          =   1695
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.PictureBox picCalComplete 
      Height          =   4215
      Left            =   720
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   4155
      ScaleWidth      =   11475
      TabIndex        =   17
      Top             =   5280
      Visible         =   0   'False
      Width           =   11535
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   600
      Top             =   9720
   End
   Begin VB.CommandButton cmdResume 
      Caption         =   "RESUME"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   16
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "SELECT"
      Enabled         =   0   'False
      Height          =   615
      Left            =   1680
      TabIndex        =   15
      Top             =   4200
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid grdSpreadsheet 
      Height          =   4215
      Left            =   840
      TabIndex        =   14
      Top             =   5280
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   136
      Cols            =   32
      FixedRows       =   2
      FixedCols       =   0
      BackColor       =   8454143
   End
   Begin MSCommLib.MSComm MSComm4 
      Left            =   6240
      Top             =   9600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm3 
      Left            =   5640
      Top             =   9600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   5040
      Top             =   9600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdHalt 
      Caption         =   "HALT"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "START"
      Enabled         =   0   'False
      Height          =   495
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.ListBox lstTasks 
      BackColor       =   &H00FFFF80&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   3600
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   2520
      Width           =   8295
   End
   Begin MSComDlg.CommonDialog cdbFile 
      Left            =   3720
      Top             =   9600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "xls"
      DialogTitle     =   "Spreadsheet Open/Save"
      FileName        =   "Current Cal.xls"
      Filter          =   "XLS (*.xls)|*.xls"
      InitDir         =   "c:\Cal Data"
   End
   Begin VB.OptionButton optOffMode 
      Caption         =   "Power OFF"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9240
      TabIndex        =   5
      Top             =   6480
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtSend 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   5280
      Width           =   6135
   End
   Begin VB.CommandButton cmdComTest 
      Caption         =   "Comm Port Test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   9000
      Width           =   1695
   End
   Begin VB.TextBox txtReceive 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   2880
      TabIndex        =   0
      Top             =   6000
      Width           =   6135
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4440
      Top             =   9600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblStatusBar 
      Caption         =   "Status"
      Height          =   495
      Left            =   0
      TabIndex        =   33
      Top             =   10200
      Width           =   12255
   End
   Begin VB.Label lblChamberTempWaitTime 
      Caption         =   "Temp Wait Time (min):"
      Height          =   375
      Left            =   7920
      TabIndex        =   29
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label lblTempSetpoint 
      Caption         =   "Temp Setpoint:"
      Height          =   375
      Left            =   8040
      TabIndex        =   28
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label lblChamberRhWaitTime 
      Caption         =   "RH wait time (min):"
      Height          =   375
      Left            =   3720
      TabIndex        =   20
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label lbl_RH_setpoint 
      Caption         =   "Chamber RH Setpoint:"
      Height          =   495
      Left            =   3720
      TabIndex        =   13
      Top             =   720
      Width           =   3375
   End
   Begin VB.Label lblMeasuredRH 
      Caption         =   "UUT RH:"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   840
      Width           =   3255
   End
   Begin VB.Label lblChamberTemp 
      Caption         =   "Temperature C:"
      Height          =   495
      Left            =   8040
      TabIndex        =   11
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label lblChamberRH 
      Caption         =   "Chamber RH:"
      Height          =   495
      Left            =   3720
      TabIndex        =   10
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label lblVoltmeter 
      Caption         =   "UUT Voltage:"
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   240
      Width           =   3015
   End
   Begin VB.Label lblSend 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   5400
      Width           =   1215
   End
   Begin VB.Label lblReceive 
      Caption         =   "Receive"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   6120
      Width           =   1215
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu subNewDataFile 
         Caption         =   "&New"
      End
      Begin VB.Menu subOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu subSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu subSaveAs 
         Caption         =   "&Save As"
      End
      Begin VB.Menu mnuSetFileFolder 
         Caption         =   "&Set File Folder"
      End
      Begin VB.Menu subExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuNormalMode 
         Caption         =   "&Normal Mode"
      End
      Begin VB.Menu mnuDiagnostics 
         Caption         =   "&Diagnostics"
      End
      Begin VB.Menu mnuTurnOnLEDS 
         Caption         =   "&Turn on LEDS"
      End
   End
   Begin VB.Menu mnuThunderMode 
      Caption         =   "&Thunder Mode"
      Enabled         =   0   'False
      Begin VB.Menu mnuRhPc 
         Caption         =   "&Rh@Pc"
      End
      Begin VB.Menu mnuRhPcTc 
         Caption         =   "&Rh@PcTc"
      End
   End
   Begin VB.Menu mnuComPort 
      Caption         =   "&Com Port"
      Enabled         =   0   'False
      Begin VB.Menu subComPort 
         Caption         =   "&Setup Com Port"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' HumCal Humidity Calibration Program
' Revision history
' 1-20-2008 JBS:    Just basic diagnostics so far.
' 1-27-2008 JBS:    Single serial Port version using frmPortSettings
' 1-27-2008 JBS:    Multiple four port version
' 1-29-2008 JBS:    First debugged version up and running at Setra.
'                   Calibration pot-tweaking not implemented yet.
'                   Data collection and pot-setting from spreadsheet works,
'                   although system crashed when we ran past midnight.
' 1-30-2008 JBS:    Fixed midnight bug in setChamber() routine.
'                   Changed setpoint commands to use Napoleon's commands.
'                   Modified checkUnits() to use two decimal places recording values on spreadsheet.
'                   Implemented Step Mode on scrollers.
' 1-31-2008 JBS:    Still crashing at midnight. Set difftime to 0 should fix that issue.
' 2-1-2008 JBS:     Fourth setpoint was 10%, changed to 50%
' 2-4-2008 JBS:     Changed port routines in frmPortSettings to be more reliable.
' 2-4-2008 JBS:     Debugged first version of calibration loop. Added fine cal loop.
'                   Also added final PASS/FAIL routine for checking accuracy results.
' 2-5-2008 JBS:     Added routines for blowing fuses on pots.
' 2-7-2008 JBS:     Cleaned up routines for hiding diagnostics.
' 2-7-2008 JBS:     New pot calibration routine uses one adjustment loop.
' 2-8-2008 JBS:     Tweaks for AdjustPot calibration: final adjustment gets written to cal
'                   no matter what. Not sure whether this is a good idea.
'                   Increased number of voltage readings to 4 per measurement.
'                   Changed acceptable pot value range to 2 to 254.
'                   Temporarily add final setpoint at 20%.
' 2-12-2008 JBS:    Modified setChamber() and SendReceiveHumilabControlCOM()
'                   to send and verfiy setpoints.
'                   Modified all routines for writing to and reading from pots to
'                   handle programming flags more intelligently.
' 2-13-2008 JBS:    Added routine for CheckFuses(). Fixed com timeout in measure_UUT_RH().
' 2-14-2008 JBS:    Modified pot initialize routine to check off Cal.SensorTip(i).Used array.
' 2-??-2008 JBS:    Added data grid.
' 2-22-2008 JBS:    TaskIndex is now a global that always determines the current task.
'                   Changed RUN button to START.
' 2-25-2008 JBS:    Cleaned up power on/off/program routines.
'                   Added SELECT TASKS button and scroll bar.
'                   Program non-failing units.
'                   GREEN, RED, WHITE status boxes.
'                   Pot initializer bug - don't see a problem - added short delay before starting.
' 2-26-2008 JBS:    Fixed bug in code that was setting colors incorrectly for pass fail units.
'                   Record "NONE" in STATUS box.
'                   Record "PROG" for already programmed units during INITIALIZE step.
'                   "NONE" units should be blank instead of showing "128" and "OK"
'                   So only "OK" units get programmed then.
' 2-28-2008 JBS:    Fixed bug in check program pots routine.
'                   Added routines for equipment check. All four serial COM ports are checked.
'                   Added progress bar to measurement loop.
'                   Added Timer1 interrupt routine for < and > arrows
'                   so that the GREEN and RED windows work properly.
' 3-4-2008 JBS:     Added Comment column for Already Programmed, etc.
'                   Modified initializePots(), ProgramAllPots()
'                   so that they all check each pot individually to see if it is already programmed.
'                   Modified CheckFuses() to check Pot1 and Pot2 columns for values.
'                   If they are blank, fuses aren't checked.
' 3-5-2008 JBS:     Fixed bug in CalibratePots() which wasn't checking pots to see whether thery were calibrated.
'                   VERSION 1 COMPLETED 3-5-2008
'
' 4-4-2008 JBS:     Switched over to Fluke Meter - polling meter using second Timer
' 4-11-2008 JBS:    LED routines: Added setCalCompleteScreen() and createLEDcommandString() routines.
'
' 4-15-2008 JBS:    Switched over to Thunder. Worked out communications with 3 COM ports.
'                   Debugged COM check routines with Thunder. Send data to Thunder using Timer 3.
'
'                   TODO: look into checking for Excel App already running:
'                   On Error Resume Next
'                   'Is Excel Running?
'                   Set XLAppFx = GetObject(, "Excel.Application")
'                   If Err.Number <> 0 Then Exit Function
'
' 4-17-2008 JBS:    Fixed 2% LED bug in setCalCompleteScreen().
'
' 4-18-2008 JBS:    Modified StartExcelAndOpenFile() and createNewSpreadsheet() to check whether
'                   Excel is already running. If it is, then the user is prompted to save and close it.
'                   Also, changed order of PASS/FAIL sequence so that FAIL comes first.
'                   Made OR in initializePots(): If (potState1 = I2C_ERROR) Or (potState2 = I2C_ERROR) Then
' 4-22-2008 JBS:    Saw runtime error at very end of cal. Wondering whether it could be TaskIndex incrementing past MaxTask.
'                   Changed code in Run to see if that fixes problem. Now TaskIndex does not increment above MaxTask.
'
' 4-22-2008 JBS:    Added "Please Wait" in Form_Load while COM ports start up.
'                   Combined ComPortCheck() with CheckCommunication() in single task #
'                   Moved all routines up one, made turnOnLEDS() last, separate task.
'
'                   For createLEDcommandString():
'                   Added special case: NO sensor is treated as a failure here since LED is turned on for empty slots as well
'
'                   Simplified use of "FAIL" indication to make code more robust.
'                   I2C failure is now noted in COMMENT column.
'                   Status column only shows the word "FAIL"
' 4-23-2008 JBS:    Changed chamberRhWaitTime to 60 minutes.
' 4-30-2008 JBS:    Changed setpoint command from "R1=" to "R2="
' 5-5-2008 JBS:     Fixed overrun bug in Timer3 - Added statement: If (Timer3Counter > 100) Then Timer3Counter = 0
'                   Made chamberRhWaitTime global which gets set at startup and can be changed to any value by user.
'
' 5-16-2008 JBS:    In function finalCheckStatus(), modified code so that previously
'                   calibrated spreadsheets would not fail, Instead of scanning for "OK"
'                   we scan for "FAIL": failCheck = InStr(1, statusString, "FAIL")
' 5-19-2008 JBS:
'                   Completed addPassFailText() using existing spreadsheet approach.
'                   Each pass percentage now has a color, and final statistics are added.
'
'                   Upgraded Timer1_Timer() to display green, yellow, orange backgrounds for 2%,3%,5% units.
'
'       TODO: Add   1)  Temperature check to spreadsheet
'                   2)  Check that chamber is making it to setpoints.
'                   3)  LED test - turn 'em all on.
'                   4)  Rename files.
'
' 5-21-2008 JBS:    Took care of 1) and 3) above - added temp measurement and LED test.
' 5-28-2008 JBS:    Created RenameAndSave routine.
' 5-29-2008 JBS:    1) Eliminated Cal Setup form.
'                   2) Added password: "gumby" upper or lower case.
' 5-30-2008 JBS:    Restored scrollbar for changing starting task.
'                   Set scrTask.value to TaskIndex any time TaskIndex gets changed.
'                   Eliminated setting TaskIndex to 0 when new spreadsheet is created.
'                   Changed password to "Humidity"
' 6-5-2008 JBS:     Eliminated old Load and Store Config routines.
'                   Created Cal record and new routines for loading/storing.
'                   Eliminated com port activation from interfaceComTest().
'                   Eliminated ComPortCheck() by opening ports in CheckCommunication()
'                   Moved Call optNormalMode_Click into more prominent place in Execute after CheckCommunication()
'                   Added frmSetup
' 6-11-2008 JBS:    Completed identifySensorRacksAndCopyToSpreadsheet() to poll the racks and see which ones are present,
'                   and write designator strings to spreadsheet.
' 6-12-2008 JBS:    Modified Timer1_Timer() to read sensors even when Used is false.
'                   Updated the following routines to work with four sensor tip racks:
'                   createLEDcommandString(), turnOnLEDS(), setCalCompleteScreen(),
'                   readPot(), writePot(), SensorNumberChange(), checkUnits(),
'                   Quit here: TODO: Look at combining routines for taking data here and in Calibrate
' 6-16-2008 JBS:    TODO: remove comments in Set Chamber
' 6-20-2008 JBS:    Commented code.
'                   Modified initializePots() so that it does not use global userSensorNumber
'                   Changed name of GLOBAL sensorNumber to userSensorNumber in the following routines
'                   SensorNumberChange(),
'                   Modified initializePots() to read pot values from previously programmed units
'                   and store the values in POT1 and POT2 columns in spreadsheet.
'                   Renamed SendReceiveInterfaceBoard() instead of Main Board.
' 6-24-2008 JBS:    Renamed to version 7.0.
' 6-25-2008 JBS:    Set allowable calibration error to 0.1% RH to speed up calibration loop.
'                   Added message box in measure_UUT_RH()
'                   for voltmeter timeout to pause the cal process
'                   if the voltmeter communication times out after VOLTMETER_TIMEOUT = 30 seconds.
'                   In checkChamberControlCommunication(), reduced time for sending Thunder RUN command.
'                   Modified measure_UUT_RH(): added extra measurement and rejected first measurement
'                   to insure integrity of data.
'                   Made error constants global and changed FIVE_PERCENT_ERROR to 4.5
'
' 6-29-2008 JBS:    Cleaned up getMaxSensors(), chkRack_Click(), getSensorString()
'                   read and write pot routines now work without open spreadsheet,
'                   by using getSensorString() to generate sensor designator string.
'                   Got rid of cal record for USED data and sensor designators.
'                   chkRack() is now loaded and stored with Get and SaveSettings().
'                   Clicking on chkRack now recalculates maxSensors each time you change racks.
'                   TODO: Add STOP command for Thunder.
' 6-30-2008 JBS:    Eliminated from Timer 1: sensorUsed = excel_app.Cells(row, SENSOR_USED_COLUMN).value
'                   so that diagnostics work with old spreadsheets.
'                   Created chkCalAllSensors feature to allow evaluation of faster calibration algorithm.
'
' 7-3-2008 JBS:     Corrected bug in Function SendReceiveInterfaceBoard() - wasn't resending command string.
' 7-7-2008 JBS:     Added RH offset to AdjustPot() cal routine: ReferenceRH = ChamberRH + SpanOffset
'                   If file is opened, then reset message on cal complete command button.
' 7-8-2008 JBS:     Fixed Span Offset bug in AdjustPot(). Add offset only to reference RH for SPAN pot calibration.
'                   Modified createLEDcommandString() and turnOnLEDS() and setCalCompleteScreen()
'                   so that one rack of LEDs is turned on at a time.
' 7-9-2008 JBS:     Fixed bugs in routines that set LEDS and also test routines for LEDS.
' 7-10-2008 JBS:    Changed COM timeouts in SendReceiveInterfaceBoard()
'                   Modified firstCheckFailStatus() and finalCheckStatus() for different pass/fail linits at 90%.
'                   TODO: Make sure that board power is ON, and LEDS are OFF when calibration starts.
' 7-11-2008 JBS:    Fixed bugs in ProgramAllPots() and BlowFuse(). Pots weren't getting blown.
'                   Fixed PROG POT command in BlowFuse().
'                   Added dummy read to SendReceiveInterfaceBoard() to clear MSCOMM1 buffer so it can recover from power outage.
'                   Improved COM tests. Fixed bugs in Thunder COM test.
'                   initializeFlukeMeter() waits set to 1 second each
' 7-15-2008 JBS:    Modfied Thunder communication failure check in Timer3:
'                   so that THUNDER_STOP_COMMAND doesn't cause serial communications to fail.
' 7-17-2008 JBS:    Modified Timer3() so that Thunder commands are sent only when
'                   RunFlag is True, in other words, during calibration.
'                   Modified finalCheckStatus() to allow different PASS/FAIL limits for 50% and 90%.
'                   Started Set Path Feature.
' 7-21-2008 JBS:    Fixed premature timeouts in SendReceiveInterfaceBoard(). Resends command every 10 seconds.
'                   Three resends allowed before stopping process.
'                   Correction: wasn't really doing resends. Fixed that.
' 8-5-2008 JBS:     Called this the definative version of V7.0 - without the Network File Path Save Feature.
'                   That form was removed from this version.
'                   VERSION 8.0:
' 8-8-2008 JBS:     Added FrmFileFolder to set local and network drives.
'                   Modified RenameAndSave( ) as well as GetSettings( ) and SaveSettings( )
'                   to store both local and network folder names: strNetworkFolder and strLocalFolder
'                   Modified frmError to display network error as well as chamber RH setpoint error.
' 8-11-2008 JBS:    Added waitForChamberToReachTemperature( ) and shutDownThunder( ) routines.
'                   Modified setup procedure to set chamber temperature.
'                   MSComm4_OnComm()
' 8-18-2008 JBS:    Added msgBox to RUN loop to pop up if temperature setpoint isn't 25 degrees C.
'                   Moved Set File Folder over to File menu.
' 8-19-2008 JBS:    Modified setChamber( ), waitForChamberToReachTemperature( ), and checkChamberControlCommunication()
'                   to read back and verify temperature setpoint.
' 9-9-2008 JBS:     Completed most of the work loading and saving pass/fail arrays and displaying
'                   them on frmSetPassLimits. Not yet actually using them for pass/fail check.
'                   Also need to deal with reading and saving new filename from window on that form.
' 9-15-2008 JBS:    Worked on Set Pass Limits file saving feature.
' 9-19-2008 JBS:    Current version saved with two windows in frmSetPasslimits.
' 9-23-2008 JBS:    Modified FinalFailCheck() and FirstFailCheck() to work with new routine
'                   checkStatus() which uses pass/fail limits stored in the arrPassLimits() array
'                   Added new variables, CustomVoutRange, VoutRange, which permits voltages other than
'                   2.5 volts to be used as the analog reference voltage.
'                   Created function calculateRh() to calculate voltage from VoutRange.
'                   Modified createNewSpreadsheet() and FinalFailCheck()
'                   to print Vout Range voltage on spreadsheet when calibration is complete.
' 9-25-2008 JBS:    Fixed verify setpoint bugs in checkChamberControlCommunication() and MSComm4_OnComm():
'                   "verifiedTemperatureSetpoint" and "verifiedRHsetpoint" are now the global variables
'                   used to verify RH and temperature setpoints for the Thunder.
'                   Fixed bugs: THUNDER_READ_RH_AND_T replaced THUNDER_READ_RH, THUNDER_READ_SETPOINTS replaced THUNDER_READ_SETPOINT.
' 9-26-2008 JBS:    Created initializeDefaultLimitsArray() in HumCalMod file, to temporarily create default limits without saving them.
'                   Modified LoadPassFailLimits()in HumCalMod to display msg box if limits file isn't found.
'                   Modified CheckCommunication() to power up rack boards and turn on compressor.
'                   Added cmdCompressor button and turnOnCompressor() and turnOffCompressor().
'                   Fixed bug in setpoint command. Allow 2 cases for Rh@Pc and Rh@PcTc: "R1=" and "R2="
'                   Fixed bug in msg box for Cancel command in frmSetPassLimits.
' 10-8-2008 JBS:    Changed five percent color to BLUE in addPassFailText()
'                   Changed Timer1_Timer() background color to blue also.
'                   Added different background colors to setCalCompleteScreen()
' 10-27-2008 JBS:   Modified checkChamberControlCommunication() so it verifies the temperature setpoint only.
'
' 5-10-2017 JBS:    Version 10.0 modifications for Windows 10:
'                   For Version 10.0: In SetUpTaskBox() don't run calibration tasks
'                   Deleted StatusBarSimple Text and substituted lblStatusBar.Caption
'                   Deleted ProgressBar.
' 5-11-2017 JBS:    Put ProgressBar back in.
'                   Created initializeSensorsUsed() for analog sensors
'                   In SetUpTaskBox(), eliminated tasks #2-7 and #9-12, changed setpoint order to: %10, %50, %90
'                   In Execute(), eliminated tasks #2-7 and #9-12, changed setpoint order to: %10, %50, %90
' 8-16-17 JBS:      Disabled MSComm3 for HumiLab reference, deleted unused diagnotics
' 9-27-17 JBS:      Recompiled at home. Stored in GitHub.

Const VERSION = "Setra Humidity Cal Program - V10.0 "
Const DEFAULT_SHEETNAME = "Current Cal.xls"
Const DEFAULT_HEIGHT = 2610
Const EXTENDED_HEIGHT = 6000

Const OFFSET_VOLTAGE = 0.5
Const SPAN_VOLTAGE = 4

'These constants determine the allowable pass/fail limits. Units are in percent RH:
Const TWO_PERCENT_ERROR = 1.5
Const THREE_PERCENT_ERROR = 2.5
Const FIVE_PERCENT_ERROR = 4.5

'These constants set the background colors for the STATUS column
'when the final PASS/FAIL data is written:
Const TWO_PERCENT_COLOR = 4     'This is GREEN
Const THREE_PERCENT_COLOR = 6   'This is YELLOW
Const FIVE_PERCENT_COLOR = 8    'This is BLUE
Const FAIL_COLOR = 3            'This is RED
Const NO_COLOR = 2              'This is WHITE - the normal background color for a spreadsheet cell
Const BLACK = 1
Const BLUE = 5

Const OFF_LED = 0
Const FAIL_LED = 1
Const TWO_PERCENT_LED = 2
Const THREE_PERCENT_LED = 3
Const FIVE_PERCENT_LED = 4
Const LED_TEST = 5

Const I2C_ERROR = 0
Const READY_TO_PROGRAM = 1
Const FUSE_BLOWN = 2
Const FUSE_BAD = 3
Const INVALID_POT_DATA = 4
Const COM_ERROR = 5
Const POWER_DOWN_TIME = 30 'This is number of seconds units are powered down after programming

Const INCREMENT = 1
Const DECREMENT = 0
Const RETRIES = 4 'Number of times serial communication gets retried.
Const BALANCE_POT = 1
Const SPAN_POT = 2

Const RACK_ROW = 5
Const VOUT_RANGE_ROW = 6
Const SETPOINT_ROW = 8

Const TITLE_ROW = SETPOINT_ROW + 1
Const ROW_OFFSET = TITLE_ROW
Const OFFSET = ROW_OFFSET

Const BLANK = "   "
Const UUT_WIDTH = 5
Const REF_WIDTH = 5
Const ERR_WIDTH = 5
Const COMMENT_WIDTH = 10

Const SENSOR_COLUMN = 1
Const COMMENT_COLUMN = 2
Const STATUS_COLUMN = 3

Const BLK1_COLUMN = 4 'Was 13 $$$$
Const REF1_COLUMN = 5
Const UUT1_COLUMN = 6
Const ERR1_COLUMN = 7

Const BLK2_COLUMN = 8
Const REF2_COLUMN = 9
Const UUT2_COLUMN = 10
Const ERR2_COLUMN = 11

Const BLK3_COLUMN = 12
Const REF3_COLUMN = 13
Const UUT3_COLUMN = 14
Const ERR3_COLUMN = 15

Const BLK4_COLUMN = 16
Const SENSOR_USED_COLUMN = 17
Const CAL_LOOPS_COLUMN = 18
Const CAL_TEST_COLUMN = 19

Const THUNDER_READ_RH_AND_T = "?"
Const THUNDER_READ_SETPOINTS = "?SP"
Const THUNDER_RUN_COMMAND = "RUN"
Const THUNDER_STOP_COMMAND = "STOP"

Const DEFAULT_VOUT_RANGE = 2.5

Dim LEDtestIndex As Integer
Dim thunderCommandString As String
Dim verifiedRHsetpoint As Integer
Dim verifiedTemperatureSetpoint As Integer
Dim Timer3Counter As Integer
Dim Timer3Timeout As Integer
Dim CalCompleteIndex As Integer
Dim rackLedIndex As Integer

Dim compressorFlag As Boolean
Dim UARTflag As Boolean
Dim VoltUARTflag As Boolean
Dim UARTbuffer As String
Dim VoltUARTbuffer As String
Dim voltage As Double
Dim Intext As String
Dim ChamberRH As Double
Dim ChamberTempC As Double
Dim ChamberRefUARTbuffer As String
Dim ChamberRefUARTflag  As Boolean
Dim ChamberControlUARTbuffer As String
Dim ChamberControlUARTflag  As Boolean
Dim userPot1Value As Integer
Dim userPot2Value As Integer
Dim userSensorNumber As Integer

Dim MaxTask As Integer
Dim RunFlag As Boolean
Dim excel_app As Object
Dim excel_sheet As Object
Dim dataFilename As String
Dim TaskIndex As Integer
Dim SelectFlag As Boolean
Dim lngStartTime As Long


'This routine initializes and opens all four COM ports and returns TRUE
'if they all open properly and FALSE if one or more doesn't open.
'If the ports are properly opened,
Public Function openComPorts _
    (BitRate As Long, _
    NewPortNumber As Integer, VoltmeterPortNumber As Integer, ChamberControlPortNumber As Integer) _
    As Boolean
'Public Function openComPorts _
'    (BitRate As Long, _
'NewPortNumber As Integer, VoltmeterPortNumber As Integer, ChamberRefPortNumber As Integer, ChamberControlPortNumber As Integer) _
'    As Boolean
Dim result As Boolean
Dim intButton As Integer

result = True 'Be optimistic - assume all ports will open properly

'Initializes the selected Com Port.
'All settings except BitRate are set explicitly in this routine.
'Some properties show alternate settings commented out.
Dim ComSettings

If MSComm1.PortOpen = True Then
    MSComm1.PortOpen = False
End If

If MSComm2.PortOpen = True Then
    MSComm2.PortOpen = False
End If

If MSComm3.PortOpen = True Then
    MSComm3.PortOpen = False
End If

If MSComm4.PortOpen = True Then
    MSComm4.PortOpen = False
End If


MSComm1.CommPort = NewPortNumber
MSComm2.CommPort = VoltmeterPortNumber
If (ChamberType = HUMILAB) Then MSComm3.CommPort = ChamberRefPortNumber
MSComm4.CommPort = ChamberControlPortNumber

PortNumber = NewPortNumber
'Use BitRate, no parity, 8 data, and 1 stop bit:
ComSettings = CStr(BitRate) & ",N,8,1"
MSComm1.Settings = ComSettings
MSComm2.Settings = ComSettings
If (ChamberType = HUMILAB) Then MSComm3.Settings = ComSettings
MSComm4.Settings = ComSettings

'Properties relating to receiving:
'Read entire buffer on Input:
MSComm1.InputLen = 0
MSComm2.InputLen = 0
If (ChamberType = HUMILAB) Then MSComm3.InputLen = 0
MSComm4.InputLen = 0
'Read one byte at a time on Input:
' MSComm1.InputLen = 1

MSComm1.InBufferSize = 1024
MSComm2.InBufferSize = 1024
If (ChamberType = HUMILAB) Then MSComm3.InBufferSize = 1024
MSComm4.InBufferSize = 1024

'Generate no OnComm event on received data:
'MSComm1.RThreshold = 0
'Generate an OnComm event on each character received:
MSComm1.RThreshold = 1
MSComm2.RThreshold = 1
If (ChamberType = HUMILAB) Then MSComm3.RThreshold = 1
MSComm4.RThreshold = 1
'The Input property stores binary data:
'MSComm1.InputMode = comInputModeBinary
'The Input property stores data as text:
MSComm1.InputMode = comInputModeText
MSComm2.InputMode = comInputModeText
If (ChamberType = HUMILAB) Then MSComm3.InputMode = comInputModeText
MSComm4.InputMode = comInputModeText
'Disable parity replacement"
'MSComm1.ParityReplace = ""

'Properties related to transmitting:

MSComm1.OutBufferSize = 16
MSComm2.OutBufferSize = 16
If (ChamberType = HUMILAB) Then MSComm3.OutBufferSize = 16
MSComm4.OutBufferSize = 16
'Generate no transmit OnComm event:
MSComm1.SThreshold = 0
MSComm2.SThreshold = 0
If (ChamberType = HUMILAB) Then MSComm3.SThreshold = 0
MSComm4.SThreshold = 0
'Generate an OnComm event when the transmit buffer
'has SThreshold bytes or fewer:
'MSComm1.SThreshold = 1

'Handshaking options:
MSComm1.Handshaking = comNone
MSComm2.Handshaking = comNone
If (ChamberType = HUMILAB) Then MSComm3.Handshaking = comNone
MSComm4.Handshaking = comNone
'MSComm1.Handshaking = comXOnXoff
'MSComm1.Handshaking = comRTS
'MSComm1.Handshaking = comRTSXOnXOff

'Try to open the ports:
On Error Resume Next
MSComm1.PortOpen = True
On Error Resume Next
MSComm2.PortOpen = True
On Error Resume Next
If (ChamberType = HUMILAB) Then MSComm3.PortOpen = True
On Error Resume Next
MSComm4.PortOpen = True


'Return success or failure
If (False = MSComm1.PortOpen) Then result = False

If (False = MSComm2.PortOpen) Then result = False

'If (ChamberType = HUMILAB) Then
'    If (False = MSComm3.PortOpen) Then result = False
'End If

If (False = MSComm4.PortOpen) Then result = False

If (result = False) Then
    intButton = MsgBox("Error opening COM ports." + vbCr + "Check COM ports", vbOKCancel)
    If (intButton = vbOK) Then
        frmPortSettings.Show
    End If
End If

'If (result = True) Then
    'Enable all command buttons now that send commands to sensors: $$$$
    'cmdWritePot1.Enabled = True
    'cmdWritePot2.Enabled = True
    'cmdReadPot1.Enabled = True
    'cmdReadPot2.Enabled = True
    'cmdDecrementSensorNumber.Enabled = True
    'cmdIncrementSensorNumber.Enabled = True
    'scrPot1.Enabled = True
    'scrPot2.Enabled = True
    'optNormalMode.Enabled = True
    'optProgramMode.Enabled = True
    'optOffMode.Enabled = True
'End If

openComPorts = result

End Function
Private Sub cmdComTest_Click()
    Call interfaceComTest
End Sub

'This routine RESUMES the calibration process
'at whatever task currently stored by TaskIndex
Private Sub cmdResume_Click()
Dim result As Integer
Dim message As String
Dim i As Integer
    lstTasks.Height = DEFAULT_HEIGHT
    If (mnuDiagnostics.Checked = False) Then grdSpreadsheet.Visible = True
    lstTasks.ListIndex = TaskIndex
    scrTasks.value = TaskIndex
    message = lstTasks.Text
    result = MsgBox("Do you want to resume calibration at this step?", vbOKCancel + vbQuestion, message)
    If (result = vbOK) Then
        RunFlag = True
        lstTasks.Enabled = False
        cmdStart.Enabled = False
        cmdResume.Enabled = False
        cmdHalt.Enabled = True
        frmMain.lblStatusBar.Caption = "Calibration in progress"
        Call Run
    End If
End Sub

'This routine enlarges the "lstsTasks" List Box to make it easier to select and deselect
'which tasks to run. This can only be done when the cal process is halted.
Private Sub cmdSelect_Click()
    Call HideDiagnostics
    cmdSelect.Enabled = True
    scrTasks.Enabled = True
    If (SelectFlag = True) Then
        SelectFlag = False
        lstTasks.Height = DEFAULT_HEIGHT
        grdSpreadsheet.Visible = True
    Else
        SelectFlag = True
        lstTasks.Height = EXTENDED_HEIGHT
        grdSpreadsheet.Visible = False
    End If
End Sub



'This routine starts the calibration process when the RUN button is pressed.
'The Task Index is reset to 0 to insure that the cal process starts at the
'top of the task list with the first checked task.
Private Sub cmdStart_Click()
Dim result As Integer
Dim message As String
Dim i As Integer
    lstTasks.Height = DEFAULT_HEIGHT
    If (mnuDiagnostics.Checked = False) Then grdSpreadsheet.Visible = True
    result = MsgBox("Are you sure you want to start from the beginning?", vbYesNo)
    DoEvents
    If (result = vbYes) Then
        TaskIndex = 0
        lstTasks.ListIndex = TaskIndex
        RunFlag = True
        lstTasks.Enabled = False
        cmdStart.Enabled = False
        cmdResume.Enabled = False
        cmdHalt.Enabled = True
        frmMain.lblStatusBar.Caption = "Calibration in progress"
        Call Run
    End If
End Sub

'This routine stops the calibration system when the HALT button is pressed
Private Sub cmdHalt_Click()
    Dim result As Integer
    
    result = MsgBox("Are you sure you want to halt the system?", vbYesNo)
    DoEvents
    If (result = vbYes) Then
        frmMain.lblStatusBar.Caption = "Stopping process, please wait."
        RunFlag = False
        lstTasks.Enabled = True
        cmdStart.Enabled = True
        cmdResume.Enabled = True
        cmdHalt.Enabled = False
    End If
End Sub


'This routine transmits command strings out the UART
'to the Humilab chamber control. A small delay had
'to be placed between characters because the Humilab
'can't handle fast input. After sending
'the command string, it waits for received input in return.
'If a return string is received, this function returns it.
'Otherwise it returns ""
Function SendReceiveHumilabControlCOM(commandString As String)
Dim i, outputLength As Integer
Dim ch, outputString As String

    SendReceiveHumilabControlCOM = ""
    outputString = commandString + vbCr
    outputLength = Len(outputString)
    
    ChamberControlUARTbuffer = "" 'Clear input buffer.
    ChamberControlUARTflag = False
        
    For i = 1 To outputLength
        ch = Mid(outputString, i, 1)
        Delay (100)
        MSComm4.Output = ch
    Next i
    
    i = 0
    'Allow 1 second for a response
    While (ChamberControlUARTflag = False) And (i < 10)
        Delay (100)
        i = i + 1
        DoEvents
    Wend
    
    SendReceiveHumilabControlCOM = ChamberControlUARTbuffer
End Function

'This is the main loop for the complete calibration and verification program.
'The program sits in this loop during the entire time that the system
'is calibrating and checking sensors. It is called when
'the RUN button is pressed.

'The main loop below executes the tasks listed
'in the lstTask list box in sequence until they are all completed.
'Only those tasks which are checked are executed. The task being
'executed is the task highlighted in BLUE. If the highlighted task
'isn't checked, then the highlight is moved ahead to the next checked task.
'Execution can begin anywhere, but unchecked tasks are always skipped,
'and the current task being executed is always highlighted.

'The user cannot uncheck or check tasks on the list while execution has
'already begun, and cannot change the current highlighted task.
'To change the current task, the HALT button can be pressed.
'This forces the Run loop to exit and stop execution.
'The first checked task can then be changed, and
'execution will begin with that new task
'when the RUN button is pressed again.

'Note that a spreadsheet must be open in order to run.
'The ExcelCheck() checks for an open spreadsheet.
'If none is open, then a new spreadsheet is created.

Private Sub Run()
    Dim strTask As String
    Dim test As Boolean
    
    chkRack(1).Enabled = False
    chkRack(2).Enabled = False
    chkRack(3).Enabled = False
    chkRack(4).Enabled = False
        
    If (temperatureSetpoint <> DEFAULT_TEMP) Then
        result = MsgBox("Temperature setpoint is: " + Format$(temperatureSetpoint) + " C. Is that correct?", vbYesNo + vbQuestion + vbDefaultButton2, "Confirm Temperature Setpoint")
        If (result = vbNo) Then
            RunFlag = False
            frmMain.lblStatusBar.Caption = "Process halted"
        End If
    End If
    
    If (RunFlag = True) Then
        'TODO: make this more robust!
        'Make sure that there is a spreadsheet to write to:
        test = ExcelCheck()
        If (dataFilename = "") Or (test = False) Then
            Call createNewSpreadsheet
        End If
    End If
    
    'RunFlag may is set to False either
    'if HALT button has been pressed
    'or calibration is completed:
    While (RunFlag = True)
        'Get highlighted task:
        lstTasks.ListIndex = TaskIndex
        scrTasks.value = TaskIndex
        DoEvents
        
        'If task is checked, execute it
        If (lstTasks.Selected(TaskIndex) = True) Then
            'Store new task number on spreadsheet:
            excel_app.Cells(4, 3).value = TaskIndex
            'Fetch task from list:
            strTask = GetTask()
            'Now run task:
            Call Execute(strTask)
            'Copy latest data to grid:
            Call copySpreadsheetToGrid
        End If
                
        'If we have reached the end of our task list, then we are done.
        'Otherwise move highlight to next task:
        If (RunFlag = True) Then
            If (TaskIndex < MaxTask) Then
                TaskIndex = TaskIndex + 1
                lstTasks.ListIndex = TaskIndex
                scrTasks.value = TaskIndex
            Else
                frmMain.lblStatusBar.Caption = "Calibration complete!"
                RunFlag = False
                TaskIndex = 0
                scrTasks.value = TaskIndex
            End If
        Else
            frmMain.lblStatusBar.Caption = "Process halted"
        End If
    Wend
    
    RunFlag = False
    lstTasks.Enabled = True
    cmdStart.Enabled = True
    cmdHalt.Enabled = False
End Sub


'This routine is called by Execute and it returns the numeric identifier
'at the beginning of each task in the task box. This makes it easy for
'Execute to determine what task to perform.
'
'For example if "[0] Com Port Check" is the current task,
'then GetTask returns "[0]". This permits easy and error-free indication
'of which task to perform while permitting the rest of the task string to be
'changed whenever desired. For example, this task could be retitled as
'"[0] Communications Port Check" so long as the "[0]" is left unmodified.
Function GetTask() As String
Dim TaskString As String
Dim FirstChPosition As Integer
Dim LastChPosition As Integer
    GetTask = ""
    TaskString = lstTasks.List(lstTasks.ListIndex)
    FirstChPosition = InStr(1, TaskString, "[")
    LastChPosition = InStr(1, TaskString, "]")
    If ((FirstChPosition > 0) And (LastChPosition > 0) And (FirstChPosition < LastChPosition)) Then
        GetTask = Mid(TaskString, FirstChPosition, LastChPosition - FirstChPosition + 1)
    End If
End Function


'This routine allows the user to select a different task from
'the task menu to start with.
Private Sub lstTasks_Click()
    lstTasks.ListIndex = TaskIndex
    scrTasks.value = TaskIndex
End Sub

Private Sub initializeFlukeMeter()
    frmMain.lblStatusBar.Caption = "Initializing Fluke meter..."
    MSComm2.Output = "VDC" + vbCr   ' $$$$
    Delay (1000)
    MSComm2.Output = "RANGE 3" + vbCr
    'MSComm2.Output = "CONF:VOLT:DC:RANG 10" + vbCr
    Delay (1000)
End Sub

'This routine checks serial com port communication
'with the voltmeter. It returns true if the voltmeter responds.
Function checkVoltmeterCommunication() As Boolean
Const VOLTMETER_TIMEOUT = 5 'This value corresponds to a 5 second timeout
Dim startTime As Variant
Dim elapsedTime As Variant
Dim ElapsedSeconds As Long
Dim previousSeconds As Long
    
    Call initializeFlukeMeter

    VoltUARTflag = False
    startTime = Timer()
    Do
        DoEvents
        elapsedTime = Timer() - startTime
        ElapsedSeconds = CLng(elapsedTime)
        If (ElapsedSeconds <> previousSeconds) Then
            previousSeconds = ElapsedSeconds
            frmMain.lblStatusBar.Caption = "Checking voltmeter communication: " + Format$(ElapsedSeconds) + " seconds."
        End If
        'Something weird just happened if elapsed time is negative!
        'Maybe midnight just reset Timer(). Clear time and start again:
        If (ElapsedSeconds < 0) Then
            startTime = Timer()
            elapsedTime = Timer() - startTime
            ElapsedSeconds = CLng(elapsedTime)
        End If
            DoEvents
    Loop While (VoltUARTflag = False) And (ElapsedSeconds < VOLTMETER_TIMEOUT)
    checkVoltmeterCommunication = VoltUARTflag
End Function
    

'This routine checks serial com port communication
'with the Humilab Reference output.
'It returns true if the Humilab responds.
Function checkHumilabRefCommunication() As Boolean
Const HUMILAB_TIMEOUT = 20 'This value corresponds to a 20 second timeout for the Humilab RH reference to send data
Dim startTime As Variant
Dim elapsedTime As Variant
Dim ElapsedSeconds As Long
Dim previousSeconds As Long
      
    lblChamberRH.Caption = ""
    startTime = Timer()
    Do
        DoEvents
        elapsedTime = Timer() - startTime
        ElapsedSeconds = CLng(elapsedTime)
        If (ElapsedSeconds <> previousSeconds) Then
            previousSeconds = ElapsedSeconds
            frmMain.lblStatusBar.Caption = "Checking Humilab reference communication: " + Format$(ElapsedSeconds) + " seconds."
        End If
        'Something weird just happened if elapsed time is negative!
        'Maybe midnight just reset Timer(). Clear time and start again:
        If (ElapsedSeconds < 0) Then
            startTime = Timer()
            elapsedTime = Timer() - startTime
            ElapsedSeconds = CLng(elapsedTime)
        End If
            DoEvents
    Loop While (lblChamberRH.Caption = "") And (ElapsedSeconds < HUMILAB_TIMEOUT)
    
    If (lblChamberRH = "") Then
        checkHumilabRefCommunication = False
    Else
        checkHumilabRefCommunication = True
    End If
End Function

Private Sub mnuNormalMode_Click()
    Call HideDiagnostics
    mnuDiagnostics.Checked = False
    mnuNormalMode.Checked = True
    mnuTurnOnLEDS.Checked = False
    
    picCalComplete.Visible = False
    cmdSetLEDS.Visible = False
    grdSpreadsheet.Visible = True
    lstTasks.Height = DEFAULT_HEIGHT
End Sub

Private Sub mnuRhPc_Click()
    mnuRhPcTc.Checked = False
    mnuRhPc.Checked = True
    thunderMode = RhPc
    lblChamberRH.Caption = "Chamber RH@Pc:"
End Sub

Private Sub mnuRhPcTc_Click()
    mnuRhPcTc.Checked = True
    mnuRhPc.Checked = False
    thunderMode = RhPcTc
    lblChamberRH.Caption = "Chamber RH@PcTc:"
End Sub

'This is called when user clicks on Options / Save settings
Private Sub mnuSaveSettings_Click()
    Call SaveSettings
End Sub

Private Sub mnuSetFileSavePath_Click()
    frmFilePath.Show
End Sub


Private Sub mnuThunder_Click()
    mnuThunder.Checked = True
    mnuHumilab.Checked = False
    ChamberType = THUNDER
End Sub

Private Sub mnuHumilab_Click()
Dim intButton As Integer
    mnuThunder.Checked = False
    mnuHumilab.Checked = True
    ChamberType = HUMILAB
    If (MSComm3.PortOpen = False) Then
        intButton = MsgBox("Humilab Reference COM port not open." + vbCr + "Open it now?", vbOKCancel)
        If (intButton = vbOK) Then
            Call subComPort_Click
        End If
    End If
End Sub

Private Sub mnuTurnOnLEDS_Click()
    Call DisplayCalCompleteScreen
End Sub

Private Sub DisplayCalCompleteScreen()
    Call HideDiagnostics
    mnuDiagnostics.Checked = False
    mnuNormalMode.Checked = False
    mnuTurnOnLEDS.Checked = True
    
    picCalComplete.Visible = True
    cmdSetLEDS.Visible = True
    grdSpreadsheet.Visible = False
    CalCompleteIndex = 0
    rackLedIndex = 0
    Call setCalCompleteScreen
    lstTasks.Height = DEFAULT_HEIGHT
End Sub


Private Sub mnuDiagnostics_Click()
    If (frmPassword.txtPassword = PASSWORD) Then
        Call SetupDiagnostics
    Else
        frmPassword.Show
    End If
End Sub

Public Sub SetupDiagnostics()
        Call DisplayDiagnostics
        mnuDiagnostics.Checked = True
        mnuNormalMode.Checked = False
        mnuTurnOnLEDS.Checked = False
    
        picCalComplete.Visible = False
        cmdSetLEDS.Visible = False
        grdSpreadsheet.Visible = False
        lstTasks.Enabled = True
        
        cmdSelect.Enabled = True
        cmdResume.Enabled = True
        scrTasks.Enabled = True
        
        mnuThunderMode.Enabled = True
        mnuFile.Enabled = True
        'mnuOptions.Enabled = True
        mnuComPort.Enabled = True
        mnuView.Enabled = True
        
        lstTasks.Height = DEFAULT_HEIGHT
End Sub


Private Sub HideDiagnostics()
        'If (ExcelCheck() = True) And (dataFilename <> "") Then
        '    excel_app.Visible = False
        'End If
        lstTasks.Height = DEFAULT_HEIGHT
        'frmMain.txtSpanOffset.Enabled = False
        mnuDiagnostics.Checked = False
        grdSpreadsheet.Visible = True
        lblSend.Visible = False
        lblReceive.Visible = False
        
        txtSend.Visible = False
        txtReceive.Visible = False
        'optNormalMode.Visible = False
        optOffMode.Visible = False
    
        picCalComplete.Visible = False
        cmdSetLEDS.Visible = False
        cmdCompressor.Visible = False
        
        txtChamberRhWaitTime.Enabled = False
        txtTemperatureSetpoint.Enabled = False
        txtTemperatureWaitTime.Enabled = False
        
        chkRack(1).Enabled = False
        chkRack(2).Enabled = False
        chkRack(3).Enabled = False
        chkRack(4).Enabled = False
        scrTasks.Enabled = False

        cmdLEDtest.Visible = False
        
End Sub

Private Sub DisplayDiagnostics()
        If (ExcelCheck() = True) And (dataFilename <> "") Then
            excel_app.Visible = True
        End If
        'frmMain.txtSpanOffset.Enabled = True
        scrTasks.Enabled = True
        cmdSelect.Enabled = True
        lstTasks.Enabled = True
        mnuDiagnostics.Checked = True
        grdSpreadsheet.Visible = False
        lblSend.Visible = True
        lblReceive.Visible = True

        txtSend.Visible = True
        txtReceive.Visible = True
        'optNormalMode.Visible = True
        
        optOffMode.Visible = True

        cmdComTest.Visible = True
        cmdLEDtest.Visible = True
        cmdCompressor.Visible = True
 
        
        txtChamberRhWaitTime.Enabled = True
        txtTemperatureSetpoint.Enabled = True
        txtTemperatureWaitTime.Enabled = True
        chkRack(1).Enabled = True
        chkRack(2).Enabled = True
        chkRack(3).Enabled = True
        chkRack(4).Enabled = True
        
End Sub

'Private Sub optNormalMode_Click()

'End Sub

'This command shuts down power to the sensor tip boards.
Private Sub optOffMode_Click()
Dim commandString As String
    'Send command to Interface board to turn on five volt sensor tip
    'power supply relays for NORMAL operation:
    commandString = ">X OFF"
    Call SendReceiveInterfaceBoard(commandString)
    frmMain.lblStatusBar.Caption = "Sensor tip supply is OFF"
End Sub



'Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)

'End Sub

Private Sub subNewDataFile_Click()
    Call createNewSpreadsheet
    Call copySpreadsheetToGrid
End Sub


'Opens an existing spreadsheet, and reads the Task Index from the appropriate cell.
'This is permits the user to resume the cal process from where it left off.
Private Sub subOpen_Click()
Dim tempFileName As String
Dim sheet As Object
Dim i As Integer
Dim row As Integer
Dim sensorDesignator As String


    'Preserve existing filename if there is one:
    tempFileName = dataFilename
 
    'If Cancel button is hit, quit this routine
    'without starting up Excel:
    cdbFile.CancelError = True
    On Error GoTo ErrHandler
    
    cdbFile.ShowOpen
    dataFilename = cdbFile.FileName
    StartExcelAndOpenFile (dataFilename)
    frmMain.Caption = VERSION + "      " + dataFilename
    frmMain.cmdSetLEDS.Caption = "Calibration Complete. Click here to remove Failed units."
    
        
    'Get index number of current task and make sure it is a valid value:
    TaskIndex = excel_app.Cells(4, 3).value
    'TaskIndex = 9 '####
    If (TaskIndex < 0) Then
        TaskIndex = 0
    ElseIf (TaskIndex > MaxTask) Then
        TaskIndex = MaxTask
    End If
    
    chkRack(1).value = 0
    chkRack(2).value = 0
    chkRack(3).value = 0
    chkRack(4).value = 0
    
    RacksUsedString = excel_app.Cells(RACK_ROW, 1).value
    
    If (InStr(1, RacksUsedString, ">A") > 0) Then
        chkRack(RACK_A).value = 1
    End If
            
    If (InStr(1, RacksUsedString, ">B") > 0) Then
        chkRack(RACK_B).value = 1
    End If

    If (InStr(1, RacksUsedString, ">C") > 0) Then
        chkRack(RACK_C).value = 1
    End If

    If (InStr(1, RacksUsedString, ">D") > 0) Then
        chkRack(RACK_D).value = 1
    End If
        
    maxSensors = getMaxSensors
    Call SaveSettings
    
    'Calibration will now be resumed at same task where it left off:
    lstTasks.ListIndex = TaskIndex
    scrTasks.value = TaskIndex
    Call copySpreadsheetToGrid
    cmdResume.Enabled = True
    Exit Sub
    
ErrHandler:
    dataFilename = tempFileName
    Exit Sub
    
End Sub



'Saves current spreadsheet
Private Sub subSave_Click()
    If (ExcelCheck() = True) Then
        With excel_app
            .ActiveWorkbook.Save
        End With
    End If
End Sub

'Returns a string with the cal date. Illegal Windows filename
'characters are removed and replaced with underscores "_"
Function GetCalDate() As String
Dim ch, calFileName, dateCreated As String
Dim length, i As Integer

    calFileName = ""
    dateCreated = Date
    length = Len(dateCreated)
    i = 1
    Do
        ch = Mid(dateCreated, i, 1)
        If (ch = "/") Then ch = "-"
        If (ch = ":") Then ch = "_"
        i = i + 1
        calFileName = calFileName + ch
    Loop Until i > length
    'calFileName = calFileName + ".xls"
    GetCalDate = calFileName
End Function

'This routine permits an existing open spreadsheet to be renamed and saved.
'As a default name, it creates a cal date string.
Private Sub subSaveAs_Click()

    cdbFile.FileName = GetCalDate + ".xls"
    
    'If Cancel button is hit, quit this routine
    'without starting up Excel:
    cdbFile.CancelError = True
    On Error GoTo SaveAsErrHandler
        
    cdbFile.ShowOpen
    dataFilename = cdbFile.FileName
    
    'excel_app.ActiveWorkbook.Save
    ActiveWorkbook.SaveAs FileName:=dataFilename
    frmMain.Caption = VERSION + "      " + dataFilename
    Exit Sub
        
SaveAsErrHandler:
    Exit Sub

End Sub

'This routine is called at the very end of the calibration run.
'It renames the CurrentCal file to a filename with the Run Number and cal date.
'The Run Number permits multiple runs in the same day.
'The default run number, obviously, is always "1".
Private Sub RenameAndSave()
Dim RunNumber As Integer
Dim CalDate As String
Dim PreviousCalDate As String
Dim networkFilename As String

    CalDate = GetCalDate
    PreviousCalDate = GetSetting(ProjectName, "RenameFile", "CalDate", "xxxxxxxx")
    SaveSetting ProjectName, "RenameFile", "CalDate", CalDate
    
    RunNumber = GetSetting(ProjectName, "RenameFile", "RunNumber", 1)
    
    'If the calibration date differs from the previous run,
    'then it must be the first run of the day, in which case
    'the run number should be reset to 1.
    'Otherwise it gets incremented to 2.
    'It is doubtful that there would ever be more than two
    'cal runs in a single day, but in any case,
    'RunNumber gets incremented as necessary:
    If (CalDate <> PreviousCalDate) Then
        RunNumber = 1
    Else
        RunNumber = RunNumber + 1
    End If
    
    SaveSetting ProjectName, "RenameFile", "RunNumber", RunNumber

    'dataFilename = "c:\Cal Data\Cal Run #" + Format$(RunNumber) + " " + CalDate + ".xls"
    
    'Save on LOCAL drive: Set output filename to new name default and delete existing file by that name:
    dataFilename = strLocalFolder + "Cal Run #" + Format$(RunNumber) + " " + CalDate + ".xls"
    ActiveWorkbook.SaveAs FileName:=dataFilename
    frmMain.Caption = VERSION + "      " + dataFilename
        
    'Now save on NETWORK drive, if network is enabled:
    If (frmFileFolder.chkNetworkEnable = Checked) Then
        If (CheckFolder(strNetworkFolder) = True) Then
            networkFilename = strNetworkFolder + "Cal Run #" + Format$(RunNumber) + " " + CalDate + ".xls"
            ActiveWorkbook.SaveAs FileName:=networkFilename
        Else
            frmError.Show
            frmError.lblLabelThree.Caption = "Network error: could not access backup drive or folder."
            frmError.lblLabelFour.Caption = "Folder name: " + strNetworkFolder
        End If
    End If
    
End Sub



'This routine looks for the folder defined by "strFolder"
'and returns TRUE if it can be found, and FALSE otherwise.
Function CheckFolder(strFolder As String) As Boolean
    Dim fs As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    CheckFolder = fs.FolderExists(strFolder)
End Function


'This routine reads the value entered by the user in the text Chamber Wait Time box and stores it.
Private Sub txtChamberRhWaitTime_Change()
    chamberRhWaitTime = Val(txtChamberRhWaitTime.Text)
    'Call SaveSettings
End Sub




Public Sub subSetupCal_Click()
    frmCalSetup.Show
End Sub

Private Sub Form_Load()

    compressorFlag = False
    dataFilename = ""
    Pathname = ""
    frmMain.Caption = VERSION + "              NO FILE LOADED"
    
    
    
    Timer3Counter = 0
    Timer3Timeout = 0
    CalCompleteIndex = 0
    
    SelectFlag = False
    TaskIndex = 0
    scrTasks.value = TaskIndex
    Call SetUpTaskBox

    ChamberRH = 0
    ChamberTempC = 0
    thunderCommandString = THUNDER_READ_RH_AND_T
        
    Call HideDiagnostics
    'Call DisplayDiagnostics
    frmMain.Show
           
    PortOpen = False
    
    frmMain.lblStatusBar.Caption = "Opening Serial Ports. Please Wait"
    Call Startup
    
    If (PortOpen = True) Then
        Call initializeFlukeMeter
        'frmMain.lblStatusBar.Caption = "Serial Ports OK..."
    Else
        frmMain.lblStatusBar.Caption = "Error opening Serial Ports."
    End If
    
    VoltUARTbuffer = ""
    
    UARTbuffer = ""
    VoltUARTbuffer = ""
    ChamberRefUARTbuffer = ""
    ChamberControlUARTbuffer = ""
    
    ChamberRefUARTflag = False
    ChamberControlUARTflag = False
    UARTflag = False
    VoltUARTflag = False
    ChamberRH = 0
    
    RunFlag = False
    barProgress.value = 0
    
    cmdStart.Enabled = True
    
    userSensorNumber = 1
    userPot1Value = 128
    userPot2Value = 128
    LEDtestIndex = 1
    

    frmMain.lblStatusBar.Caption = "Ready to run calibration."
End Sub


'This routine determines whether Excel is running,
'and if so, is there an open spreadsheet named as the string dataFilename?
'If not, then it closes Excel so that another routine can start it up from scratch.
Function ExcelCheck() As Boolean
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

'Starts up Excel and opens an existing spreadsheet.
Private Sub StartExcelAndOpenFile(spreadsheetName As String)
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

'This routine is called when the program is shut down.
'It insures that Excel is properly closed.
'It also closes the COM ports, and saves all startup setting information.
'The setup file is also saved. This stores the array that indicates
'which of 128 possible sesnor tip sockets are being used,
'and which of four racks are being used.
Private Sub Form_Unload(Cancel As Integer)
Dim test As Boolean
    test = ExcelCheck()
    If ((dataFilename <> "") And (test = True)) Then
        ' Close the open workbook:
        With excel_app
            .ActiveWorkbook.Close True
            .Quit
            'excel_app.Quit
            Set excel_sheet = Nothing
            Set excel_app = Nothing
        End With
        'Screen.MousePointer = vbDefault
    End If

    Call ComPortShutDown
    Call SaveSettings
    End
End Sub

Private Sub subExit_Click()
    Call Form_Unload(0)
End Sub



'This routine receives incoming UART characters for the Chamber Ref COM port
'used for communicating with the Humilab Chamber reference port
Public Sub MSComm3_OnComm()
    Dim StartNumber, EndNumber, LengthNumber As Integer
    Dim NumberString As String
    Dim ChamberRefIntext As String
    Const CR = vbCr
    
   Select Case MSComm3.CommEvent
   ' Handle each event or error by placing
   ' code below each case statement

   ' Errors
      Case comEventBreak   ' A Break was received.
      Case comEventFrame   ' Framing Error
      Case comEventOverrun   ' Data Lost.
      Case comEventRxOver   ' Receive buffer overflow.
      Case comEventRxParity   ' Parity Error.
      Case comEventTxFull   ' Transmit buffer full.
      Case comEventDCB   ' Unexpected error retrieving DCB]

   ' Events
      Case comEvCD   ' Change in the CD line.
      Case comEvCTS   ' Change in the CTS line.
      Case comEvDSR   ' Change in the DSR line.
      Case comEvRing   ' Change in the Ring Indicator.
      Case comEvReceive   ' Received RThreshold # of chars.
            ChamberRefUARTbuffer = ChamberRefUARTbuffer + MSComm3.Input
            
            'When Control text is received, incoming text from Humilab reference is complete.
            If InStr(1, ChamberRefUARTbuffer, "Control") > 0 Then
                ChamberRefIntext = ChamberRefUARTbuffer
                ChamberRefUARTbuffer = ""
                
                'Now pull RH string out and get RH value first:
                i = InStr(1, ChamberRefIntext, "RH")
                If i > 0 Then
                    StartNumber = InStr(i, ChamberRefIntext, "=") + 1
                    EndNumber = InStr(i, ChamberRefIntext, CR)
                    LengthNumber = EndNumber - StartNumber
                    If (LengthNumber > 0) Then
                        NumberString = Mid(ChamberRefIntext, StartNumber, LengthNumber)
                        ChamberRH = Val(NumberString)
                        lblChamberRH.Caption = "Chamber RH: " + Format$(ChamberRH, "###.#") + "%"
                    End If
                End If
                
                'Now pull temperature string out and get temperature:
                i = InStr(1, ChamberRefIntext, "TMP C")
                If i > 0 Then
                    StartNumber = InStr(i, ChamberRefIntext, "=") + 1
                    EndNumber = InStr(i, ChamberRefIntext, CR)
                    LengthNumber = EndNumber - StartNumber
                    If (LengthNumber > 0) Then
                        NumberString = Mid(ChamberRefIntext, StartNumber, LengthNumber)
                        ChamberTempC = Val(NumberString)
                        lblChamberTemp.Caption = "Temperature C: " + Format$(ChamberTempC, "###.#")
                    End If
                End If
                
                'If the chamber control input port is open, send it the whole
                'control string from the reference port:
                ChamberRefUARTflag = True
            End If
      Case comEvSend   ' There are SThreshold number of
                     ' characters in the transmit
                     ' buffer.
      Case comEvEOF   ' An EOF character was found in
                     ' the input stream
   End Select
End Sub


'This routine searches an input string from the Thunder COM port
'and returns a value specified by the input numberIndex.
'
' Example: if numberIndex is a 3, then this routine would return a 45.3 from the following string:
' 10.3 , 10.11 , 45.3 , 54.00 , 32.77 , 25.77 , 22.6 , 1
'
' ASCII strings from the Thunder control COM port are terminated
' with a carriage return (vbCr) followed by a line feed (vbLf)
' Note that each value in the input string is followed by a comma
' except for the last value.

Function getValueFromThunderString(numberIndex As Integer, inputString As String) As Double
Dim value As Double
Dim ptrStartVal As Integer    'points to position of first digit value, ie "4" if value is "45.3 ,"
Dim ptrEndVal As Integer         'points to position of comma following value:  "45.3 ,"
Dim valueCount As Integer       'Keeps count of how many numbers have been
                                'encountered while searching through string.
Dim ptrNextComma As Integer
Dim valueString As String
Dim valueStringLength As Integer
Const CR = vbCr

    valueCount = 0
    ptrEndVal = 0
    Do
        ptrNextComma = InStr(ptrEndVal + 1, inputString, ",")
        If (ptrNextComma > 0) Then
            ptrStartVal = ptrEndVal + 1
            ptrEndVal = ptrNextComma
            valueCount = valueCount + 1
        Else
            'If we are searching for the last value in string,
            'then there is a carriage return instead of a comma:
            ptrStartVal = ptrEndVal + 1
            ptrEndVal = InStr(1, inputString, CR)
            If (ptrEndVal > 0) Then valueCount = valueCount + 1
        End If
    Loop While (valueCount < numberIndex) And (ptrNextComma > 0)
    
    valueStringLength = ptrEndVal - ptrStartVal
    If (valueStringLength > 0) Then
        valueString = Mid(inputString, ptrStartVal, valueStringLength)
        value = Val(valueString)
        getValueFromThunderString = value
    Else
        getValueFromThunderString = 0
    End If
    
End Function


'This routine creates a delay equal to
'the number of milliseconds specified by Msec
Public Sub Delay(Msec As Single)
Dim t As Single
t = Msec + timeGetTime()
Do Until timeGetTime() >= t
DoEvents
Loop
End Sub

'This routine brings up the com port form which selects the ports
'required to run this program. The Humilab requires four ports
'and the Thunder requires three. So one of the port settings on this
'form is left invisible for the THUNDER option.
Private Sub subComPort_Click()
    If (ChamberType = HUMILAB) Then
        frmPortSettings.fraReference.Visible = True
    Else
        frmPortSettings.fraReference.Visible = False 'For Thunder: leave one port invisible
    End If
        
    Call frmPortSettings.initializePortSettings
    frmPortSettings.Show
End Sub







'This routine attempts to insure that all four COM ports
'can send and receive data.
Private Sub CheckCommunication()
Dim result As Integer
Dim comPortTest As Boolean

    'If COM ports aren't open, try opening them:
    If (False = PortOpen) Then
        'PortOpen = frmMain.openComPorts(1200, PortNumber, VoltmeterPortNumber, ChamberRefPortNumber, ChamberControlPortNumber)
        PortOpen = frmMain.openComPorts(1200, PortNumber, VoltmeterPortNumber, ChamberControlPortNumber)
    End If

    'If COM ports still aren't open, then quit routine:
    If (False = PortOpen) Then
        RunFlag = False
        MsgBox ("COM ports not working." + vbCr + "Try Com Port pull down menu")
    Else
        'Test Main (Interface) Board communication first:
        comPortTest = False
        While (comPortTest = False) And (RunFlag = True)
            If (interfaceComTest() = False) Then
                result = MsgBox("Make sure power supply is turned on. Try again?", vbOKCancel + vbCritical + vbDefaultButton1, "Interface Board Not Responding")
                If (result = vbCancel) Then RunFlag = False
            Else
                                            'If Interface board is working send command
                                            'to make sure power to racks is turned on.
                Call turnOnCompressor       'Make sure that compressor is on.
                comPortTest = True
                frmMain.lblStatusBar.Caption = "Success - main board communication works!"
            End If
        Wend
    
        'Test voltmeter communication:
        comPortTest = False
        While (comPortTest = False) And (RunFlag = True)
            Call initializeFlukeMeter
            If (checkVoltmeterCommunication() = False) Then
              result = MsgBox("Make sure voltmeter is turned on. Try again?", vbOKCancel + vbCritical + vbDefaultButton1, "Voltmeter Not Responding")
              If (result = vbCancel) Then RunFlag = False
            Else
                comPortTest = True
                frmMain.lblStatusBar.Caption = "Success - voltmeter communication works!"
            End If
        Wend
        
        'Test Humilab Reference communication:
        If (ChamberType = HUMILAB) Then
            comPortTest = False
            While (comPortTest = False) And (RunFlag = True)
                If (checkHumilabRefCommunication() = False) Then
                    result = MsgBox("Make sure Humilab Reference is connected. Try again?", vbOKCancel + vbCritical + vbDefaultButton1, "Humilab Not Responding")
                    If (result = vbCancel) Then RunFlag = False
                Else
                    comPortTest = True
                    frmMain.lblStatusBar.Caption = "Success - Humilab Reference communication works!"
                End If
            Wend
        End If
    
        'Test Chamber Control communication:
        comPortTest = False
        While (comPortTest = False) And (RunFlag = True)
            If (checkChamberControlCommunication() = False) Then
                result = MsgBox("Make sure Chamber Control is connected. Try again?", vbOKCancel + vbCritical + vbDefaultButton1, "Chamber Not Responding")
                If (result = vbCancel) Then RunFlag = False
            Else
                comPortTest = True
                frmMain.lblStatusBar.Caption = "Success - Chamber Control communication works!"
            End If
        Wend
    End If
    
    If (RunFlag = False) Then result = MsgBox("Calibration has been aborted.", vbOK + vbCritical + vbDefaultButton1, "PROCESS HALTED")
End Sub


'Set chamber loop: set chamber setpoint, wait for elapsed time.
'Elapsed time is computed using Timer(),
'which returns time in seconds since previous midnight.
'
'StartTime divides this value by 60 to get starting minutes
'at the beginning of the routine
'when the chamber is set to the new setpoint.
'DiffTime is the difference between the start and current time.
'elapsedTime takes into account the rollover which occurs
'if the chamber runs past midnight, in which case the elapsed time
'up until that point is copied to OffsetTime.
Private Sub setChamber(setpoint As Integer)
Const AllowableSetpointError = 2
Dim intResult As Integer
Dim setpointString As String
Dim receivedString As String
Dim verifyDouble As Double
Dim setpointFlag As Boolean

Dim verifySetpoint As Integer
Dim verifyFlag As Boolean
Dim displayString As String
Dim currentTime As Variant
Dim intStartTime, intDiffTime, intElapsedTime, intOffsetTime, intPreviousTime As Integer
Dim i As Integer
Dim j As Integer
Dim intChamberRH As Integer

    'Progress bar will reach end when chamberRhWaitTime has elapsed
    barProgress.value = 0
    barProgress.Max = chamberRhWaitTime
    Delay (100)
    
    i = 0
    displayString = " "
    'Keep sending setpoint to chamber control until it returns it:
    Do
        lbl_RH_setpoint.Caption = "Setpoint = "
        If (ChamberType = HUMILAB) Then
            setpointString = "*+++" + Format$(setpoint)
            receivedString = SendReceiveHumilabControlCOM(setpointString)
        Else
            thunderCommandString = THUNDER_READ_SETPOINTS
        End If
        
        If (ChamberType = HUMILAB) Then
            'Valid input string from chamber should contain a decimal point
            'and at least one digit:
            j = InStr(1, receivedString, ".")
            If (j > 0) Then
                verifySetpoint = CInt(Val(receivedString))
                'Replace decimal point with %RH:
                If (j > 1) Then displayString = Left$(receivedString, j - 1)
                lbl_RH_setpoint.Caption = "Setpoint = " + displayString + "% RH"
            End If
            
            If (verifySetpoint = setpoint) Then
                verifyFlag = True
            Else
                verifyFlag = False
            End If
            
        Else
        
            'Send temperature setpoint string to Thunder:
            frmMain.lblStatusBar.Caption = "Setting temperature to: " + Format$(temperatureSetpoint) + " degrees C, " + "Trial #" + Format$(i)
            thunderCommandString = "TS=" + Format$(temperatureSetpoint)
            Timer3Counter = 0
            Do
                DoEvents
            Loop While (Timer3Counter < 2)
        
            'Send RH setpoint string to Thunder:
            If (thunderMode = RhPc) Then
                thunderCommandString = "R1=" + Format$(setpoint)
                frmMain.lblStatusBar.Caption = "Setting Rh@Pc to: " + Format$(setpoint) + " %RH, " + "Trial #" + Format$(i)
            Else
                thunderCommandString = "R2=" + Format$(setpoint)
                frmMain.lblStatusBar.Caption = "Setting Rh@PcTc to: " + Format$(setpoint) + " %RH, " + "Trial #" + Format$(i)
            End If
            
            Timer3Counter = 0
            Do
                DoEvents
            Loop While (Timer3Counter < 2)
                        
            'Now read back setpoint:
            thunderCommandString = THUNDER_READ_SETPOINTS
            Timer3Counter = 0
            Do
                DoEvents
            Loop While (Timer3Counter < 3)
           
            If ((verifiedRHsetpoint = setpoint) And (verifiedTemperatureSetpoint = temperatureSetpoint)) Then
                verifyFlag = True
            Else
                verifyFlag = False
            End If
        End If
        i = i + 1
    Loop While ((i <= RETRIES) And (verifyFlag = False))
    
    'If (verifySetpoint <> setpoint) Then
    '    intResult = MsgBox("Check chamber setpoint." + vbCr + "Is it " + Format$(setpoint) + "?", vbYesNo)
    '    If (intResult = vbNo) Then
    '        MsgBox ("Calibration halted." + vbCr + "Check chamber control COM port.")
    '        RunFlag = False
    '    End If
    'End If
    
    thunderCommandString = THUNDER_READ_RH_AND_T
    
    setpointFlag = False
    If (RunFlag = True) Then
        'StartTime is the time in minutes when chamber
        'is set to the new setpoint:
        currentTime = Timer()
        intStartTime = CInt(currentTime / 60#)
        intOffsetTime = 0
        intPreviousTime = 0
        intElapsedTime = 0
        frmMain.lblStatusBar.Caption = "Waiting for chamber to reach " + Format$(setpoint) + "%RH,  Time: " + Format$(intElapsedTime) + " minutes"
        Do
            DoEvents
            'DiffTime is the elapsed time since chamber is set to new setpoint,
            'assuming that we haven't just passed midnight:
            currentTime = Timer()
            intDiffTime = (CInt(currentTime / 60#)) - intStartTime
            
            'If DiffTime is negative, then we must have just passed midnight.
            'So we need to store the elapsed time thus far
            'and record a new StartTime.
            'Offset time is now the time elapsed before midnight,
            'DiffTime will be the time elapsed after midnight, and
            'the total ElapsedTime will be the sum of the two:
            If (intDiffTime < 0) Then
                    intOffsetTime = intElapsedTime
                    currentTime = Timer()
                    intStartTime = CInt(currentTime / 60#)
                    intDiffTime = 0
            End If
            
            intElapsedTime = intDiffTime + intOffsetTime
            'If another minute has just passed by, update time displayed:
            If (intElapsedTime <> intPreviousTime) Then
                intPreviousTime = intElapsedTime
                frmMain.lblStatusBar.Caption = "Waiting for chamber to reach " + Format$(setpoint) + "%RH,  Time: " + Format$(intElapsedTime) + " minutes"
                If (intElapsedTime < barProgress.Max) Then
                    barProgress.value = intElapsedTime
                End If
                
                If (intElapsedTime >= chamberRhWaitTime) Then
                    intChamberRH = CInt(ChamberRH)
                    If (Abs(setpoint - intChamberRH) > AllowableSetpointError) Then
                        frmError.Show
                        frmError.lblLabelOne.Caption = "This chamber did not reach " + Format$(setpoint) + "% within " + Format$(chamberRhWaitTime) + " minutes."
                        frmError.lblLabelTwo.Caption = "At " + Format$(intElapsedTime) + " minutes, the RH = " + Format$(ChamberRH, "##.##") + "%."
                    Else
                       setpointFlag = True
                    End If
                End If
            End If
            DoEvents
            
        Loop While ((setpointFlag = False) And (RunFlag = True))
    End If
End Sub

Private Sub scrTasks_Change()
    TaskIndex = scrTasks.value
    lstTasks.ListIndex = TaskIndex
End Sub


'This routine polls the four sensor tip boards to see which ones are present.
'If they are connected to the Interface Board and the power supply is on,
'then this routine should be able to detect the ones that are hooked up
'and record them in the Cal.rackUsed(i) array.
'Also the chkRack check boxes are updated to indicate which ones are being used.
Public Sub identifySensorRacksAndCopyToSpreadsheet()
Dim rackNumber As Integer
Dim socketNumber As Integer
Dim sensorNumber As Integer
Dim row As Integer
Dim command As String
Dim socketDesignator As String
Dim RacksUsedString As String

Const RACK_A = 1
Const RACK_B = 2
Const RACK_C = 3
Const RACK_D = 4
    
    For rackNumber = RACK_A To RACK_D
            If rackNumber = RACK_A Then
                command = ">A LEDS 0 0 0 0 0"
            ElseIf rackNumber = RACK_B Then
                command = ">B LEDS 0 0 0 0 0"
            ElseIf rackNumber = RACK_C Then
                command = ">C LEDS 0 0 0 0 0"
            Else
                command = ">D LEDS 0 0 0 0 0"
            End If
            
            If (SendReceiveInterfaceBoard(command) = True) Then
                If (InStr(1, Intext, "OK")) > 0 Then
                    chkRack(rackNumber).value = 1
                Else
                    chkRack(rackNumber).value = 0
                End If
            Else
                MsgBox ("Interface board not responding. Check power and try again")
                RunFlag = False
                GoTo Quit
            End If
    Next rackNumber
    
    If ((chkRack(RACK_A).value = 0) And (chkRack(RACK_B).value = 0) And (chkRack(RACK_C).value = 0) And (chkRack(RACK_D).value = 0)) Then
        MsgBox ("No sensor rackNumbers detected!" + vbCrLf + "Make sure ribbon cables are connected and power is on.")
        RunFlag = False
        GoTo Quit
    End If
    
    RacksUsedString = excel_app.Cells(RACK_ROW, 1).value
    If (chkRack(RACK_A).value = 1) Then RacksUsedString = RacksUsedString + " >A,"
    If (chkRack(RACK_B).value = 1) Then RacksUsedString = RacksUsedString + " >B,"
    If (chkRack(RACK_C).value = 1) Then RacksUsedString = RacksUsedString + " >C,"
    If (chkRack(RACK_D).value = 1) Then RacksUsedString = RacksUsedString + " >D,"
    excel_app.Cells(RACK_ROW, 1).value = RacksUsedString
    
    'TODO: think about this a little bit.
    'Now assign a designator string to every sensor.
    'This string is transmitted to the Interface board
    'for every command that addresses a sensor.
    'The designator includes a letter to indicate the rack:
    '"A" is rack number 1
    '"B" is rack number 2
    '"C" is rack number 3
    '"D" is rack number 4
      
    maxSensors = getMaxSensors
    Call SaveSettings
    
    For sensorNumber = 1 To 128
        row = sensorNumber + OFFSET
        excel_app.Cells(row, SENSOR_COLUMN).value = getSensorString(sensorNumber)
    Next sensorNumber
    
    For sensorNumber = 1 To maxSensors
            excel_app.Cells(sensorNumber + ROW_OFFSET, SENSOR_USED_COLUMN).value = True
            excel_app.Cells(sensorNumber + ROW_OFFSET, STATUS_COLUMN).value = "OK"
    Next sensorNumber
    
Quit:
End Sub

'This routine is called after sending a read pot command
'to the Interface Board. It searches incoming text strings
'for fuse status information such as "BLOWN" from correctly
'programmed pots. It then returns a single integer to indicate that status.
' The return values are:
'
' I2C_ERROR = 0
' READY_TO_PROGRAM = 1
' FUSE_BLOWN = 2
' FUSE_BAD = 3
' INVALID_POT_DATA = 4
' COM_ERROR = 5

Function processReceivedFuseStatus(potSelect As Integer) As Integer
Dim statusString As String
Dim potValue As Integer
    
    processReceivedFuseStatus = 0
    If InStr(1, txtReceive.Text, "ERROR") Then
        processReceivedFuseStatus = I2C_ERROR
        COMflag = False
        statusString = "I2C ERROR"
    ElseIf InStr(1, txtReceive.Text, "READY") Then
        processReceivedFuseStatus = READY_TO_PROGRAM
        statusString = "Ready to program"
    ElseIf InStr(1, txtReceive.Text, "BLOWN") Then
        processReceivedFuseStatus = FUSE_BLOWN
        statusString = "PROGRAMMED"
    ElseIf InStr(1, txtReceive.Text, "BAD") Then
        processReceivedFuseStatus = FUSE_BAD
        statusString = "FUSE PROGRAM ERROR"
    Else
        processReceivedFuseStatus = COM_ERROR
        statusString = "COM ERROR"
    End If
    
    If (potSelect = 1) Then
        lblPot1Status.Caption = statusString
    Else
        lblPot2Status.Caption = statusString
    End If
End Function

Function getSensorString(sensorNumber As Integer) As String
Dim rackNumber As Integer
Dim rackUsed(1 To 4) As String
Dim j As Integer
Dim socketNumber As String
Dim sensorDesignator As String
Dim Dina As Integer

    Dina = chkRack(1).value
    
    For j = 1 To 4
        rackUsed(j) = ""
    Next j

    If (sensorNumber > 128) Then
        getSensorString = "ERROR"
    Else
        sensorDesignator = ""
        rackNumber = ((sensorNumber - 1) \ 32) + 1 'Integer devision - round off to nearest integer to find out which rack to use: rack #1 to#4
        socketNumber = ((sensorNumber - 1) Mod 32) + 1 'Get remainder - this will give us the sensor socket number: #1 to #32
        j = 1
        For rack = 1 To 4
            If (chkRack(rack).value = 1) Then
                If (rack = 1) Then
                    rackUsed(j) = ">A "
                    j = j + 1
                ElseIf (rack = 2) Then
                    rackUsed(j) = ">B "
                    j = j + 1
                ElseIf (rack = 3) Then
                    rackUsed(j) = ">C "
                    j = j + 1
                ElseIf (rack = 4) Then
                    rackUsed(j) = ">D "
                    j = j + 1
                End If
            End If
        Next rack
        
        'If (rackUsed() is blank, then the rack isn't being used,
        'and there is no sensor tip corresponding to this sensor number.
        'So for that case, this routine returns a blank string: ""
        If (rackUsed(rackNumber) = "") Then
            getSensorString = ""
        Else
            sensorDesignator = rackUsed(rackNumber) + "#S" + Format$(socketNumber)
        End If
    End If
    
    getSensorString = sensorDesignator
End Function

'For this rev, active racks cannot be selected by the user.
Private Sub chkRack_Click(Index As Integer)
    maxSensors = getMaxSensors
'    Call SaveSettings
End Sub

'Determine the maximum number of sensor sockets that can have sensors in them.
'This will be a multiple of 32: 32, 64, 96, or 128
Function getMaxSensors() As Integer
Dim i As Integer
Dim maxNumberOfSensors As Integer
    maxNumberOfSensors = 0
    For i = 1 To 4
        If (chkRack(i).value = 1) Then maxNumberOfSensors = maxNumberOfSensors + 32
    Next i
    getMaxSensors = maxNumberOfSensors
End Function

           
'Send the STOP command to shut down the Thunder.
Public Sub stopThunder()
    frmMain.lblStatusBar.Caption = "Shutting down Thunder"
    thunderCommandString = THUNDER_STOP_COMMAND
    Timer3Counter = 0
    Do
        DoEvents
    Loop While (Timer3Counter < 8)
    thunderCommandString = THUNDER_READ_RH_AND_T
End Sub




'This routine sends a command string out to the MAIN (interface) cal board,
'and waits for a response. It will make up to 3 repeat attempts.
'If a string is received from the Interface board,
'this function returns True. After 3 failed attempts,
'it stops the process and displays a com failure message box.
'The routine then returns False.
Function SendReceiveInterfaceBoard(commandString As String) As Boolean
Dim i As Integer
Dim j As Integer
Dim dummy As String
Dim seconds As Integer


    dummy = MSComm1.Input 'Do a dummy read to clear incoming com port
    
    UARTbuffer = ""
    Intext = ""
    txtReceive.Text = ""
    UARTflag = False
    txtSend.Text = commandString
    
    i = 1
    Do
        'Outgoing command is sent here. Make sure that port is open first:
        If (MSComm1.PortOpen = True) Then MSComm1.Output = commandString + vbCr
    
        'Now wait for a response from Interface board. Resend command every 10 seconds.
        j = 0
        While (UARTflag = False) And (j < 100)
            seconds = j / 10
            frmMain.lblStatusBar.Caption = "Sending " + commandString + " to Interface board, " + Format$(seconds) + " seconds, try #" + Format$(i)
            Delay (100)
            j = j + 1
        Wend
        frmMain.lblStatusBar.Caption = "Sending " + commandString + " to Interface board, try #" + Format$(i)
        i = i + 1
        If (i = 4) Then
            'If Interface board isn't working, then either fix problem or abort process:
            result = MsgBox("Make sure power supply is on. Try again?", vbYesNo + vbCritical + vbDefaultButton1, "Interface Board Not Responding")
            'Try again:
            If (result = vbYes) Then
                i = 1
                UARTbuffer = ""
                Intext = ""
                txtReceive.Text = ""
                UARTflag = False
                dummy = MSComm1.Input 'Do a dummy read to clear incoming com port
            '...otherwise abort process:
            Else
                RunFlag = False
            End If
        End If
    Loop While (UARTflag = False) And (i <= 3)
    
    If (UARTflag = False) Then
        SendReceiveInterfaceBoard = False
    Else
        SendReceiveInterfaceBoard = True
        txtReceive.Text = Intext
    End If
End Function


'This routine receives incoming UART characters for the main COM port
'used for communicating with the calibration INTERFACE board.
'Incoming strings are terminated with a carriage return.
'This sets a global flag "UARTflag" to True,
'thereby indicating that the incoming message buffer is true.
'The incoming string received from the Interface board is then
'transferred to the global variable Intext, so that it can be processed by
'the routine SendReceiveInterfaceBoard()
Public Sub MSComm1_OnComm()
    Const CR = vbCr
    
    Select Case MSComm1.CommEvent
    ' Handle each event or error by placing
    ' code below each case statement

    ' Errors
      Case comEventBreak   ' A Break was received.
      Case comEventFrame   ' Framing Error
      Case comEventOverrun   ' Data Lost.
      Case comEventRxOver   ' Receive buffer overflow.
      Case comEventRxParity   ' Parity Error.
      Case comEventTxFull   ' Transmit buffer full.
      Case comEventDCB   ' Unexpected error retrieving DCB]

   ' Events
      Case comEvCD   ' Change in the CD line.
      Case comEvCTS   ' Change in the CTS line.
      Case comEvDSR   ' Change in the DSR line.
      Case comEvRing   ' Change in the Ring Indicator.
      Case comEvReceive   ' Received RThreshold # of chars.
            UARTbuffer = UARTbuffer + MSComm1.Input
            If InStr(1, UARTbuffer, CR) > 0 Then
                UARTflag = True
                Intext = UARTbuffer
                UARTbuffer = ""
            End If
      Case comEvSend   ' There are SThreshold number of
                     ' characters in the transmit
                     ' buffer.
      Case comEvEOF   ' An EOF character was found in
                     ' the input stream
   End Select
End Sub



'This routine receives incoming UART characters for the Chamber Control COM port
'used for communicating with the Humilab Chamber control port
'or with the Thunder port.
Public Sub MSComm4_OnComm()
Const LINEFEED = vbLf
Const CR_LINEFEED = vbCrLf
Dim thunderReceiveString As String

   Select Case MSComm4.CommEvent
   ' Handle each event or error by placing
   ' code below each case statement

   ' Errors
      Case comEventBreak   ' A Break was received.
      Case comEventFrame   ' Framing Error
      Case comEventOverrun   ' Data Lost.
      Case comEventRxOver   ' Receive buffer overflow.
      Case comEventRxParity   ' Parity Error.
      Case comEventTxFull   ' Transmit buffer full.
      Case comEventDCB   ' Unexpected error retrieving DCB]
   ' Events
      Case comEvCD   ' Change in the CD line.
      Case comEvCTS   ' Change in the CTS line.
      Case comEvDSR   ' Change in the DSR line.
      Case comEvRing   ' Change in the Ring Indicator.
      Case comEvReceive   ' Received RThreshold # of chars.
            
            ChamberControlUARTbuffer = ChamberControlUARTbuffer + MSComm4.Input
            Timer3Timeout = 0
                        
            If (ChamberType = HUMILAB) Then
                If InStr(1, ChamberControlUARTbuffer, ".") > 0 Then
                    ChamberControlUARTflag = True
                End If
            Else
                If InStr(1, ChamberControlUARTbuffer, LINEFEED) > 0 Then
                    thunderReceiveString = ChamberControlUARTbuffer
                    ChamberControlUARTbuffer = ""
                    
                    If (thunderCommandString = THUNDER_READ_RH_AND_T) Then
                        If (thunderMode = RhPc) Then
                            ChamberRH = getValueFromThunderString(1, thunderReceiveString)
                            lblChamberRH.Caption = "Chamber RH@Pc: " + Format$(ChamberRH, "###.#") + "%"
                        ElseIf (thunderMode = RhPcTc) Then
                            ChamberRH = getValueFromThunderString(2, thunderReceiveString)
                            lblChamberRH.Caption = "Chamber RH@PcTc: " + Format$(ChamberRH, "###.#") + "%"
                        End If

                        ChamberTempC = getValueFromThunderString(6, thunderReceiveString)
                        lblChamberTemp.Caption = "Temperature C: " + Format$(ChamberTempC, "###.#")
                        ChamberControlUARTflag = True
                    ElseIf (thunderCommandString = THUNDER_READ_SETPOINTS) Then
                        If (thunderMode = RhPc) Then
                            verifiedRHsetpoint = CInt(getValueFromThunderString(1, thunderReceiveString))
                        ElseIf (thunderMode = RhPcTc) Then
                            verifiedRHsetpoint = CInt(getValueFromThunderString(2, thunderReceiveString))
                        End If
                        verifiedTemperatureSetpoint = CInt(getValueFromThunderString(4, thunderReceiveString))
                        lbl_RH_setpoint.Caption = "Setpoint: " + Format$(verifiedRHsetpoint)
                        ChamberControlUARTflag = True
                    ElseIf (thunderCommandString = THUNDER_RUN_COMMAND) Then
                        If InStr(1, thunderReceiveString, CR_LINEFEED) > 0 Then ChamberControlUARTflag = True
                    End If
                End If
            End If
            
      Case comEvSend   ' There are SThreshold number of
                     ' characters in the transmit
                     ' buffer.
      Case comEvEOF   ' An EOF character was found in
                     ' the input stream
   End Select
End Sub



'This routine captures received data from the voltmeter Com Port
Public Sub MSComm2_OnComm()
    Dim VoltText As String
    Dim endNum As Integer   'Place in incoming voltmeter number string where number ends
    Dim startNum As Integer 'Place in incoming voltmeter number string where number begins
    Dim lengthNum As Integer 'Length of numeric portion of string
    Dim measuredRH As Double
    Const CR = vbCr
    
   Select Case MSComm2.CommEvent
   ' Handle each event or error by placing
   ' code below each case statement

   ' Errors
      Case comEventBreak   ' A Break was received.
      Case comEventFrame   ' Framing Error
      Case comEventOverrun   ' Data Lost.
      Case comEventRxOver   ' Receive buffer overflow.
      Case comEventRxParity   ' Parity Error.
      Case comEventTxFull   ' Transmit buffer full.
      Case comEventDCB   ' Unexpected error retrieving DCB]

   ' Events
      Case comEvCD   ' Change in the CD line.
      Case comEvCTS   ' Change in the CTS line.
      Case comEvDSR   ' Change in the DSR line.
      Case comEvRing   ' Change in the Ring Indicator.
      Case comEvReceive   ' Received RThreshold # of chars.
            VoltUARTbuffer = VoltUARTbuffer + MSComm2.Input
            If InStr(1, VoltUARTbuffer, ">") > 0 Then
                endNum = InStr(1, VoltUARTbuffer, "=") - 1
                startNum = InStr(1, VoltUARTbuffer, "MEAS?") + 5
                lengthNum = endNum - startNum
                If lengthNum > 1 Then
                    VoltText = Mid$(VoltUARTbuffer, startNum, lengthNum)
                    voltage = Val(VoltText)
                    measuredRH = calculateRh(voltage)
                    lblVoltmeter.Caption = "UUT Volts: " + Format$(voltage, "##.###") + "v"
                    lblMeasuredRH.Caption = "UUT RH:    " + Format$(measuredRH, "##.##") + "%"
                    VoltUARTflag = True
                End If
                VoltUARTbuffer = ""
            End If
      Case comEvSend   ' There are SThreshold number of
                     ' characters in the transmit
                     ' buffer.
      Case comEvEOF   ' An EOF charater was found in
                     ' the input stream
   End Select
End Sub

'This routine sends the MEASURE command to the voltmeter
'approximately every second.
Private Sub Timer2_Timer()
    If (MSComm2.PortOpen = True) Then
        MSComm2.Output = "MEAS?" + vbCr
        'MSComm2.Output = "MEAS:DC?" + vbCr
    End If
End Sub

'This routine reads the RH span offset entered by the user in the RH offset text box and stores it. $$$$
'Private Sub txtSpanOffset_Change()
 '   SpanOffset = Val(txtSpanOffset.Text)
    'Call SaveSettings
'End Sub


'This routine determines which of 32 LEDs on a cal rack will get turned on
'to indicate Pass/Fail status when calibration is completed.
'It must get called for each of the four pass/fail categories: 2%, 3%, 5%, and FAIL,
'and there are four racks of sensor tip boards, so for a full load of 128 sensor tips,
'this means 4 x 4 = 16 times this routine is called.
'
'The output of this routine is a command string to be sent to the Interface board.
'The Interface board in turn passes this command string on to the sensor tip board
'which then turns on the indicated LEDS.
'
'
'The "passStatus" arguement selects the desired category: PASS 2%, PASS 3%, PASS 5%, or FAIL.
'
'The rackString arguement selects rack "A", "B", "C", or "D"
'
'LED command string syntax is as follows:
'                  >RACK, COMMAND, BANK 1, BANK 2, BANK 3, BANK 4
' COMMAND byte may have following values:
' 0 = LEDS OFF
' 1 = FAIL
' 2 = 2%
' 3 = 3%
' 4 = 5%
' 5 = All pass/fail LEDS on (for testing only)
Function createLEDcommandString(rackString As String, passStatus As String) As String
Dim row As Integer
Dim i As Integer
Dim SensorTip As Integer
Dim bank As Integer
Dim commandString As String
Dim LEDbank As Byte
Dim mask As Byte
Dim sensorStatus As String
Dim command As Integer
Dim sensorString As String
Dim findFirstSensor As Boolean
       
    If (passStatus = "PASS 2%") Then
        command = TWO_PERCENT_LED
    ElseIf (passStatus = "PASS 3%") Then
        command = THREE_PERCENT_LED
    ElseIf (passStatus = "PASS 5%") Then
        command = FIVE_PERCENT_LED
    Else
        command = FAIL_LED
    End If
    
    commandString = rackString + " LEDS " + Format$(command)
    
    
    'Find the row for the first sensor in the rack
    'This should be for sensor #1, 33, 65, or 97
    'If we don't end up with one of those four valid values,
    'then something is wrong!
    findFirstSensor = False
    i = 1
    Do
        row = i + ROW_OFFSET
        sensorString = excel_app.Cells(row, SENSOR_COLUMN).value
        If (InStr(1, sensorString, rackString) > 0) Then
            findFirstSensor = True
        Else
            i = i + 32
        End If
    Loop While (findFirstSensor = False) And (i < 128)
    
    If (findFirstSensor = False) Then
        createLEDcommandString = "ERROR"
    ElseIf (command = OFF_LED) Then
        commandString = commandString + " 0 0 0 0 0"  'Five zeros = OFF command plus all four LED banks off.
        createLEDcommandString = commandString
    Else
        For bank = 1 To 4
            LEDbank = 0
            mask = &H80
            For SensorTip = 1 To 8
                sensorStatus = excel_app.Cells(row, STATUS_COLUMN).value
                If (sensorStatus = passStatus) Then
                    LEDbank = LEDbank Or mask
                End If
            
                'Special case: NO sensor is treated as a failure here since LED is turned on for empty slots as well
                If (passStatus = "FAIL") And (sensorStatus = "NONE") Then
                    LEDbank = LEDbank Or mask
                End If
            
                mask = mask / 2
                row = row + 1
            Next SensorTip
            commandString = commandString + Str$(LEDbank) + " "
        Next bank
        createLEDcommandString = commandString
    End If
    
End Function


' This routine sends a command to the rack designated by "rackString"
' to turn on the LEDS indicating the pass/fail status of sensor tips
' on that rack. rackStr should be ">A", ">B", ">C", or ">D"
' The "passFailStatus" input is also a string set to "FAIL","PASS 2%", etc.
Private Sub turnOnLEDS(rackString As String, passFailStatus As String)
Dim commandString As String
Dim result As Integer

    cmdSetLEDS.Caption = "Please wait..."
    
    ' First turn off LEDS on any racks that may already have LEDS turned on:
    If (chkRack(RACK_A).value = 1) Then Call SendReceiveInterfaceBoard(">A LEDS 0 0 0 0 0")
    If (chkRack(RACK_B).value = 1) Then Call SendReceiveInterfaceBoard(">B LEDS 0 0 0 0 0")
    If (chkRack(RACK_C).value = 1) Then Call SendReceiveInterfaceBoard(">C LEDS 0 0 0 0 0")
    If (chkRack(RACK_D).value = 1) Then Call SendReceiveInterfaceBoard(">D LEDS 0 0 0 0 0")
        
    'If we are turning on LEDs, do it now:
    If (passFailStatus = "TEST") Then
        commandString = rackString + " LEDS 5 255 255 255 255"
        SendReceiveInterfaceBoard (commandString)
    ElseIf (passFailStatus <> "OFF") Then
        commandString = createLEDcommandString(rackString, passFailStatus)
    
        If (commandString = "ERROR") Then
            result = MsgBox("An error calculating pass/fail LEDS has occured." + vbCrLf + "Check spreadsheet for pass/fail data", vbOK + vbCritical, "ERROR - LED CALCULATION")
        Else
            SendReceiveInterfaceBoard (commandString)
        End If
    End If
    
End Sub

'This command button appears at completion of calibration.
'It is clicked on to step through the pass/fail LED sequence.
Private Sub cmdSetLEDS_Click()
    cmdSetLEDS.Enabled = False
    Call setCalCompleteScreen
End Sub

'This routine steps through the Pass/Fail LED sequence, and turns on the rack LEDS
'to indicate the passing sensor tips in each category, FAIL, 2%,3%,and 5%.
'The global CalCompleteIndex keeps track of which pass/fail category
'to indicate next.
'Modified for version 8.0: power to sensor tip boards is shut off using optOffMode_Click
Private Sub setCalCompleteScreen()
Dim rackString As String
Dim rackUsed As Boolean

Const RED = &HFF&
Const YELLOW = &HFFFF&
Const GREEN = &HFF00&
Const NO_COLOR = &HFFFFFF
Const BLUE = &HFFFF00
    

    If (rackLedIndex = RACK_A) Then
        rackString = ">A"
    ElseIf (rackLedIndex = RACK_B) Then
        rackString = ">B"
    ElseIf (rackLedIndex = RACK_C) Then
        rackString = ">C"
    ElseIf (rackLedIndex = RACK_D) Then
        rackString = ">D"
    Else
        rackLedIndex = RACK_A
        rackString = ""
    End If


    If (ExcelCheck() = True) And (dataFilename <> "") Then
        Call optOffMode_Click
        'Call optNormalMode_Click
        If (CalCompleteIndex = 0) Then
            cmdSetLEDS.Caption = "Calibration Complete. Click here to remove FAILED units."
            cmdSetLEDS.BackColor = NO_COLOR
            rackLedIndex = 0
            CalCompleteIndex = 1
            cmdSetLEDS.Enabled = True
        ElseIf (CalCompleteIndex = 1) Then
            Call turnOnLEDS(rackString, "FAIL")
            cmdSetLEDS.Caption = "Remove FAILED units from rack " + rackString + " now. Then click here for next batch."
            cmdSetLEDS.BackColor = RED
            cmdSetLEDS.Enabled = True
        ElseIf (CalCompleteIndex = 2) Then
            Call turnOnLEDS(rackString, "PASS 2%")
            cmdSetLEDS.Caption = "Remove 2% units from rack " + rackString + " now. Then click here for next batch."
            cmdSetLEDS.BackColor = GREEN
            cmdSetLEDS.Enabled = True
        ElseIf (CalCompleteIndex = 3) Then
            Call turnOnLEDS(rackString, "PASS 3%")
            cmdSetLEDS.Caption = "Remove 3% units from rack " + rackString + " now. Then click here for next batch."
            cmdSetLEDS.BackColor = YELLOW
            cmdSetLEDS.Enabled = True
        ElseIf (CalCompleteIndex = 4) Then
            Call turnOnLEDS(rackString, "PASS 5%")
            cmdSetLEDS.Caption = "Remove 5% units from rack " + rackString + " now. Then click here for next batch."
            cmdSetLEDS.BackColor = BLUE
            cmdSetLEDS.Enabled = True
        Else
            'Turn off all LEDS:
            cmdSetLEDS.BackColor = NO_COLOR
            Call turnOnLEDS("", "OFF")
            CalCompleteIndex = 0
            rackLedIndex = 0
            Call mnuNormalMode_Click
            'Call optOffMode_Click
            cmdSetLEDS.Enabled = True
        End If
    Else
        cmdSetLEDS.Caption = "No Calibration Spreadsheet loaded."
        cmdSetLEDS.Enabled = True
    End If
    
    'Before turning on the next batch of LEDs,
    'we need to step through the rack check boxes
    'to see which racks are being used.
    'The LEDs will then be turned on one rack at a time.
    'When the user presses the command button, the next statement steps
    'to the next rack. If there are no more racks,
    'then increment the CalCompleteIndex to move to the next
    'pass/fail category.
    If (rackLedIndex < 0) Then rackLedIndex = 0
    rackUsed = False
    While (rackUsed = False) And (CalCompleteIndex < 5)
        rackLedIndex = rackLedIndex + 1
        If rackLedIndex > RACK_D Then
            CalCompleteIndex = CalCompleteIndex + 1
            rackLedIndex = RACK_A
        End If
        If (chkRack(rackLedIndex).value = 1) Then rackUsed = True
    Wend
    
    
End Sub

'This routine gets called when the LED TEST command button is pressed.
'It turns on all the LEDs on one rack at a time.
'The LEDtestIndex determines which rack is being tested,
'and is incremented each time this button is pushed.
'The chkRack checkboxes are polled to see whether the Index
'is selecting a rack that is plugged in and being used.
'If it isn't being used, then the Index jumps ahead to the next
'checked rack. If there are no more checked racks,
'then the OFF command is sent to shut off all the racks.
Private Sub cmdLEDtest_Click()
Dim LEDcommandString As String
Dim rackBeingUsed As Boolean
Dim rackString As String

    If (LEDtestIndex < 1) Then LEDtestIndex = 1
    If (LEDtestIndex > 5) Then LEDtestIndex = 5
    If (cmdLEDtest.Caption = "LED TEST OFF") Then LEDtestIndex = 1
    
    rackBeingUsed = False
         
    While (rackBeingUsed = False) And LEDtestIndex < 5
        If (frmMain.chkRack(LEDtestIndex).value = 1) Then
            rackBeingUsed = True
        Else
            LEDtestIndex = LEDtestIndex + 1
        End If
    Wend
        
    If (LEDtestIndex = 1) Then
        rackString = ">A"
        cmdLEDtest.Caption = "TESTING RACK A"
    ElseIf (LEDtestIndex = 2) Then
        rackString = ">B"
        cmdLEDtest.Caption = "TESTING RACK B"
    ElseIf (LEDtestIndex = 3) Then
        rackString = ">C"
        cmdLEDtest.Caption = "TESTING RACK C"
    ElseIf (LEDtestIndex = 4) Then
        rackString = ">D"
        cmdLEDtest.Caption = "TESTING RACK D"
    Else
        rackString = ""
        cmdLEDtest.Caption = "LED TEST OFF"
    End If
            
    If (rackString <> "") Then
        Call turnOnLEDS(rackString, "TEST")
    Else
        Call turnOnLEDS(rackString, "OFF")
    End If
    
    LEDtestIndex = LEDtestIndex + 1
End Sub

        

'This routine sends a command string to the Thunder
'approximately every two seconds. Generally this "thunderCommandString"
'is set to THUNDER_READ_RH_AND_T to prompt the Thunder to update the reference RH.
'The MSComm4_OnComm() interrupt receives the serial responses from the Thunder.
'Each time an incoming character is received, the Timer3Timeout is reset to 0.
'If nothing is received, then this routine times out in about 20 seconds.
'If a calibration process is in progress, it is halted, and an error message box pops up.

Private Sub Timer3_Timer()
Dim result As Integer

    Timer3Counter = Timer3Counter + 1
    If (Timer3Counter > 100) Then Timer3Counter = 0
    
    Timer3Timeout = Timer3Timeout + 1
    If (Timer3Timeout > RETRIES) Then
        Timer3Timeout = 0
        If ((RunFlag = True) And (ChamberType = THUNDER) And (thunderCommandString = THUNDER_READ_RH_AND_T)) Then
            result = MsgBox("Make sure the Thunder serial cable is connected. Try again?", vbYesNo + vbCritical + vbDefaultButton1, "Thunder Not Responding")
            If (result = vbNo) Then
                RunFlag = False
                result = MsgBox("Press RESUME or START to continue calibration again.", vbOKOnly, "Calibration halted.")
                frmMain.lblStatusBar.Caption = "Calibration halted."
            End If
        End If
    End If
    
    If (MSComm4.PortOpen = True) Then
        If (ChamberType = THUNDER) Then
            MSComm4.Output = thunderCommandString + vbCr
        End If
    End If

End Sub

'This routine checks serial com port communication
'with the Humilab or Thunder Control output. It returns true if data is received
Function checkChamberControlCommunication() As Boolean
'Dim RHsetpoint As Integer
Dim i As Integer
Dim setpointString As String
Dim receivedString As String
Dim loopCount As Integer
Dim verifyFlag As Boolean

    RHsetpoint = 20
    
    verifiedRHsetpoint = -1#                  '$$$$
    verifiedTemperatureSetpoint = -1#
    ChamberControlUARTflag = False
    verifyFlag = False
       
    i = 0
    
    Do
        If (ChamberType = HUMILAB) Then
            frmMain.lblStatusBar.Caption = "Checking Humilab control communication. Try #" + Format$(i)
            lbl_RH_setpoint.Caption = "Setpoint = "
            setpointString = "*+++" + Format$(RHsetpoint)
            receivedString = SendReceiveHumilabControlCOM(setpointString)
            verifiedRHsetpoint = CInt(Val(receivedString))
            If (verifiedRHsetpoint = RHsetpoint) Then
                verifyFlag = True
            Else
                verifyFlag = False
            End If
        Else
            frmMain.lblStatusBar.Caption = "Sending Thunder RUN command. Try #" + Format$(i)
            
            'Send the startup command to activate the Thunder.
            'The ChamberControlUARTflag will be set to TRUE if command is successful
            'and Thunder responds:
            thunderCommandString = THUNDER_RUN_COMMAND
            Timer3Counter = 0
            Do
                DoEvents
            Loop While (Timer3Counter < 2)
            
            'Send temperature setpoint string to Thunder:
            frmMain.lblStatusBar.Caption = "Sending Thunder TEMPERATURE SETPOINT command. Try #" + Format$(i)
            thunderCommandString = "TS=" + Format$(temperatureSetpoint)
            Timer3Counter = 0
            Do
                DoEvents
            Loop While (Timer3Counter < 2)
            
            'CHANGE: We check the temperature setpoint only:
            'Send RH setpoint string to Thunder:
'            If (thunderMode = RhPc) Then
'                thunderCommandString = "R1=" + Format$(RHsetpoint)
'                frmMain.lblStatusBar.Caption = "Setting Rh@Pc to: " + Format$(RHsetpoint) + " %RH, " + "Trial #" + Format$(i)
'            Else
'                thunderCommandString = "R2=" + Format$(RHsetpoint)
'                frmMain.lblStatusBar.Caption = "Setting Rh@PcTc to: " + Format$(RHsetpoint) + " %RH, " + "Trial #" + Format$(i)
'            End If
            
'            Timer3Counter = 0
'            Do
'                DoEvents
'            Loop While (Timer3Counter < 2)
                        
            'Now read back RH and temperature setpoints:
            thunderCommandString = THUNDER_READ_SETPOINTS
            frmMain.lblStatusBar.Caption = "Sending Thunder READ SETPOINT command. Try #" + Format$(i)
            Timer3Counter = 0
            Do
                DoEvents
            Loop While (Timer3Counter < 2)
            
            'Make sure a valid Temperature setpoint string has been read back:
            If (verifiedTemperatureSetpoint = temperatureSetpoint) Then
                verifyFlag = True
            Else
                verifyFlag = False
            End If
        End If
        i = i + 1
    Loop While (i < RETRIES) And (verifyFlag = False)
    
    thunderCommandString = THUNDER_READ_RH_AND_T
    checkChamberControlCommunication = verifyFlag
    
End Function




'This routine pauses the cal process for a user-specified period
'to allow the chamber to reach a temperature setpoint.
'Nothing very exciting happens during this period, except that
'the temperature gets read and displayed.

'Elapsed time is computed using Timer(),
'which returns time in seconds since previous midnight.
'
'StartTime divides this value by 60 to get starting minutes
'at the beginning of the routine
'when the chamber is set to the new setpoint.
'DiffTime is the difference between the start and current time.
'elapsedTime takes into account the rollover which occurs
'if the chamber runs past midnight, in which case the elapsed time
'up until that point is copied to OffsetTime.

Sub waitForChamberToReachTemperature()

Const AllowableTempSetpointError = 10
Dim intResult As Integer
Dim setpointString As String
Dim receivedString As String
Dim setpointFlag As Boolean

Dim displayString As String
Dim currentTime As Variant
Dim intStartTime, intDiffTime, intElapsedTime, intOffsetTime, intPreviousTime As Integer
Dim i As Integer
Dim j As Integer
Dim intChamberTc As Integer

    'Progress bar will reach end when temperatureWaitTime has elapsed
    barProgress.value = 0
    barProgress.Max = temperatureWaitTime
    Delay (100)
    
    i = 0
    displayString = " "
    'Keep sending setpoint to chamber control until it returns it:
    Do
        'Send temperature setpoint string to Thunder:
        frmMain.lblStatusBar.Caption = "Setting temperature to: " + Format$(temperatureSetpoint) + " degrees C, " + "Trial #" + Format$(i)
        thunderCommandString = "TS=" + Format$(temperatureSetpoint)
        Timer3Counter = 0
        Do
            DoEvents
        Loop While (Timer3Counter < 2)
                        
        'Now read back setpoint:
        thunderCommandString = THUNDER_READ_SETPOINTS
        Timer3Counter = 0
        Do
            DoEvents
        Loop While (Timer3Counter < 3)
    
        i = i + 1
    Loop While ((i <= RETRIES) And (verifiedTemperatureSetpoint <> temperatureSetpoint))
    
    'If (verifySetpoint <> setpoint) Then
    '    intResult = MsgBox("Check chamber setpoint." + vbCr + "Is it " + Format$(setpoint) + "?", vbYesNo)
    '    If (intResult = vbNo) Then
    '        MsgBox ("Calibration halted." + vbCr + "Check chamber control COM port.")
    '        RunFlag = False
    '    End If
    'End If
    
    thunderCommandString = THUNDER_READ_RH_AND_T
    
    setpointFlag = False
    If (RunFlag = True) Then
        'StartTime is the time in minutes when chamber
        'is set to the new setpoint:
        currentTime = Timer()
        intStartTime = CInt(currentTime / 60#)
        intOffsetTime = 0
        intPreviousTime = 0
        intElapsedTime = 0
        frmMain.lblStatusBar.Caption = "Waiting for chamber to reach " + Format$(temperatureSetpoint) + " degrees C,  Time: " + Format$(intElapsedTime) + " minutes"
        Do
            DoEvents
            'DiffTime is the elapsed time since chamber is set to new setpoint,
            'assuming that we haven't just passed midnight:
            currentTime = Timer()
            intDiffTime = (CInt(currentTime / 60#)) - intStartTime
            
            'If DiffTime is negative, then we must have just passed midnight.
            'So we need to store the elapsed time thus far
            'and record a new StartTime.
            'Offset time is now the time elapsed before midnight,
            'DiffTime will be the time elapsed after midnight, and
            'the total ElapsedTime will be the sum of the two:
            If (intDiffTime < 0) Then
                    intOffsetTime = intElapsedTime
                    currentTime = Timer()
                    intStartTime = CInt(currentTime / 60#)
                    intDiffTime = 0
            End If
            
            intElapsedTime = intDiffTime + intOffsetTime
            'If another minute has just passed by, update time displayed:
            If (intElapsedTime <> intPreviousTime) Then
                intPreviousTime = intElapsedTime
                frmMain.lblStatusBar.Caption = "Waiting for chamber to reach " + Format$(temperatureSetpoint) + " degrees C,  Time: " + Format$(intElapsedTime) + " minutes"
                If (intElapsedTime < barProgress.Max) Then
                    barProgress.value = intElapsedTime
                End If
      
                If (intElapsedTime >= temperatureWaitTime) Then
                    intChamberTc = CInt(ChamberTempC)
                    If (Abs(temperatureSetpoint - intChamberTc) > AllowableTempSetpointError) Then
                        frmError.Show
                        frmError.lblLabelOne.Caption = "This chamber did not reach " + Format$(temperatureSetpoint) + " degrees C within " + Format$(temperatureWaitTime) + " minutes."
                        frmError.lblLabelTwo.Caption = "At " + Format$(intElapsedTime) + " minutes, the Temperature = " + Format$(ChamberTempC, "##.#") + " C."
                    Else
                       setpointFlag = True
                    End If
                End If
            End If
            DoEvents
            
        Loop While ((setpointFlag = False) And (RunFlag = True))
    End If
End Sub

Private Sub txtTemperatureSetpoint_Change()
    temperatureSetpoint = Val(txtTemperatureSetpoint.Text)
End Sub

Private Sub txtTemperatureWaitTime_Change()
    temperatureWaitTime = Val(txtTemperatureWaitTime.Text)
End Sub

Private Sub mnuSetFileFolder_Click()
    Call frmFileFolder.Show
    Call frmFileFolder.setupFileFolder
End Sub

Private Sub mnuSetPassLimits_Click()
    Call frmSetPassLimits.Form_Load
    frmSetPassLimits.Show
End Sub


' This routine checks the accuracy for a single sensor tip,
' by scanning across the row and checking the calibration and validation data columns.

' If completeFlag is TRUE, then all the data columns are checked,
' including the calibration tests at 20% and 80%,
' and all four validation points at 90%, 50%, 10%, and 50%.
' Otherwise, if completeFlag is false, the last column at 50% is not checked.

' Pass/Fail limits are stored in the arrPassLimits() array in the following order:
' 10%, 20%, 50%, 80%, 90%
'
' The percent error columns store the error data in the following order:
' ERR1_COLUMN: 10% RH
' ERR2_COLUMN: 50% RH
' ERR3_COLUMN: 90% RH
Function checkStatus(row As Integer) As String
Dim column As Integer
Dim lastColumn As Integer
Dim statusString As String
Dim passFailBin As Integer
Dim percentError As Double

Const TWO_PERCENT_BIN = 1
Const THREE_PERCENT_BIN = 2
Const FIVE_PERCENT_BIN = 3
Const FAIL_BIN = 4

    lastColumn = ERR3_COLUMN
    
    'Start with first cal data column, and assume unit is "OK"
    'If we encounter FAIL status at any point, we jump out of loop
    'without checking the rest of the columns:
    column = ERR1_COLUMN
    passFailBin = TWO_PERCENT_BIN
    Do
        'Check for valid data and make sure unit hasn't been marked as failed:
        statusString = excel_app.Cells(row, STATUS_COLUMN).value
        
        If (statusString <> "OK") Then '$$$$
            passFailBin = FAIL_BIN
        End If
        
        If (passFailBin < FAIL_BIN) Then
            percentError = excel_app.Cells(row, column).value
            
            If (passFailBin = TWO_PERCENT_BIN) Then
                If (percentError > 2# Or percentError < -2#) Then passFailBin = passFailBin + 1
            End If
            
            If (passFailBin = THREE_PERCENT_BIN) Then
                If (percentError > 3# Or percentError < -3#) Then passFailBin = passFailBin + 1
            End If
            
            If (passFailBin = FIVE_PERCENT_BIN) Then
                If (percentError > 5# Or percentError < -5#) Then passFailBin = passFailBin + 1
            End If
        
        End If
        column = column + 4
        
    Loop While (column <= lastColumn) And (passFailBin < FAIL_BIN)
    
    checkStatus = "FAIL" 'Assume worst case
    
    'Now that testing is complete and all cal and validation points have been checked,
    'we mark unit status as 2%, 3%, 5%, or FAIL.
    'Otherwise, if testing isn't complete, then we mark it as either "OK" or "FAIL":
     If (passFailBin = TWO_PERCENT_BIN) Then checkStatus = "PASS 2%"
     If (passFailBin = THREE_PERCENT_BIN) Then checkStatus = "PASS 3%"
     If (passFailBin = FIVE_PERCENT_BIN) Then checkStatus = "PASS 5%"

End Function


Public Sub shutDownCompressor()
Dim command As String
    command = ">X COMPRESSOR_OFF"
    cmdCompressor.Caption = "Turn Compressor On"
    compressorFlag = False
    Call SendReceiveInterfaceBoard(command)
End Sub

Public Sub turnOnCompressor()
Dim command As String
    command = ">X COMPRESSOR_ON"
    cmdCompressor.Caption = "Turn Compressor Off"
    compressorFlag = True
    Call SendReceiveInterfaceBoard(command)
End Sub

Private Sub cmdCompressor_Click()
    If (compressorFlag = True) Then
        compressorFlag = False
        cmdCompressor.Caption = "Turn Compressor On"
        Call shutDownCompressor
    Else
        compressorFlag = True
        cmdCompressor.Caption = "Turn Compressor Off"
        Call turnOnCompressor
    End If
End Sub

'turnOnCompressor()
'This routine checks communication with the interface board
Function interfaceComTest()
Dim responseString As String
Dim commandString As String
    interfaceComTest = False
    commandString = ">X COM"
    Delay (100)
    If (SendReceiveInterfaceBoard(commandString) = True) Then
            responseString = Intext
            If (InStr(1, responseString, "COM PORT WORKS")) Then
                interfaceComTest = True
            End If
    End If
End Function


'This routine calculates pass/fail statistics and records them on the spreadsheet.
'It is called at end of process after determining pass/fail status.
Private Sub addPassFailText()
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
    
    If (ExcelCheck() = True) Then
        With excel_app
            For i = 1 To maxSensors
                row = i + OFFSET
                If excel_app.Cells(row, SENSOR_USED_COLUMN).value = True Then
                    totalSensors = totalSensors + 1
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
        
            row = OFFSET + maxSensors + 2
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


'This routine starts up Excel and creates a new blank spreadsheet.
'It then adds the column titles, and saves the result to the filename
'defined by DEFAULT_SHEETNAME, which is "Current Cal.xls"
'Before doing all this it checks Excel to make sure that it isn't
'already running with a speadsheet of the same name.
'Otherwise Excel would create the new spreadsheet with a different name like "Book1"
'
'If Excel is already running, then the user is prompted to save and close Excel.
'While this is not the most elegant solution to the problem,
'it is hopefuly foolproof.
'
'The resulting spreadsheet must remain open throughout the cal process.
'At the end of the process it will be renamed with cal date and run number.

Private Sub createNewSpreadsheet()
Dim ColumnTitle(1 To 32) As String
Dim ColumnWidth(1 To 32) As Integer
Dim sheet As Object
Dim currentDateTime As String
Dim SensorTipNumber As Integer
Dim setpointColumn As Integer
Dim setpoint(1 To 3) As Integer
Dim i As Integer
Dim result As Integer
Dim Color As Integer


frmMain.lblStatusBar.Caption = "Preparing spreadhseet. Please wait..."

ColumnTitle(1) = "Sensor"
ColumnWidth(1) = Len("Sensor")

ColumnTitle(2) = "Comments"
ColumnWidth(2) = Len("Comments")

ColumnTitle(3) = "Status"
ColumnWidth(3) = Len("Status    ") 'Allow a little extra room for comments, ie: "FAIL POT ERROR", etc.

ColumnTitle(4) = BLANK          '10% RH
ColumnWidth(4) = Len(BLANK)
ColumnTitle(5) = "Ref"
ColumnWidth(5) = REF_WIDTH
ColumnTitle(6) = "UUT"
ColumnWidth(6) = UUT_WIDTH
ColumnTitle(7) = "Error"
ColumnWidth(7) = ERR_WIDTH

ColumnTitle(8) = BLANK          '50% RH
ColumnWidth(8) = Len(BLANK)
ColumnTitle(9) = "Ref"
ColumnWidth(9) = REF_WIDTH
ColumnTitle(10) = "UUT"
ColumnWidth(10) = UUT_WIDTH
ColumnTitle(11) = "Error"
ColumnWidth(11) = ERR_WIDTH

ColumnTitle(12) = BLANK         '90% RH
ColumnWidth(12) = Len(BLANK)
ColumnTitle(13) = "Ref"
ColumnWidth(13) = REF_WIDTH
ColumnTitle(14) = "UUT"
ColumnWidth(14) = UUT_WIDTH
ColumnTitle(15) = "Error"
ColumnWidth(15) = ERR_WIDTH

ColumnTitle(16) = BLANK
ColumnWidth(16) = Len(BLANK)

ColumnTitle(17) = "Used"
ColumnWidth(17) = Len("Used")
ColumnTitle(18) = "# Loops"
ColumnWidth(18) = Len("# Loops")
ColumnTitle(19) = "Repeats?"
ColumnWidth(19) = Len("Repeats?")

    'Set output filename to default and delete existing file by that name:
    dataFilename = strLocalFolder + DEFAULT_SHEETNAME
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

        'Fifth row is the racks used:
        row = RACK_ROW
        .Cells(row, 1).Select
        .ActiveCell.FormulaR1C1 = "Racks Used: "
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
        
        'Sixth row is Vout Range:
        row = 6
        .Cells(row, 1).Select
        .ActiveCell.FormulaR1C1 = "Vout Range: "
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
        
        
        'Seventh row is additional notes:
        row = 7
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
        For i = 1 To 32
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
        'for each setpoint ie: "Val 1: 10%" etc.
        'The first setpoint column begins at the REFERENCE 1 COLUMN
        row = SETPOINT_ROW
        
        'These are the three validation setpoints:
        setpoint(1) = 10
        setpoint(2) = 50
        setpoint(3) = 90
        setpointColumn = REF1_COLUMN
        For i = 1 To 3
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
                        
        'Initialize TASK cell and task list index with starting task = 0
        excel_app.Cells(4, 3) = 0
    End With

    frmMain.lblStatusBar.Caption = "Spreadsheet ready"
    
    

Quit:
End Sub

'This initializes the Task Box and scrollbar.
'The TaskIndex is set to a default of 0.
'For Version 10, eliminated #2-7 and #9-12
Private Sub SetUpTaskBox()
Dim i As Integer

    lstTasks.Height = DEFAULT_HEIGHT

    lstTasks.AddItem "[0] Com Port Check"
    lstTasks.AddItem "[1] Check Sensor Racks"
    lstTasks.AddItem "[2] Set chamber to 10% RH" '$$$$
    lstTasks.AddItem "[3] Check units at 10% RH"
    lstTasks.AddItem "[4] Set chamber to 50% RH"
    lstTasks.AddItem "[5] Check units at 50% RH"
    lstTasks.AddItem "[6] Set chamber to 90% RH"
    lstTasks.AddItem "[7] Check units at 90% RH"
    lstTasks.AddItem "[8] Set chamber to 30% RH"
    lstTasks.AddItem "[9] Check off failed units"
    lstTasks.AddItem "[10] Save Pass/Fail data."
    lstTasks.AddItem "[11] Turn on Pass/Fail LEDs."
    lstTasks.AddItem "[12] Shut down chamber."
    
    MaxTask = lstTasks.ListCount - 1
    scrTasks.Max = MaxTask
    scrTasks.Min = 0
    For i = 0 To MaxTask
            lstTasks.Selected(i) = True 'For Version 10, check all tasks
    Next i
    lstTasks.ListIndex = 0
    TaskIndex = 0
    scrTasks.value = TaskIndex
End Sub

'This routine refereshes the data displayed in the grid by copying the entire spreadsheet
'The number of rows equals the number of sensor sockets plus two column title rows.
Public Sub copySpreadsheetToGrid()
Dim spreadsheetRow As Integer
Dim i As Integer
Dim column As Integer

    If (ExcelCheck() = True) And (dataFilename <> "") Then
        For i = -1 To maxSensors
            spreadsheetRow = i + ROW_OFFSET
            For column = 0 To (CAL_TEST_COLUMN - 1)
                grdSpreadsheet.Col = column
                grdSpreadsheet.row = i + 1
                grdSpreadsheet.Text = excel_app.Cells(spreadsheetRow, column + 1).value
            Next column
        Next i
    End If
End Sub


' This routine calculates the percent RH from the measured voltage
Function calculateRh(measuredVoltage As Double)
    calculateRh = ((measuredVoltage - OFFSET_VOLTAGE) / SPAN_VOLTAGE) * 100
End Function

Public Sub initializeSensorsUsed() '$$$$
Dim row As Integer
Dim i As Integer

    Delay (4000)
    i = 1
    DoEvents
    barProgress.value = 0
    barProgress.Max = maxSensors
    Do
            barProgress.value = i
            DoEvents
            row = i + ROW_OFFSET
            excel_app.Cells(row, STATUS_COLUMN).value = "OK"
            excel_app.Cells(row, SENSOR_USED_COLUMN).value = True
        DoEvents
    i = i + 1
    Loop While (i <= maxSensors) And (RunFlag = True)
End Sub


'This routine takes voltage readings from the Unit Under Test and averages them.
'The average voltage is then used to calculate the measured UUT RH.
'
'The numberOfReadings determines how many readings
'are included in each average. If a timeout occurs
'then a message box pops on and the calibration process is paused
'until the user corrects the problem.
'This would occur for example if the voltmeter
'became unplugged or the COM port were closed.
'
'The first reading is rejected to insure integrity of the measurements.
Function measure_UUT_RH(intNumberOfReadings As Integer) As Double
Const MAXREADINGS = 16
Const VOLTMETER_TIMEOUT = 15
Dim dblSum As Double
Dim i As Integer
Dim dblNumberOfReadings As Double
Dim dblMeasuredVoltage As Double
Dim dblAverageRH As Double
Dim startTime As Variant
Dim elapsedTime As Variant
Dim ElapsedSeconds As Long
Dim result As Integer

    measure_UUT_RH = 0#
    If (intNumberOfReadings > MAXREADINGS) Then intNumberOfReadings = MAXREADINGS
    If (intNumberOfReadings < 1) Then intNumberOfReadings = 1
    
    dblNumberOfReadings = CDbl(intNumberOfReadings) 'Convert to double
    dblSum = 0
    
    i = 0
    Do
        VoltUARTflag = False
        startTime = Timer()
        Do
            DoEvents
            elapsedTime = Timer() - startTime
            ElapsedSeconds = CLng(elapsedTime)
            'Something weird just happened if elapsed time is negative!
            'Maybe midnight just reset Timer(). Clear time and start again:
            If (ElapsedSeconds < 0) Then
                startTime = Timer()
                elapsedTime = Timer() - startTime
                ElapsedSeconds = CLng(elapsedTime)
            End If
            DoEvents
        Loop While (VoltUARTflag = False) And (ElapsedSeconds < VOLTMETER_TIMEOUT)
        
        If (ElapsedSeconds = VOLTMETER_TIMEOUT) Then
            result = MsgBox("Make sure voltmeter is turned on. Try again?", vbYesNo + vbCritical + vbDefaultButton1, "Voltmeter Not Responding")
            If (result = vbNo) Then
                RunFlag = False
            Else
                VoltUARTflag = False
                startTime = Timer()
                i = 0
                dblSum = 0
            End If
            dblSum = 0
        Else
            'Reject first reading when i=0,
            'in case measurement is left over from previous device.
            'Otherwise, take reading and add it to runnning sum:
            If (i > 0) Then dblSum = dblSum + voltage
            i = i + 1
        End If
    Loop While (i <= intNumberOfReadings) And (RunFlag = True)
    
    If (RunFlag = True) Then
        dblMeasuredVoltage = dblSum / dblNumberOfReadings
        If dblMeasuredVoltage < OFFSET_VOLTAGE Then
            measure_UUT_RH = 0
        Else
            dblAverageRH = calculateRh(dblMeasuredVoltage)
            measure_UUT_RH = dblAverageRH
        End If
    End If
    
End Function

'This routine reads the UUT and reference RH and records them
'in the spreadsheet. This gets called once for each validation point
'after calibration.
'
'The input variable firstDataColumn points to the first spreadsheet column
'in which to begin data. This is the Reference column, followed by the
'UUT measured RH column, and finally the Percent Error column.
Private Sub checkUnits(firstDataColumn As Integer)
Dim potValue As Integer
Dim i As Integer
Dim row As Integer
Dim commandString As String
Dim measuredRH As Double
Dim ReferenceRH As Double
Dim percentError As Double
Dim setpointString As String
Dim MainComm As Boolean
Dim commRetries As Integer
 
    barProgress.value = 0
    barProgress.Max = maxSensors
    i = 1
    DoEvents
    
    'Add temperature string to setpoint cell:
    setpointString = excel_app.Cells(SETPOINT_ROW, firstDataColumn).value
    excel_app.Cells(SETPOINT_ROW, firstDataColumn).value = setpointString + lblChamberTemp.Caption
    
    Do
        row = i + ROW_OFFSET
        If excel_app.Cells(row, SENSOR_USED_COLUMN).value = True Then
            barProgress.value = i
            frmMain.lblStatusBar.Caption = "Testing sensor " + excel_app.Cells(row, SENSOR_COLUMN).value
            
            DoEvents
            
            'Send Interface (Main board) the command to set multiplexer for next sensor tip
            'so that voltage can be read. Make sure that Interface board  echoes back
            'the correct command string. If it doesn't, 8 retries are made before aborting cal run:
            commandString = excel_app.Cells(row, SENSOR_COLUMN).value
            MainComm = False
            commRetries = 0
            Do
                If (SendReceiveInterfaceBoard(commandString) = True) Then
                    If (InStr(1, Intext, commandString) > 0) Then
                        MainComm = True
                    End If
                    commRetries = commRetries + 1
                End If
            Loop While ((commRetries < RETRIES) And (MainComm = False)) And (RunFlag = True)
            
            
            If (MainComm = False) Then
                MsgBox ("Serial communication with Interface board has failed." + vbCrLf + "Cal run has been aborted")
                RunFlag = False
            Else
                'Read chamber RH:
                ReferenceRH = ChamberRH
                'Read device under test. Take 4 readings and get average:
                measuredRH = measure_UUT_RH(4)
                If measuredRH = 0 Then
                    excel_app.Cells(row, SENSOR_USED_COLUMN).value = False
                    excel_app.Cells(row, STATUS_COLUMN).value = " "
                Else
                    percentError = measuredRH - ReferenceRH
                    DoEvents
            
                    excel_app.Cells(row, firstDataColumn).NumberFormat = "0.00"
                    excel_app.Cells(row, firstDataColumn).value = Format$(ReferenceRH, "0.00")
            
                    excel_app.Cells(row, firstDataColumn + 1).NumberFormat = "0.00"
                    excel_app.Cells(row, firstDataColumn + 1).value = Format$(measuredRH, "0.00")
            
                    excel_app.Cells(row, firstDataColumn + 2).NumberFormat = "0.00"
                    excel_app.Cells(row, firstDataColumn + 2).value = Format$(percentError, "0.00")
                End If
            End If
        End If
        DoEvents
        i = i + 1
    Loop While (i <= maxSensors) And (RunFlag = True)
End Sub

'This routine is called from the Run() routine above.   $$$$
'It is called each time the task index increments ahead
'to the next task. The "InStr(1, strTask, "[0]") > 0)"
'statements search for the selected task indicated by the
'input string, and call the appropriate routine(s) to
'execute that task. After each task is completed,
'the spreadsheet is saved to preserve the latest data.

Private Sub Execute(strTask As String)
        
    'This checks that the com ports are open and working
    'communication on the com ports
    If (InStr(1, strTask, "[0]") > 0) Then
        frmMain.lblStatusBar.Caption = lstTasks.Text
        Call CheckCommunication
    End If
    
    'Poll sensor racks to see how many are present,
    'then initialize spreadsheet:
    If (InStr(1, strTask, "[1]") > 0) Then
        frmMain.lblStatusBar.Caption = lstTasks.Text
        Call identifySensorRacksAndCopyToSpreadsheet
        Call initializeSensorsUsed
    End If
    
            
    'First validation point is at 10% RH:  '$$$$
    If (InStr(1, strTask, "[2]") > 0) Then
        frmMain.lblStatusBar.Caption = lstTasks.Text
        Call setChamber(10)
    End If
    
    If (InStr(1, strTask, "[3]") > 0) Then
        frmMain.lblStatusBar.Caption = lstTasks.Text
        Call checkUnits(REF1_COLUMN)
    End If


    
    'Second validation point is at 50% RH:
    If (InStr(1, strTask, "[4]") > 0) Then
        frmMain.lblStatusBar.Caption = lstTasks.Text
        Call setChamber(50)
    End If
    
    If (InStr(1, strTask, "[5]") > 0) Then
        frmMain.lblStatusBar.Caption = lstTasks.Text
        Call checkUnits(REF2_COLUMN)
    End If
    
        'Third validation point is at 90% RH:  $$$$
    If (InStr(1, strTask, "[6]") > 0) Then
        frmMain.lblStatusBar.Caption = lstTasks.Text
        Call setChamber(90)
    End If
    
    If (InStr(1, strTask, "[7]") > 0) Then
        frmMain.lblStatusBar.Caption = lstTasks.Text
        Call checkUnits(REF3_COLUMN)
    End If

    'Set chamber to 30% RH"
    If (InStr(1, strTask, "[8]") > 0) Then
        frmMain.lblStatusBar.Caption = lstTasks.Text
        Call setChamber(30)
    End If
    
    'Now determine which units pass and which fail.
    'Look at all four columns of error data,
    'and assign "PASS 2%", "PASS 3%, "PASS 5%, and "FAIL"
    'to the STATUS column for each.
    'Then rename file to today's date and save it.
    'Finally send STOP command to shut down Thunder:
    If (InStr(1, strTask, "[9]") > 0) Then
        frmMain.lblStatusBar.Caption = lstTasks.Text
        Call FinalFailCheck
        If (ExcelCheck() = True) Then Call RenameAndSave
    End If
        
    'Then enable CAL COMPLETE command button
    'so user can click on it to turn
    'on PASS/FAIL LEDS
    If (InStr(1, strTask, "[10]") > 0) Then
        frmMain.lblStatusBar.Caption = lstTasks.Text
        Call DisplayCalCompleteScreen
    End If
    
    'Finally, shut down chamber, if desired:
    If (InStr(1, strTask, "[11]") > 0) Then
        frmMain.lblStatusBar.Caption = lstTasks.Text
        If (ChamberType = THUNDER) Then
            Call stopThunder
            If (chkCompressorShutdown.value = 1) Then Call shutDownCompressor
        End If
    End If
    
    If (ExcelCheck() = True) Then
        With excel_app
            .ActiveWorkbook.Save
        End With
    End If
    
End Sub

'This routine is called at the completion of the calibration process.
'It is calls the finalCheckStatus() routine below
'to check all of the cal and validation data for all the units.
'The end result is that the STATUS column for all USED sockets is marked:
'"PASS 2%", "PASS 3%", "PASS 5%", or "FAIL"
Public Sub FinalFailCheck()
Dim row As Integer
Dim i As Integer
Dim passFailStatus As String
Dim currentDateTime As String

    'Record calibration completion time on spreadsheet:
    row = 3
    currentDateTime = Now
    excel_app.Cells(row, 1).Select
    excel_app.ActiveCell.FormulaR1C1 = "Calibration completion time: " + currentDateTime 'TODO: can this be fixed?
    
    i = 1
    Do
        row = i + ROW_OFFSET
        'Now check accuracy for all units. Record results in STATUS column
        If excel_app.Cells(row, SENSOR_USED_COLUMN).value = True Then
            passFailStatus = checkStatus(row)
            DoEvents
            
            excel_app.Cells(row, STATUS_COLUMN).value = passFailStatus
            
            'Now set the cell backround color to match the pass/fail status:
            passColor = NO_COLOR 'Default background color for cells is white
            If (passFailStatus = "PASS 2%") Then
                passColor = TWO_PERCENT_COLOR
            ElseIf (passFailStatus = "PASS 3%") Then
                passColor = THREE_PERCENT_COLOR
            ElseIf (passFailStatus = "PASS 5%") Then
                passColor = FIVE_PERCENT_COLOR
            ElseIf (passFailStatus = "FAIL") Then
                passColor = FAIL_COLOR
            End If
            
            excel_app.Cells(row, STATUS_COLUMN).Interior.ColorIndex = passColor
'            End With
'           excel_app.Cells(row, STATUS_COLUMN).value = passFailStatus
            
            DoEvents
        'At this point, we blank out data in cells with no sensors in them:
'        Else
            'excel_app.Cells(i + 9, STATUS_COLUMN).value = ""
            'excel_app.Cells(i + 9, POT1_COLUMN).value = ""
            'excel_app.Cells(i + 9, POT2_COLUMN).value = ""
        End If
        DoEvents
        i = i + 1
    Loop While (i <= maxSensors)
    
    Call addPassFailText
End Sub

