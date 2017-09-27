Attribute VB_Name = "GlobalMod"
Global Const CONFIG_FILE = "c:\CalConfig.txt"
Global Const CONFIG_FILENAME = "CalConfig.txt"
Global Const DEFAULT_SHEETNAME = "c:\Cal Data\Current Cal.xls"
Global Const DEFAULT_HEIGHT = 2535
Global Const EXTENDED_HEIGHT = 6435

'These constants set the background colors for the STATUS column
'when the final PASS/FAIL data is written:
Global Const TWO_PERCENT_COLOR = 4     'This is GREEN
Global Const THREE_PERCENT_COLOR = 6   'This is YELLOW
Global Const FIVE_PERCENT_COLOR = 46   'This is ORANGE
Global Const FAIL_COLOR = 3            'This is RED
Global Const NO_COLOR = 2              'This is WHITE - the normal background color for a spreadsheet cell
Global Const BLACK = 1
Global Const BLUE = 5

Global Const BOARD_A = 1
Global Const BOARD_B = 2
Global Const BOARD_C = 3
Global Const BOARD_D = 4

Global Const OFF_LED = 0
Global Const FAIL_LED = 1
Global Const TWO_PERCENT_LED = 2
Global Const THREE_PERCENT_LED = 3
Global Const FIVE_PERCENT_LED = 4
Global Const LED_TEST = 5

Global Const I2C_ERROR = 0
Global Const READY_TO_PROGRAM = 1
Global Const FUSE_BLOWN = 2
Global Const FUSE_BAD = 3
Global Const INVALID_POT_DATA = 4
Global Const COM_ERROR = 5
Global Const POWER_DOWN_TIME = 5

Global Const INCREMENT = 1
Global Const DECREMENT = 0
Global Const RETRIES = 5
Global Const BALANCE_POT = 1
Global Const SPAN_POT = 2
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

Global Const THUNDER_READ_RH = "?"
Global Const THUNDER_READ_SETPOINT = "?SP"
Global Const THUNDER_RUN_COMMAND = "RUN"
Global Const THUNDER_STOP_COMMAND = "STOP"

Public thunderCommandString As String
Public thunderSetpointValue As Integer
Public Timer3Counter As Integer
Public CalCompleteIndex As Integer
Public TestSetpoint As Integer
Public ConfigFileNumber As Integer
Public UARTflag As Boolean
Public VoltUARTflag As Boolean
Public UARTbuffer As String
Public VoltUARTbuffer As String
Public Voltage As Double
Public Intext As String
Public ChamberRH As Double
Public ChamberTempC As Double
Public ChamberRefUARTbuffer As String
Public ChamberRefUARTflag  As Boolean
Public ChamberControlUARTbuffer As String
Public ChamberControlUARTflag  As Boolean
Public Pot1Value As Integer
Public Pot2Value As Integer
Public sensorNumber As Integer
Public previousSensorNumber As Integer
Public MaxTask As Integer
Public RunFlag As Boolean

Public excel_sheet As Object
Public dataFilename As String
Public TaskIndex As Integer
Public SelectFlag As Boolean
Public lngStartTime As Long
Public Timer1Counter As Integer




