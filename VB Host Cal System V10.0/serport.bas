Attribute VB_Name = "GeneralSerialPortRoutines"
Option Explicit
'General routines used by applications that access the serial port.
'Some routines access forms and variables in template.vbp.
'The following constants are from win32api.txt:
'Constants used in DCB access
'
' Revision history:
' 1-27-2008 JBS - Multiple Port version with VoltmeterPort added
' 1-27-2008 JBS - Multiple four Port version with Chamber ports added
' 4-11-2008 JBS - Added ChamberType to GetSettings() and SaveSettings()
'                 Modified checkNewPortNumbers() to ignore ChamberRefPortNumber if Chamber Type is THUNDER
' 6-20-2008 JBS - Modified checkNewPortNumbers() so it doesn't check for duplicates.
'                 chamberRhWaitTime = 60 minutes

Global Const HUMILAB = 1
Global Const THUNDER = 0

'Parity
Global Const NOPARITY = 0
Global Const ODDPARITY = 1
Global Const EVENPARITY = 2
Global Const MARKPARITY = 3
Global Const SPACEPARITY = 4
'Stop bits
Global Const ONESTOPBIT = 0
Global Const ONE5STOPBITS = 1
Global Const TWOSTOPBITS = 2

'Errors
Global Const CE_RXOVER = &H1
Global Const CE_OVERRUN = &H2
Global Const CE_RXPARITY = &H4
Global Const CE_FRAME = &H8
Global Const CE_BREAK = &H10
Global Const CE_CTSTO = &H20
Global Const CE_DSRTO = &H40
Global Const CE_RLSDTO = &H80
Global Const CE_TXFULL = &H100
Global Const CE_PTO = &H200
Global Const CE_IOE = &H400
Global Const CE_DNS = &H800
Global Const CE_OOP = &H1000
Global Const CE_MODE = &H8000

Global Const IE_BADID = (-1)
Global Const IE_OPEN = (-2)
Global Const IE_NOPEN = (-3)
Global Const IE_MEMORY = (-4)
Global Const IE_DEFAULT = (-5)
Global Const IE_HARDWARE = (-10)
Global Const IE_BYTESIZE = (-11)
Global Const IE_BAUDRATE = (-12)

'CommEventMask bits
Global Const EV_RXCHAR = &H1
Global Const EV_RXFLAG = &H2
Global Const EV_TXEMPTY = &H4
Global Const EV_CTS = &H8
Global Const EV_DSR = &H10
Global Const EV_RLSD = &H20
Global Const EV_BREAK = &H40
Global Const EV_ERR = &H80
Global Const EV_RING = &H100
Global Const EV_PERR = &H200
Global Const EV_CTSS = &H400
Global Const EV_DSRS = &H800
Global Const EV_RLSDS = &H1000

'EscapeCommFunction values
Global Const SETXOFF = 1
Global Const SETXON = 2
Global Const SETRTS = 3
Global Const CLRRTS = 4
Global Const SETDTR = 5
Global Const CLRDTR = 6
Global Const RESETDEV = 7
Global Const GETMAXLPT = 8
Global Const GETMAXCOM = 9
Global Const GETBASEIRQ = 10

'Bit rates
Global Const CBR_110 = &HFF10
Global Const CBR_300 = &HFF11
Global Const CBR_600 = &HFF12
Global Const CBR_1200 = &HFF13
Global Const CBR_2400 = &HFF14
Global Const CBR_4800 = &HFF15
Global Const CBR_9600 = &HFF16
Global Const CBR_14400 = &HFF17
Global Const CBR_19200 = &HFF18
Global Const CBR_38400 = &HFF1B
Global Const CBR_56000 = &HFF1F
Global Const CBR_128000 = &HFF23
Global Const CBR_256000 = &HFF27

Global Const CN_RECEIVE = &H1
Global Const CN_TRANSMIT = &H2
Global Const CN_EVENT = &H4
Global Const CSTF_CTSHOLD = &H1
Global Const CSTF_DSRHOLD = &H2
Global Const CSTF_RLSDHOLD = &H4
Global Const CSTF_XOFFHOLD = &H8
Global Const CSTF_XOFFSENT = &H10
Global Const CSTF_EOF = &H20
Global Const CSTF_TXIM = &H40
Global Const LPTx = &H80

'Public Const OPEN_EXISTING = 3

'  DTR Control Flow Values.
Public Const DTR_CONTROL_DISABLE = &H0
Public Const DTR_CONTROL_ENABLE = &H1
Public Const DTR_CONTROL_HANDSHAKE = &H2

'  RTS Control Flow Values
Public Const RTS_CONTROL_DISABLE = &H0
Public Const RTS_CONTROL_ENABLE = &H1
Public Const RTS_CONTROL_HANDSHAKE = &H2
Public Const RTS_CONTROL_TOGGLE = &H3

Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000

'DCB Bits values:
Public Const FLAG_fBinary& = &H1
Public Const FLAG_fParity& = &H2
Public Const FLAG_fOutxCtsFlow = &H4
Public Const FLAG_fOutxDsrFlow = &H8
Public Const FLAG_fDtrControl = &H30
Public Const FLAG_fDsrSensitivity = &H40
Public Const FLAG_fTXContinueOnXoff = &H80
Public Const FLAG_fOutX = &H100
Public Const FLAG_fInX = &H200
Public Const FLAG_fErrorChar = &H400
Public Const FLAG_fNull = &H800
Public Const FLAG_fRtsControl = &H3000
Public Const FLAG_fAbortOnError = &H4000

'End of win32api.txt constants.

Public Type COMMTIMEOUTS
     ReadIntervalTimeout As Long
    ReadTotalTimoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type

Public Type dcbType
        DCBlength As Long
        BaudRate As Long
        Bits1 As Long
        wReserved As Integer
        XonLim As Integer
        XoffLim As Integer
        ByteSize As Byte
        Parity As Byte
        StopBits As Byte
        XonChar As Byte
        XoffChar As Byte
        ErrorChar As Byte
        EofChar As Byte
        EvtChar As Byte
        wReserved2 As Integer
End Type

'Global variables & constants used by the application:

Public Const ProjectName = "HumCal"
Public Const RhPcTc = 0
Public Const RhPc = 1

Public strLimitsFilename As String
Public strNetworkFolder As String
Public strLocalFolder As String
Public chamberRhWaitTime As Integer
Public temperatureSetpoint As Integer
Public temperatureWaitTime As Integer
Public ChamberType As Integer
Public BitRate As Long
Public Buffer As Variant
Public CommDCB As dcbType
Public CommPorts() As String
Public OneByteDelay As Single

Public PortOpen As Boolean
Public SaveDataInFile As Boolean
Public TimedOut As Boolean
Public ValidPort As Boolean

Public PortNumber As Integer
Public VoltmeterPortNumber As Integer
Public ChamberRefPortNumber As Integer
Public ChamberControlPortNumber As Integer

Public thunderMode As Integer

'API declares:
Public Declare Function apiGetCommState _
    Lib "kernel32" _
    Alias "GetCommState" _
    (ByVal nCid As Long, _
    lpDCB As dcbType) _
    As Long
Public Declare Function apiSetCommState _
    Lib "kernel32" _
    Alias "SetCommState" _
    (ByVal hCommDev As Long, _
    lpDCB As dcbType) _
    As Long
Public Declare Function EscapeCommFunction _
    Lib "kernel32" _
    (ByVal nCid As Long, _
    ByVal nFunc As Long) _
    As Long
Public Declare Function GetCommTimeouts _
    Lib "kernel32" _
    (ByVal hFile As Long, _
    lpCommTimeouts As COMMTIMEOUTS) _
    As Long
Public Declare Function SetCommTimeouts _
    Lib "kernel32" _
    (ByVal hFile As Long, _
    lpCommTimeouts As COMMTIMEOUTS) _
    As Long
Public Declare Function timeGetTime _
    Lib "winmm.dll" () _
    As Long
Public Declare Function TransmitCommChar _
    Lib "kernel32" _
    (ByVal nCid As Long, _
    ByVal cChar As Byte) _
    As Long

Public Function fncAddChecksumToAsciiHexString _
    (UserString As String) _
    As String
'Calculates a checksum for a string containing
'a series bytes in Ascii Hex format.
'Places the checksum in Ascii Hex format
'at the end of the string.
Dim Count As Integer
Dim Sum As Long
Dim Checksum As Byte
Dim ChecksumAsAsciiHex As String
'Add the values of each Ascii Hex pair:
For Count = 1 To Len(UserString) - 1 Step 2
    Sum = Sum + Val("&h" & Mid(UserString, Count, 2))
Next Count
'The checksum is the low byte of the sum.
Checksum = Sum - (CInt(Sum / 256)) * 256
ChecksumAsAsciiHex = fncByteToAsciiHex(Checksum)
'Add the checksum to the end of the string.
fncAddChecksumToAsciiHexString = UserString & ChecksumAsAsciiHex
End Function

Public Function fncByteToAsciiHex _
    (ByteToConvert As Byte) _
    As String
'Converts a byte to a 2-character ASCII Hex string
Dim AsciiHex As String
AsciiHex = Hex$(ByteToConvert)
If Len(AsciiHex) = 1 Then
    AsciiHex = "0" & AsciiHex
End If
fncByteToAsciiHex = AsciiHex
End Function

Public Function fncDisplayDateAndTime() As String
'Date and time formatting.
fncDisplayDateAndTime = _
    CStr(Format(Date, "General Date")) & ", " & _
    (Format(Time, "Long Time"))
End Function

Public Function fncGetHighestComPortNumber() As Integer
'Returns the number of the system's highest Com port.
'Also shows how to use the EscapeCommFunction API call.
Dim ClosePortOnExit As Boolean
Dim PortCount As Long
Dim handle As Long
'The API call requires a CommID of an open port.
If frmMain.MSComm1.PortOpen = False Then
    frmMain.MSComm1.PortOpen = True
    ClosePortOnExit = True
Else
    ClosePortOnExit = False
End If
handle = frmMain.MSComm1.CommID
PortCount = GETMAXCOM
'Add 1 because EscapeCommFunction begins counting at 0.
fncGetHighestComPortNumber = _
    EscapeCommFunction(handle, PortCount) + 1
If ClosePortOnExit = True Then
    frmMain.MSComm1.PortOpen = False
End If
End Function

Public Function fncOneByteDelay(BitRate As Long) As Single
'Calculate the time in milliseconds to transmit
'8 bits + 1 Start & 1 Stop bit.
Dim DelayTime As Integer
DelayTime = 10000 / BitRate
fncOneByteDelay = DelayTime
End Function

Public Function fncVerifyChecksum(UserString As String) As Boolean
'Verifies data by comparing a received checksum
'to the calculated value.
'UserString is a series of bytes in Ascii Hex format,
'Ending in a checksum.
Dim Count As Integer
Dim Sum As Long
Dim Checksum As Byte
Dim ChecksumAsAsciiHex As String
'Add the values of each Ascii Hex pair:
For Count = 1 To Len(UserString) - 3 Step 2
    Sum = Sum + Val("&h" & Mid(UserString, Count, 2))
Next Count
'The checksum is the low byte of the sum.
Checksum = Sum - (CInt(Sum / 256)) * 256
ChecksumAsAsciiHex = fncByteToAsciiHex(Checksum)
'Compare the calculated checksum to the received checksum.
If Checksum = Val("&h" & (Right(UserString, 2))) Then
    fncVerifyChecksum = True
Else
    fncVerifyChecksum = False
End If
End Function

Public Sub Delay(DelayInMilliseconds As Single)
'Delay timer with approximately 1-msec. resolution.
'Uses the API function timeGetTime.
'Rolls over 24 days after the last Windows startup.
Dim Timeout As Single
Timeout = DelayInMilliseconds + timeGetTime()
Do Until timeGetTime() >= Timeout
    DoEvents
Loop
End Sub

Public Sub EditDCB()
'Enables changes to a port's DCB.
'The port must be open.
Dim Success As Boolean
Dim PortID As Long
PortID = frmMain.MSComm1.CommID
Success = apiGetCommState(PortID, CommDCB)

'To change a value, uncomment and revise the appropriate line:
'CommDCB.BaudRate = 2400
'CommDCB.Bits1 = &H11
'CommDCB.XonLim = 64
'CommDCB.XoffLim = 64
'CommDCB.ByteSize = 8
'CommDCB.Parity = 0
'CommDCB.StopBits = 0
'CommDCB.XonChar = &H12
'CommDCB.XoffChar = &H13
'CommDCB.ErrorChar = 0
'CommDCB.EofChar = &H1A
'CommDCB.EvtChar = 0

'Write the values to the DCB.
Success = apiSetCommState(PortID, CommDCB)

'Read the values back to verify changes.
Success = apiGetCommState(PortID, CommDCB)

Debug.Print "DCBlength: ", Hex$(CommDCB.DCBlength)
Debug.Print "BaudRate: ", CommDCB.BaudRate
Debug.Print "Bits1: ", Hex$(CommDCB.Bits1); "h"
Debug.Print "wReserved: ", Hex$(CommDCB.wReserved)
Debug.Print "XonLim: ", CommDCB.XonLim
Debug.Print "XoffLim: ", CommDCB.XoffLim
Debug.Print "ByteSize: ", CommDCB.ByteSize
Debug.Print "Parity: ", CommDCB.Parity
Debug.Print "StopBits: ", CommDCB.StopBits
Debug.Print "XonChar: ", Hex$(CommDCB.XonChar); "h"
Debug.Print "XoffChar: ", Hex$(CommDCB.XoffChar); "h"
Debug.Print "ErrorChar: ", Hex$(CommDCB.ErrorChar); "h"
Debug.Print "EofChar: ", Hex$(CommDCB.EofChar); "h"
Debug.Print "EvtChar: ", Hex$(CommDCB.EvtChar); "h"
Debug.Print "wReserved2: ", Hex$(CommDCB.wReserved2)

End Sub

'This routine makes sure that the port numbers are 1-16
'since Visual Basic routines won't accept higher (or lower) port numbers.
Function checkNewPortNumbers() As Boolean

checkNewPortNumbers = True
frmPortSettings.txtStatus.Text = ""

If (PortNumber > 16) Or (VoltmeterPortNumber > 16) Or (ChamberRefPortNumber > 16) Or (ChamberControlPortNumber > 16) Then
    frmPortSettings.txtStatus.Text = frmPortSettings.txtStatus.Text + "COM port numbers must be between 1 and 16" + vbCrLf
    checkNewPortNumbers = False
End If

If (PortNumber < 1) Or (VoltmeterPortNumber < 1) Or (ChamberRefPortNumber < 1) Or (ChamberControlPortNumber < 1) Then
    frmPortSettings.txtStatus.Text = frmPortSettings.txtStatus.Text + "COM port numbers must be between 1 and 16" + vbCrLf
    checkNewPortNumbers = False
End If

End Function

Public Sub GetSettings()
Dim Rack_A_used As Integer
Dim Rack_B_used As Integer
Dim Rack_C_used As Integer
Dim Rack_D_used As Integer
Dim networkSaveEnable As Boolean
Dim compressorShutoffEnable As Boolean

'Get user settings from last time.
BitRate = GetSetting(ProjectName, "Startup", "BitRate", 1200)
PortNumber = GetSetting(ProjectName, "Startup", "PortNumber", 1)
VoltmeterPortNumber = GetSetting(ProjectName, "Startup", "VoltmeterPortNumber", 1)
ChamberRefPortNumber = GetSetting(ProjectName, "Startup", "ChamberRefPortNumber", 1)
ChamberControlPortNumber = GetSetting(ProjectName, "Startup", "ChamberControlPortNumber", 1)
ChamberType = GetSetting(ProjectName, "Startup", "ChamberType", THUNDER)
chamberRhWaitTime = GetSetting(ProjectName, "Startup", "chamberRhWaitTime", 60) '$$$$

temperatureSetpoint = GetSetting(ProjectName, "Startup", "temperatureSetpoint", DEFAULT_TEMP)
frmMain.txtTemperatureSetpoint.Text = Val(temperatureSetpoint)
temperatureWaitTime = GetSetting(ProjectName, "Startup", "temperatureWaitTime", 300)
frmMain.txtTemperatureWaitTime.Text = Val(temperatureWaitTime)

maxSensors = GetSetting(ProjectName, "Startup", "maxSensors", 128)
'SpanOffset = GetSetting(ProjectName, "Startup", "SpanOffset", 1#)
'frmMain.txtSpanOffset.Text = Format$(SpanOffset)
strLocalFolder = GetSetting(ProjectName, "Startup", "LocalFolder", "C:\Cal Data\")
strNetworkFolder = GetSetting(ProjectName, "Startup", "NetworkPath", "S:\SRH_HUMIDITY\Sensor Tip Calibration Run Data\")
networkSaveEnable = GetSetting(ProjectName, "Startup", "NetworkSaveEnable", False)
strLimitsFilename = GetSetting(ProjectName, "Startup", "strLimitsFilename", "C:\Cal Data\Default Limits.dat")

compressorShutoffEnable = GetSetting(ProjectName, "Startup", "CompressorShutoffEnable", False)
If (compressorShutoffEnable = True) Then
    frmMain.chkCompressorShutdown = Checked
Else
    frmMain.chkCompressorShutdown = Unchecked
End If

If (networkSaveEnable = True) Then
    frmFileFolder.chkNetworkEnable = Checked
Else
    frmFileFolder.chkNetworkEnable = Unchecked
End If

thunderMode = GetSetting(ProjectName, "Startup", "thunderMode", RhPcTc)

If (thunderMode = RhPc) Then
    frmMain.mnuRhPc.Checked = True
    frmMain.mnuRhPcTc.Checked = False
    frmMain.lblChamberRH.Caption = "Chamber RH@Pc:"
Else
    frmMain.mnuRhPc.Checked = False
    frmMain.mnuRhPcTc.Checked = True
    frmMain.lblChamberRH.Caption = "Chamber RH@PcTc:"
End If

Rack_A_used = GetSetting(ProjectName, "Startup", "Rack_A_Used", 0)
Rack_B_used = GetSetting(ProjectName, "Startup", "Rack_B_Used", 0)
Rack_C_used = GetSetting(ProjectName, "Startup", "Rack_C_Used", 0)
Rack_D_used = GetSetting(ProjectName, "Startup", "Rack_D_Used", 0)

frmMain.chkRack(RACK_A).value = Rack_A_used
frmMain.chkRack(RACK_B).value = Rack_B_used
frmMain.chkRack(RACK_C).value = Rack_C_used
frmMain.chkRack(RACK_D).value = Rack_D_used

frmMain.txtChamberRhWaitTime.Text = Format$(chamberRhWaitTime)

'Defaults in case values retrieved are invalid:
If BitRate < 300 Then BitRate = 1200
If PortNumber < 1 Then PortNumber = 1
If VoltmeterPortNumber < 1 Then VoltmeterPortNumber = 1
If ChamberRefPortNumber < 1 Then ChamberRefPortNumber = 1
If ChamberControlPortNumber < 1 Then ChamberControlPortNumber = 1
End Sub



Public Sub SaveSettings()
Dim Rack_A_used As Integer
Dim Rack_B_used As Integer
Dim Rack_C_used As Integer
Dim Rack_D_used As Integer
Dim networkSaveEnable As Boolean
Dim compressorShutoffEnable As Boolean

Rack_A_used = frmMain.chkRack(RACK_A).value
Rack_B_used = frmMain.chkRack(RACK_B).value
Rack_C_used = frmMain.chkRack(RACK_C).value
Rack_D_used = frmMain.chkRack(RACK_D).value

'Save user settings for next time.
SaveSetting ProjectName, "Startup", "BitRate", BitRate
SaveSetting ProjectName, "Startup", "PortNumber", PortNumber
SaveSetting ProjectName, "Startup", "VoltmeterPortNumber", VoltmeterPortNumber
SaveSetting ProjectName, "Startup", "ChamberRefPortNumber", ChamberRefPortNumber
SaveSetting ProjectName, "Startup", "ChamberControlPortNumber", ChamberControlPortNumber
SaveSetting ProjectName, "Startup", "ChamberType", ChamberType
SaveSetting ProjectName, "Startup", "chamberRhWaitTime", chamberRhWaitTime
'SaveSetting ProjectName, "Startup", "SpanOffset", SpanOffset
SaveSetting ProjectName, "Startup", "LocalFolder", strLocalFolder
SaveSetting ProjectName, "Startup", "NetworkPath", strNetworkFolder
SaveSetting ProjectName, "Startup", "temperatureSetpoint", temperatureSetpoint
SaveSetting ProjectName, "Startup", "temperatureWaitTime", temperatureWaitTime
SaveSetting ProjectName, "Startup", "strLimitsFilename", strLimitsFilename

If (frmMain.chkCompressorShutdown = Checked) Then
    compressorShutoffEnable = True
Else
    compressorShutoffEnable = False
End If
SaveSetting ProjectName, "Startup", "CompressorShutoffEnable", compressorShutoffEnable


If (frmFileFolder.chkNetworkEnable = Checked) Then
    networkSaveEnable = True
Else
    networkSaveEnable = False
End If
SaveSetting ProjectName, "Startup", "NetworkSaveEnable", networkSaveEnable


SaveSetting ProjectName, "Startup", "Rack_A_Used", Rack_A_used
SaveSetting ProjectName, "Startup", "Rack_B_Used", Rack_B_used
SaveSetting ProjectName, "Startup", "Rack_C_Used", Rack_C_used
SaveSetting ProjectName, "Startup", "Rack_D_Used", Rack_D_used

SaveSetting ProjectName, "Startup", "maxSensors", maxSensors
SaveSetting ProjectName, "Startup", "thunderMode", thunderMode


End Sub

Sub ImmediateTransmit(ByteToSend As Byte)
'Places a byte at the top of the transmit buffer
'for immediate sending.
Dim Success As Boolean
Success = TransmitCommChar(frmMain.MSComm1.CommID, ByteToSend)
End Sub

Public Sub LowResDelay(DelayInMilliseconds As Single)
'Uses the system timer, with resolution of about 56 milliseconds.
Dim Timeout As Single
'Add the delay to the current time.
Timeout = Timer + DelayInMilliseconds / 1000
If Timeout > 86399 Then
    'If the end of the delay spans midnight,
    'subtract 24 hrs. from the Timeout count:
    Timeout = Timeout - 86399
    'and wait for midnight:
    Do Until Timer < 100
        DoEvents
    Loop
End If
'Wait for the Timeout count.
Do Until Timer >= Timeout
    DoEvents
Loop
End Sub



Public Sub ComPortShutDown()
'Close the port.
If frmMain.MSComm1.PortOpen = True Then
    frmMain.MSComm1.PortOpen = False
End If

If frmMain.MSComm2.PortOpen = True Then
    frmMain.MSComm2.PortOpen = False
End If

If frmMain.MSComm3.PortOpen = True Then
    frmMain.MSComm3.PortOpen = False
End If

If frmMain.MSComm4.PortOpen = True Then
    frmMain.MSComm4.PortOpen = False
End If


End Sub


'All four COM ports are opened here.
'Not that the baud rate is fixed at 1200.
'If a different baud rate is desired,
'then a variable must be substituted for the 1200:
Public Sub Startup()
    Call GetSettings

    ' $$$$ PortOpen = frmMain.openComPorts(1200, PortNumber, VoltmeterPortNumber, ChamberRefPortNumber, ChamberControlPortNumber)
    PortOpen = frmMain.openComPorts(1200, PortNumber, VoltmeterPortNumber, ChamberControlPortNumber)

    If (PortOpen = False) Then
        Call frmPortSettings.initializePortSettings
        frmPortSettings.Show
    End If

End Sub

Public Sub VbSetCommTimeouts(BitRate As Long)
'The default timeout for serial-port operations is 5 seconds.
'This routine sets the timeout so that
'the requested number of bytes can transmit or be read
'at the current bit rate.
'Uses the GetCommTimeouts and SetCommTimeouts API functions.
Dim Timeouts As COMMTIMEOUTS
Dim Success As Long
Dim OneByteTimeout As Long
Success = GetCommTimeouts(frmMain.MSComm1.CommID, Timeouts)
OneByteTimeout = CLng(fncOneByteDelay(BitRate))
If frmMain.MSComm1.PortOpen = True Then
    'All values are milliseconds
    'Maximum time between two received characters:
    Timeouts.ReadIntervalTimeout = OneByteTimeout
    'Maximum time for a character to arrive:
    Timeouts.ReadTotalTimoutMultiplier = OneByteTimeout
    'Provide enough time for the bytes to arrive + 1 second.
    Timeouts.ReadTotalTimeoutConstant = 1000
    'Maximum time for a character to transmit:
    Timeouts.WriteTotalTimeoutMultiplier = OneByteTimeout
    'Provide enough time for the bytes to transmit + 1 second.
    Timeouts.WriteTotalTimeoutConstant = 1000
    Success = SetCommTimeouts(frmMain.MSComm1.CommID, Timeouts)
End If
'For debugging/verifying:
'Success = GetCommTimeouts(frmMain.MSComm1.CommID, Timeouts)
'Debug.Print Timeouts.ReadIntervalTimeout
'Debug.Print Timeouts.ReadTotalTimoutMultiplier
'Debug.Print Timeouts.ReadTotalTimeoutConstant
'Debug.Print Timeouts.WriteTotalTimeoutMultiplier
'Debug.Print Timeouts.WriteTotalTimeoutConstant
End Sub
