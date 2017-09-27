Attribute VB_Name = "HumCalMod"
'// API Declarations
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'// API Structures
Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type


'// API constants
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const OPEN_EXISTING = 3
Public Const FILE_ATTRIBUTE_NORMAL = &H80

'Other constants
Public Const RACK_A = 1
Public Const RACK_B = 2
Public Const RACK_C = 3
Public Const RACK_D = 4

Public maxSensors As Integer
'Public SpanOffset As Double
Public Const PASSWORD = "Humidity"

Public Const DEFAULT_TEMP = 23





