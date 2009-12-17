Attribute VB_Name = "ConfigUtilityModule"
Option Explicit

Public Const CONFIG_VERSION = 16

Public Const INPUT_COUNT = 6

Public Const RX_CMD_SET_TX_MODE = 0
Public Const RX_CMD_CONFIG = 1

Public Const TX_MODE_INFO = 0
Public Const TX_MODE_CONFIG = 1


Public Type UInt
    lo As Byte
    hi As Byte
End Type


'NOTE: we use LONG because VB6 has no unsigned int
'SO STRUCTURES BELOW  ARE NOT PACKED THE SAME WAY AS THE C VERSION !

Public Type InfoType
    sequence As Byte        'pcaket sequence, increments on each packet (overflows at 255)
    interval As Long        'interval since last sample in us
    adc(0 To 5) As Long          'raw adc readings
    adcGyroMid(0 To 2) As Long 'measured gyro adc values for 0 deg/s
End Type


Public Type ConfigType
    version As Byte         'version of config (if version in eeprom doesn't match defaults are loaded)
    inpInvert As Byte       'invert input bit flags 0..5, 6 - swap buttons
    inpAnNum(0 To 5) As Byte     'analog port number for inputs AN 0-5, 255 - N/A
    zeroLevel(0 To 5) As Long       'accelerometer/gyro  zero level mV
    inpSens(0 To 5) As Long      'input sensitivity 0..2 mv/g , 3..5 uV/deg/s = mV/deg/ms
    outScale(0 To 1) As Long     'output scale in 1/1000 units
    outSmoothing(0 To 1) As Byte 'output smoothing
    vdd As Long                      '        VDD voltage used for ADC as reference (mV)
    gyroNoise(0 To 2) As Long '  gyro expected noise at zero level (mV)
    gyroDrift(0 To 2) As Long  ' gyro expected drift from spec zero-level
    gyroAutoZero As Byte  ' automatically detect gyro zero level when device is idle
End Type



Function DeviceTypeToStr(deviceType As Byte)
    Dim str As String
    Select Case deviceType
        Case &H80: str = "PIC18LF13K50"   ' 0b10000000
        Case &H84: str = "PIC18LF14K50"   ' 0b10000100
        Case &H0: str = "PIC18F13K50"     ' 0b00000000
        Case &H4: str = "PIC18F14K50"     ' 0b00000100
        Case &H9: str = "PIC18F2550"      ' 0b00001001
        Case &HE: str = "PIC18F4550"      ' 0b00001110
        Case Else: str = "Unknown"
    End Select
    DeviceTypeToStr = str
End Function


Sub Str2File(pStr As String, pFilePath As String, Optional RaiseErrors = True)
    Dim f, str As String
    f = FreeFile
    On Error Resume Next
    Kill pFilePath
    On Error GoTo 0
    
    If Not RaiseErrors Then On Error Resume Next
    Open pFilePath For Binary Access Write As f
        Put #f, , pStr
    Close f
End Sub

Public Function FileExists(strDest As String) As Boolean
  ' Comments  : Determines if the named file exists
  ' Parameters: strDest - File to check
  ' Returns   : True if the file exists, false otherwise
  ' Source    : Total VB SourceBook 6
  '
  Dim intLen As Integer
 
  If strDest <> vbNullString Then
    On Error Resume Next
    intLen = Len(Dir$(strDest))
    On Error GoTo PROC_ERR
    FileExists = (Not Err And intLen > 0)
  Else
    FileExists = False
  End If

PROC_EXIT:
  Exit Function
  
PROC_ERR:
  MsgBox "Error: " & Err.Number & ". " & Err.Description, , _
    "FileExists"
  Resume PROC_EXIT
  
End Function

Function StrToDouble(str As String, errorValue As Double)
    StrToDouble = errorValue
    On Error Resume Next
    StrToDouble = CDbl(str)
End Function

Function StrToInt(str As String, errorValue As Double)
    StrToInt = errorValue
    On Error Resume Next
    StrToInt = CInt(str)
End Function


Function StrToLong(str As String, errorValue As Double)
    StrToLong = errorValue
    On Error Resume Next
    StrToLong = CLng(str)
End Function
