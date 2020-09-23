Attribute VB_Name = "modX10"
'CM17 only supports these commands
Global Const ALL_UNITS_OFF = &H0
Global Const ALL_LIGHTS_ON = &H1
Global Const C_ON = &H2
Global Const C_OFF = &H3
Global Const C_DIM = &H4
Global Const C_BRIGHT = &H5
Global Const ALL_LIGHTS_OFF = &H6
'possible future support or for CM11
Global Const C_EXTENDED = &H7
Global Const C_HAIL_REQ = &H8
Global Const C_HAIL_ACK = &H9
Global Const C_PRE_SET_DIM1 = &HA
Global Const C_PRE_SET_DIM2 = &HB
Global Const C_EXTENDED_DATA_TRANSFER = &HC
Global Const C_STATUS_ON = &HD
Global Const C_STATUS_OFF = &HE
Global Const C_STATUS_REQUEST = &HF
Global Const C_CLEAR_MEM = &H10
'
'Global X10HouseCode As Integer
'Global X10DeviceCode As Integer
Global X10Init As Byte
Global LastDim(16) As Integer
Global IgnScrl As Byte
Global SelSwitch As Byte
Global XCommand As Integer
Global X10Command$(17)
Public Type X10Info
    DeviceName(1 To 16) As String * 20
    Configured As Boolean
End Type
Global X10(0 To 15) As X10Info
Public Type X10Status
    Device(1 To 16) As Byte
End Type
Global X10Out(0 To 15) As X10Status
Global UnitTable(16) As Integer
Global HouseTable(16) As Integer
Global HouseUnitCode(16) As Integer
Global StatMess$
Global ComByte() As Byte
Global ComIn$
Global Monitored(16) As Byte
Global MonitoredWords As Long
Global OnStatus(16) As Byte
Global OnStatusWords As Long
Global DimStatus(16) As Byte
Global DimStatusWords As Long
'
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const HWND_TOPMOST = -1
Public Nid As NOTIFYICONDATA
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Global ComInProgress As Byte

Sub Main()
XPMain
Load frmCM11A
frmCM11A.Show
End Sub
