VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsX10CM11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'I created this class to communicate with the CM11A X10 module.
'Some of the inspiration came from Keware's CM17A firecracker module
'
'HouseDeviceBinary
'A  1   0110    6
'B  2   1110    E
'C  3   0010    2
'D  4   1010    A
'E  5   0001    1
'F  6   1001    9
'G  7   0101    5
'H  8   1101    D
'I  9   0111    7
'J  10  1111    E
'K  11  0011    3
'L  12  1011    B
'M  13  0000    0
'N  14  1000    8
'O  15  0100    4
'P  16  1100    C
'
' 0 =All Units Off
' 1 =All Lights On
' 2 =On
' 3 =Off
' 4 =Dim
' 5 =Bright
' 6 =All Lights Off
' 7 =Extended
' 8 =Hail Request
' 9 =Hail Ack
' 10 =Pre-set dim1
' 11 =Pre-set dim2
' 12 =Extended Data
' 13 =Status On
' 14 =Status Off
' 15 =Status Request
' 16 =-1
'
Option Explicit
Private Type LARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
'
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Dim UseComport As Integer
Dim gComOpen As Boolean
Private WithEvents Comm1 As MSCommLib.MSComm
Attribute Comm1.VB_VarHelpID = -1
Private mHouseCode As Integer
Private mDeviceCode As Integer
Private UnitTable(16)
Private HouseTable(16)
Private HouseUnitCode(16)
'these are the events that are generated in our class file and
'are accessed by the calling module
Public Event Initialized()
Public Event X10SingleEvent(Devices As String, HouseCode As String, Command As Integer, Extra As String, Data2 As String)
'devices: A single device string, like "A1"
'housecode: The housecode the address was addressed to
'(should always be the same as the housecode preceding the device number.
'command: This will be a -1, which designates the event as an address
'extra: Will always be an empty string
'data2: Will always be an empty string
'
Public Event X10Event(Devices As String, HouseCode As String, Command As Integer, Extra As String, Data2 As String)
'devices: A string of devices if more than 1 was seen.
'Each device is separated with a space. An example would be: "B1 B2 B3".
'You would not see different house codes in this string.
'For example, you would not see "A1 B2 C3", as the devices are in
'different house codes.
'housecode: A string containing a single housecode.
'command: An integer representing the command detected.
'See the exec method for a list of commands.
'extra: This is a string that includes the brightness (or dimness) level that was
'sent by a Dim or Bright command.
'   It is already converted to %. A dim value of 50% would be returned as "50".
'   Note that with the dim command, you may not receive a device in the devices parameter,
'   it may be an empty string. If the dim button is pressed on an X10 controller,
'   it only sends the command to the previously set housecode, and the last addressed
'   device will be dimmed. If the command is Preset Dim command, the extra parameter
'   will be in the range of 0-31.
'data2: If the command is the X10 Extended Code, this is the second data byte of the
'command. The first data byte is in the extra parameter.
'
Private eTitle$
Private EMess$
Private mError As Long
Private lRet As Long
Dim i As Integer
Dim UpperByte As Byte
Dim LowerByte As Byte
Dim HouseStatus$
Dim DeviceStatus$
Dim CommandStatus$
Dim LastHouseCode As Integer
Dim LastDeviceCode As Integer
Const AckStr$ = "Ã" 'Chr$(195)
Dim ComIn$
Dim j As Integer
'
'In order to poll the PC, the interface will continually send:
'    Poll:      7   6   5   4   3   2   1   0
'    Value:     0   1   0   1   1   0   1   0        (0x5a)= 90 = "Z"
'This signal will be repeated once every second until the PC responds.
'4.2.    PC Response to the Poll Signal.
'To terminate the interface's polling and initiate the data transfer,
'the PC must send an acknowledgment to the interface's poll signal.
'This acknowledgment is:
'    Poll:      7   6   5   4   3   2   1   0
'    Value:     1   1   0   0   0   0   1   1        (0xc3)= 195 = "Ã"
'Notice that bit #2 of the PC transmission is not set, indicating that this
'cannot be the beginning of a transmission from the PC.
'
'Power-fail Macro Download Poll Code.
'In order to poll the PC, the interface will continually send:
'    Poll:      7   6   5   4   3   2   1   0
'    Value:     1   0   1   0   0   1   0   1            (0xa5)
'This signal will be repeated once every second until the PC responds.
'5.2.    PC Response to Macro Download Poll Code.
'To stop the polling, the PC must respond with:
'    PC Response:   7   6   5   4   3   2   1   0
'    Value      :   1   1   1   1   1   0   1   1            (0xfb)
'Once this has been transmitted, the macro must be immediately downloaded.
'At this stage, the interface will wait until the 42 byte macro has been
'received before any X-10 transmissions can occur.

Private Sub Class_Initialize()
Set Comm1 = New MSCommLib.MSComm
X10Command$(0) = "All Units Off"
X10Command$(1) = "All Lights On"
X10Command$(2) = "On"
X10Command$(3) = "Off"
X10Command$(4) = "Dim"
X10Command$(5) = "Bright"
X10Command$(6) = "All Lights Off"
X10Command$(7) = "Extended"
X10Command$(8) = "Hail Request"
X10Command$(9) = "Hail Ack"
X10Command$(10) = "Pre-set Dim1"
X10Command$(11) = "Pre-set Dim2"
X10Command$(12) = "Extended Data"
X10Command$(13) = "Status On"
X10Command$(14) = "Status Off"
X10Command$(15) = "Status Request"
X10Command$(16) = "-1"
'Look at FireCracker datasheet for reference
UnitTable(1) = &H0
UnitTable(2) = &H10
UnitTable(3) = &H8
UnitTable(4) = &H18
UnitTable(5) = &H40
UnitTable(6) = &H50
UnitTable(7) = &H48
UnitTable(8) = &H58
UnitTable(9) = &H400
UnitTable(10) = &H410
UnitTable(11) = &H408
UnitTable(12) = &H418
UnitTable(13) = &H440
UnitTable(14) = &H450
UnitTable(15) = &H448
UnitTable(16) = &H458
'
HouseTable(0) = &H60
HouseTable(1) = &H70
HouseTable(2) = &H40
HouseTable(3) = &H50
HouseTable(4) = &H80
HouseTable(5) = &H90
HouseTable(6) = &HA0
HouseTable(7) = &HB0
HouseTable(8) = &HE0
HouseTable(9) = &HF0
HouseTable(10) = &HC0
HouseTable(11) = &HD0
HouseTable(12) = &H0
HouseTable(13) = &H10
HouseTable(14) = &H20
HouseTable(15) = &H30
'below is the mirror of above
HouseUnitCode(1) = &H6
HouseUnitCode(2) = &HE
HouseUnitCode(3) = &H2
HouseUnitCode(4) = &HA
HouseUnitCode(5) = &H1
HouseUnitCode(6) = &H9
HouseUnitCode(7) = &H5
HouseUnitCode(8) = &HD
HouseUnitCode(9) = &H7
HouseUnitCode(10) = &HE
HouseUnitCode(11) = &H3
HouseUnitCode(12) = &HB
HouseUnitCode(13) = &H0
HouseUnitCode(14) = &H8
HouseUnitCode(15) = &H4
HouseUnitCode(16) = &HC
End Sub

Private Sub Class_Terminate()
If Comm1.PortOpen = True Then
    ResetCom
End If
End Sub

Public Property Get ComPort() As Integer
ComPort = UseComport
End Property

Public Property Let ComPort(ByVal newport As Integer)
UseComport = newport
End Property

Public Property Get HouseCode() As Integer
HouseCode = mHouseCode
End Property

Public Property Let HouseCode(ByVal NewHouse As Integer)
mHouseCode = NewHouse
End Property

Public Property Get DeviceCode() As Integer
DeviceCode = mDeviceCode
End Property

Public Property Let DeviceCode(ByVal NewDevice As Integer)
mDeviceCode = NewDevice
End Property

Sub WaitMicroSecs(ms As Long)
Dim s As LARGE_INTEGER
Dim e As LARGE_INTEGER
Dim total
Dim err As Long
total = ms / 0.8
err = QueryPerformanceCounter(s)
Do
    err = QueryPerformanceCounter(e)
Loop While e.LowPart < s.LowPart + total
'Debug.Print Str(s.highpart) + " " + Str(s.lowpart)
'Debug.Print Str(e.highpart) + " " + Str(e.lowpart)
End Sub

Public Sub ResetCom()
' simply closes the comm port
Comm1.PortOpen = False
End Sub

Public Function Init() As Integer
On Error GoTo Oops
X10Init = 0
With Comm1
    .DTREnable = True
    .RThreshold = 1
    .CommPort = UseComport
    ' 4800 baud, no parity, 8 data, and 1 stop bit.
    .Settings = "4800,N,8,1"
    ' Tell the control to read entire buffer when Input is used.
    .InputLen = 0
    ' Open the port.
    .PortOpen = True
    DoEvents
    'get status
    .Output = "‹" 'i think "`" works also
End With
Init = 0
gComOpen = True
GoTo Exit_Init
Oops:
gComOpen = False
Init = err.Number
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine Init "
EMess$ = "Error # " & err.Number & " - " & err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in Init"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_Init:
End Function

Public Sub Exec(HouseCodeStr As String, DeviceCodeStr As String, Command As Integer, Optional Brightness As Integer)
' 2=60 = A
' 2=70 = B
' 3=98 = dim
' 3=88 = bright
' 3=10=1 on
' 3=18=3 on
'--------------------------------------------------------------------------
'Each transmission consists of 40 bits. This is made up of the following:
' Header: 16 bits
' Data: 16 bits
' Footer: 8 bits
'The header and footer never change and are defined as follows:
' Header: 11010101 10101010 => &HD5 &HAA
' Footer: 10101101 => &HAD
'--------------------------------------------------------------------------
Dim SendStr$
Dim HUnit As Integer
Dim CheckSum As Byte
Dim NumDims As Byte
Dim CommandByte As Byte
If Not gComOpen Then Exit Sub
ComInProgress = 1
HouseCodeStr = UCase(HouseCodeStr)
HouseCode = Asc(HouseCodeStr) - Asc("A") '0-15
DeviceCode = Val(DeviceCodeStr)
'If D16 is a "1" the code is a function (command).
'Otherwise D1-D8 is an 'address'.
'Once addressed, a module responds to any command code.
'It becomes 'unaddressed' by the first 'address' message after a command,
'or by 'All Units Off'.
'it seems the firecracker is different from below
'It actually appears to mirror the below table, which is for cm11
'Unit           D8  D4  D2  D1     Hex
'1               0   1   1   0      6h
'2               1   1   1   0      Eh
'3               0   0   1   0      2h
'4               1   0   1   0      Ah
'5               0   0   0   1      1h
'6               1   0   0   1      9h
'7               0   1   0   1      5h
'8               1   1   0   1      Dh
'9               0   1   1   1      7h
'10              1   1   1   1      Fh
'11              0   0   1   1      3h
'12              1   0   1   1      Bh
'13              0   0   0   0      0h
'14              1   0   0   0      8h
'15              0   1   0   0      4h
'16              1   1   0   0      Ch
'0  0  0  x  x     House-Unit "On"
'0  0  1  x  x     House-Unit "Off"
'0  1  0  0  0     House "All Units Off"
'0  1  1  0  0     House "All Lights Off"
'0  1  x  1  0     House "All Lights On"
'0  1  x  0  1     House "Bright 005"
'0  1  x  1  1     House "Dim 005"
'
'all lights on = command &h90?
'all lights/units off = command &ha0?
'
' 0 =All Units Off
' 1 =All Lights On
' 2 =On
' 3 =Off
' 4 =Dim
' 5 =Bright
' 6 =All Lights Off
' 7 =Extended
' 8 =Hail Request
' 9 =Hail Ack
' 10 =Pre-set dim1
' 11 =Pre-set dim2
' 12 =Extended Data
' 13 =Status On
' 14 =Status Off
' 15 =Status Request
' 16 =-1
'
StartSend:
HUnit = (HouseUnitCode(HouseCode + 1) * 16) + HouseUnitCode(DeviceCode)
Comm1.Output = Chr$(4) & Chr$(HUnit)
Debug.Print "Sending 4"; Chr$(HUnit); " ("; Hex$(HUnit); "h)"
CheckSum = 4 + HUnit
'wait for response, which is the checksum
ComIn$ = ""
Do
    DoEvents
Loop Until ComIn$ <> ""
'When the interface receives a transmission from the PC, it will sum all of the bytes,
'and then return a byte checksum. If the checksum is correct, the PC should return a
'value of 0x00 to indicate that the transmission should take place.
'If however, the checksum is incorrect, then the PC should again attempt to transmit
'the Header:Code combination and await a new checksum.
'
'verify the checksum
If ComByte(0) <> CheckSum Then
    frmCM11A.txtEvent.Text = frmCM11A.txtEvent.Text & vbCrLf & "CheckSum not correct - "
    frmCM11A.txtEvent.Text = frmCM11A.txtEvent.Text & "ComByte(0) = " & ComByte(0) & " CheckSum = " & CheckSum & vbCrLf
End If
'send a null
Debug.Print "sending null"
Comm1.Output = Chr$(0)
ComIn$ = ""
Do
    DoEvents
Loop Until ComIn$ <> ""
'we should receive 1 character back acknowledging the null
'It is a "U" (This is a checksum also)
'send the command
SendStr$ = "Exec: " & HouseCodeStr & DeviceCodeStr & " " & X10Command$(Command) & " => "
HUnit = (HouseUnitCode(HouseCode + 1) * 16) + Command
'we must invert the brightness to be dimness if we are dimming
If Brightness > 0 And Command = 4 Then
    Brightness = 100 - Brightness
End If
If Command = 4 Or Command = 5 Then 'dim=4 or bright=5
    If Command = 4 And X10Out(HouseCode).Device(DeviceCode) = 0 Then
        'first turn the light on
        Comm1.Output = Chr$(6) & Chr$(HUnit)
        CheckSum = 6 + HUnit
        ComIn$ = ""
        Do
            DoEvents
        Loop Until ComIn$ <> ""
        Comm1.Output = Chr$(0)
        ComIn$ = ""
        Do
            DoEvents
        Loop Until ComIn$ <> ""
        X10Out(HouseCode).Device(DeviceCode) = 1
    End If
    'send the number of dims in the upper 5 bits and 06 in the lower 3
    NumDims = Brightness \ 5
    SendStr$ = SendStr$ & " NumDims = " & NumDims
    'multiply the dims by 8
    CommandByte = (NumDims * 8) + 6
Else
    CommandByte = 6
End If
Comm1.Output = Chr$(CommandByte) & Chr$(HUnit)
Debug.Print "Sending "; Chr$(CommandByte); Chr$(HUnit); " ("; Hex$(CommandByte); " "; Hex$(HUnit); "h)"
CheckSum = CommandByte + HUnit
ComIn$ = ""
Do
    DoEvents
Loop Until ComIn$ <> ""
'we now get a character back acknowledging the command
'"i" is for off; "h" is for on for housecode a
'frmCM11A.txtEvent.Text = frmCM11A.txtEvent.Text & ComIn$ & " "
'send a null
If ComByte(0) <> CheckSum Then
    frmCM11A.txtEvent.Text = frmCM11A.txtEvent.Text & vbCrLf & "CheckSum not correct - "
    frmCM11A.txtEvent.Text = frmCM11A.txtEvent.Text & "ComByte(0) = " & ComByte(0) & " CheckSum = " & CheckSum & vbCrLf
End If
Comm1.Output = Chr$(0)
Debug.Print "sending null"
'
'don't show byte 0 or 1 or 4 as the headers and footers are fixed
frmCM11A.txtEvent.Text = frmCM11A.txtEvent.Text & SendStr$ & vbCrLf
Debug.Print SendStr$
If Command = 0 Or Command = 6 Then
    Erase X10Out(LastHouseCode).Device
    'For j = 1 To 16
    '    X10Out(LastHouseCode).Device(j) = 0
    'Next j
End If
'example of device table
'A1 ON           0110 0000   0000 0000    60 00
'A1 OFF          0110 0000   0010 0000    60 20
'A2 ON           0110 0000   0001 0000    60 10
'A2 OFF          0110 0000   0011 0000    60 30
'A3 ON           0110 0000   0000 1000    60 08
'A3 OFF          0110 0000   0010 1000    60 28
'A4 ON           0110 0000   0001 1000    60 18
'A4 OFF          0110 0000   0011 1000    60 38
'A5 ON           0110 0000   0100 0000    60 40
'A5 OFF          0110 0000   0110 0000    60 60
'A6 ON           0110 0000   0101 0000    60 50
'A6 OFF          0110 0000   0111 0000    60 70
'A7 ON           0110 0000   0100 1000    60 48
'A7 OFF          0110 0000   0110 1000    60 68
'A8 ON           0110 0000   0101 1000    60 58
'A8 OFF          0110 0000   0111 1000    60 78
'A9 ON           0110 0100   0000 0000    64 00
'A9 OFF          0110 0100   0010 0000    64 20
'A10 ON          0110 0100   0001 0000    64 10
'A10 OFF         0110 0100   0011 0000    64 30
'A11 ON          0110 0100   0000 1000    64 08
'A11 OFF         0110 0100   0010 1000    64 28
'A12 ON          0110 0100   0001 1000    64 18
'A12 OFF         0110 0100   0011 1000    64 38
'A13 ON          0110 0100   0100 0000    64 40
'A13 OFF         0110 0100   0110 0000    64 60
'A14 ON          0110 0100   0101 0000    64 50
'A14 OFF         0110 0100   0111 0000    64 70
'A15 ON          0110 0100   0100 1000    64 48
'A15 OFF         0110 0100   0110 1000    64 68
'A16 ON          0110 0100   0101 1000    64 58
'A16 OFF         0110 0100   0111 1000    64 78
'A BRIGHT 005    0110 0000   1000 1000    64 88
'A DIM 005       0110 0000   1001 1000    64 98
'--------------------------------------------------------------------------
ComInProgress = 0
End Sub

Function GetUnit(CodeIn As Byte) As String
'A  1   0110    6
'B  2   1110    E
'C  3   0010    2
'D  4   1010    A
'E  5   0001    1
'F  6   1001    9
'G  7   0101    5
'H  8   1101    D
'I  9   0111    7
'J  10  1111    E
'K  11  0011    3
'L  12  1011    B
'M  13  0000    0
'N  14  1000    8
'O  15  0100    4
'P  16  1100    C
'pick out the the unitcode that matches our table
For i = 1 To 16
    If CodeIn = HouseUnitCode(i) Then
        GetUnit = Trim$(Str$(i))
        LastDeviceCode = i
        'Debug.Print "LastDeviceCode = "; LastDeviceCode
        Exit For
    End If
Next i
End Function

Function GetHouse(CodeIn As Byte) As String
'pick out the the housecode that matches our table
For i = 1 To 16
    If CodeIn \ 16 = HouseUnitCode(i) Then
        GetHouse = Chr$(64 + i)
        LastHouseCode = i - 1
        'Debug.Print "LastHouseCode = "; LastHouseCode
        Exit For
    End If
Next i
End Function

Sub DecodeAddress(DataIn As Byte)
'separate the house and device from the combined word
UpperByte = DataIn And &HF0  'house code
LowerByte = DataIn And &HF  'device code
HouseStatus$ = GetHouse(UpperByte)
DeviceStatus$ = GetUnit(LowerByte)
Debug.Print "X10 Single Event - "; HouseStatus$; DeviceStatus$
RaiseEvent X10SingleEvent(HouseStatus$ & DeviceStatus$, HouseStatus$, -1, "", "")
End Sub

Sub DecodeCommand(DataIn As Byte)
'separate the house and command
Dim CommandIn As Integer
UpperByte = DataIn And &HF0  'house code
LowerByte = DataIn And &HF  'Command code
CommandIn = CInt(LowerByte)
CommandStatus$ = X10Command$(LowerByte)
Debug.Print "X10 Event - "; HouseStatus$; " "; CommandStatus$
If CommandIn = 3 Then 'off code
    X10Out(LastHouseCode).Device(LastDeviceCode) = 0
    Debug.Print "Changing Device " & HouseStatus$ & DeviceStatus$ & " to Off"
ElseIf CommandIn = 2 Or CommandIn = 4 Or CommandIn = 5 Or CommandIn = 10 Or CommandIn = 11 Then
    X10Out(LastHouseCode).Device(LastDeviceCode) = 1
    Debug.Print "Changing Device " & HouseStatus$ & DeviceStatus$ & " to On"
ElseIf CommandIn = 1 Then
    For j = 1 To 16
        X10Out(LastHouseCode).Device(j) = 1
    Next j
    Debug.Print "Changing Device " & HouseStatus$ & " All On"
    'set the device status to null since all on or all off don't have a device code
    DeviceStatus$ = ""
ElseIf CommandIn = 0 Or CommandIn = 6 Then 'all lights off or all units off
    Erase X10Out(LastHouseCode).Device
    'For j = 1 To 16
    '    X10Out(LastHouseCode).Device(j) = 0
    'Next j
    Debug.Print "Changing Device " & HouseStatus$ & " All Off"
    'set the device status to null since all on or all off don't have a device code
    DeviceStatus$ = ""
End If
RaiseEvent X10Event(DeviceStatus$, HouseStatus$, CommandIn, "", "")
End Sub

Private Sub Comm1_OnComm()
' The OnComm event is used for trapping communications events and errors.
Dim EVMsg$
Dim ERMsg$
Dim cLen
Dim Data1 As Integer
Dim Data2 As Integer
ComInProgress = 1
frmCM11A.Timer1.Enabled = False
' Branch according to the CommEvent property.
Select Case Comm1.CommEvent    ' Event messages.
    Case comEvReceive
        Sleep (100)
        cLen = Comm1.InBufferCount
        ComIn$ = (Comm1.Input)
        'Power-fail Macro Download Poll Code = chr$(165)
        If Left$(ComIn$, 1) = "Z" Or Left$(ComIn$, 1) = "¥" Then
            'acknowledge the hail
            Comm1.Output = AckStr$
            X10Init = 1
            RaiseEvent Initialized
        Else
            'i think that byte 0 is the number of bytes
            'determine event
            ComByte = StrConv(ComIn$, vbFromUnicode)
            Debug.Print "----------------------------------"
            For i = 0 To UBound(ComByte)
                Debug.Print Hex$(ComByte(i)); " ";
            Next i
            Debug.Print
            ShowData ComIn$
            'A  1   0110    6
            'B  2   1110    E
            'C  3   0010    2
            'D  4   1010    A
            'E  5   0001    1
            'F  6   1001    9
            'G  7   0101    5
            'H  8   1101    D
            'I  9   0111    7
            'J  10  1111    E
            'K  11  0011    3
            'L  12  1011    B
            'M  13  0000    0
            'N  14  1000    8
            'O  15  0100    4
            'P  16  1100    C
            '
            ' 0 =All Units Off
            ' 1 =All Lights On
            ' 2 =On
            ' 3 =Off
            ' 4 =Dim
            ' 5 =Bright
            ' 6 =All Lights Off
            ' 7 =Extended
            ' 8 =Hail Request
            ' 9 =Hail Ack
            ' 10 =Pre-set dim1
            ' 11 =Pre-set dim2
            ' 12 =Extended Data
            ' 13 =Status On
            ' 14 =Status Off
            ' 15 =Status Request
            '
            'i think that the all lights on and off don't specify a device,
            'but may have a 0 there instead.
            'when the cm11 timer turned on light A1, it sent  91 129 221
            If UBound(ComByte) > 0 Then
                If UBound(ComByte) = 13 Then DecodeStatus
                If UBound(ComByte) > 2 Then Data1 = CInt(ComByte(3))
                If UBound(ComByte) > 3 Then Data2 = CInt(ComByte(4))
                If ComByte(0) = 2 Then
                    'single events have byte 0=2 and byte 1=0
                    'events have byte 0=2 and byte 1=1
                    'events give housecode and device and its new status
                    'single events give only the housecode and device
                    If ComByte(1) = 0 Then
                        'we have an address string, (Single event happening)
                        'decode the address
                        DecodeAddress ComByte(2)
                    ElseIf ComByte(1) = 1 Then
                        'we have an command event happening
                        DecodeCommand ComByte(2)
                    End If
                    For i = 2 To UBound(ComByte)
                        'Debug.Print i; "="; Hex$(ComByte(i)); " "; Hex$(UpperByte); " " & Hex$(LowerByte)
                    Next i
                    'Debug.Print
                    'ShowData ComIn$
                ElseIf ComByte(1) >= 2 And UBound(ComByte) > 2 Then
                    frmCM11A.txtEvent.SelStart = Len(frmCM11A.txtEvent.Text)
                    frmCM11A.txtEvent.SelText = "---------------------------------" & vbCrLf
                    frmCM11A.txtEvent.SelText = (ComByte(0) - 1) & " Total events" & vbCrLf
                    Dim b As Integer
                    'byte 1 tells us if this is an address string (0) or command(1)
                    b = 2
                    Do
                        'we have both and event and a Single event happening
                        'this might be happening when we first initialize
                        DecodeAddress ComByte(b)
                        'we have to get the data for the event happening
                        If b + 1 > ComByte(0) Then Exit Do
                        DecodeCommand ComByte(b + 1)
                        b = b + 2
                    Loop Until b >= UBound(ComByte) ' ComByte(0)
                    '
                    'ShowData ComIn$
                Else
                    ShowData ComIn$
                End If
            End If
        End If
        ' Events
    Case comEvSend
    Case comEvCTS
        EVMsg$ = "Change in CTS Detected"
    Case comEvDSR
        EVMsg$ = "Change in DSR Detected"
    Case comEvCD
        EVMsg$ = "Change in CD Detected"
    Case comEvRing
        EVMsg$ = "The Phone is Ringing"
    Case comEvEOF
        EVMsg$ = "End of File Detected"
        ' Error messages.
    Case comEventDCB
        EVMsg$ = "Unexpected error retrieving DCB]"
    Case comEventBreak
        EVMsg$ = "Break Received"
    Case comEventFrame
        EVMsg$ = "Framing Error"
    Case comEventOverrun
        ERMsg$ = "Overrun Error"
    Case comEventRxOver
        ERMsg$ = "Receive Buffer Overflow"
    Case comEventTxFull
        ERMsg$ = "Transmit Buffer Full"
    Case comEventRxParity
        EVMsg$ = "Parity Error"
    Case comEventCTSTO
        ERMsg$ = "CTS Timeout"
    Case comEventDSRTO
        ERMsg$ = "DSR Timeout"
    Case comEventCDTO
        ERMsg$ = "Carrier Detect Timeout"
    Case Else
        ERMsg$ = "Unknown error or event => " & Comm1.CommEvent
End Select
If Len(EVMsg$) Then
    ' Display event messages in the label control.
    Debug.Print EVMsg$
    EVMsg$ = ""
ElseIf Len(ERMsg$) Then
    ' Display error messages in an alert message box.
    ERMsg$ = ERMsg$ & vbCrLf & "Click Cancel to quit, OK to ignore."
    lRet = MsgBox(ERMsg$, 1, "Comm Error")
    ERMsg$ = ""
    ' If the user clicks Cancel (2)...
    If lRet = 2 Then
        Comm1.PortOpen = False    ' Close the port and quit.
    End If
End If
ComInProgress = 0
End Sub

Public Static Sub ShowData(Dta$)
' This procedure adds data to the txtterm control's Text property.
' It also filters control characters, such as BACKSPACE,
' carriage return, and line feeds, and writes data to
' an open log file.
' BACKSPACE characters delete the character to the left,
' either in the Text property, or the passed string.
' Line feed characters are appended to all carriage
' returns.  The size of the txtterm control's Text
' property is also monitored so that it never
' exceeds 16384 characters.
'
On Error Resume Next
Dim Nd
Dim dLen
Dim dStr$
Dim dStrH$
Dim aStr$
Dim bInStr$
Dim Added$
Dim Char$
With frmCM11A.txtEvent
    ' Make sure the existing text doesn't get too large.
    Nd = Len(.Text)
    If Nd >= 16384 Then
        .Text = Mid$(.Text, 4097)
        Nd = Len(.Text)
    End If
    ' Point to the end of txtEvent's data.
    .SelStart = Nd
    .SelColor = vbBlack
    GoTo nofilt
    ' Filter/handle BACKSPACE characters.
    Do
        i = InStr(Dta$, Chr$(8))
        If i Then
            If i = 1 Then
                .SelStart = Nd - 1
                .SelLength = 1
                Dta$ = Mid$(Dta$, i + 1)
            Else
                Dta$ = Left$(Dta$, i - 2) & Mid$(Dta$, i + 1)
            End If
        End If
    Loop While i
    ' Eliminate line feeds and carriage returns
    Do
        i = InStr(Dta$, Chr$(10))
        If i Then
            Dta$ = Left$(Dta$, i - 1) & Mid$(Dta$, i + 1)
        End If
    Loop While i
    Do
        i = InStr(Dta$, Chr$(13))
        If i Then
            Dta$ = Left$(Dta$, i - 1) & " " & Mid$(Dta$, i + 1)
        End If
    Loop While i
    ' Make sure all carriage returns have a line feed.
    i = 1
    Do
        i = InStr(i, Dta$, Chr$(13))
        If i Then
            'Dta$ = Left$(Dta$, i) & Chr$(10) & Mid$(Dta$, i + 1)
            i = i + 1
        End If
    Loop While i
nofilt:
    dLen = Len(Dta$)
    dStrH$ = "Hex = "
    dStr$ = "Dec = "
    aStr$ = "Asc = " & Dta$
    bInStr$ = "Bin ="
    For i = 1 To dLen
        Char$ = Mid$(Dta$, i, 1)
        '    If Asc(char$) <> 0 Then
        dStrH$ = dStrH$ & Hex$(Asc(Char$)) & " "
        dStr$ = dStr$ & Trim$(Str$(Asc(Char$))) & " "
        bInStr$ = bInStr$ & HexStrtoBitStr(Hex$(Asc(Char$)), 8) & " "
        '    End If
    Next i
    If ComByte(0) = 4 And Len(Dta$) = 2 Then
        'we have our command coming in from the cm11
        'now we must decode the house and unit
        DecodeAddress (Asc(Right$(Dta$, 1)))
        frmCM11A.txtEvent.SelStart = Len(frmCM11A.txtEvent.Text)
        .SelText = "House Code " & HouseStatus$ & " - Device Code " & DeviceStatus$ & vbCrLf
    End If
    If ComByte(0) = 6 And Len(Dta$) = 2 Then
        'we have our command coming in from the cm11
        'now we must decode the house and unit
        DecodeCommand (Asc(Right$(Dta$, 1)))
        frmCM11A.txtEvent.SelStart = Len(frmCM11A.txtEvent.Text)
        .SelText = "House Code " & HouseStatus$ & " - Command = " & CommandStatus$ & vbCrLf
    End If
    Select Case Dta$
'        Case Is = "f"
'            .SelText = "House Code A - Device Code 1" & vbCrLf
'        Case Is = "n"
'            .SelText = "House Code A - Device Code 2" & vbCrLf
'        Case Is = "b"
'            .SelText = "House Code A - Device Code 3" & vbCrLf
'        Case Is = "j"
'            .SelText = "House Code A - Device Code 4" & vbCrLf
'        Case Is = "a"
'            .SelText = "House Code A - Device Code 5" & vbCrLf
'        Case Is = "i"
'            .SelText = "House Code A - Device Code 6" & vbCrLf
'        Case Is = "e"
'            .SelText = "House Code A - Device Code 7" & vbCrLf
'        Case Is = "m"
'            .SelText = "House Code A - Device Code 8" & vbCrLf
'        Case Is = "g"
'            .SelText = "House Code A - Device Code 9" & vbCrLf
'        Case Is = "o"
'            .SelText = "House Code A - Device Code 10" & vbCrLf
'        Case Is = "c"
'            .SelText = "House Code A - Device Code 11" & vbCrLf
'        Case Is = "k"
'            .SelText = "House Code A - Device Code 12" & vbCrLf
'        Case Is = "`"
'            .SelText = "House Code A - Device Code 13" & vbCrLf
'        Case Is = "h"
'            .SelText = "House Code A - Device Code 14" & vbCrLf
'        Case Is = "d"
'            .SelText = "House Code A - Device Code 15" & vbCrLf
'        Case Is = "b"
'            .SelText = "On Code" & vbCrLf
'        Case Is = "c"
'            .SelText = "Off Code" & vbCrLf
        Case Is = "‹"
            .SelText = "answering HailRequest with status request" & vbCrLf
            Comm1.Output = Chr$(155) '›
        Case Is = "j" '6Ah = 0110 1010
            .SelText = "got a j - sending a null" & vbCrLf
            'Comm1.Output = Chr$(0) ' & vbCrLf
        Case Is = "U"
            .SelText = "got a U - sending the on command" & vbCrLf
            'Comm1.Output = Chr$(6) & "b" '& vbCrLf
        Case Is = "h" '6Ah = 0110 1010
            .SelText = "got a h - sending a null" & vbCrLf
            'Comm1.Output = Chr$(0) ' & vbCrLf
            'light should now be on
        Case Else
            GoTo AddCh
    End Select
    GoTo noadd
AddCh:
    Added$ = ""
    ' Add the filtered data to the Text property.
    If Len(Dta$) = 1 Then
        'Add Ascii
        .SelText = Dta$
        .SelStart = Len(.Text)
        'Add Dec
        .SelText = vbTab & dStr$
        'Add Hex
        .SelText = vbTab & dStrH$
        'Add Bin
        .SelText = vbTab & bInStr$ & vbCrLf & Chr$(10)
    ElseIf Len(Dta$) <> 14 Then
        'Add Ascii
        .SelText = Dta$
        .SelStart = Len(.Text)
        .SelText = vbCrLf
        'Add Dec
        .SelText = dStr$ & vbCrLf
        'Add Hex
        .SelText = dStrH$ & vbCrLf
        'Add Bin
        .SelText = bInStr$ & vbCrLf & Chr$(10)
    End If
End With
Added$ = dStr$ & vbCrLf & dStrH$ & vbCrLf & bInStr$ & vbCrLf & Chr$(10)
'Clipboard.Clear
'Clipboard.SetText Added$
'Comm1.Output = Dta$
noadd:
'.SelText = astr$ & vbcrlf & Chr$(10)
'add carriage return
'.Text = .Text & vbcrlf
End Sub

Function DecodeStatus() As String
On Error GoTo Oops
'For a CM11, the status request is: 0x8b. "›"
'The status request is immediately followed by:
'Dec = 5 0 45 10 4 84 1 103 64 0 64 0 0 0
'Hex = 5 0 2D A 4 54 1 67 40 0 40 0 0 0
'Bin =0000 0101  0000 0000  0010 1101  0000 1010  0000 0100  0101 0100  0000 0001
'0110 0111  0100 0000  0000 0000  0100 0000  0000 0000  0000 0000  0000 0000
'
'0-15   16-31   32-47   48-63   64-79   80-95
'
'Byte   Bit range  Description
'0/1    111 to 96  Battery timer (set to 0xffff on reset)
'
'2      95 to 88   Current time (seconds)
'3      87 to 80   Current time (minutes ranging from 0 to 119)
'
'4      79 to 72   Current time (hours/2, ranging from 0 to 11)
'5      71 to 63   Current year day (MSB bit 63)
'
'the daymask uses a 1 in the position corresponding to the day!
'so sun is 1; mon is 10; tues is 100; wed is 1000; thurs is 10000; etc to 1000000
'6      62 to 56   Day mask (SMTWTFS)
'
'7      55 to 52   Monitored house code
'7      51 to 48   Firmware revision level 0 to 15
'
'8/9    47 to 32   Currently addressed monitored devices
'
'10/11  31 to 16   On / Off status of the monitored devices
'
'12/13  15 to 0    Dim status of the monitored devices
'
'    Battery timer      seconds   min 0-119   hours/2    YearDay---|DayMask   HouseFirm
'    5          0         29          55         9          57         8         67
'0000 0101  0000 0000  0000 1100  0101 0010  0000 1001  0101 0111  0000 1000  0110 0111
' Monitored Devices        On/Off status          Dim status
'     2          0         62          2        40          0
'0000 0010  0000 0000  0110 0010  0000 0010  0100 0000  0000 0000
'
Dim BattTimer As Integer
Dim TimeSec As Byte
Dim TimeMin As Byte
Dim TimeHour As Byte
Dim YearDay As Integer
Dim YearDayMSB As Byte
Dim DayMask As Byte
Dim DayOfWeek As Byte
Dim FirmRev As Byte
Dim ThisHouse As Byte
Dim ThisHouseCode As String * 1
Dim InterDate As Date
For i = 0 To UBound(ComByte)
    Debug.Print Hex$(ComByte(i)); " ";
Next i
Debug.Print
'
BattTimer = (ComByte(0) * 256) + ComByte(1)
TimeSec = ComByte(2)
TimeMin = ComByte(3)
TimeHour = ComByte(4) * 2
If TimeMin > 59 Then
    TimeMin = TimeMin - 59
    TimeHour = TimeHour + 1
End If
YearDay = ComByte(5) 'this only gives 0-255 and we need 365
'the msb for the day of the year is bit 7 of the next byte
'isolate the msb of the byte for the daymask
YearDayMSB = ((ComByte(6) And &H80) / 128)
'if the msb is 1, then add 256 to the day of the year
YearDay = YearDay + (YearDayMSB * 256)
'remove the msb from the daymask
DayMask = ComByte(6) And &H7F
'now find which bit is on
For i = 0 To 7
    If DayMask And (2 ^ i) Then
        DayOfWeek = i + 1
        Exit For
    End If
Next i
FirmRev = ComByte(7) And 15
ThisHouse = (ComByte(7) And 240)
ThisHouseCode = GetHouse(ThisHouse)
InterDate = DateAdd("d", YearDay, ("01-01-" & Year(Date)))
'now decode individual device info
MonitoredWords = Val("&H" & Hex$(ComByte(8)) & Hex$(ComByte(9)))
OnStatusWords = Val("&H" & Hex$(ComByte(10)) & Hex$(ComByte(11)))
DimStatusWords = Val("&H" & Hex$(ComByte(12)) & Hex$(ComByte(13)))
If OnStatusWords = 0 Then
    Erase X10Out(LastHouseCode).Device
    'For j = 1 To 16
    '    X10Out(LastHouseCode).Device(j) = 0
    'Next j
End If
'
Erase Monitored()
Erase OnStatus()
Erase DimStatus()
'decode Monitored status to binary
For i = 0 To 7
    Monitored(i + 8) = ((2 ^ i) And ComByte(8)) / (2 ^ i)
Next i
For i = 0 To 7
    Monitored(i) = ((2 ^ i) And ComByte(9)) / (2 ^ i)
Next i
'decode device on  status to binary
'I thought I could figure out how devices are coded, but it does not seem to follow
'a set order, i.e.  when device A1 is the only thing on, I get 40 00 hex.,
'but, if i also turn on device a3, it changes to something else.
'Maybe i can figure out what the method is
For i = 0 To 7
    OnStatus(i + 8) = ((2 ^ i) And ComByte(10)) / (2 ^ i)
Next i
For i = 0 To 7
    OnStatus(i) = ((2 ^ i) And ComByte(11)) / (2 ^ i)
Next i
'decode dim status to binary
For i = 0 To 7
    DimStatus(i + 8) = ((2 ^ i) And ComByte(12)) / (2 ^ i)
Next i
For i = 0 To 7
    DimStatus(i) = ((2 ^ i) And ComByte(13)) / (2 ^ i)
Next i
'
StatMess$ = "Battery Timer = " & BattTimer
StatMess$ = StatMess$ & vbCrLf '& "Interface Time is "
StatMess$ = StatMess$ & Format$(TimeHour & ":" & TimeMin & ":" & TimeSec, "h:mm:ss am/pm")
'StatMess$ = StatMess$ & vbCrLf & "Interface date is "
StatMess$ = StatMess$ & " " & WeekdayName(DayOfWeek)
StatMess$ = StatMess$ & " " & InterDate
StatMess$ = StatMess$ & vbCrLf & "HouseCode = " & ThisHouseCode
StatMess$ = StatMess$ & ": Firmware Rev. " & FirmRev
Debug.Print "-----------------------------------------------"
Debug.Print StatMess$
Debug.Print "-----------------------------------------------"
'show bits
StatMess$ = StatMess$ & vbCrLf & "Monitored = " & HexStrtoBitStr(Hex$(ComByte(8)) & Hex$(ComByte(9)), 16)
StatMess$ = StatMess$ & vbCrLf & "On/Off Stat = " & HexStrtoBitStr(Hex$(ComByte(10)) & Hex$(ComByte(11)), 16)
StatMess$ = StatMess$ & vbCrLf & "Dim Stat  = " & HexStrtoBitStr(Hex$(ComByte(12)) & Hex$(ComByte(13)), 16)
Debug.Print "Monitored ";
For i = 15 To 0 Step -1
    Debug.Print Trim$(Monitored(i));
    If i = 8 Then Debug.Print " ";
Next i
Debug.Print
Debug.Print "On/Off "; Hex$(ComByte(10)); "; "; Hex$(ComByte(11)); ""
For i = 15 To 0 Step -1
    Debug.Print Trim$(OnStatus(i));
    If i = 8 Then Debug.Print " ";
Next i
Debug.Print
Debug.Print "Dim "; Hex$(ComByte(12)); "; "; Hex$(ComByte(13)); ""
For i = 15 To 0 Step -1
    Debug.Print Trim$(DimStatus(i));
    If i = 8 Then Debug.Print " ";
Next i
Debug.Print
frmCM11A.lblStatus.Caption = StatMess$
X10Init = 1
RaiseEvent Initialized
GoTo Exit_DecodeStatus
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine DecodeStatus "
EMess$ = "Error # " & err.Number & " - " & err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in DecodeStatus"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
Alarm
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_DecodeStatus:
frmCM11A.Timer1.Enabled = True
End Function

Sub GetStatus()
'send the command to acquire the status from the module.
Comm1.Output = "‹"
DoEvents
End Sub
