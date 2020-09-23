Attribute VB_Name = "CommonSubs"
Option Explicit
Declare Function GetWindowsDirectoryA Lib "kernel32" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Declare Function INP Lib "INOUT.dll" (ByVal ADDRESS&) As Integer
'Declare Sub OUT Lib "INOUT.dll" (ByVal ADDRESS&, ByVal value%)
'Declare Function timeGetTime Lib "MMSYSTEM" () As Long
'Declare Function INPORTB Lib "ACCES32.DLL" Alias "InPortB" (ByVal baddr As Integer) As Integer
'Declare Sub outportb Lib "ACCES32.DLL" Alias "OutPortB" (ByVal baddress As Integer, ByVal value As Integer)
'Declare Function InPort Lib "ACCES32.DLL" Alias "InPortB" (ByVal baddr As Integer) As Integer
'Declare Sub OutPort Lib "ACCES32.DLL" Alias "OutPortB" (ByVal baddress As Integer, ByVal value As Integer)
'below is for latest version of access32, hopefully it will work in win xp
'it does not work with the old function calls, so i have to recompile it for each project...ahhhhh
'
Public Declare Function Inp Lib "inpout32.dll" Alias "Inp32" (ByVal PortAddress As Integer) As Integer
Public Declare Sub Out Lib "inpout32.dll" Alias "Out32" (ByVal PortAddress As Integer, ByVal Value As Integer)
'
'Public Declare Function InPortB Lib "ACCESXP" Alias "VBInPortB" (ByVal Port As Long) As Integer
'Public Declare Function OutPortB Lib "ACCESXP" Alias "VBOutPortB" (ByVal Port As Long, ByVal Value As Byte) As Integer
'Public Declare Function InPort Lib "ACCESXP" Alias "VBInPort" (ByVal Port As Long) As Integer
'Public Declare Function OutPort Lib "ACCESXP" Alias "VBOutPort" (ByVal Port As Long, ByVal Value As Integer) As Integer
'Public Declare Function InPortL Lib "ACCESXP" Alias "VBInPortL" (ByVal Port As Long) As Long
'Public Declare Function OutPortL Lib "ACCESXP" Alias "VBOutPortL" (ByVal Port As Long, ByVal Value As Long) As Integer
'Public Declare Function InPortDWord Lib "ACCESXP" Alias "VBInPortDWord" (ByVal Port As Long) As Long
'Public Declare Function OutPortDWord Lib "ACCESXP" Alias "VBOutPortDWord" (ByVal Port As Long, ByVal Value As Long) As Integer
'
Private Declare Sub GetMem1 Lib "msvbvm60" (ByVal Addr As Long, retval As Byte)
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal Addr As Long, retval As Integer)
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, retval As Long)
Private Declare Sub GetMem8 Lib "msvbvm60" (ByVal Addr As Long, retval As Currency)
Private Declare Sub PutMem1 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Byte)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Integer)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Long)
Private Declare Sub PutMem8 Lib "msvbvm60" (ByVal Addr As Long, ByVal NewVal As Currency)
'
Declare Function timeGetTime Lib "winmm.dll" () As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function FlashWindow Lib "user32" (ByVal hwnd As Long, ByVal bInvert As Long) As Long
Declare Function GetActiveWindow Lib "user32" () As Long
Declare Function SetActiveWindow Lib "user32" () As Long
Declare Function GetDesktopWindow& Lib "user32" ()
'message box api
Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
'
'***************************************************************
'Windows API/Global Declarations for :CapsLock and NumLock
'***************************************************************
Public Type KeyboardBytes
    kbByte(0 To 255) As Byte
End Type
Public kbArray As KeyboardBytes
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Long
Public Declare Function GetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Public Declare Function SetKeyboardState Lib "user32" (kbArray As KeyboardBytes) As Long
Public Enum Key
    NumLock = vbKeyNumlock
    Capital = vbKeyCapital
    ScrollLock = vbKeyScrollLock
    Insert = vbKeyInsert
End Enum
Private lngNumLockState As Long
Private lngCapsLockState As Long
Private lngScrollLockState As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
' Colors
Global Const Gray = &H8000000F '&hc0c0c0
Global Const Black = &H0&
Global Const Red = &HFF&
Global Const Green = &HFF00&
Global Const Yellow = &HFFFF&
Global Const Blue = &HC00000   '&hff0000
Global Const Magenta = &HFF00FF
Global Const Cyan = &HFFFF00
Global Const White = &HFFFFFF
'---------------------------------------
'Key Status Control
'---------------------------------------
'Style
Global Const KEYSTAT_CAPSLOCK = 0
Global Const KEYSTAT_NUMLOCK = 1
Global Const KEYSTAT_INSERT = 2
Global Const KEYSTAT_SCROLLLOCK = 3
Global Const WINDOWSERROR = 99
'---------------------------------------
'MESSAGE BOX
Global Const MB_OK = 0, MB_OKCANCEL = 1    ' Define buttons.
Global Const MB_YESNOCANCEL = 3, MB_YESNO = 4
Global Const MB_ICONSTOP = 16, MB_ICONQUESTION = 32    ' Define Icons.
Global Const MB_ICONEXCLAMATION = 48, MB_ICONINFORMATION = 64
Global Const MB_DEFBUTTON2 = 256, IDYES = 6, IDNO = 7  ' Define other.
' Values from CONSTANT.TXT.
Global Const KEY_F1 = &H70
Global Const KEY_F2 = &H71
Global Const KEY_F3 = &H72
Global Const KEY_F4 = &H73
Global Const PGUP = 33
Global Const PGDN = 34
Global Const LArrow = 37
Global Const UArrow = 38
Global Const RArrow = 39
Global Const DArrow = 40
Global Const Home = 36
'
'taskbar manipulation
Global hWnd1 As Long
Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
'note you either have to know the original text of the textbox OR when this control was placed on the form
'ie: the first,second,last etc... then you can find it with findwindowex's second parameter
'
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Declare Function ShowWindowAsync Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Global Const SW_HIDE = 0
Global Const SW_SHOWNORMAL = 1
Global Const SW_SHOWMINIMIZED = 2
Global Const SW_SHOWMAXIMIZED = 3
Global Const SW_SHOWNOACTIVATE = 4
Global Const SW_SHOW = 5
Global Const SW_MINIMIZE = 6
Global Const SW_SHOWMINNOACTIVE = 7
Global Const SW_SHOWNA = 8
Global Const SW_RESTORE = 9
Global Const SW_MAX = 10
Global Const SW_NORMAL = 1
Global Const SW_MAXIMIZE = 3
Global Const HWND_TOP = 0
Global Const HWND_TOPMOST = -1
Global Const HWND_NOTOPMOST = -2
Global Const HWND_BOTTOM = 1
Global Const SWP_NOSIZE = &H1
Global Const SWP_NOMOVE = &H2
Global Const SWP_NOZORDER = &H4
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'below works for combo box
'Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Declare Function SendMessageBynum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal sParam$) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
'
Global Const SWP_HIDEWINDOW = &H80
Global Const SWP_SHOWWINDOW = &H40
Global TaskOn As Byte
'
' // Undocumented Native API to get ShutDown Privilege
'Declare Function RtlAdjustPrivilege& Lib "ntdll" (ByVal Privilege&, ByVal NewValue&, ByVal NewThread&, OldValue&)
' // Native API to ShutDown the System
'Declare Function NtShutdownSystem Lib "ntdll" (ByVal ShutDownAction&)
' // *************
' // Constants
' // *************
' // The ShutDown Privilege
Global Const SE_SHUTDOWN_PRIVILEGE& = 19
'shut down windows
Global Const EWX_FORCE As Long = 4
Global Const EWX_LOGOFF As Long = 0
Global Const EWX_REBOOT As Long = 2
Global Const EWX_ShutDown As Long = 1
Global Const PWR_HIBERNATE = 5
Global Const PWR_SUSPEND = 6
Global Const EWX_POWEROFF = 8
Declare Function ExitWindows Lib "user32" Alias "ExitWindowsEx" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
'
'print spooling
Declare Function ClosePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function EndDocPrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function EndPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function OpenPrinter Lib "winspool.drv" Alias "OpenPrinterA" (ByVal pPrinterName As String, phPrinter As Long, ByVal pDefault As Long) As Long
Declare Function StartDocPrinter Lib "winspool.drv" Alias "StartDocPrinterA" (ByVal hPrinter As Long, ByVal Level As Long, pDocInfo As DOCINFO) As Long
Declare Function StartPagePrinter Lib "winspool.drv" (ByVal hPrinter As Long) As Long
Declare Function WritePrinter Lib "winspool.drv" (ByVal hPrinter As Long, pBuf As Any, ByVal cdBuf As Long, pcWritten As Long) As Long
'
Global lhPrinter As Long
Global PrintSel As Byte
Global lReturn As Long
'
'Global Const WM_CLOSE = &H10
Global Const GW_HWNDNEXT = 2
Global Const GW_OWNER = 4
Global MsgTitle As String
Global MsgInterval As Integer
'
'Listbox API constant
'Global Const LB_ITEMFROMPOINT = &H1A9
'Global Const LB_SETTOPINDEX = &H197
'Global Const LB_FINDSTRING = &H18F
'Global Const LB_SELECTSTRING = &H18C
'Global Const LB_SELITEMRANGEEX = &H183
'Global Const LB_SETHORIZONTALEXTENT = &H194
'Combo Box
Global Const CB_FINDSTRING = &H14C
'
'FileAttribute constants
Global Const ATTR_READONLY = 1    'Read-only file
Global Const ATTR_VOLUME = 8  'Volume label
Global Const ATTR_ARCHIVE = 32    'File has changed since last back-up
Global Const ATTR_NORMAL = 0  'Normal files
Global Const ATTR_HIDDEN = 2  'Hidden files
Global Const ATTR_SYSTEM = 4  'System files
Global Const ATTR_DIRECTORY = 16  'Directory
Global Const ATTR_DIR_ALL = ATTR_DIRECTORY + ATTR_READONLY + ATTR_ARCHIVE + ATTR_HIDDEN + ATTR_SYSTEM
Global Const ATTR_ALL_FILES = ATTR_NORMAL Or ATTR_HIDDEN Or ATTR_SYSTEM Or ATTR_READONLY Or ATTR_ARCHIVE
'
'below is not really used
'Declare Function ConfigurePort Lib "winspool.drv" Alias "ConfigurePortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pPortName As String) As Long
'usage MsgBox ConfigurePort("", Me.hwnd, "COM1")
'return is 1 for success and 0 for error
'
'***************************************************************
'Windows API/Global Declarations for :ListBox Functions
'***************************************************************
Global Const NUL = 0&
Global gSelected() As Long
Global gTotalSelected As Long
Global gItemToInsertBefore As Long
'
Declare Function GetFocus Lib "user32" () As Long
Declare Function GetSelItems Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, Selected&) As Long
'
'for moving stuff
Public Type POINTAPI
    X As Long
    Y As Long
End Type
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Global iTPPY As Long
Global iTPPX As Long
'
Global Mess$
Global Disp
Global CurYear
Global CurDate
Global SERStart
Global DayNum  As Integer
Global Dat$
Global TimeHM$
Global TimeNOW$
Global DayNumSTR$
Global Years
Global mnth
Global Daycode$
Global LastInput
Global OffLine
Global AppDir As String
Global AppName As String
Global PWORD(6)
Global ValidPass
Global LastPass
Type PassData
    WORD As String * 8
End Type
Global Pass As PassData
Global Scratch$
Global Scratch2$
Global SUPER
Global OldDest
Global LastState(48) As Byte
Type DOCINFO
    pDocName As String
    pOutputFile As String
    pDatatype As String
End Type
Global MyDocInfo As DOCINFO
'
'for recycle bin
Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String '  only used if FOF_SIMPLEPROGRESS
End Type
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
'Public Declare Function SHFileOperation Lib "shell32.dll" (ByRef lpFileOp As SHFILEOPSTRUCT) As Long
Public Const ERROR_SUCCESS = 0&
Public Const FO_MOVE = &H1
Public Const FO_COPY = &H2
Public Const FO_DELETE = &H3
Public Const FO_RENAME = &H4
Public Const FOF_ALLOWUNDO = &H40
Public Const FOF_CONFIRMMOUSE = &H2
Public Const FOF_FILESONLY = &H80
Public Const FOF_MULTIDESTFILES = &H1
Public Const FOF_NOCONFIRMATION = &H10
Public Const FOF_NOCONFIRMMKDIR = &H200
Public Const FOF_RENAMEONCOLLISION = &H8
Public Const FOF_SILENT = &H4
Public Const FOF_SIMPLEPROGRESS = &H100
Public Const FOF_WANTMAPPINGHANDLE = &H20
'for titlebar font info
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const LOGPIXELSY = 90
Public Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Public Type LogFont
    FontHeight As Long
    FontWidth As Long
    FontEscapement As Long
    FontOrientation As Long
    FontWeight As Long
    FontItalic As Byte
    FontUnderline As Byte
    FontStrikeOut As Byte
    FontCharSet As Byte
    FontOutPrecision As Byte
    FontClipPrecision As Byte
    FontQuality As Byte
    FontPitchAndFamily As Byte
    FontFaceName As String * 32
End Type
Public Const SPI_GETNONCLIENTMETRICS = 41
Public Type NONCLIENTMETRICS
    cbSize As Long
    iBorderWidth As Long
    iScrollWidth As Long
    iScrollHeight As Long
    iCaptionWidth As Long
    iCaptionHeight As Long
    lfCaptionFont As LogFont
    iSMCaptionWidth As Long
    iSMCaptionHeight As Long
    lfSMCaptionFont As LogFont
    iMenuWidth As Long
    iMenuHeight As Long
    lfMenuFont As LogFont
    lfStatusFont As LogFont
    lfMessageFont As LogFont
End Type
Global FileOnly$
Global CaptWidth As Integer
Global IDEMode As Boolean
Global Confirm As Long
Global eTitle$
Global EMess$
Global mError As Long
Global Cap$
Global Capt$
Global Cap2$
Dim i As Integer
Global sTime#
Global WinDir$
'common dialog printer flags
Global Const cdlPDAllPages = &H0                 'Returns or sets the state of the All Pagesoption button.
Global Const cdlPDCollate = &H10                 'Returns or sets the state of the Collatecheck box.
Global Const cdlPDDisablePrintToFile = &H80000   'Disables the Print To File check box.
Global Const cdlPDHelpButton = &H800             'Causes the dialog box to display the Help button.
Global Const cdlPDHidePrintToFile = &H100000     'Hides the Print To File check box.
Global Const cdlPDNoPageNums = &H8               'Disables the Pages option button and the associated edit control.
Global Const cdlPDNoSelection = &H4              'Disables the Selection option button.
Global Const cdlPDNoWarning = &H80               'Prevents a warning message from being displayed when there is no default printer.
Global Const cdlPDPageNums = &H2                 'Returns or sets the state of the Pages option button.
Global Const cdlPDPrintSetup = &H40              'Causes the system to display the Print Setup dialog box rather than the Print dialog box.
Global Const cdlPDPrintToFile = &H20             'Returns or sets the state of the Print To File check box.
Global Const cdlPDReturnDC = &H100               'Returns adevice context for the printer selection made in the dialog box. The device context is returned in the dialog box's hDC property.
Global Const cdlPDReturnDefault = &H400          'Returns default printer name.
Global Const cdlPDReturnIC = &H200               'Returns an information context for the printer selection made in the dialog box. An information context provides a fast way to get information about the device without creating a device context. The information context is returned in the dialog box's hDC property.
Global Const cdlPDSelection = &H1                'Returns or sets the state of the Selection option button. If neither cdlPDPageNums nor cdlPDSelection is specified, the All option button is in the selected state.
Global Const cdlPDUseDevModeCopies = &H40000     'If a printer driver doesn't support multiple copies, setting this flag disables the Number of copies spinner control in the Print dialog. If a driver does support multiple copies, setting this flag indicates that the dialog box stores the requested number of copies in the Copies property.
'This line is for NT2000 platforms
Public Declare Function NTBeep Lib "kernel32" Alias "Beep" (ByVal FreqHz As Long, ByVal DurationMs As Long) As Long
'adapted from Microsoft's Knowledgebase article (Q189249)
Public Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Enum OsType
    Nt2000
    Win9xMe
    OsUnknown
End Enum
Global WinXP As Byte
'C:\WINDOWS\system32>shutdown.exe /h
'Usage: shutdown.exe [-i | -l | -s | -r | -a] [-f] [-m \\computername][-t xx] [-'c "comment"] [-d up:xx:yy]
'
'        No args                 Display this message (same as -?)
'        -i                      Display GUI interface, must be the first option
'        -l                      Log off (cannot be used with -m option)
'        -s                      Shutdown the computer
'        -r                      Shutdown and restart the computer
'        -a                      Abort a system shutdown
'        -m \\computername       Remote computer to shutdown/restart/abort
'        -t xx                   Set timeout for shutdown to xx seconds
'        -c "comment"            Shutdown comment (maximum of 127 characters)
'        -f                      Forces running applications to close without warning
'        -d [u][p]:xx:yy         The reason code for the shutdown
'                                u is the user code
'                                p is a planned shutdown code
'                                xx is the major reason code (positive integer less than 256)
'                                yy is the minor reason code (positive integer less than 65536)
'ShellExecute GetDesktopWindow, "open", "shutdown", "-s -t 00", "C:\", 0
'Does the trick nicely.  But i dunno if this only works on XP.
'(-s is shutdown, -r is restart and -l is logoff)
'
'APIs to access INI files and retrieve data
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal Filename$)
'
Private Type SHELLITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHELLITEMID
End Type
Public Enum SpecialFolderTypes
    sftCDBurningCache = 59&
    sftCommonAdminTools = 47&
    sftCommonApplicationData = 35&
    sftCommonDesktop = 25&
    sftCommonDocumentTemplates = 45&
    sftCommonFavorites = 31&
    sftCommonMyDocuments = 46&
    sftCommonMyPictures = 54&
    sftCommonProgramFiles = 43&
    sftCommonStartMenu = 22&
    sftCommonStartMenuPrograms = 23&
    sftCommonStartup = 24&
    sftFonts = 20&
    sftProgramFiles = 38&
    sftSystem32Folder = 41&
    sftSystemFolder = 37&
    sftThemes = 56&
    sftUserAdminTools = 48&
    sftUserApplicationData = 26&
    sftUserCookies = 33&
    sftUserDesktop = 16&
    sftUserDocumentTemplates = 21&
    sftUserFavorites = 6&
    sftUserHistory = 34&
    sftUserLocalApplicationData = 28&
    sftUserMyDocuments = 5&
    sftUserMyMusic = 13&
    sftUserMyPictures = 39&
    sftUserNetHood = 19&
    sftUserPrintHood = 27&
    sftUserProfileFolder = 40&
    sftUserRecentDocuments = 8&
    sftUserSendTo = 9&
    sftUserStartMenu = 11&
    sftUserStartMenuPrograms = 2&
    sftUserStartup = 7&
    sftUserTempInternetFiles = 32&
    sftWindowsFolder = 36&
End Enum
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Global MyDocs$
Global AppData$
Global DeskTop$
'usage
'DeskTop$ = SpecialFolderPath(sftUserDesktop)
'jpegFile$ = DeskTop$ & "Stats.jpeg" ' "C:\Documents and Settings\hp\Desktop\Stats.jpeg"

Public Function GetCaptionFont(CapForm As Form) As StdFont
'returns the font used for displaying window captions
'usage
'Set JoyMain.Font = GetCaptionFont(JoyMain)
'
Dim WinFont As LogFont
Dim TargetFont As Font
Dim NCM As NONCLIENTMETRICS
NCM.cbSize = Len(NCM)
Call SystemParametersInfo(SPI_GETNONCLIENTMETRICS, 0, NCM, 0)
If NCM.iCaptionHeight = 0 Then
    WinFont.FontHeight = 0
Else
    WinFont = NCM.lfCaptionFont
End If
Set TargetFont = New StdFont
TargetFont.Charset = WinFont.FontCharSet
TargetFont.Weight = WinFont.FontWeight
TargetFont.Name = WinFont.FontFaceName
TargetFont.Strikethrough = WinFont.FontStrikeOut
TargetFont.Underline = WinFont.FontUnderline
TargetFont.Italic = WinFont.FontItalic
TargetFont.Bold = (WinFont.FontWeight = 700)
TargetFont.Size = -(WinFont.FontHeight * (72 / GetDeviceCaps(CapForm.hdc, LOGPIXELSY)))
Set GetCaptionFont = TargetFont
End Function

Sub SetCapt(CapForm As Form)
'usage
'Set JoyMain.Font = GetCaptionFont(JoyMain)
'SetCapt JoyMain
'
'setup initial caption
Dim tWid As Long
Cap2$ = Trim$(Format(Date, "ddd.  mmm. d, yyyy (y) ")) & Capt$ & Trim$(Time)
GetCaptWidth CapForm, Cap2$
'calculate width of initial caption in twips
tWid = (CapForm.TextWidth(Cap2$))
While tWid < CaptWidth
    Capt$ = " " & Capt$ & " "
    Cap2$ = Trim$(Format(Date, "ddd.  mmm. d, yyyy (y) ")) & Capt$ & Format(Time, "h:nn AM/PM")
    tWid = (CapForm.TextWidth(Cap2$))
Wend
'only update the caption if it is different 10/2002
If CapForm.Caption <> Cap2$ Then
    CapForm.Caption = Cap2$
End If
'CapForm.Refresh
End Sub

Sub GetCaptWidth(CapForm As Form, CaptIn As String)
'gets the available caption area in twips
'
Dim cwid As Integer
Dim metDat As Integer
Dim metDat2 As Integer
Dim controlboxsize As Integer
'get average size of character in twips
cwid = (CapForm.TextWidth(CaptIn)) / Len(CaptIn) 'in twips (scalemode)
'get the Height of windows caption
metDat = GetSystemMetrics(4) * Screen.TwipsPerPixelX ' - 1 '(for some reason it is 1 over the actual size)
'get the width of titlebar bitmap
metDat2 = GetSystemMetrics(30) * Screen.TwipsPerPixelX '
'there are normally 3 control boxes (min; restore; close)
'there is also space between the 3 boxes so add some and add titlebar bitmap size
controlboxsize = ((3 * metDat)) + metDat2 + 200
'calculate character caption area
If CapForm.WindowState = vbMinimized Then
    CaptWidth = 2000
Else
    CaptWidth = (CapForm.ScaleWidth - controlboxsize)
End If
End Sub

Sub CapsOff()
lngCapsLockState = GetKeyState(vbKeyCapital)
GetPlatform
If WinXP = 0 Then
    kbArray.kbByte(vbKeyCapital) = 0
    lReturn = SetKeyboardState(kbArray)
    Debug.Print lReturn
Else
    If lngCapsLockState = 1 Then
        'toggle it on
        'Simulate Key Press
        keybd_event vbKeyCapital, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        'Simulate Key Release
        keybd_event vbKeyCapital, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If
End If
End Sub

Sub CapsOn()
lngCapsLockState = GetKeyState(vbKeyCapital)
GetPlatform
If WinXP = 0 Then
    kbArray.kbByte(vbKeyCapital) = 1
    lReturn = SetKeyboardState(kbArray)
    Debug.Print lReturn
Else
    If lngCapsLockState = 0 Then
        'toggle it on
        'Simulate Key Press
        keybd_event vbKeyCapital, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        'Simulate Key Release
        keybd_event vbKeyCapital, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If
End If
End Sub

Sub NumsOff()
lngNumLockState = GetKeyState(vbKeyNumlock)
GetPlatform
If WinXP = 0 Then
    kbArray.kbByte(vbKeyNumlock) = 0 '
    SetKeyboardState kbArray
Else
    If lngNumLockState = 1 Then
        'toggle it off
        'Simulate Key Press
        keybd_event vbKeyNumlock, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        'Simulate Key Release
        keybd_event vbKeyNumlock, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If
End If
End Sub

Sub NumsOn()
'Because the SetKeyboardState function alters the input state of the calling thread
'and not the global input state of the system, an application cannot use
'SetKeyboardState to set the NUM LOCK, CAPS LOCK, or SCROLL LOCK (or the Japanese KANA)
'indicator lights on the keyboard.
'These can be set or cleared using SendInput to simulate keystrokes.
'Windows NT/2000/XP:
'The keybd_event function can also toggle the NUM LOCK, CAPS LOCK, and SCROLL LOCK keys.
'Windows 95 / 98 / Me:
'The keybd_event function can toggle only the CAPS LOCK and SCROLL LOCK keys.
'It cannot toggle the NUM LOCK key.
'
'so we must get the current state and toggle it if it is not what we want
lngNumLockState = GetKeyState(vbKeyNumlock)
GetPlatform
If WinXP = 0 Then
    kbArray.kbByte(vbKeyNumlock) = 1
    lReturn = SetKeyboardState(kbArray)
    Debug.Print lReturn
Else
    If lngNumLockState = 0 Then
        'toggle it on
        'Simulate Key Press
        keybd_event vbKeyNumlock, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        'Simulate Key Release
        keybd_event vbKeyNumlock, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If
End If
End Sub

Sub PrintLine(PrStr$)
Dim lpcWritten As Long
Dim sWrittenData As String
Dim lDoc As Long
'
'If PrintSel = 0 Then
'    JoyMain.CommonDialog1.ShowPrinter
'    PrintSel = 1
'End If
'above moved to print click event
'if printer was not selected,
'we still print to old one
'
lReturn = OpenPrinter(Printer.DeviceName, lhPrinter, 0)
If lReturn = 0 Then
    MsgBox "The Printer Name you typed wasn't recognized."
    Exit Sub
End If
MyDocInfo.pDocName = App.Title
MyDocInfo.pOutputFile = vbNullString
MyDocInfo.pDatatype = vbNullString
lDoc = StartDocPrinter(lhPrinter, 1, MyDocInfo)
Call StartPagePrinter(lhPrinter)
'
sWrittenData = PrStr$ & vbCrLf '"How's that for Magic !!!!" & vbCrLf
lReturn = WritePrinter(lhPrinter, ByVal sWrittenData, Len(sWrittenData), lpcWritten)
lReturn = EndPagePrinter(lhPrinter)
lReturn = EndDocPrinter(lhPrinter)
lReturn = ClosePrinter(lhPrinter)
End Sub

Sub SetupPrinter(ctlName As Control)
'usage SetupPrinter frm.CommonDialog1
On Error GoTo NOSET
ctlName.CancelError = True
If PrintSel = 0 Then
    'ctlName.Flags = &H40 Or ctlName.Flags
    'cdlPDPrintSetup shows the printer setup, not the page or number of copies
    ctlName.flags = ctlName.flags Or cdlPDPrintSetup And Not cdlPDReturnDefault
    'above caused the dialog to not show after a save before a print
    'i believe that the flags returned after an export was &h400 which is
    'CdlOFNExtensionDifferent.  this set of flags prints to the default printer
    'hence no dialog is displayed!
    ctlName.ShowPrinter
    DoEvents
    'now show the selection and copy dialog
    ctlName.flags = ctlName.flags And Not cdlPDPrintSetup
    ctlName.ShowPrinter
    PrintSel = 1
End If
GoTo pOK
'-----------------
NOSET:
ctlName.flags = 0
PrintSel = 0
'-----------------
pOK:
End Sub

Function AddZero(StrIn$, intDigits)
'we can actually do this by
'AddZero = Format(StrIn$, String(intDigits, "0"))
'
StrIn$ = Trim$(StrIn$)
If Len(StrIn$) >= intDigits Then
    AddZero = StrIn$
    Exit Function
End If
AddZero = String(intDigits - Len(StrIn$), "0") & StrIn$
End Function

Function AddSpace(StrIn$, intDigits)
StrIn$ = Trim$(StrIn$)
If Len(StrIn$) >= intDigits Then
    AddSpace = StrIn$
    Exit Function
End If
AddSpace = String(intDigits - Len(StrIn$), " ") & StrIn$
End Function

Function PadSpace(StrIn$, intDigits)
StrIn$ = Trim$(StrIn$)
If Len(StrIn$) >= intDigits Then
    PadSpace = StrIn$
    Exit Function
End If
PadSpace = StrIn$ & String(intDigits - Len(StrIn$), " ")
End Function

Function CenterText(StrIn$, intDigits)
Dim PreSp As String
Dim PostSp As String
StrIn$ = Trim$(StrIn$)
If Len(StrIn$) > intDigits Then
    CenterText = StrIn$
    Exit Function
End If
PreSp = String((intDigits - Len(StrIn$)) \ 2, " ")
PostSp = String(intDigits - Len(PreSp & StrIn$), " ")
CenterText = PreSp & StrIn$ & PostSp
End Function

Sub FocusMe(ctlName As Control)
Const EM_SETSEL = &HB1
'usage FocusMe Text1
'Static Gcurrent As Object
'Static Gpast As Object
'Static OldFore As Long
'Static OldBack As Long
'If Gcurrent Is Nothing Then
'    Set Gcurrent = Screen.ActiveControl
'    OldBack = Gcurrent.BackColor
'    OldFore = Gcurrent.ForeColor
'Else
'    Set Gpast = Gcurrent
'    Gpast.BackColor = OldBack
'    Gpast.ForeColor = OldFore
'    Set Gcurrent = Screen.ActiveControl
'    OldBack = Gcurrent.BackColor
'    OldFore = Gcurrent.ForeColor
'    Gcurrent.BackColor = vbBlue
'    Gcurrent.ForeColor = vbYellow
'End If
With ctlName
    '.SelStart = 0
    '.SelLength = Len(ctlName.Text)
    SendMessage .hwnd, EM_SETSEL, 0, -1
    'why did i rem out below?
    'because it loses focus with the keypad loading
    '.ForeColor = vbYellow
    '.BackColor = BLUE
    .Refresh
End With
End Sub

Sub CenterForm(CurrentForm As Form)
'usage = CenterForm Me
'
Screen.MousePointer = 11
CurrentForm.Move (Screen.Width - CurrentForm.Width) \ 2, (Screen.Height - CurrentForm.Height) \ 2
Load CurrentForm
CurrentForm.Show
Screen.MousePointer = 0
End Sub

Sub Alarm()
PcSpeakerBeep 100, 50
PcSpeakerBeep 200, 50
PcSpeakerBeep 300, 50
PcSpeakerBeep 400, 50
PcSpeakerBeep 500, 50
PcSpeakerBeep 600, 50
PcSpeakerBeep 700, 50
PcSpeakerBeep 800, 50
PcSpeakerBeep 100, 50
PcSpeakerBeep 200, 50
PcSpeakerBeep 300, 50
PcSpeakerBeep 400, 50
PcSpeakerBeep 500, 50
PcSpeakerBeep 600, 50
PcSpeakerBeep 700, 50
PcSpeakerBeep 800, 50
' Turn speaker off
'OUT 97, INP(97) And &HFC
DoEvents
End Sub
'below can be replaced by using inpout32.dll for winXP
'Public Function Inp(ADDRESS)
'Inp = InPort(ADDRESS)
'End Function
'
'Sub Out(ADDRESS, Value%)
'OutPort ADDRESS, Value%
'End Sub

Sub Beeep()
'now do the message beep for winxp
PcSpeakerBeep 500, 50
PcSpeakerBeep 600, 50
PcSpeakerBeep 700, 50
PcSpeakerBeep 800, 50
' Turn speaker off
'below taken out because of above usage
'Out 97, Inp(97) And &HFC
End Sub

Sub MakeSound(Frequency&)
PcSpeakerBeep Frequency&, 50
Exit Sub
'below now works with inpout32.dll
'declare variables
Dim ClockTicks% 'number of clock ticks
Dim loopcount%  'loop counting variable
'calculate clicks -> Clock/sound frequency = clock ticks
ClockTicks% = CInt(1193280 \ Frequency&)
'ClockTicks% = 1000 '?
'initialize system timer
'prepare for data
Out 67, 182
' Send data
Out 66, ClockTicks% And &HFF
Out 66, ClockTicks% \ 256
'turn speaker on
Out 97, Inp(97) Or &H3 'turns bits 1 & 2 on
'make sure sound plays for a period of time
Call Sleep(50)
' Turn speaker off
Out 97, Inp(97) And &HFC 'turns bits 1 & 2 off
End Sub

Private Sub Win9xBeep(ByVal Freq As Integer, ByVal Length As Single)
Dim LoByte As Integer
Dim HiByte As Integer
Dim Clicks As Integer
Dim SpkrOn As Integer
Dim SpkrOff As Integer
Dim TimeEnd As Single
TimeEnd = Timer + Length / 1000
'Ports 66, 67, and 97 control timer and speaker
'
'Divide clock frequency by sound frequency
'to get number of "clicks" clock must produce.
Clicks = CInt(1193280 / Freq)
LoByte = Clicks And &HFF
HiByte = Clicks \ 256
'Tell timer that data is coming
Out 67, 182
'Send count to timer
Out 66, LoByte
Out 66, HiByte
'Turn speaker on by setting bits 0 and 1 of PPI chip.
SpkrOn = Inp(97) Or &H3
Out 97, SpkrOn
'Leave speaker on (while timer runs)
Do While Timer < TimeEnd
    'Let processor do other tasks
    DoEvents
Loop
'Turn speaker off.
SpkrOff = Inp(97) And &HFC
Out 97, SpkrOff
End Sub

Public Sub PcSpeakerBeep(ByVal FreqHz As Integer, ByVal LengthMs As Single)
Select Case GetPlatform
    Case Win9xMe
        Call Win9xBeep(FreqHz, LengthMs)
    Case Nt2000
        Call NTBeep(CLng(FreqHz), CLng(LengthMs))
    Case OsUnknown
        Beep    'use the default beep routine, probably the sound card
End Select
End Sub
'following routine largely by Jorge Loubet

Public Sub Warble(ByVal FreqHz As Integer, ByVal DurationMs As Single)
Dim EndTime As Single
EndTime = Timer + DurationMs / 1000
If FreqHz < 100 Then FreqHz = 100
Do While EndTime > Timer
    Call PcSpeakerBeep(FreqHz, 10)
    Call PcSpeakerBeep(FreqHz / 1.1, 10)
    Call PcSpeakerBeep(FreqHz / 1.2, 10)
    Call PcSpeakerBeep(FreqHz / 1.3, 10)
    Call PcSpeakerBeep(FreqHz / 1.2, 10)
    Call PcSpeakerBeep(FreqHz / 1.1, 10)
Loop
End Sub

Sub GetTime()
'*************************************************************************
Dim MonthNum, MonthSuffix, msg, SERIAL  ' Declare variables.
SERIAL = Now    ' Get date serial number.
'MonthNum = Month(SERIAL)    ' Get current month number.
Dim Days As String
mnth = Format(SERIAL, "mm")
Days = Format(SERIAL, "DD") '
Dat$ = Format(SERIAL, "ddd  MMMM D, YYYY")
TimeHM$ = Format(SERIAL, "HH:MM AM/PM")
TimeNOW$ = Format(SERIAL, "HH:MM:SS AM/PM")
CurYear = Format(SERIAL, "YYYY") 'Mid$(Date$, 9, 2)
CurDate = DateSerial(CurYear, mnth, Days)
SERStart = (DateValue("01-01-" & CurYear) - 1)
CurDate = (DateValue(Now)) 'DateSerial(95, 5, 5)
DayNum = (CurDate - SERStart)
DayNumSTR$ = Trim$(Right$(Str$(DayNum), 3))
DayNumSTR$ = AddZero(DayNumSTR$, 3)
'While Len(DAYNUMSTR$) < 3: DAYNUMSTR$ = "0" & DAYNUMSTR$: Wend
'DAYNUM = ((CURDATE - DATEVALUE("01-01-" & CURYEAR$)) + 1)
Years = Right$(CurYear, 1)
End Sub

Function HexStrtoBitStr(HexStr As String, BitMax As Integer) As String
' Convert Hex string to binary string Pattern (MSB is first)
On Error GoTo Oops
Dim bNegative As Byte
Dim Exp1 As Long
Dim Exp2 As Long
Dim BitStr$
If BitMax > 31 Then ' GoTo Strings
    Dim WordIn As Long
    'this part fixes the 32nd bit
    'thanks to
    ' D R Lambert 2001 - http://www.drldev.co.uk
    ' Since we can't actually do anything with a positive value >= 2147483648 (&H80000000)
    ' without causing an overflow error, I test whether the number is negative then AND
    ' it with &H7FFFFFFF if it is. This means we only ever have to deal with what appears
    ' to be a positive value. We "bolt" the the value of the Negate bit back onto the front
    ' of the string representing the binary value just before it is returned.
    WordIn = Val("&H" & (HexStr))
    bNegative = (WordIn < 0)  ' Note whether lngValue is negative
    If bNegative Then           ' Convert lngValue into a positive number
        WordIn = WordIn And &H7FFFFFFF
        BitStr$ = "1"             ' The "Negate" bit to be prepended to the result string
    Else
        BitStr$ = "0"             ' The "Negate" bit to be prepended to the result string
    End If
    BitMax = BitMax - 1
Else
    WordIn = Val(Abs("&H" & (HexStr)))
    'why does above give -1 for FFFF instead of 65535???
    'abs fixed it
    BitStr$ = ""
End If
'above messes up at 32 bit
'Hex_To_Dec = CLng(CDec("&H" & (HexStr)))
'above is another way to convert
' Calculate bit pattern from MSB to LSB
'this does not work for anything over 31 bits
For i = BitMax - 1 To 0 Step -1
    Exp1 = (2 ^ i)
    Exp2 = (Exp1 And WordIn)
    BitStr$ = BitStr$ & (Exp2 / Exp1)
    If i Mod 4 = 0 Then BitStr$ = BitStr$ & " "
Next i
HexStrtoBitStr = BitStr$
GoTo Exit_HexStrtoBitStr
'
Strings:
'below works for all hex (over 32 bits too!)
Dim j As Integer
Dim nDec As Long
Const HexChar As String = "0123456789ABCDEF"
HexStr = UCase(HexStr)
BitStr$ = ""
For i = 1 To Len(HexStr)
    nDec = InStr(1, HexChar, Mid$(HexStr, i, 1)) - 1
    For j = 3 To 0 Step -1
        BitStr$ = BitStr$ & (nDec \ (2 ^ j))
        nDec = nDec Mod (2 ^ j)
    Next j
    BitStr$ = BitStr$ & " "
Next i
HexStrtoBitStr = BitStr$
GoTo Exit_HexStrtoBitStr
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine HexStrtoBitStr "
EMess$ = "Error # " & err.Number & " - " & err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in HexStrtoBitStr"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
Alarm
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_HexStrtoBitStr:
End Function

Function ConBin(bNum As Long) As String
Dim X As Long
Dim bIn As String
X = CLng(bNum)
bIn = ""
Do
    bIn = (X And 1) & bIn
    X = X \ 2
Loop While X
Do Until Len(bIn) > 32
    bIn = 0 & bIn
Loop
ConBin = bIn
End Function

Function HexStrtoDec(HexStr As String) As Long
'Dim i As Integer
'Dim nDec As Long
'Const HexChar As String = "0123456789ABCDEF"
'For i = Len(sHex) To 1 Step -1
'    nDec = nDec + (InStr(1, HexChar, Mid$(sHex, i, 1)) - 1) * 16 ^ (Len(sHex) - i)
'Next i
'HexStrtoDec = CStr(nDec)
'this is easier
On Error Resume Next
HexStrtoDec = Val(Abs("&H" & (HexStr)))
'Below works too
'Hex_To_Dec = CLng(CDec("&H" & HexStr))
'
'Dim N As double
'Dim tstr As String
'Dim ustr As String
'ustr = UCase$(HexStr)
'' Convert A-F to 1 - 6 and add to 9.
'' "A" = &H41, - &H40 -> 1
'' Assumes:  lower case hex not used (via InputHex)
'For i = Len(ustr) To 1 Step -1
'    tstr = Mid$(ustr, i, 1)
'    If (tstr >= "A") Then
'        N = N + (9 + (Asc(tstr) - &H40)) * (16 ^ (Len(ustr) - i))
'    Else
'        N = N + Val(tstr) * (16 ^ (Len(ustr) - i))
'    End If
'Next i
'HexStrtoDec = N
End Function

Function Bin2Dec(ByVal sBin As String) As Long
Dim i As Integer
For i = 1 To Len(sBin)
    Bin2Dec = Bin2Dec + CLng(CInt(Mid$(sBin, Len(sBin) - i + 1, 1)) * 2 ^ (i - 1))
Next i
End Function

Sub DelMS()
Call Sleep(1)
End Sub

Sub DelSEC(Dwell As Integer)
If Dwell < 1 Then Dwell = 1
sTime# = Timer
While Timer - sTime# < Dwell: DoEvents: Wend
End Sub

Sub Del10()
Call Sleep(10)
End Sub

Sub Del50()
Call Sleep(50)
End Sub

Sub Del100()
Call Sleep(100)
End Sub

Sub Del250()
Call Sleep(250)
End Sub

Sub Del500()
Call Sleep(500)
End Sub

Sub UnloadForms()
Dim F As Form
For Each F In Forms
    Unload F
    Set F = Nothing
Next
End
End Sub

Function StripPath(FilePathIn As String)
'Dim arrslash
'If FilePathIn = "" Then GoTo pathend
'arrslash = Split(FilePathIn, "\")
'lastslash = UBound(arrslash)
'arrslash(lastslash) = ""
'StripPath = Join(arrslash, "\")
'
'below starts from the right side and may be quicker
Dim ppos As Integer
ppos = InStrRev(FilePathIn, "\")
StripPath = Left$(FilePathIn, ppos)
FileOnly$ = Mid$(FilePathIn, ppos + 1)
pathend:
End Function

Public Function IsFormLoaded(ByVal pObjForm As Form) As Boolean
Dim tmpForm As Form
For Each tmpForm In Forms
    If tmpForm Is pObjForm Then
        IsFormLoaded = True
        Exit For
    End If
Next
End Function

Public Sub Recycle(ByVal Filename As String, Optional Confirm As Boolean)
Dim CFileStruct As SHFILEOPSTRUCT
Dim retval As Long
With CFileStruct
    '.hwnd = Me.hwnd
    If Confirm Then
        .fFlags = FOF_ALLOWUNDO
    Else
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
    End If
    .pFrom = Filename
    .wFunc = FO_DELETE
End With
'Send this file to the bin
retval = SHFileOperation(CFileStruct)
DoEvents
If retval <> ERROR_SUCCESS Then
    Alarm
    MsgBox "Recycle returned error " & retval
    'An error occurred.
End If
End Sub

Function IsBitSet(iBitString As Integer, lBitNo As Integer) As Boolean
IsBitSet = iBitString And (2 ^ lBitNo)
End Function

Sub SetupPath()
On Error GoTo Oops
Dim AppDrive As String
Dim drloc As Integer
If AppDir = "" Then AppDir = App.Path
ChDir AppDir
AppDrive = "C"
drloc = InStr(1, AppDir, ":")
If drloc > 0 Then AppDrive = Mid$(AppDir, drloc - 1, 1)
ChDrive AppDrive
GoTo Exit_SetupPath
Oops:
'Abort=3,Retry=4,Ignore=5
eTitle$ = App.Title & ": Error in Subroutine SetupPath "
EMess$ = "Error # " & err.Number & " - " & err.Description & vbCrLf
EMess$ = EMess$ & "Occurred in SetupPath"
EMess$ = EMess$ & IIf(Erl <> 0, vbCrLf & " at line " & CStr(Erl) & ".", ".")
If err.Number = 76 Then
    EMess$ = EMess$ & vbCrLf & "Drive=" & AppDrive
    EMess$ = EMess$ & vbCrLf & "Path=" & AppDir
End If
Alarm
mError = MsgBox(EMess$, vbAbortRetryIgnore, eTitle$)
If mError = vbRetry Then Resume
If mError = vbIgnore Then Resume Next
Exit_SetupPath:
End Sub

Public Sub FormDrag(TheForm As Form)
ReleaseCapture
Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Function DirExists(sDir As String) As Integer
'check if valid path (valid folder)
Dim tmp As String
Dim iResult As Integer
iResult = 0
On Error GoTo nodir
DirExists = False
If Dir$(sDir, ATTR_DIR_ALL) <> "" Then
    iResult = GetAttr(sDir) And ATTR_DIRECTORY
End If
'could also use DirExists = (Dir$(Path & "\nul") <> "")
If iResult = 0 Then   'Directory not found, or the passed argument is a filename not a directory
    DirExists = False
Else
    DirExists = True
End If
nodir:
End Function

Public Function IsInIDE() As Boolean
'-----------------------------------
'-IsInIDE()
'-
'- It'll return true if running in IDE
'-
Dim X As Long
On Error Resume Next
X = VB.App.LogMode()
If X = 1 Then
    IsInIDE = False
Else
    IsInIDE = True
End If
IDEMode = IsInIDE
End Function

Public Function GetWindowsDirectory() As String
Dim s As String
Dim i As Integer
'or do this
'MsgBox Environ$("windir") & IIf(Len(Environ$("OS")), "\SYSTEM32", "\SYSTEM")
'
i = GetWindowsDirectoryA("", 0)
s = Space(i)
Call GetWindowsDirectoryA(s, i)
If Len(s) > 0 Then
    s = (Left$(s, i - 1))
End If
GetWindowsDirectory = s
End Function

Public Function GetVersion() As String
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)
With osinfo
    Select Case .dwPlatformId
        Case 1
            If .dwMinorVersion = 0 Then
                GetVersion = "Windows 95"
            ElseIf .dwMinorVersion = 10 Then
                GetVersion = "Windows 98"
            ElseIf .dwMinorVersion = 90 Then
                GetVersion = "Windows Me"
            End If
            WinXP = 0
        Case 2
            If .dwMajorVersion = 3 Then
                GetVersion = "Windows NT 3.51"
            ElseIf .dwMajorVersion = 4 Then
                GetVersion = "Windows NT 4.0"
            ElseIf .dwMajorVersion = 5 Then
                GetVersion = "Windows 2000/XP"
                If .dwMinorVersion > 0& Then
                    GetVersion = "Windows XP"
                ElseIf .dwMinorVersion = 0& Then
                    GetVersion = "Windows 2000"
                End If
            End If
            WinXP = 1
        Case Else
            GetVersion = "Failed"
    End Select
End With
End Function
'for speaker beep function, only the platform type is relevant

Public Function GetPlatform() As OsType
'desktop for xp = C:\Documents and Settings\hp\Desktop\
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)
Select Case osinfo.dwPlatformId
    Case 1
        GetPlatform = Win9xMe
        WinXP = 0
    Case 2
        GetPlatform = Nt2000
        WinXP = 1
    Case Else
        GetPlatform = OsUnknown
End Select
End Function

Sub POKE(ByVal Address As Variant, ByVal Value As Variant, Optional ByVal NumBits As Byte = 32)
'from http://www.planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=54863
Select Case NumBits
    Case 8
        PutMem1 Address, Value
    Case 16
        PutMem2 Address, Value
    Case 32
        PutMem4 Address, Value
    Case 64
        PutMem8 Address, Value
    Case Else
        MsgBox "Invalid value length" & vbCr & vbCr & "Must be one from: 8/16/32/64" & vbCr & vbCr & vbTab & "8 - Byte (unsigned)" & vbCr & vbTab & "16 - Word/Integer" & vbCr & vbTab & "32 - Dword/Long" & vbCr & vbTab & "64 - Qword/Currency"
End Select
End Sub

Function PEEK(ByVal Address As Long, Optional ByVal NumBits As Byte = 32) As Variant
Dim Value As Variant
Select Case NumBits
    Case 8
        GetMem1 Address, Value
    Case 16
        GetMem2 Address, Value
    Case 32
        GetMem4 Address, Value
    Case 64
        GetMem8 Address, Value
    Case Else
        MsgBox "Invalid value length" & vbCr & vbCr & "Must be one from: 8/16/32/64" & vbCr & vbCr & vbTab & "8 - Byte (unsigned)" & vbCr & vbTab & "16 - Word/Integer" & vbCr & vbTab & "32 - Dword/Long" & vbCr & vbTab & "64 - Qword/Currency"
        Exit Function
End Select
PEEK = Value
End Function

Sub TestPeek()
Dim Var_Byte As Byte
Dim Var_Int As Integer
Dim Var_Lng As Long
Dim Var_Curr As Currency
Var_Byte = 123
Var_Int = 1234
Var_Lng = 123456
Var_Curr = CDec(5234567890#)
Dim strMsg As String
strMsg = "Get value of variables by address With PEEK:" & vbCr
strMsg = strMsg & "BYTE: " & PEEK(VarPtr(Var_Byte), 8) & vbCr
strMsg = strMsg & "INTEGER: " & PEEK(VarPtr(Var_Int), 16) & vbCr
strMsg = strMsg & "LONG: " & PEEK(VarPtr(Var_Lng)) & vbCr
strMsg = strMsg & "CURRENCY: " & PEEK(VarPtr(Var_Curr), 64)
MsgBox strMsg
POKE VarPtr(Var_Byte), 210, 8
POKE VarPtr(Var_Int), 4321, 16
POKE VarPtr(Var_Lng), 654321
POKE VarPtr(Var_Curr), CDec(9999999999#), 64
strMsg = "Values of variables was changed With POKE:" & vbCr
strMsg = strMsg & "BYTE: " & Var_Byte & vbCr
strMsg = strMsg & "INTEGER: " & Var_Int & vbCr
strMsg = strMsg & "LONG: " & Var_Lng & vbCr
strMsg = strMsg & "CURRENCY: " & Var_Curr & vbCr
MsgBox strMsg
End Sub

Function GetSettingIni(AppName As String, ByVal Section As String, ByVal Key As String, Optional DefValue As String) As Variant
'usage
'EncoderRes = GetSettingIni(App.Title, "Settings", "EncoderRes", 720)
'Returns info from an INI file
Dim Buffer As String
Dim iniFileName$
iniFileName$ = App.Path & "\" & AppName & ".ini"
If Dir(iniFileName$) = "" Then
    EMess$ = iniFileName$ & " not found!"
    GetSettingIni = DefValue
    'MsgBox EMess$, vbCritical
    Exit Function
End If
'we should use an ini file instead of the registry
Buffer = String$(255, 0)
lReturn = GetPrivateProfileString(Section, Key, DefValue, Buffer, Len(Buffer), iniFileName$)
If lReturn = 0 Then
    GetSettingIni = ""
Else
    GetSettingIni = Left(Buffer, InStr(Buffer, Chr(0)) - 1)
End If
End Function

Function SaveSettingIni(AppName As String, ByVal Section As String, ByVal Key As String, ByVal KeyValue As String) As Long
'Function returns 0 if successful and error number if unsuccessful
'usage
'SaveSettingIni App.Title, "Settings", "TallyLeft", frmTally.Left
Dim iniFileName$
SaveSettingIni = 1
iniFileName$ = App.Path & "\" & AppName & ".ini"
If Dir(iniFileName$) = "" Then
    EMess$ = iniFileName$ & " not found!"
    'MsgBox EMess$, vbCritical
    'Exit Function
End If
WritePrivateProfileString Section, Key, KeyValue, iniFileName$
SaveSettingIni = 0
End Function

Public Function SpecialFolderPath(ByVal lngFolderType As SpecialFolderTypes) As String
Dim strPath As String
Dim IDL As ITEMIDLIST
SpecialFolderPath = ""
If SHGetSpecialFolderLocation(0&, lngFolderType, IDL) = 0& Then
    strPath = Space$(255)
    If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal strPath) Then
        SpecialFolderPath = Left$(strPath, InStr(strPath, vbNullChar) - 1&) & "\"
    End If
End If
End Function
