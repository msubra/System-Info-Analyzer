Attribute VB_Name = "Genral"
Option Explicit

Dim FileSys As New FileSystemObject
Public Type SYSTEM_INFO
        dwOemID As Long
         
        dwPageSize As Long
        lpMinimumApplicationAddress As Long
        lpMaximumApplicationAddress As Long
        dwActiveProcessorMask As Long
        dwNumberOrfProcessors As Long
        dwProcessorType As Long
        dwAllocationGranularity As Long
        dwReserved As Long
End Type

Public Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Public Const PROCESSOR_ALPHA_21064 = 21064

Public Const PROCESSOR_ARCHITECTURE_ALPHA = 2

Public Const PROCESSOR_ARCHITECTURE_INTEL = 0

Public Const PROCESSOR_ARCHITECTURE_MIPS = 1

Public Const PROCESSOR_ARCHITECTURE_PPC = 3

Public Const PROCESSOR_ARCHITECTURE_UNKNOWN = &HFFFF

Public Const PROCESSOR_INTEL_386 = 386

Public Const PROCESSOR_INTEL_486 = 486

Public Const PROCESSOR_INTEL_PENTIUM = 586

Public Const PROCESSOR_MIPS_R4000 = 4000


Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" _
    (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, _
    Arguments As Long) As Long
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Public Declare Function GetProcessor Lib "getcpu.dll" (ByVal strCpu As String, ByVal strVendor As String, ByVal strL2Cache As String) As Long
Public Declare Function GetProcessorRawSpeed Lib "getcpu.dll" (ByVal RawSpeed As String) As Long
Public Declare Function GetProcessorNormSpeed Lib "getcpu.dll" (ByVal NormSpeed As String) As Long

'This module contains functions needed by XPButton control

Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function ProgIDFromClsID Lib "ole32.dll" Alias "ProgIDFromCLSID" (pCLSID As _
    Any, lpszProgID As Long) As Long
Public Declare Function CLSIDFromString Lib "ole32.dll" (ByVal lpszProgID As _
    Long, pCLSID As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal bytes As Long)

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As _
    Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
    "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
    ByVal lpReserved As Long, lpType As Long, lpData As Any, _
    lpcbData As Long) As Long

Public Const KEY_READ = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                          ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                          ' SYNCHRONIZE))

Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_MULTI_SZ = 7
Public Const ERROR_MORE_DATA = 234
Public Const MAX_COMPUTERNAME_LENGTH = 15

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Type LARGE_INTEGER
    hiword As Long
    loword As Long
End Type

Public Enum SystemMetricsIndexes
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CXDLGFRAME = 7
    SM_CYDLGFRAME = 8
    SM_CYVTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_DEBUG = 22
    SM_SWAPBUTTON = 23
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_CXMIN = 28
    SM_CYMIN = 29
    SM_CXSIZE = 30
    SM_CYSIZE = 31
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXMINTRACK = 34
    SM_CYMINTRACK = 35
    SM_CXDOUBLECLK = 36
    SM_CYDOUBLECLK = 37
    SM_CXICONSPACING = 38
    SM_CYICONSPACING = 39
    SM_MENUDROPALIGNMENT = 40
    SM_PENWINDOWS = 41
    SM_DBCSENABLED = 42
    SM_CMOUSEBUTTONS = 43
    SM_SECURE = 44
    SM_CXEDGE = 45
    SM_CYEDGE = 46
    SM_CXMINSPACING = 47
    SM_CYMINSPACING = 48
    SM_CXSMICON = 49
    SM_CYSMICON = 50
    SM_CYSMCAPTION = 51
    SM_CXSMSIZE = 52
    SM_CYSMSIZE = 53
    SM_CXMENUSIZE = 54
    SM_CYMENUSIZE = 55
    SM_ARRANGE = 56
    SM_CXMINIMIZED = 57
    SM_CYMINIMIZED = 58
    SM_CXMAXTRACK = 59
    SM_CYMAXTRACK = 60
    SM_CXMAXIMIZED = 61
    SM_CYMAXIMIZED = 62
    SM_NETWORK = 63
    SM_CLEANBOOT = 67
    SM_CXDRAG = 68
    SM_CYDRAG = 69
    SM_SHOWSOUNDS = 70
    SM_CXMENUCHECK = 71           '/* Use instead of GetMenuCheckMarkDimensions()! */
    SM_CYMENUCHECK = 72
    SM_SLOWMACHINE = 73
    SM_MIDEASTENABLED = 74
    SM_MOUSEWHEELPRESENT = 75
    SM_XVIRTUALSCREEN = 76
    SM_YVIRTUALSCREEN = 77
    SM_CXVIRTUALSCREEN = 78
    SM_CYVIRTUALSCREEN = 79
    SM_CMONITORS = 80
    SM_SAMEDISPLAYFORMAT = 81
End Enum
' Parameter for SystemParametersInfo()
Public Enum SysParameters
        SPI_GETBEEP = 1
        SPI_SETBEEP = 2
        SPI_GETMOUSE = 3
        SPI_SETMOUSE = 4
        SPI_GETBORDER = 5
        SPI_SETBORDER = 6
        SPI_GETKEYBOARDSPEED = 10
        SPI_SETKEYBOARDSPEED = 11
        SPI_LANGDRIVER = 12
        SPI_ICONHORIZONTALSPACING = 13
        SPI_GETSCREENSAVETIMEOUT = 14
        SPI_SETSCREENSAVETIMEOUT = 15
        SPI_GETSCREENSAVEACTIVE = 16
        SPI_SETSCREENSAVEACTIVE = 17
        SPI_GETGRIDGRANULARITY = 18
        SPI_SETGRIDGRANULARITY = 19
        SPI_SETDESKWALLPAPER = 20
        SPI_SETDESKPATTERN = 21
        SPI_GETKEYBOARDDELAY = 22
        SPI_SETKEYBOARDDELAY = 23
        SPI_ICONVERTICALSPACING = 24
        SPI_GETICONTITLEWRAP = 25
        SPI_SETICONTITLEWRAP = 26
        SPI_GETMENUDROPALIGNMENT = 27
        SPI_SETMENUDROPALIGNMENT = 28
        SPI_SETDOUBLECLKWIDTH = 29
        SPI_SETDOUBLECLKHEIGHT = 30
        SPI_GETICONTITLELOGFONT = 31
        SPI_SETDOUBLECLICKTIME = 32
        SPI_SETMOUSEBUTTONSWAP = 33
        SPI_SETICONTITLELOGFONT = 34
        SPI_GETFASTTASKSWITCH = 35
        SPI_SETFASTTASKSWITCH = 36
        SPI_SETDRAGFULLWINDOWS = 37
        SPI_GETDRAGFULLWINDOWS = 38
        SPI_GETNONCLIENTMETRICS = 41
        SPI_SETNONCLIENTMETRICS = 42
        SPI_GETMINIMIZEDMETRICS = 43
        SPI_SETMINIMIZEDMETRICS = 44
        SPI_GETICONMETRICS = 45
        SPI_SETICONMETRICS = 46
        SPI_SETWORKAREA = 47
        SPI_GETWORKAREA = 48
        SPI_SETPENWINDOWS = 49
        SPI_GETFILTERKEYS = 50
        SPI_SETFILTERKEYS = 51
        SPI_GETTOGGLEKEYS = 52
        SPI_SETTOGGLEKEYS = 53
        SPI_GETMOUSEKEYS = 54
        SPI_SETMOUSEKEYS = 55
        SPI_GETSHOWSOUNDS = 56
        SPI_SETSHOWSOUNDS = 57
        SPI_GETSTICKYKEYS = 58
        SPI_SETSTICKYKEYS = 59
        SPI_GETACCESSTIMEOUT = 60
        SPI_SETACCESSTIMEOUT = 61
        SPI_GETSERIALKEYS = 62
        SPI_SETSERIALKEYS = 63
        SPI_GETSOUNDSENTRY = 64
        SPI_SETSOUNDSENTRY = 65
        SPI_GETHIGHCONTRAST = 66
        SPI_SETHIGHCONTRAST = 67
        SPI_GETKEYBOARDPREF = 68
        SPI_SETKEYBOARDPREF = 69
        SPI_GETSCREENREADER = 70
        SPI_SETSCREENREADER = 71
        SPI_GETANIMATION = 72
        SPI_SETANIMATION = 73
        SPI_GETFONTSMOOTHING = 74
        SPI_SETFONTSMOOTHING = 75
        SPI_SETDRAGWIDTH = 76
        SPI_SETDRAGHEIGHT = 77
        SPI_SETHANDHELD = 78
        SPI_GETLOWPOWERTIMEOUT = 79
        SPI_GETPOWEROFFTIMEOUT = 80
        SPI_SETLOWPOWERTIMEOUT = 81
        SPI_SETPOWEROFFTIMEOUT = 82
        SPI_GETLOWPOWERACTIVE = 83
        SPI_GETPOWEROFFACTIVE = 84
        SPI_SETLOWPOWERACTIVE = 85
        SPI_SETPOWEROFFACTIVE = 86
        SPI_SETCURSORS = 87
        SPI_SETICONS = 88
        SPI_GETDEFAULTINPUTLANG = 89
        SPI_SETDEFAULTINPUTLANG = 90
        SPI_SETLANGTOGGLE = 91
        SPI_GETWINDOWSEXTENSION = 92
        SPI_SETMOUSETRAILS = 93
        SPI_GETMOUSETRAILS = 94
        SPI_SCREENSAVERRUNNING = 97
End Enum

' SystemParametersInfo flags
Public Const SPIF_UPDATEINIFILE = &H1
Public Const SPIF_SENDWININICHANGE = &H2


Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As LARGE_INTEGER) As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As LARGE_INTEGER) As Long
Public Declare Function GetAsyncKeyState Lib "User32" (ByVal vKey As Long) As Integer
Public Declare Function GetDoubleClickTime Lib "User32" () As Long
Public Declare Function SetDoubleClickTime Lib "User32" (ByVal wCount As Long) As Long
Public Declare Function GetSystemMetrics Lib "User32" (ByVal nIndex As Long) As Long

Public mIsWin95 As Boolean
Public mIsWin95Initialized As Boolean
Global WinDir As String
Global SpFolder As SpecialFolderConst
Public Declare Function SystemParametersInfo Lib "User32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long

Global RegVs As RegValues
Global RegV As RegValue
Global RegK As RegKey
Global RegKs As RegKeys
Global Reg As New Registry

Global Ts As TextStream
Global TxtFso As New FileSystemObject
Global Details As String
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Const HKEY_LOCAL_MACHINE = &H80000002

Public Const HKEY_PERFORMANCE_DATA = &H80000004

Public Const HKEY_USERS = &H80000003

Public Const HKEY_CLASSES_ROOT = &H80000000

Public Const HKEY_CURRENT_CONFIG = &H80000005

Public Const HKEY_CURRENT_USER = &H80000001

Public Const HKEY_DYN_DATA = &H80000006





'Public Type LARGE_INTEGER
'   HiInt As Long
'   LoInt As Long
'End Type

'Returns if the currently running operating system is Microsoft Windows 95
Property Get IsWin95() As Boolean
    If Not mIsWin95Initialized Then
        Dim Os As OSVERSIONINFO, Ret As Long
        Os.dwOSVersionInfoSize = Len(Os)
        
        Ret = GetVersionEx(Os)
        
        mIsWin95 = Ret = 0 Or (Os.dwMajorVersion = 4 And Os.dwMinorVersion = 0)
    End If
    IsWin95 = mIsWin95
End Property

'If value "v" is lower "Min" or higher "Max" it is aligned to this bounds
Function ToBounds(ByVal v As Long, ByVal min As Long, ByVal Max As Long) As Long
    If v < min Then
        ToBounds = min
    ElseIf v > Max Then
        ToBounds = Max
    Else
        ToBounds = v
    End If
End Function

Function SystemErrorDescription(ByVal ErrCode As Long) As String
    Dim Buffer As String * 1024
    Dim Ret As Long
    ' return value is the length of the result message
    Ret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, ErrCode, 0, Buffer, _
        Len(Buffer), 0)
    SystemErrorDescription = Left(Buffer, Ret)
End Function
Public Function hiByte(ByVal W As Integer) As Byte
   If W And &H8000 Then
      hiByte = &H80 Or ((W And &H7FFF) \ &HFF)
   Else
      hiByte = W \ 256
    End If
End Function

Public Function hiword(dw As Long) As Integer
 If dw And &H80000000 Then
      hiword = (dw \ 65535) - 1
 Else
    hiword = dw \ 65535
 End If
End Function

Public Function LoByte(W As Integer) As Byte
 LoByte = W And &HFF
End Function

Public Function loword(dw As Long) As Integer
  If dw And &H8000& Then
      loword = &H8000 Or (dw And &H7FFF&)
   Else
      loword = dw And &HFFFF&
   End If
End Function

Public Function MakeInt(ByVal LoByte As Byte, _
   ByVal hiByte As Byte) As Integer

MakeInt = ((hiByte * &H100) + LoByte)

End Function

Public Function MakeLong(ByVal loword As Integer, _
  ByVal hiword As Integer) As Long

MakeLong = ((hiword * &H10000) + loword)

End Function

Public Function SystemMetrics(sMetrics As SystemMetricsIndexes) As Long
SystemMetrics = GetSystemMetrics(sMetrics)
End Function
Public Function SystemParaMeters(sPara As SysParameters) As Long
SystemParaMeters = sPara
End Function
Public Function LargeIntToCurrency(liInput As LARGE_INTEGER) As Currency
    'copy 8 bytes from the large integer to an ampty currency
    CopyMemory LargeIntToCurrency, liInput, LenB(liInput)
    'adjust it
    LargeIntToCurrency = LargeIntToCurrency * 10000
End Function

Public Function ComputerName() As String
    Dim sTmp As String
    Dim lret As Long
    
    sTmp = String(MAX_COMPUTERNAME_LENGTH, 0)
    lret = GetComputerName(sTmp, Len(sTmp))
    ComputerName = Left(sTmp, InStr(sTmp, Chr$(0)) - 1)
End Function

Function EncryptorDecrypt(Data As String) As String
Dim l, l1 As Long
Dim Key(0 To 5) As String * 1
Dim EncData As String
Dim i As Long



l = Len(Data)

Key(0) = "M"
Key(1) = "A"
Key(2) = "H"
Key(3) = "E"
Key(4) = "S"
Key(5) = "H"

l1 = 6

For i = 1 To l
    EncData = EncData & Chr(Asc(Mid(Data, i, 1)) Xor Asc(Key(i Mod l1)))
Next i

EncryptorDecrypt = EncData
End Function


