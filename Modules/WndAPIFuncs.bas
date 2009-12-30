Attribute VB_Name = "WndAPIFuncs"
Option Explicit
'\\ --[WndAPIFuncs]------------------------------------------------------------------------
'\\ API declarations and utility functions for getting/setting information
'\\ about windows.
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

'\\ Window CLASS information....
Public Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetClassInfo Lib "user32" Alias "GetClassInfoA" (ByVal hInstance As Long, ByVal lpClassName As String, lpWndClass As WNDCLASS) As Long

'\\ Get Class Long Index constants....
Public Const GCL_CBCLSEXTRA = (-20)
Public Const GCL_CBWNDEXTRA = (-18)
Public Const GCL_STYLE = (-26)
Public Const GCL_WNDPROC = (-24)

'\\ Class style constants....
Public Enum ClassStyleConstants
    CS_BYTEALIGNCLIENT = &H1000
    CS_BYTEALIGNWINDOW = &H2000
    CS_CLASSDC = &H40
    CS_DBLCLKS = &H8
    CS_HREDRAW = &H2
    CS_INSERTCHAR = &H2000
    CS_KEYCVTWINDOW = &H4
    CS_NOCLOSE = &H200
    CS_NOKEYCVT = &H100
    CS_NOMOVECARET = &H4000
    CS_OWNDC = &H20
    CS_PARENTDC = &H80
    CS_PUBLICCLASS = &H4000
    CS_SAVEBITS = &H800
    CS_VREDRAW = &H1
End Enum

'\\ Window specific information
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'\\ Get the window text....
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long

'\\ Get Window Long Indexes...
Public Const GWL_EXSTYLE = (-20)
Public Const GWL_HINSTANCE = (-6)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_ID = (-12)
Public Const GWL_STYLE = (-16)
Public Const GWL_USERDATA = (-21)
Public Const GWL_WNDPROC = (-4)

'\\ Get relative window
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long

'\\ Get relative window constants
Public Const GW_CHILD = 5
Public Const GW_HWNDFIRST = 0
Public Const GW_HWNDLAST = 1
Public Const GW_HWNDNEXT = 2
Public Const GW_HWNDPREV = 3
Public Const GW_MAX = 5
Public Const GW_OWNER = 4

'\\ Child window enumerations....
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

'\\ Window message constants - for WndProc(wMsg).
Public Enum WindowMessages
 WM_ACTIVATE = &H6             '(LOWORD) wp = WA_, (HIWORD) > 0 if minimized, lp =hwnd
 WM_ACTIVATEAPP = &H1C
 WM_ASKCBFORMATNAME = &H30C
 WM_CANCELJOURNAL = &H4B
 WM_CANCELMODE = &H1F
 WM_CAPTURECHANGED = &H1F       'wParam = 0, lParam = New capture hWnd
 WM_CHANGECBCHAIN = &H30D
 WM_CHAR = &H102
 WM_CHARTOITEM = &H2F
 WM_CHILDACTIVATE = &H22
 WM_CHOOSEFONT_GETLOGFONT = (&H400 + 1)
 WM_CHOOSEFONT_SETFLAGS = (&H400 + 102)
 WM_CHOOSEFONT_SETLOGFONT = (&H400 + 101)
 WM_CLEAR = &H303
 WM_CLOSE = &H10
 WM_COMMAND = &H111
 WM_COMPACTING = &H41
 WM_COMPAREITEM = &H39
 WM_CONVERTREQUESTEX = &H108
 WM_COPY = &H301
 WM_COPYDATA = &H4A
 WM_CREATE = &H1
 WM_CTLCOLORBTN = &H135
 WM_CTLCOLORDLG = &H136
 WM_CTLCOLOREDIT = &H133
 WM_CTLCOLORLISTBOX = &H134
 WM_CTLCOLORMSGBOX = &H132
 WM_CTLCOLORSCROLLBAR = &H137
 WM_CTLCOLORSTATIC = &H138
 WM_CUT = &H300
 WM_DDE_ACK = (&H3E0 + 4)
 WM_DDE_ADVISE = (&H3E0 + 2)
 WM_DDE_DATA = (&H3E0 + 5)
 WM_DDE_EXECUTE = (&H3E0 + 8)
 WM_DDE_FIRST = &H3E0
 WM_DDE_INITIATE = &H3E0
 WM_DDE_LAST = (&H3E0 + 8)
 WM_DDE_POKE = (&H3E0 + 7)
 WM_DDE_REQUEST = (&H3E0 + 6)
 WM_DDE_TERMINATE = (&H3E0 + 1)
 WM_DDE_UNADVISE = (&H3E0 + 3)
 WM_DEADCHAR = &H103
 WM_DELETEITEM = &H2D
 WM_DESTROY = &H2
 WM_DESTROYCLIPBOARD = &H307
 WM_DEVMODECHANGE = &H1B
 WM_DRAWCLIPBOARD = &H308
 WM_DRAWITEM = &H2B
 WM_DROPFILES = &H233
 WM_ENABLE = &HA
 WM_ENDSESSION = &H16
 WM_ENTERIDLE = &H121
 WM_ENTERMENULOOP = &H211
 WM_ERASEBKGND = &H14  'wParam = 0, lParam = hDC of window.  Return 0 if intercepted...
 WM_EXITMENULOOP = &H212
 WM_FONTCHANGE = &H1D
 WM_GETDLGCODE = &H87
 WM_GETFONT = &H31
 WM_GETHOTKEY = &H33
 WM_GETMINMAXINFO = &H24
 WM_GETTEXT = &HD
 WM_GETTEXTLENGTH = &HE
 WM_HOTKEY = &H312
 WM_HSCROLL = &H114
 WM_HSCROLLCLIPBOARD = &H30E
 WM_ICONERASEBKGND = &H27
 WM_IME_CHAR = &H286
 WM_IME_COMPOSITION = &H10F
 WM_IME_COMPOSITIONFULL = &H284
 WM_IME_CONTROL = &H283
 WM_IME_ENDCOMPOSITION = &H10E
 WM_IME_KEYDOWN = &H290
 WM_IME_KEYLAST = &H10F
 WM_IME_KEYUP = &H291
 WM_IME_NOTIFY = &H282
 WM_IME_SELECT = &H285
 WM_IME_SETCONTEXT = &H281
 WM_IME_STARTCOMPOSITION = &H10D
 WM_INITDIALOG = &H110
 WM_INITMENU = &H116
 WM_INITMENUPOPUP = &H117
 WM_KEYDOWN = &H100
 WM_KEYFIRST = &H100
 WM_KEYLAST = &H108
 WM_KEYUP = &H101
 WM_KILLFOCUS = &H8 'wParam = hWnd of window about to lose focus.
 WM_LBUTTONDBLCLK = &H203
 WM_LBUTTONDOWN = &H201
 WM_LBUTTONUP = &H202
 WM_MBUTTONDBLCLK = &H209
 WM_MBUTTONDOWN = &H207
 WM_MBUTTONUP = &H208
 WM_MDIACTIVATE = &H222
 WM_MDICASCADE = &H227
 WM_MDICREATE = &H220
 WM_MDIDESTROY = &H221
 WM_MDIGETACTIVE = &H229
 WM_MDIICONARRANGE = &H228
 WM_MDIMAXIMIZE = &H225
 WM_MDINEXT = &H224
 WM_MDIREFRESHMENU = &H234
 WM_MDIRESTORE = &H223
 WM_MDISETMENU = &H230
 WM_MDITILE = &H226
 WM_MEASUREITEM = &H2C
 WM_MENUCHAR = &H120
 WM_MENUSELECT = &H11F
 WM_MOUSEACTIVATE = &H21
 WM_MOUSEFIRST = &H200
 WM_MOUSELAST = &H209
 WM_MOUSEMOVE = &H200
 WM_MOVE = &H3
 WM_NCACTIVATE = &H86
 WM_NCCALCSIZE = &H83
 WM_NCCREATE = &H81
 WM_NCDESTROY = &H82
 WM_NCHITTEST = &H84
 WM_NCLBUTTONDBLCLK = &HA3
 WM_NCLBUTTONDOWN = &HA1
 WM_NCLBUTTONUP = &HA2
 WM_NCMBUTTONDBLCLK = &HA9
 WM_NCMBUTTONDOWN = &HA7
 WM_NCMBUTTONUP = &HA8
 WM_NCMOUSEMOVE = &HA0
 WM_NCPAINT = &H85
 WM_NCRBUTTONDBLCLK = &HA6
 WM_NCRBUTTONDOWN = &HA4
 WM_NCRBUTTONUP = &HA5
 WM_NEXTDLGCTL = &H28
 WM_NULL = &H0
 WM_PAINT = &HF
 WM_PAINTCLIPBOARD = &H309
 WM_PAINTICON = &H26
 WM_PALETTECHANGED = &H311
 WM_PALETTEISCHANGING = &H310
 WM_PARENTNOTIFY = &H210
 WM_PASTE = &H302
 WM_PENWINFIRST = &H380
 WM_PENWINLAST = &H38F
 WM_POWER = &H48
 WM_PSD_ENVSTAMPRECT = (&H400 + 5)
 WM_PSD_FULLPAGERECT = (&H400 + 1)
 WM_PSD_GREEKTEXTRECT = (&H400 + 4)
 WM_PSD_MARGINRECT = (&H400 + 3)
 WM_PSD_MINMARGINRECT = (&H400 + 2)
 WM_PSD_PAGESETUPDLG = (&H400)
 WM_PSD_YAFULLPAGERECT = (&H400 + 6)
 WM_QUERYDRAGICON = &H37
 WM_QUERYENDSESSION = &H11
 WM_QUERYNEWPALETTE = &H30F
 WM_QUERYOPEN = &H13
 WM_QUEUESYNC = &H23
 WM_QUIT = &H12
 WM_RBUTTONDBLCLK = &H206
 WM_RBUTTONDOWN = &H204
 WM_RBUTTONUP = &H205
 WM_RENDERALLFORMATS = &H306
 WM_RENDERFORMAT = &H305
 WM_SETCURSOR = &H20
 WM_SETFOCUS = &H7
 WM_SETFONT = &H30
 WM_SETHOTKEY = &H32
 WM_SETREDRAW = &HB
 WM_SETTEXT = &HC
 WM_SETTINGCHANGE = &H1A
 WM_SHOWWINDOW = &H18
 WM_SIZE = &H5
 WM_SIZECLIPBOARD = &H30B
 WM_SPOOLERSTATUS = &H2A
 WM_SYSCHAR = &H106
 WM_SYSCOLORCHANGE = &H15
 WM_SYSCOMMAND = &H112
 WM_SYSDEADCHAR = &H107
 WM_SYSKEYDOWN = &H104
 WM_SYSKEYUP = &H105
 WM_TIMECHANGE = &H1E
 WM_TIMER = &H113
 WM_UNDO = &H304
 WM_USER = &H400
 WM_VKEYTOITEM = &H2E
 WM_VSCROLL = &H115
 WM_VSCROLLCLIPBOARD = &H30A
 WM_WINDOWPOSCHANGED = &H47
 WM_WINDOWPOSCHANGING = &H46
 WM_WININICHANGE = &H1A
End Enum

Public Enum EditMessages
    EM_CANUNDO = &HC6
    EM_EMPTYUNDOBUFFER = &HCD
    EM_FMTLINES = &HC8
    EM_GETFIRSTVISIBLELINE = &HCE
    EM_GETHANDLE = &HBD
    EM_GETLINE = &HC4
    EM_GETLINECOUNT = &HBA
    EM_GETMODIFY = &HB8
    EM_GETPASSWORDCHAR = &HD2
    EM_GETRECT = &HB2
    EM_GETSEL = &HB0
    EM_GETTHUMB = &HBE
    EM_LIMITTEXT = &HC5
    EM_LINEFROMCHAR = &HC9
    EM_LINEINDEX = &HBB
    EM_LINELENGTH = &HC1
    EM_LINESCROLL = &HB6
    EM_REPLACESEL = &HC2
    EM_SCROLL = &HB5
    EM_SCROLLCARET = &HB7
    EM_SETHANDLE = &HBC
    EM_SETMODIFY = &HB9
    EM_SETPASSWORDCHAR = &HCC
    EM_SETREADONLY = &HCF
    EM_SETRECT = &HB3
    EM_SETRECTNP = &HB4
    EM_SETSEL = &HB1
    EM_SETTABSTOPS = &HCB
    EM_UNDO = &HC7
    EM_GETWORDBREAKPROC = &HD1 'wp = 0L, lp = procaddress
    EM_SETWORDBREAKPROC = &HD0 'wp = 0L, lp = procaddress
End Enum

'\\ Edit word break code constants
Public Const WB_ISDELIMITER = 2
Public Const WB_LEFT = 0
Public Const WB_RIGHT = 1

Public Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Public Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Public Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Public Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Public Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type

Private Const DWL_DLGPROC = 4
Private Const DWL_MSGRESULT = 0
Private Const DWL_USER = 8

Public Declare Function EnumWindowStations Lib "user32" Alias "EnumWindowStationsA" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumDesktops Lib "user32" Alias "EnumDesktopsA" (ByVal hwinsta As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

Public Declare Function GetProcessWindowStation Lib "user32" () As Long
Public Declare Function EnumProps Lib "user32" Alias "EnumPropsA" (ByVal hWnd As Long, ByVal lpEnumFunc As Long) As Long
Public Declare Function EnumPropsEx Lib "user32" Alias "EnumPropsExA" (ByVal hWnd As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumResourceLanguages Lib "kernel32" Alias "EnumResourceLanguagesA" (ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumResourceNames Lib "kernel32" Alias "EnumResourceNamesA" (ByVal hModule As Long, ByVal lpType As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function EnumResourceTypes Lib "kernel32" Alias "EnumResourceTypesA" (ByVal hModule As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function LoadResource Lib "kernel32" (ByVal hInstance As Long, ByVal hResInfo As Long) As Long
Public Declare Function LoadLibraryEx Lib "kernel32" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Public Declare Function EnumDesktopWindows Lib "user32" (ByVal hDesktop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Public Declare Function GetThreadDesktop Lib "user32" (ByVal dwThread As Long) As Long

Public Const LOAD_LIBRARY_AS_DATAFILE = &H2

Public Enum enResourceTypes
    RT_ACCELERATOR = 9&
    RT_BITMAP = 2&
    RT_CURSOR = 1&
    RT_DIALOG = 5&
    RT_FONT = 8&
    RT_FONTDIR = 7&
    RT_ICON = 3&
    RT_MENU = 4&
    RT_RCDATA = 10&
    RT_STRING = 6&
End Enum

Public Declare Function SendMessageMove Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As RECT) As Long
Public Declare Function SendMessageProc Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Long) As Long

Public Declare Function IsIconic Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SendMessageRepaint Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Type POINTAPI
        x As Long
        y As Long
End Type
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Const WA_ACTIVE = 1
Public Const WA_INACTIVE = 0
Public Const WA_CLICKACTIVE = 2

Public Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long

Public Enum enSystemParametersInfo
    SPI_NOMESSAGE = 0 '\\ To cope with Win95/WinNT differences!
    SPI_GETACCESSTIMEOUT = 60
    SPI_GETANIMATION = 72
    SPI_GETBEEP = 1
    SPI_GETBORDER = 5
    SPI_GETDEFAULTINPUTLANG = 89
    SPI_GETDRAGFULLWINDOWS = 38
    SPI_GETFASTTASKSWITCH = 35
    SPI_GETFILTERKEYS = 50
    SPI_GETFONTSMOOTHING = 74
    SPI_GETGRIDGRANULARITY = 18
    SPI_GETHIGHCONTRAST = 66
    SPI_GETICONMETRICS = 45
    SPI_GETICONTITLELOGFONT = 31
    SPI_GETICONTITLEWRAP = 25
    SPI_GETKEYBOARDDELAY = 22
    SPI_GETKEYBOARDPREF = 68
    SPI_GETKEYBOARDSPEED = 10
    SPI_GETLOWPOWERACTIVE = 83
    SPI_GETLOWPOWERTIMEOUT = 79
    SPI_GETMENUDROPALIGNMENT = 27
    SPI_GETMINIMIZEDMETRICS = 43
    SPI_GETMOUSE = 3
    SPI_GETMOUSEKEYS = 54
    SPI_GETMOUSETRAILS = 94
    SPI_GETNONCLIENTMETRICS = 41
    SPI_GETPOWEROFFACTIVE = 84
    SPI_GETPOWEROFFTIMEOUT = 80
    SPI_GETSCREENREADER = 70
    SPI_GETSCREENSAVEACTIVE = 16
    SPI_GETSCREENSAVETIMEOUT = 14
    SPI_GETSERIALKEYS = 62
    SPI_GETSHOWSOUNDS = 56
    SPI_GETSOUNDSENTRY = 64
    SPI_GETSTICKYKEYS = 58
    SPI_GETTOGGLEKEYS = 52
    SPI_GETWINDOWSEXTENSION = 92
    SPI_GETWORKAREA = 48
    SPI_ICONHORIZONTALSPACING = 13
    SPI_ICONVERTICALSPACING = 24
    SPI_LANGDRIVER = 12
    SPI_SCREENSAVERRUNNING = 97
    SPI_SETACCESSTIMEOUT = 61
    SPI_SETANIMATION = 73
    SPI_SETBEEP = 2
    SPI_SETBORDER = 6
    SPI_SETCURSORS = 87
    SPI_SETDEFAULTINPUTLANG = 90
    SPI_SETDESKPATTERN = 21
    SPI_SETDESKWALLPAPER = 20
    SPI_SETDOUBLECLICKTIME = 32
    SPI_SETDOUBLECLKHEIGHT = 30
    SPI_SETDOUBLECLKWIDTH = 29
    SPI_SETDRAGFULLWINDOWS = 37
    SPI_SETDRAGHEIGHT = 77
    SPI_SETDRAGWIDTH = 76
    SPI_SETFASTTASKSWITCH = 36
    SPI_SETFILTERKEYS = 51
    SPI_SETFONTSMOOTHING = 75
    SPI_SETGRIDGRANULARITY = 19
    SPI_SETHANDHELD = 78
    SPI_SETHIGHCONTRAST = 67
    SPI_SETICONMETRICS = 46
    SPI_SETICONS = 88
    SPI_SETICONTITLELOGFONT = 34
    SPI_SETICONTITLEWRAP = 26
    SPI_SETKEYBOARDDELAY = 23
    SPI_SETKEYBOARDPREF = 69
    SPI_SETKEYBOARDSPEED = 11
    SPI_SETLANGTOGGLE = 91
    SPI_SETLOWPOWERACTIVE = 85
    SPI_SETLOWPOWERTIMEOUT = 81
    SPI_SETMENUDROPALIGNMENT = 28
    SPI_SETMINIMIZEDMETRICS = 44
    SPI_SETMOUSE = 4
    SPI_SETMOUSEBUTTONSWAP = 33
    SPI_SETMOUSEKEYS = 55
    SPI_SETMOUSETRAILS = 93
    SPI_SETNONCLIENTMETRICS = 42
    SPI_SETPENWINDOWS = 49
    SPI_SETPOWEROFFACTIVE = 86
    SPI_SETPOWEROFFTIMEOUT = 82
    SPI_SETSCREENREADER = 71
    SPI_SETSCREENSAVEACTIVE = 17
    SPI_SETSCREENSAVETIMEOUT = 15
    SPI_SETSERIALKEYS = 63
    SPI_SETSHOWSOUNDS = 57
    SPI_SETSOUNDSENTRY = 65
    SPI_SETSTICKYKEYS = 59
    SPI_SETTOGGLEKEYS = 53
    SPI_SETWORKAREA = 47
End Enum

'\\ Windows Menu related API calls
Public Declare Function GetMenu Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetMenuCheckMarkDimensions Lib "user32" () As Long
Public Declare Function GetMenuContextHelpId Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Public Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Public Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Public Declare Function GetMenuItemInfo Lib "user32" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Boolean, lpMenuItemInfo As MENUITEMINFO) As Long
Public Declare Function GetMenuItemRect Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long, ByVal uItem As Long, lprcItem As RECT) As Long
Public Declare Function GetMenuState Lib "user32" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Public Declare Function GetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Public Declare Function SetMenu Lib "user32" (ByVal hWnd As Long, ByVal hMenu As Long) As Long
Public Declare Function SetMenuContextHelpId Lib "user32" (ByVal hMenu As Long, ByVal dw As Long) As Long
Public Declare Function SetMenuDefaultItem Lib "user32" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Public Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Public Declare Function SetMenuItemInfo Lib "user32" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal bool As Boolean, lpcMenuItemInfo As MENUITEMINFO) As Long
Public Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type

'\\ Sending messages to a window....
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageString Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lpstrParam As String) As Long

'\\ Windows hooks...
'SetWindowsHookEx
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long


'\\ Special structures passed by the hook proc
Public Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    y As Long
    x As Long
    style As Long
    lpszName As Long
    lpszClass As Long
    ExStyle As Long
End Type


Public Type CBTACTIVATESTRUCT
     fMouse As Long
     hWndActive As Long
End Type

Public Type CBT_CREATEWND
    lpcs As Long 'pointer to CREATESTRUCT
    hWndInsertAfter As Long
End Type

Public Type CBT_CREATEWND_FULL
    csThis As CREATESTRUCT
    hWndInsertAfter As Long
End Type


'\\ Shell execute - running external applications
'\\ Declaration
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'\\ API Error decoding
Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

'\\ Scrollbar messages
Public Enum enScrollMessages
    SB_ENDSCROLL = 8
    SB_TOP = 6
    SB_LINELEFT = 0
    SB_LINERIGHT = 1
    SB_LINEDOWN = 1
    SB_LINEUP = 0
    SB_PAGEDOWN = 3
    SB_PAGELEFT = 2
    SB_PAGERIGHT = 3
    SB_PAGEUP = 2
    SB_RIGHT = 7
    SB_LEFT = 6
    SB_THUMBPOSITION = 4
    SB_THUMBTRACK = 5
End Enum

'\\ System command messages
Public Enum enSystemCommands
    SC_ARRANGE = &HF110
    SC_CLOSE = &HF060
    SC_HOTKEY = &HF150
    SC_HSCROLL = &HF080
    SC_KEYMENU = &HF100
    SC_MAXIMIZE = &HF030
    SC_MINIMIZE = &HF020
    SC_MOUSEMENU = &HF090
    SC_MOVE = &HF010
    SC_NEXTWINDOW = &HF040
    SC_PREVWINDOW = &HF050
    SC_RESTORE = &HF120
    SC_SCREENSAVE = &HF140
    SC_SIZE = &HF000
    SC_TASKLIST = &HF130
    SC_VSCROLL = &HF070
End Enum


'\\ ShowWindow constants
Public Enum enShowWindow
    SW_ERASE = &H4
    SW_HIDE = 0
    SW_INVALIDATE = &H2
    SW_MAX = 10
    SW_MAXIMIZE = 3
    SW_MINIMIZE = 6
    SW_NORMAL = 1
    SW_OTHERUNZOOM = 4
    SW_OTHERZOOM = 2
    SW_PARENTCLOSING = 1
    SW_PARENTOPENING = 3
    SW_RESTORE = 9
    SW_SCROLLCHILDREN = &H1
    SW_SHOW = 5
    SW_SHOWDEFAULT = 10
    SW_SHOWMAXIMIZED = 3
    SW_SHOWMINIMIZED = 2
    SW_SHOWMINNOACTIVE = 7
    SW_SHOWNA = 8
    SW_SHOWNOACTIVATE = 4
    SW_SHOWNORMAL = 1
End Enum

'\\ Window Style
Public Enum enWindowStyles
    WS_BORDER = &H800000
    WS_CAPTION = &HC00000                  '  WS_BORDER Or WS_DLGFRAME
    WS_CHILD = &H40000000
    WS_CLIPCHILDREN = &H2000000
    WS_CLIPSIBLINGS = &H4000000
    WS_DISABLED = &H8000000
    WS_DLGFRAME = &H400000
    WS_EX_ACCEPTFILES = &H10&
    WS_EX_DLGMODALFRAME = &H1&
    WS_EX_NOPARENTNOTIFY = &H4&
    WS_EX_TOPMOST = &H8&
    WS_EX_TRANSPARENT = &H20&
    WS_GROUP = &H20000
    WS_MAXIMIZE = &H1000000
    WS_MAXIMIZEBOX = &H10000
    WS_MINIMIZE = &H20000000
    WS_MINIMIZEBOX = &H20000
    WS_OVERLAPPED = &H0&
    WS_POPUP = &H80000000
    WS_SYSMENU = &H80000
    WS_TABSTOP = &H10000
    WS_THICKFRAME = &H40000
    WS_VISIBLE = &H10000000
    WS_VSCROLL = &H200000
End Enum

Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Enum enHWND
    HWND_BOTTOM = 1
    HWND_BROADCAST = &HFFFF&
    HWND_DESKTOP = 0
    HWND_NOTOPMOST = -2
    HWND_TOP = 0
    HWND_TOPMOST = -1
End Enum

Public Enum enSetWindowPos
    SWP_FRAMECHANGED = &H20        '  The frame changed: send WM_NCCALCSIZE
    SWP_HIDEWINDOW = &H80
    SWP_NOACTIVATE = &H10
    SWP_NOCOPYBITS = &H100
    SWP_NOMOVE = &H2
    SWP_NOOWNERZORDER = &H200      '  Don't do owner Z ordering
    SWP_NOREDRAW = &H8
    SWP_NOSIZE = &H1
    SWP_NOZORDER = &H4
    SWP_SHOWWINDOW = &H40
End Enum

'\\ Pointer validation in StringFromPointer
Declare Function IsBadStringPtrByLong Lib "kernel32" Alias "IsBadStringPtrA" (ByVal lpsz As Long, ByVal ucchMax As Long) As Long

'\\ Module handles....
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Type CWPSTRUCT
    lParam As Long
    wParam As Long
    Message As Long
    hWnd As Long
End Type

Type CWPRETSTRUCT
    lResult As Long
    lParam As Long
    wParam As Long
    Message As Long
    hWnd As Long
End Type

Type DEBUGHOOKINFO
    hModuleHook As Long
    Reserved As Long
    lParam As Long
    wParam As Long
    code As Long
End Type

Type MSG
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Type EVENTMSG
    Message As Long
    paramL As Long
    paramH As Long
    time As Long
    hWnd As Long
End Type

Type MOUSEHOOKSTRUCT
        pt As POINTAPI
        hWnd As Long
        wHitTestCode As Long
        dwExtraInfo As Long
End Type







'\\ --[sGetMessageName]---------------------------------------------------------------------
'\\ Returns the text name of a windows message from its long number...
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function sGetMessageName(ByVal nMsg As Long) As String

Select Case nMsg
Case WM_ACTIVATE
    sGetMessageName = "WM_ACTIVATE"
Case WM_ACTIVATEAPP
    sGetMessageName = "WM_ACTIVATEAPP"
Case WM_ASKCBFORMATNAME
    sGetMessageName = "WM_ASKCBFORMATNAME"
Case WM_CANCELJOURNAL
    sGetMessageName = "WM_CANCELJOURNAL"
Case WM_CANCELMODE
    sGetMessageName = "WM_CANCELMODE"
Case WM_CAPTURECHANGED
    sGetMessageName = "WM_CAPTURECHANGED"
Case WM_CHANGECBCHAIN
    sGetMessageName = "WM_CHANGECBCHAIN"
Case WM_CHAR
    sGetMessageName = "WM_CHAR"
Case WM_CHARTOITEM
    sGetMessageName = "WM_CHARTOITEM"
Case WM_CHILDACTIVATE
    sGetMessageName = "WM_CHILDACTIVATE"
Case WM_CHOOSEFONT_GETLOGFONT
    sGetMessageName = "WM_CHOOSEFONT_GETLOGFONT"
Case WM_CHOOSEFONT_SETFLAGS
    sGetMessageName = "WM_CHOOSEFONT_SETFLAGS"
Case WM_CHOOSEFONT_SETLOGFONT
    sGetMessageName = "WM_CHOOSEFONT_SETLOGFONT"
Case WM_CLEAR
    sGetMessageName = "WM_CLEAR"
Case WM_CLOSE
    sGetMessageName = "WM_CLOSE"
Case WM_COMMAND
    sGetMessageName = "WM_COMMAND"
Case WM_COMPACTING
    sGetMessageName = "WM_COMPACTING"
Case WM_COMPAREITEM
    sGetMessageName = "WM_COMPAREITEM"
Case WM_CONVERTREQUESTEX
    sGetMessageName = "WM_CONVERTREQUESTEX"
Case WM_COPY
    sGetMessageName = "WM_COPY"
Case WM_COPYDATA
    sGetMessageName = "WM_COPYDATA"
Case WM_CREATE
    sGetMessageName = "WM_CREATE"
Case WM_CTLCOLORBTN
    sGetMessageName = "WM_CTLCOLORBTN"
Case WM_CTLCOLORDLG
    sGetMessageName = "WM_CTLCOLORDLG"
Case WM_CTLCOLOREDIT
    sGetMessageName = "WM_CTLCOLOREDIT"
Case WM_CTLCOLORLISTBOX
    sGetMessageName = "WM_CTLCOLORLISTBOX"
Case WM_CTLCOLORMSGBOX
    sGetMessageName = "WM_CTLCOLORMSGBOX"
Case WM_CTLCOLORSCROLLBAR
    sGetMessageName = "WM_CTLCOLORSCROLLBAR"
Case WM_CTLCOLORSTATIC
    sGetMessageName = "WM_CTLCOLORSTATIC"
Case WM_CUT
    sGetMessageName = "WM_CUT"
Case WM_DDE_ACK
    sGetMessageName = "WM_DDE_ACK"
Case WM_DDE_ADVISE
    sGetMessageName = "WM_DDE_ADVISE"
Case WM_DDE_DATA
    sGetMessageName = "WM_DDE_DATA"
Case WM_DDE_EXECUTE
    sGetMessageName = "WM_DDE_EXECUTE"
Case WM_DDE_FIRST
    sGetMessageName = "WM_DDE_FIRST"
Case WM_DDE_INITIATE
    sGetMessageName = "WM_DDE_INITIATE"
Case WM_DDE_LAST
    sGetMessageName = "WM_DDE_LAST"
Case WM_DDE_POKE
    sGetMessageName = "WM_DDE_POKE"
Case WM_DDE_REQUEST
    sGetMessageName = "WM_DDE_REQUEST"
Case WM_DDE_TERMINATE
    sGetMessageName = "WM_DDE_TERMINATE"
Case WM_DDE_UNADVISE
    sGetMessageName = "WM_DDE_UNADVISE"
Case WM_DEADCHAR
    sGetMessageName = "WM_DEADCHAR"
Case WM_DELETEITEM
    sGetMessageName = "WM_DELETEITEM"
Case WM_DESTROY
    sGetMessageName = "WM_DESTROY"
Case WM_DESTROYCLIPBOARD
    sGetMessageName = "WM_DESTROYCLIPBOARD"
Case WM_DEVMODECHANGE
    sGetMessageName = "WM_DEVMODECHANGE"
Case WM_DRAWCLIPBOARD
    sGetMessageName = "WM_DRAWCLIPBOARD"
Case WM_DRAWITEM
    sGetMessageName = "WM_DRAWITEM"
Case WM_DROPFILES
    sGetMessageName = "WM_DROPFILES"
Case WM_ENABLE
    sGetMessageName = "WM_ENABLE"
Case WM_ENDSESSION
    sGetMessageName = "WM_ENDSESSION"
Case WM_ENTERIDLE
    sGetMessageName = "WM_ENTERIDLE"
Case WM_ENTERMENULOOP
    sGetMessageName = "WM_ENTERMENULOOP"
Case WM_ERASEBKGND
    sGetMessageName = "WM_ERASEBKGND"
Case WM_EXITMENULOOP
    sGetMessageName = "WM_EXITMENULOOP"
Case WM_FONTCHANGE
    sGetMessageName = "WM_FONTCHANGE"
Case WM_GETDLGCODE
    sGetMessageName = "WM_GETDLGCODE"
Case WM_GETFONT
    sGetMessageName = "WM_GETFONT"
Case WM_GETHOTKEY
    sGetMessageName = "WM_GETHOTKEY"
Case WM_GETMINMAXINFO
    sGetMessageName = "WM_GETMINMAXINFO"
Case WM_GETTEXT
    sGetMessageName = "WM_GETTEXT"
Case WM_GETTEXTLENGTH
    sGetMessageName = "WM_GETTEXTLENGTH"
Case WM_HOTKEY
    sGetMessageName = "WM_HOTKEY"
Case WM_HSCROLL
    sGetMessageName = "WM_HSCROLL"
Case WM_HSCROLLCLIPBOARD
    sGetMessageName = "WM_HSCROLLCLIPBOARD"
Case WM_ICONERASEBKGND
    sGetMessageName = "WM_ICONERASEBKGND"
Case WM_IME_CHAR
    sGetMessageName = "WM_IME_CHAR"
Case WM_IME_COMPOSITION
    sGetMessageName = "WM_IME_COMPOSITION"
Case WM_IME_COMPOSITIONFULL
    sGetMessageName = "WM_IME_COMPOSITIONFULL"
Case WM_IME_CONTROL
    sGetMessageName = "WM_IME_CONTROL"
Case WM_IME_ENDCOMPOSITION
    sGetMessageName = "WM_IME_ENDCOMPOSITION"
Case WM_IME_KEYDOWN
    sGetMessageName = "WM_IME_KEYDOWN"
Case WM_IME_KEYLAST
    sGetMessageName = "WM_IME_KEYLAST"
Case WM_IME_KEYUP
    sGetMessageName = "WM_IME_KEYUP"
Case WM_IME_NOTIFY
    sGetMessageName = "WM_IME_NOTIFY"
Case WM_IME_SELECT
    sGetMessageName = "WM_IME_SELECT"
Case WM_IME_SETCONTEXT
    sGetMessageName = "WM_IME_SETCONTEXT"
Case WM_IME_STARTCOMPOSITION
    sGetMessageName = "WM_IME_STARTCOMPOSITION"
Case WM_INITDIALOG
    sGetMessageName = "WM_INITDIALOG"
Case WM_INITMENU
    sGetMessageName = "WM_INITMENU"
Case WM_INITMENUPOPUP
    sGetMessageName = "WM_INITMENUPOPUP"
Case WM_KEYDOWN
    sGetMessageName = "WM_KEYDOWN"
Case WM_KEYFIRST
    sGetMessageName = "WM_KEYFIRST"
Case WM_KEYLAST
    sGetMessageName = "WM_KEYLAST"
Case WM_KEYUP
    sGetMessageName = "WM_KEYUP"
Case WM_KILLFOCUS
    sGetMessageName = "WM_KILLFOCUS"
Case WM_LBUTTONDBLCLK
    sGetMessageName = "WM_LBUTTONDBLCLK"
Case WM_LBUTTONDOWN
    sGetMessageName = "WM_LBUTTONDOWN"
Case WM_LBUTTONUP
    sGetMessageName = "WM_LBUTTONUP"
Case WM_MBUTTONDBLCLK
    sGetMessageName = "WM_MBUTTONDBLCLK"
Case WM_MBUTTONDOWN
    sGetMessageName = "WM_MBUTTONDOWN"
Case WM_MBUTTONUP
    sGetMessageName = "WM_MBUTTONUP"
Case WM_MDIACTIVATE
    sGetMessageName = "WM_MDIACTIVATE"
Case WM_MDICASCADE
    sGetMessageName = "WM_MDICASCADE"
Case WM_MDICREATE
    sGetMessageName = "WM_MDICREATE"
Case WM_MDIDESTROY
    sGetMessageName = "WM_MDIDESTROY"
Case WM_MDIGETACTIVE
    sGetMessageName = "WM_MDIGETACTIVE"
Case WM_MDIICONARRANGE
    sGetMessageName = "WM_MDIICONARRANGE"
Case WM_MDIMAXIMIZE
    sGetMessageName = "WM_MDIMAXIMIZE"
Case WM_MDINEXT
    sGetMessageName = "WM_MDINEXT"
Case WM_MDIREFRESHMENU
    sGetMessageName = "WM_MDIREFRESHMENU"
Case WM_MDIRESTORE
    sGetMessageName = "WM_MDIRESTORE"
Case WM_MDISETMENU
    sGetMessageName = "WM_MDISETMENU"
Case WM_MDITILE
    sGetMessageName = "WM_MDITILE"
Case WM_MEASUREITEM
    sGetMessageName = "WM_MEASUREITEM"
Case WM_MENUCHAR
    sGetMessageName = "WM_MENUCHAR"
Case WM_MENUSELECT
    sGetMessageName = "WM_MENUSELECT"
Case WM_MOUSEACTIVATE
    sGetMessageName = "WM_MOUSEACTIVATE"
Case WM_MOUSEFIRST
    sGetMessageName = "WM_MOUSEFIRST"
Case WM_MOUSELAST
    sGetMessageName = "WM_MOUSELAST"
Case WM_MOUSEMOVE
    sGetMessageName = "WM_MOUSEMOVE"
Case WM_MOVE  ' &H3
    sGetMessageName = "WM_MOVE"
Case WM_NCACTIVATE  ' &H86
    sGetMessageName = "WM_NCACTIVATE"
Case WM_NCCALCSIZE  ' &H83
    sGetMessageName = "WM_NCCALCSIZE"
Case WM_NCCREATE  ' &H81
    sGetMessageName = "WM_NCCREATE"
Case WM_NCDESTROY  ' &H82
    sGetMessageName = "WM_NCDESTROY"
Case WM_NCHITTEST  ' &H84
    sGetMessageName = "WM_NCHITTEST"
Case WM_NCLBUTTONDBLCLK  ' &HA3
    sGetMessageName = "WM_NCLBUTTONDBLCLK"
Case WM_NCLBUTTONDOWN  ' &HA1
    sGetMessageName = "WM_NCLBUTTONDOWN"
Case WM_NCLBUTTONUP  ' &HA2
    sGetMessageName = "WM_NCLBUTTONUP"
Case WM_NCMBUTTONDBLCLK  ' &HA9
    sGetMessageName = "WM_NCMBUTTONDBLCLK"
Case WM_NCMBUTTONDOWN  ' &HA7
    sGetMessageName = "WM_NCMBUTTONDOWN"
Case WM_NCMBUTTONUP  ' &HA8
    sGetMessageName = "WM_NCMBUTTONUP"
Case WM_NCMOUSEMOVE  ' &HA0
    sGetMessageName = "WM_NCMOUSEMOVE"
Case WM_NCPAINT  ' &H85
    sGetMessageName = "WM_NCPAINT"
Case WM_NCRBUTTONDBLCLK  ' &HA6
    sGetMessageName = "WM_NCRBUTTONDBLCLK"
Case WM_NCRBUTTONDOWN  ' &HA4
    sGetMessageName = "WM_NCRBUTTONDOWN"
Case WM_NCRBUTTONUP  ' &HA5
    sGetMessageName = "WM_NCRBUTTONUP"
Case WM_NEXTDLGCTL  ' &H28
    sGetMessageName = "WM_NEXTDLGCTL"
Case WM_NULL  ' &H0
    sGetMessageName = "WM_NULL"
Case WM_PAINT  ' &HF
    sGetMessageName = "WM_PAINT"
Case WM_PAINTCLIPBOARD  ' &H309
    sGetMessageName = "WM_PAINTCLIPBOARD"
Case WM_PAINTICON  ' &H26
    sGetMessageName = "WM_PAINTICON"
Case WM_PALETTECHANGED  ' &H311
    sGetMessageName = "WM_PALETTECHANGED"
Case WM_PALETTEISCHANGING  ' &H310
    sGetMessageName = "WM_PALETTEISCHANGING"
Case WM_PARENTNOTIFY  ' &H210
    sGetMessageName = "WM_PARENTNOTIFY"
Case WM_PASTE  ' &H302
    sGetMessageName = "WM_PASTE"
Case WM_PENWINFIRST  ' &H380
    sGetMessageName = "WM_PENWINFIRST"
Case WM_PENWINLAST  ' &H38F
    sGetMessageName = "WM_PENWINLAST"
Case WM_POWER  ' &H48
    sGetMessageName = "WM_POWER"
Case WM_PSD_ENVSTAMPRECT  ' (&H400 + 5)
    sGetMessageName = "WM_PSD_ENVSTAMPRECT"
Case WM_PSD_FULLPAGERECT  ' (&H400 + 1)
    sGetMessageName = "WM_PSD_FULLPAGERECT"
Case WM_PSD_GREEKTEXTRECT  ' (&H400 + 4)
    sGetMessageName = "WM_PSD_GREEKTEXTRECT"
Case WM_PSD_MARGINRECT  ' (&H400 + 3)
    sGetMessageName = "WM_PSD_MARGINRECT"
Case WM_PSD_MINMARGINRECT  ' (&H400 + 2)
    sGetMessageName = "WM_PSD_MINMARGINRECT"
Case WM_PSD_PAGESETUPDLG  ' (&H400)
    sGetMessageName = "WM_PSD_PAGESETUPDLG"
Case WM_PSD_YAFULLPAGERECT  ' (&H400 + 6)
    sGetMessageName = "WM_PSD_YAFULLPAGERECT"
Case WM_QUERYDRAGICON  ' &H37
    sGetMessageName = "WM_QUERYDRAGICON"
Case WM_QUERYENDSESSION  ' &H11
    sGetMessageName = "WM_QUERYENDSESSION"
Case WM_QUERYNEWPALETTE  ' &H30F
    sGetMessageName = "WM_QUERYNEWPALETTE"
Case WM_QUERYOPEN  ' &H13
    sGetMessageName = "WM_QUERYOPEN"
Case WM_QUEUESYNC  ' &H23
    sGetMessageName = "WM_QUEUESYNC"
Case WM_QUIT  ' &H12
    sGetMessageName = "WM_QUIT"
Case WM_RBUTTONDBLCLK  ' &H206
    sGetMessageName = "WM_RBUTTONDBLCLK"
Case WM_RBUTTONDOWN  ' &H204
    sGetMessageName = "WM_RBUTTONDOWN"
Case WM_RBUTTONUP  ' &H205
    sGetMessageName = "WM_RBUTTONUP"
Case WM_RENDERALLFORMATS  ' &H306
    sGetMessageName = "WM_RENDERALLFORMATS"
Case WM_RENDERFORMAT  ' &H305
    sGetMessageName = "WM_RENDERFORMAT"
Case WM_SETCURSOR  ' &H20
    sGetMessageName = "WM_SETCURSOR"
Case WM_SETFOCUS  ' &H7
    sGetMessageName = "WM_SETFOCUS"
Case WM_SETFONT  ' &H30
    sGetMessageName = "WM_SETFONT"
Case WM_SETHOTKEY  ' &H32
    sGetMessageName = "WM_SETHOTKEY"
Case WM_SETREDRAW  ' &HB
    sGetMessageName = "WM_SETREDRAW"
Case WM_SETTEXT  ' &HC
    sGetMessageName = "WM_SETTEXT"
Case WM_SETTINGCHANGE  ' &H1A
    sGetMessageName = "WM_SETTINGCHANGE"
Case WM_SHOWWINDOW  ' &H18
    sGetMessageName = "WM_SHOWWINDOW"
Case WM_SIZE  ' &H5
    sGetMessageName = "WM_SIZE"
Case WM_SIZECLIPBOARD  ' &H30B
    sGetMessageName = "WM_SIZECLIPBOARD"
Case WM_SPOOLERSTATUS  ' &H2A
    sGetMessageName = "WM_SPOOLERSTATUS"
Case WM_SYSCHAR  ' &H106
    sGetMessageName = "WM_SYSCHAR"
Case WM_SYSCOLORCHANGE  ' &H15
    sGetMessageName = "WM_SYSCOLORCHANGE"
Case WM_SYSCOMMAND  ' &H112
    sGetMessageName = "WM_SYSCOMMAND"
Case WM_SYSDEADCHAR  ' &H107
    sGetMessageName = "WM_SYSDEADCHAR"
Case WM_SYSKEYDOWN  ' &H104
    sGetMessageName = "WM_SYSKEYDOWN"
Case WM_SYSKEYUP  ' &H105
    sGetMessageName = "WM_SYSKEYUP"
Case WM_TIMECHANGE  ' &H1E
    sGetMessageName = "WM_TIMECHANGE"
Case WM_TIMER  ' &H113
    sGetMessageName = "WM_TIMER"
Case WM_UNDO  ' &H304
    sGetMessageName = "WM_UNDO"
Case WM_USER  ' &H400
    sGetMessageName = "WM_USER"
Case WM_VKEYTOITEM  ' &H2E
    sGetMessageName = "WM_VKEYTOITEM"
Case WM_VSCROLL  ' &H115
    sGetMessageName = "WM_VSCROLL"
Case WM_VSCROLLCLIPBOARD  ' &H30A
    sGetMessageName = "WM_VSCROLLCLIPBOARD"
Case WM_WINDOWPOSCHANGED  ' &H47
    sGetMessageName = "WM_WINDOWPOSCHANGED"
Case WM_WINDOWPOSCHANGING  ' &H46
    sGetMessageName = "WM_WINDOWPOSCHANGING"
Case WM_WININICHANGE  ' &H1A
    sGetMessageName = "WM_WININICHANGE"
Case Else
    sGetMessageName = "UNKNOWN MESSAGE : " & Hex(nMsg)
End Select

End Function

'\\ --[sGetShowWindowName]---------------------------------------------------------------------
'\\ Returns the text name of a show windpow constant from its long number...
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function sGetShowWindowName(nSWNum As Long) As String

Select Case nSWNum
Case SW_ERASE
    sGetShowWindowName = "SW_ERASE"
Case SW_HIDE
    sGetShowWindowName = "SW_HIDE"
Case SW_INVALIDATE
    sGetShowWindowName = "SW_INVALIDATE"
Case SW_MAX
    sGetShowWindowName = "SW_MAX"
Case SW_MAXIMIZE
    sGetShowWindowName = "SW_MAXIMIZE"
Case SW_MINIMIZE
    sGetShowWindowName = "SW_MINIMIZE"
Case SW_NORMAL
    sGetShowWindowName = "SW_NORMAL"
Case SW_OTHERUNZOOM
    sGetShowWindowName = "SW_OTHERUNZOOM"
Case SW_OTHERZOOM
    sGetShowWindowName = "SW_OTHERZOOM"
Case SW_PARENTCLOSING
    sGetShowWindowName = "SW_PARENTCLOSING"
Case SW_PARENTOPENING
    sGetShowWindowName = "SW_PARENTOPENING"
Case SW_RESTORE
    sGetShowWindowName = "SW_RESTORE"
Case SW_SCROLLCHILDREN
    sGetShowWindowName = "SW_SCROLLCHILDREN"
Case SW_SHOW
    sGetShowWindowName = "SW_SHOW"
Case SW_SHOWDEFAULT
    sGetShowWindowName = "SW_SHOWDEFAULT"
Case SW_SHOWMAXIMIZED
    sGetShowWindowName = "SW_SHOWMAXIMIZED"
Case SW_SHOWMINIMIZED
    sGetShowWindowName = "SW_SHOWMINIMIZED"
Case SW_SHOWMINNOACTIVE
    sGetShowWindowName = "SW_SHOWMINNOACTIVE"
Case SW_SHOWNA
    sGetShowWindowName = "SW_SHOWNA"
Case SW_SHOWNOACTIVATE
    sGetShowWindowName = "SW_SHOWNOACTIVATE"
Case SW_SHOWNORMAL
    sGetShowWindowName = "SW_SHOWNORMAL"
Case Else
    sGetShowWindowName = "UNKNOWN SW_ CONSTANT: " & Hex(nSWNum)
End Select

End Function

'\\ --[sGetSystemCommandName]---------------------------------------------------------------------
'\\ Returns the text name of a system command message from its long number...
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function sGetSystemCommandName(ByVal nCmd As Long) As String

nCmd = LoWord(nCmd)

Select Case True
Case nCmd = LoWord(SC_ARRANGE)
    sGetSystemCommandName = "SC_ARRANGE"
Case nCmd = LoWord(SC_CLOSE)
    sGetSystemCommandName = "SC_CLOSE"
Case nCmd = LoWord(SC_HOTKEY)
    sGetSystemCommandName = "SC_HOTKEY"
Case nCmd = LoWord(SC_HSCROLL)
    sGetSystemCommandName = "SC_HSCROLL"
Case nCmd = LoWord(SC_KEYMENU)
    sGetSystemCommandName = "SC_KEYMENU"
Case nCmd = LoWord(SC_MAXIMIZE)
    sGetSystemCommandName = "SC_MAXIMIZE"
Case nCmd = LoWord(SC_MINIMIZE)
    sGetSystemCommandName = "SC_MINIMIZE"
Case nCmd = LoWord(SC_MOUSEMENU)
    sGetSystemCommandName = "SC_MOUSEMENU"
Case nCmd = LoWord(SC_MOVE)
    sGetSystemCommandName = "SC_MOVE"
Case nCmd = LoWord(SC_NEXTWINDOW)
    sGetSystemCommandName = "SC_NEXTWINDOW"
Case nCmd = LoWord(SC_PREVWINDOW)
    sGetSystemCommandName = "SC_PREVWINDOW"
Case nCmd = LoWord(SC_RESTORE)
    sGetSystemCommandName = "SC_RESTORE"
Case nCmd = LoWord(SC_SCREENSAVE)
    sGetSystemCommandName = "SC_SCREENSAVE"
Case nCmd = LoWord(SC_SIZE)
    sGetSystemCommandName = "SC_SIZE"
Case nCmd = LoWord(SC_TASKLIST)
    sGetSystemCommandName = "SC_TASKLIST"
Case nCmd = LoWord(SC_VSCROLL)
    sGetSystemCommandName = "SC_VSCROLL"
Case Else
     sGetSystemCommandName = "UNKNOWN SYSTEM COMMAND : " & Hex(LoWord(nCmd))
End Select

End Function


'\\ --[SubclassDialog]---------------------------------------------------------------
'\\ Subclasses the dialog proc. whose hwnd is given with the proc at address
'\\ lpfnDlgRox
'\\ ---------------------------------------------------------------------------------
Public Sub SubclassDialog(hwndDlg As Long, lpfnDlgProc As Long)

Dim lRet As Long

lRet = SetWindowLong(hwndDlg, DWL_DLGPROC, lpfnDlgProc)
If lRet > 0 Then
    Eventhandler.hOldDlgProc = lRet
End If

End Sub


'\\ --[VB_GetClassname]---------------------------------------------------------------------
'\\ Returns the class name of the given window
'\\ Parameters:
'\\    hwndTest - The window handle of the window for which you require the class name.
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function VB_GetClassname(ByVal hwndTest As Long) As String

Dim lRet As Long
Dim sClassname As String

'\\ Get the window class name....
sClassname = String$(256, 0)
lRet = GetClassName(hwndTest, sClassname, 255)
If lRet > 0 Then
    sClassname = Left$(sClassname, lRet)
End If

VB_GetClassname = sClassname

End Function

Public Function sGetWindowInformation(hwndTest As Long) As String

Dim sRet As String
Dim sClassname As String
Dim sWindowText As String

Dim sNL As String
sNL = Chr$(13) & Chr$(10)

sClassname = VB_GetClassname(hwndTest)
sWindowText = VB_GetWindowText(hwndTest)

'\\ Compose the information together...
sRet = "Classname : " & sClassname & sNL
sRet = sRet & "WindowText : " & sWindowText & sNL
If VB_IsWindowClassSet(hwndTest, CS_BYTEALIGNCLIENT) Then sRet = sRet & " + BYTEALIGNCLIENT " & sNL
If VB_IsWindowClassSet(hwndTest, CS_BYTEALIGNWINDOW) Then sRet = sRet & " + BYTEALIGNWINDOW " & sNL
If VB_IsWindowClassSet(hwndTest, CS_CLASSDC) Then sRet = sRet & " + CLASSDC " & sNL
If VB_IsWindowClassSet(hwndTest, CS_OWNDC) Then sRet = sRet & " + OWNDC " & sNL
If VB_IsWindowClassSet(hwndTest, CS_PARENTDC) Then sRet = sRet & " + PARENTDC " & sNL
If VB_IsWindowClassSet(hwndTest, CS_DBLCLKS) Then sRet = sRet & " + DBLCLKS " & sNL
If VB_IsWindowClassSet(hwndTest, CS_HREDRAW) Then sRet = sRet & " + HREDRAW " & sNL
If VB_IsWindowClassSet(hwndTest, CS_VREDRAW) Then sRet = sRet & " + VREDRAW " & sNL
If VB_IsWindowClassSet(hwndTest, CS_INSERTCHAR) Then sRet = sRet & " + INSERTCHAR " & sNL
If VB_IsWindowClassSet(hwndTest, CS_KEYCVTWINDOW) Then sRet = sRet & " + KEYCVTWINDOW " & sNL
If VB_IsWindowClassSet(hwndTest, CS_NOCLOSE) Then sRet = sRet & " + NOCLOSE " & sNL
If VB_IsWindowClassSet(hwndTest, CS_NOKEYCVT) Then sRet = sRet & " + NOKEYCVT " & sNL
If VB_IsWindowClassSet(hwndTest, CS_NOMOVECARET) Then sRet = sRet & " + NOMOVECARET " & sNL
If VB_IsWindowClassSet(hwndTest, CS_PUBLICCLASS) Then sRet = sRet & " + PUBLICCLASS " & sNL
If VB_IsWindowClassSet(hwndTest, CS_SAVEBITS) Then sRet = sRet & " + SAVEBITS " & sNL

sGetWindowInformation = sRet

End Function

'\\ --[VB_GetWindowText]---------------------------------------------------------------------
'\\ Returns the window text of the given window
'\\ Parameters:
'\\    hwndTest - The window handle of the window for which you require the text.
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function VB_GetWindowText(ByVal hwndTest As Long) As String

Dim lRet As Long
Dim sRet As String

lRet = GetWindowTextLength(hwndTest)
If lRet > 0 Then
    sRet = String$(lRet + 1, 0)
    lRet = GetWindowText(hwndTest, sRet, Len(sRet))
    '\\ Returns length up to NULL terminator
    If lRet > 0 Then
        sRet = Left$(sRet, lRet)
    End If
End If

VB_GetWindowText = sRet

End Function



'\\ --[VB_IsWindowClassSet]---------------------------------------------------------------------
'\\ Returns true if the given window has the given class style set
'\\ Parameters:
'\\    hwndTest - The window handle of the window for which you require the text.
'\\    ClassStyle - The class style you are testing for
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function VB_IsWindowClassSet(ByVal hwndTest As Long, ClassStyle As ClassStyleConstants) As Boolean

Dim lRet As Long
lRet = GetClassLong(hwndTest, GCL_STYLE)

VB_IsWindowClassSet = (lRet And ClassStyle)

End Function


'\\ --[HiWord]-----------------------------------------------------------------------------
'\\ Returns the high word component of a long value
'\\ Parameters:
'\\   dw - The long of which we need the HiWord
'\\
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function HiWord(dw As Long) As Integer
 If dw And &H80000000 Then
      HiWord = (dw \ 65535) - 1
 Else
    HiWord = dw \ 65535
 End If
End Function

'\\ --[LoByte]-----------------------------------------------------------------------------
'\\ Returns the low byte component of an integer value
'\\ Parameters:
'\\   w - The integer of which we need the loWord
'\\
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function LoByte(w As Integer) As Byte
 LoByte = w And &HFF
End Function

'\\ --[LoWord]-----------------------------------------------------------------------------
'\\ Returns the low word component of a long value
'\\ Parameters:
'\\   dw - The long of which we need the LoWord
'\\
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function LoWord(dw As Long) As Integer
  If dw And &H8000& Then
      LoWord = &H8000 Or (dw And &H7FFF&)
   Else
      LoWord = dw And &HFFFF&
   End If
End Function

'\\ --[HiByte]-----------------------------------------------------------------------------
'\\ Returns the high byte component of an integer
'\\ Parameters:
'\\   w - The integer of which we need the HiByte
'\\
'\\ ----------------------------------------------------------------------------------------
'\\ You have a royalty free right to use, reproduce, modify, publish and mess with this code
'\\ I'd like you to visit http://www.merrioncomputing.com for updates, but won't force you
'\\ ----------------------------------------------------------------------------------------
Public Function HiByte(ByVal w As Integer) As Byte
   If w And &H8000 Then
      HiByte = &H80 Or ((w And &H7FFF) \ &HFF)
   Else
      HiByte = w \ 256
    End If
End Function
