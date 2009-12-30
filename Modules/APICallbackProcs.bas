Attribute VB_Name = "APICallbackProcs"
'\\ --[APICallbackProcs]---------------------------------------------------------------
'\\ Series of visual basic functions whose addresses can be passed as lpfnProcAddress
'\\ parameter of windows API callback functions using the AddressOf operator.
'\\ NOTE:
'\\ When creating a new callback proc, don't forget to declare the parameters ByVal,
'\\ or VB's type conversion will fail with GPF consequences
'\\ -----------------------------------------------------------------------------------

'typedef BOOL (CALLBACK* GRAYSTRINGPROC)(HDC, LPARAM, int);
'typedef VOID (CALLBACK* SENDASYNCPROC)(HWND, UINT, DWORD, LRESULT);
'typedef BOOL (CALLBACK* PROPENUMPROCA)(HWND, LPCSTR, HANDLE);
'typedef BOOL (CALLBACK* PROPENUMPROCEXA)(HWND, LPSTR, HANDLE, DWORD);
'typedef int (CALLBACK* EDITWORDBREAKPROCA)(LPSTR lpch, int ichCurrent, int cch, int code);
'typedef BOOL (CALLBACK* NAMEENUMPROCA)(LPSTR, LPARAM);
'typedef BOOL (CALLBACK* ENUMRESTYPEPROC)(HMODULE hModule, LPTSTR lpType, LONG lParam);
'typedef BOOL (CALLBACK* ENUMRESNAMEPROC)(HMODULE hModule, LPCTSTR lpType, LPTSTR lpName, LONG lParam);
'typedef BOOL (CALLBACK* ENUMRESLANGPROC)(HMODULE hModule, LPCTSTR lpType, LPCTSTR lpName, WORD  wLanguage, LONG lParam);

Option Explicit

'\\ Application global variables....
Public Eventhandler As EnumHandler
Public APIDispenser As APIFunctions
Public AllSubclassedWindows As colSubclassedWindows
Public AllTopLevelWindows As Collection
Public AllInstalledLocales As Collection

'\\ Windows hooks...
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long

'\\ Enumerating windows....
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function IsWindowApi Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Long

'\\ Enumerating system locales
Private Enum SystemLocaleEnumerationTypes
     LCID_INSTALLED = &H1        ' ##  installed locale ids
     LCID_SUPPORTED = &H2        ' ##  supported locale ids
     LCID_ALTERNATE_SORTS = &H4        ' ##  alternate sort locale ids
End Enum
Private Declare Function EnumSystemLocales Lib "kernel32" Alias "EnumSystemLocalesA" (ByVal lpLocaleEnumProc As Long, ByVal dwFlags As Long) As Long

'\\ Speeding up error reporting
'\\ API Error decoding
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public Function GetLastSystemError() As String

Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Dim sError As String * 500 '\\ Preinitilise a string buffer to put any error message into
Dim lErrNum As Long
Dim lErrMsg As Long

lErrNum = Err.LastDllError

lErrMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lErrNum, 0, sError, Len(sError), 0)

GetLastSystemError = Trim(sError)

End Function


Public Sub Main()

Set APIDispenser = New APIFunctions

End Sub


Public Sub RefreshInstalledLocales()

Dim lret As Long

Set AllInstalledLocales = New Collection

lret = EnumSystemLocales(AddressOf VB_ENUMLOCALESPROC, LCID_INSTALLED)


End Sub

Public Sub RefreshTopLevelWindows()

Dim lret As Long

Set AllTopLevelWindows = New Collection

lret = EnumWindows(AddressOf TopLevelWndEnumProc, 0)
If Err.LastDllError <> 0 Then
    ReportError Err.LastDllError, "System:TopLevelWindows", GetLastSystemError
End If

End Sub

Public Sub ReportError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)

If APIDispenser Is Nothing Then
    Err.Raise Number, Source, Description
Else
    APIDispenser.RaiseError Number, Source, Description
End If

Err.Clear

End Sub
'\\ --[VB_DLGPROC]----------------------------------------------------------------------------
'\\ typedef BOOL (CALLBACK* DLGPROC)(HWND, UINT, WPARAM, LPARAM)
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_DLGPROC(ByVal hwnd As Long, ByVal uint As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Dim Params() As Variant

ReDim Params(1 To 5) As Variant
Params(1) = hwnd
Params(2) = uint
Params(3) = wParam
Params(4) = lParam
Params(5) = 0

If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent DLGPROC, Params()
End If

VB_DLGPROC = Params(5)

End Function

'\\ --[VB_EDITWORDBREAKPROCA]------------------------------------------------------------
'\\ 'typedef int (CALLBACK* EDITWORDBREAKPROCA)(LPSTR lpch, int ichCurrent, int cch, int code);
'\\ This gets called by an edit control when a line of text has filled up the available
'\\ space.
'\\ By default, a text edit box breaks on spaces.
'\\ (This version prevents numbers being broken up if the digit grouping sepeartor is a space.)
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_EDITWORDBREAKPROCA(ByVal lpch As Long, _
                                      ByVal ichCurrent As Long, _
                                      ByVal cch As Long, _
                                      ByVal Code As Long) As Long
                                      
On Local Error Resume Next

Dim sCharacters As String

Dim lCharPos As Long

sCharacters = StringFromPointer(lpch, 1024)

Select Case Code
Case WB_ISDELIMITER
    '\\ Edit control is asking if this character is a wordbreak char...
    '\\ Reply FALSE is it is not a space, or if the characters either side of it
    '\\ are numbers....
    If Mid$(sCharacters, ichCurrent, 1) = " " Then
        VB_EDITWORDBREAKPROCA = 1
        If (ichCurrent > 0) And (ichCurrent < Len(sCharacters)) Then
            If IsNumeric(Mid$(sCharacters, ichCurrent - 1, 1)) And IsNumeric(Mid$(sCharacters, ichCurrent + 1, 1)) Then
                VB_EDITWORDBREAKPROCA = 0
            End If
        End If
    Else
        VB_EDITWORDBREAKPROCA = 0
    End If

Case WB_LEFT
  '\\ Find the begining of a word to the left of this position....
  For lCharPos = ichCurrent To 1 Step -1
    If Mid$(sCharacters, lCharPos, 1) = " " Then
        If Not (IsNumeric(Mid$(sCharacters, lCharPos - 1, 1)) And IsNumeric(Mid$(sCharacters, lCharPos + 1, 1))) Then
            VB_EDITWORDBREAKPROCA = lCharPos
            Exit For
        End If
    End If
  Next lCharPos
  
Case WB_RIGHT
'\\ Find the begining of a word to the right of this position....
  For lCharPos = ichCurrent To Len(sCharacters)
    If Mid$(sCharacters, lCharPos, 1) = " " Then
        If Not (IsNumeric(Mid$(sCharacters, lCharPos - 1, 1)) And IsNumeric(Mid$(sCharacters, lCharPos + 1, 1))) Then
            VB_EDITWORDBREAKPROCA = lCharPos
            Exit For
        End If
    End If
  Next lCharPos
End Select

End Function


'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_EnumDesktops(ByVal lpstrName As String, ByVal lParam As Long) As Long

Dim Params() As Variant

ReDim Params(1 To 2) As Variant
Params(1) = lpstrName
Params(2) = lParam

If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent DESKTOPENUMPROC, Params()
End If

VB_EnumDesktops = 1

End Function

'\\ --[VB_ENUMLOCALESPROC]-----------------------------------------------------------
'\\ BOOL CALLBACK EnumLocalesProc(
'\\   LPTSTR lpLocaleString   // locale identifier string
'\\   };
'\\ ---------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing Ltd.  All rights reserved
'\\ ---------------------------------------------------------------------------------
Public Function VB_ENUMLOCALESPROC(ByVal lpLocales As Long) As Long

Dim sLocale As String
Dim LCid As Long

Dim lcThis As New ApiLocale

sLocale = StringFromPointer(lpLocales, 9)

LCid = Val("&H" & sLocale)
lcThis.LocaleId = LCid
AllInstalledLocales.Add lcThis, "&H" & sLocale

VB_ENUMLOCALESPROC = 1

End Function

'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_ENUMPROC(ByVal hwnd As Long, ByVal lpStrPropName As String, ByVal hHandle As Long) As Long

Dim Params() As Variant

ReDim Params(1 To 3) As Variant
Params(1) = hwnd
Params(2) = lpStrPropName
Params(3) = hHandle
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent PROPENUMPROC, Params()
End If

VB_ENUMPROC = 1

End Function

'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_ENUMPROCEX(ByVal hwnd As Long, ByVal lpStr As String, ByVal hHandle As Long, ByVal dWord As Long) As Long

Dim Params() As Variant

ReDim Params(1 To 3) As Variant
Params(1) = hwnd
Params(2) = lpStr
Params(3) = hHandle
Params(4) = dWord
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent PROPENUMPROC, Params()
End If

VB_ENUMPROCEX = 1

End Function

'\\ --[VB_ENUMRESLANGPROC]---------------------------------------------
'\\ Decl:
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_ENUMRESLANGPROC(ByVal hModule As Long, ByVal lpType As String, ByVal lpName As String, ByVal wLanguage As Long, ByVal lParam As Long) As Long

Dim Params() As Variant

ReDim Params(1 To 5) As Variant
Params(1) = hModule
Params(2) = lpType
Params(3) = lpName
Params(4) = wLanguage
Params(5) = lParam
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent ENUMRESLANGPROC, Params()
End If

End Function

'\\ --[VB_ENUMRESNAMEPROC]------------------------------------------------------------
'\\ Decl:
'\\
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_ENUMRESNAMEPROC(ByVal hModule As Long, ByVal lpType As String, _
                                    ByVal lpName As String, ByVal lParam As Long) As Long

Dim Params() As Variant

ReDim Params(1 To 4) As Variant
Params(1) = hModule
Params(2) = lpType
Params(3) = lpName
Params(4) = lParam
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent ENUMRESNAMEPROC, Params()
End If

End Function

'\\ --[VB_ENUMRESTYPEPROC]----------------------------------------------
'\\ Enumerates the resource types in a module
'\\ Decl:
'\\ BOOL (CALLBACK* ENUMRESTYPEPROC)(HMODULE hModule, LPTSTR lpType, LONG lParam);
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_ENUMRESTYPEPROC(ByVal hModule As Long, ByVal lpType As Long, ByVal lParam As Long) As Long

Dim Params() As Variant

ReDim Params(1 To 3) As Variant
Params(1) = hModule
Params(2) = lpType
Params(3) = lParam

If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent ENUMRESTYPEPROC, Params()
End If

End Function

'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_EnumWinstations(ByVal lpstrName As String, ByVal lParam As Long) As Long

Dim Params() As Variant

ReDim Params(1 To 2) As Variant
Params(1) = lpstrName
Params(2) = lParam

If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent WINSTATIONENUMPROC, Params()
End If

VB_EnumWinstations = 1

End Function

'\\ -[VB_TimerProc]------------------------------------------------------
'\\ 'typedef VOID (CALLBACK* TIMERPROC)(HWND, UINT, UINT, DWORD);
'\\ parameters:
'\\ hWnd - The window handle to which the timer is attached...
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Sub VB_TIMERPROC(ByVal hwnd As Long, _
                        ByVal uint1 As Long, _
                        ByVal nEventId As Long, _
                        ByVal dwParam As Long)

On Error Resume Next

Dim Params() As Variant

ReDim Params(1 To 4) As Variant
Params(1) = hwnd
Params(2) = uint1
Params(3) = nEventId
Params(4) = dwParam
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent TIMERPROC, Params()
End If

End Sub

'\\ --[VB_WindowProc]-------------------------------------------------------------------
'\\ 'typedef LRESULT (CALLBACK* WNDPROC)(HWND, UINT, WPARAM, LPARAM);
'\\ Parameters:
'\\   hwnd - window handle receiving message
'\\   wMsg - The window message (WM_..etc.)
'\\   wParam - First message parameter
'\\   lParam - Second message parameter
'\\ Note:
'\\    When subclassing a window proc using this, set the eventhandler's
'\\    hOldWndProc property to the window's previous window proc address.
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_WindowProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next
Dim lret As Long
Dim Params() As Variant
ReDim Params(1 To 5) As Variant

Params(1) = hwnd
Params(2) = wMsg
Params(3) = wParam
Params(4) = lParam
Params(5) = lret
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent WNDPROC, Params()
End If
lret = Params(5)

VB_WindowProc = lret

End Function

'\\ [VB_WndEnumProc]---------------------------------------------------------------------------
'\\ 'typedef BOOL (CALLBACK* WNDENUMPROC)(HWND, LPARAM);
'\\ Used in EnumWindows and EnumChildWindows
'\\ hwnd - Window handle of the enumerated window,
'\\ lparam - passed into the enumwindows proc by programmer...
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_WndEnumProc(ByVal hwnd As Long, ByVal lParam As Long) As Long

Dim Params() As Variant

'\\ 1 - Pack the param array.....
ReDim Params(1 To 2) As Variant
Params(1) = hwnd
Params(2) = lParam

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent WNDENUMPROC, Params()
End If

VB_WndEnumProc = 1

End Function

Private Function TopLevelWndEnumProc(ByVal hwnd As Long, ByVal lParam As Long) As Long


Dim window As ApiWindow

If IsWindowApi(hwnd) Then
    '\\ If its already subclassed, return that instance...
    On Error Resume Next
    Set window = AllSubclassedWindows.Item(hwnd)
    On Error GoTo 0
    If window Is Nothing Then
        Set window = New ApiWindow
        window.hwnd = hwnd
    End If
    AllTopLevelWindows.Add window, "hwnd:" & window.hwnd
    TopLevelWndEnumProc = True
Else
    '\\ A bad hwnd has been returned - do not continue
    TopLevelWndEnumProc = False
End If


End Function

'\\ [VB_HOOKCALLWNDPROC]----------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ NOTE:
'\\  This code has been kept for backwards compatibility only.
'\\  You should use the specific CBTHookProc, ShellHookProc etc...
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_HOOKCALLWNDPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKCALLWNDPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_CALLWNDPROC), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_CALLWNDPROC, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_CALLWNDPROC), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKCALLWNDPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKCALLWNDPROC = lret
End If

End Function

'\\ [VB_HOOKCBTPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKCBTPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKCBTPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_CBT), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_CBT, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_CBT), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKCBTPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKCBTPROC = lret
End If

End Function

'\\ [VB_HOOKDEBUGPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKDEBUGPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKDEBUGPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_DEBUG), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_DEBUG, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_DEBUG), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKDEBUGPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKDEBUGPROC = lret
End If

End Function

'\\ [VB_HOOKFOREGROUNDIDLEPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKFOREGROUNDIDLEPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKFOREGROUNDIDLEPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_FOREGROUNDIDLE), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_FOREGROUNDIDLE, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_FOREGROUNDIDLE), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKFOREGROUNDIDLEPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKFOREGROUNDIDLEPROC = lret
End If
End Function


'\\ [VB_HOOKGETMESSAGEPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKGETMESSAGEPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKGETMESSAGEPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_GETMESSAGE), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_GETMESSAGE, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_GETMESSAGE), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKGETMESSAGEPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKGETMESSAGEPROC = lret
End If

End Function

'\\ [VB_HOOKJOURNALPLAYBACKPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKJOURNALPLAYBACKPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKJOURNALPLAYBACKPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_JOURNALPLAYBACK), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_JOURNALPLAYBACK, Params()
    lMsgRet = Params(4)
End If

'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_JOURNALPLAYBACK), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKJOURNALPLAYBACKPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKJOURNALPLAYBACKPROC = lret
End If


End Function

'\\ [VB_HOOKJOURNALRECORDPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKJOURNALRECORDPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKJOURNALRECORDPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_JOURNALRECORD), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_JOURNALRECORD, Params()
    lMsgRet = Params(4)
End If

'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_JOURNALRECORD), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKJOURNALRECORDPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKJOURNALRECORDPROC = lret
End If
End Function

'\\ [VB_HOOKMOUSEPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKMOUSEPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKMOUSEPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_MOUSE), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_MOUSE, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_MOUSE), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKMOUSEPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKMOUSEPROC = lret
End If

End Function


'\\ [VB_HOOKLOWLEVELMOUSEPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKLOWLEVELMOUSEPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKLOWLEVELMOUSEPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_MOUSE_LL), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_MOUSE_LL, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_MOUSE_LL), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKLOWLEVELMOUSEPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKLOWLEVELMOUSEPROC = lret
End If

End Function

'\\ [VB_HOOKMESSAGEFILTERPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKMESSAGEFILTERPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKMESSAGEFILTERPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_MSGFILTER), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_MESSAGEFILTER, Params()
    lMsgRet = Params(4)
End If

'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_MSGFILTER), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKMESSAGEFILTERPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKMESSAGEFILTERPROC = lret
End If

End Function

'\\ [VB_HOOKSHELLPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKSHELLPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKSHELLPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_SHELL), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_SHELL, Params()
    lMsgRet = Params(4)
End If

'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_SHELL), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKSHELLPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKSHELLPROC = lret
End If

End Function

'\\ [VB_HOOKSYSMESSAGEFILTERPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKSYSMESSAGEFILTERPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKSYSMESSAGEFILTERPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_SYSMSGFILTER), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_SYSMESSAGEFILTER, Params()
    lMsgRet = Params(4)
End If

'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_SYSMSGFILTER), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKSYSMESSAGEFILTERPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKSYSMESSAGEFILTERPROC = lret
End If

End Function

'\\ [VB_HOOKLOWLEVELKEYBOARDPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKLOWLEVELKEYBOARDPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKLOWLEVELKEYBOARDPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_KEYBOARD_LL), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam 'virtual key (WM_KEYDOWN, WM_KEYUP, WM_SYSKEYDOWN, WM_SYSKEYUP)
Params(3) = lParam 'pointer to a KBDLLHOOKSTRUCT object
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_KEYBOARD_LL, Params()
    lMsgRet = Params(4)
End If

'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_KEYBOARD_LL), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKLOWLEVELKEYBOARDPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKLOWLEVELKEYBOARDPROC = lret
End If

End Function

'\\ [VB_HOOKKEYBOARDPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKKEYBOARDPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKKEYBOARDPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_KEYBOARD), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_KEYBOARD, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_KEYBOARD), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKKEYBOARDPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKKEYBOARDPROC = lret
End If

End Function

'\\ [VB_HOOKHARDWAREPROC]----------------------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ ----------------------------------------------------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------------------------------------------------
Public Function VB_HOOKHARDWAREPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKHARDWAREPROC = CallNextHookEx(Eventhandler.HookIdByType(WH_HARDWARE), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    Eventhandler.TriggerEvent HOOKPROC_HARDWARE, Params()
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(WH_HARDWARE), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKHARDWAREPROC ", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKHARDWAREPROC = lret
End If

End Function

'\\ [VB_HookProc]----------------------------------------------------------------------------------
'\\ typedef LRESULT (CALLBACK* HOOKPROC)(int code, WPARAM wParam, LPARAM lParam);
'\\ code - type of hook,
'\\ Wparam, Lparam - message specific
'\\ lMsgRet = The message to pass to the calling code
'\\ NOTE:
'\\  This code has been kept for backwards compatibility only.
'\\  You should use the specific CBTHookProc, ShellHookProc etc...
'\\ ----------------------------------------------------------------------------------------
'\\ (c) 2001 - Merrion Computing.  All rights  to use, reproduce or publish this code reserved
'\\ Please check http://www.merrioncomputing.com for updates.
'\\ ----------------------------------------------------------------------------------------
Public Function VB_HOOKPROC(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

On Local Error Resume Next

Dim Params() As Variant
Dim lret As Long
Dim lMsgRet As Long

'\\ Note: If the code passed in is less than zero, it must be passed direct to the next hook proc
If Code < 0 Then
    VB_HOOKPROC = CallNextHookEx(Eventhandler.HookIdByType(Eventhandler.CurrentHookType), Code, wParam, lParam)
End If

'\\ 1 - Pack the param array.....
ReDim Params(1 To 4) As Variant
Params(1) = Code
Params(2) = wParam
Params(3) = lParam
Params(4) = lMsgRet

'\\ 2 - Call the event firer....
If Not Eventhandler Is Nothing Then
    If Eventhandler.CurrentHookType = WH_MOUSE Then
        Eventhandler.TriggerEvent HOOKPROC_MOUSE, Params()
    ElseIf Eventhandler.CurrentHookType = WH_MOUSE_LL Then
        Eventhandler.TriggerEvent HOOKPROC_MOUSE_LL, Params()
    Else
        Eventhandler.TriggerEvent HOOKPROC, Params()
    End If
    lMsgRet = Params(4)
End If


'\\ 3 - Pass this message on to the next hook proc in the chain (if any)
lret = CallNextHookEx(Eventhandler.HookIdByType(Eventhandler.CurrentHookType), Code, wParam, lParam)
If Err.LastDllError > 0 Then
    Call ReportError(Err.LastDllError, "VB_HOOKPROC", GetLastSystemError)
End If

'\\ If the message isn't cancelled, return the next hook's message...
If Not (lMsgRet) Then
    '\\ Return value to calling code....
    VB_HOOKPROC = lret
End If

End Function
