VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Memory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long

Private Const HKEY_DYN_DATA = &H80000006
Private Const KEY_ALL_ACCESS = 131135

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private pUdtMemStatus As MEMORYSTATUS

Private Declare Sub GlobalMemoryStatus Lib _
"kernel32" (lpBuffer As MEMORYSTATUS)

Public Property Get AvailablePhysicalMemory() As Double
'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    AvailablePhysicalMemory = BytesToMegabytes(dblAns)
    
End Property

Public Property Get TotalPhysicalMemory() As Double
'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwTotalPhys
    TotalPhysicalMemory = BytesToMegabytes(dblAns)
End Property

Public Property Get PercentMemoryFree() As Double
   PercentMemoryFree = Format(AvailableMemory / TotalMemory * _
   100, "0#")
End Property

Public Property Get AvailablePageFile() As Double
'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPageFile
    AvailablePageFile = BytesToMegabytes(dblAns)
End Property

Public Property Get PageFileSize() As Double
'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwTotalPageFile
    PageFileSize = BytesToMegabytes(dblAns)

End Property

Public Property Get AvailableMemory() As Double
'Return Value in Megabytes
     AvailableMemory = AvailablePhysicalMemory + AvailablePageFile
End Property

Public Property Get TotalMemory() As Double
'Return Value in Megabytes
    TotalMemory = PageFileSize + TotalPhysicalMemory
End Property

Private Function BytesToMegabytes(bytes As Double) As Double
 
  Dim dblAns As Double
  dblAns = (bytes / 1024) / 1024
  BytesToMegabytes = Format(dblAns, "###,###,##0.00")
  
End Function

'Public Property Get GetCPUUsage() As Long
'    Dim rc As Long
'    Dim hKey As Long
'    Dim KeyData As String
'    Dim KeyValType As Long
'    Dim lAns As Long
'    Dim KeyValSize As Long
'    rc = RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StatData", 0, KEY_ALL_ACCESS, hKey)
'    If (rc <> 0) Then GoTo GetKeyError
'    KeyData = String$(1024, 0)
'    KeyValSize = 1024
'    rc = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, KeyValType, KeyData, Len(KeyData))
'    If (rc <> 0) Then GoTo GetKeyError
'
'    If (Asc(Mid(KeyData, Len(KeyData), 1)) = 0) Then
'        KeyData = Left(KeyData, Len(KeyData) - 1)
'    Else
'        KeyData = Left(KeyData, Len(KeyData))
'    End If
    
'    lAns = Asc(KeyData) * -1 + 100
'    GetCPUUsage = 100 - lAns
'Exit Property
'GetKeyError:
'    GetCPUUsage = -1
'    rc = RegCloseKey(hKey)
'End Property

Private Sub Class_Initialize()
GlobalMemoryStatus pUdtMemStatus
End Sub
