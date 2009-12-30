Attribute VB_Name = "FreeMemory"
Option Explicit

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

Public Function AvailablePhysicalMemory() As Double
'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPhys
    AvailablePhysicalMemory = BytesToMegabytes(dblAns)
    
End Function

Public Function TotalPhysicalMemory() As Double
'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwTotalPhys
    TotalPhysicalMemory = BytesToMegabytes(dblAns)
End Function

Public Function PercentMemoryFree() As Double

   PercentMemoryFree = Format(AvailableMemory / TotalMemory * _
   100, "0#")
End Function

Public Function AvailablePageFile() As Double
'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwAvailPageFile
    AvailablePageFile = BytesToMegabytes(dblAns)
End Function

Public Function PageFileSize() As Double
'Return Value in Megabytes
    Dim dblAns As Double
    GlobalMemoryStatus pUdtMemStatus
    dblAns = pUdtMemStatus.dwTotalPageFile
    PageFileSize = BytesToMegabytes(dblAns)

End Function

Public Function AvailableMemory() As Double
'Return Value in Megabytes
     AvailableMemory = AvailablePhysicalMemory + AvailablePageFile
End Function

Public Function TotalMemory() As Double
'Return Value in Megabytes
    TotalMemory = PageFileSize + TotalPhysicalMemory
End Function

Private Function BytesToMegabytes(Bytes As Double) As Double
 
  Dim dblAns As Double
  dblAns = (Bytes / 1024) / 1024
  BytesToMegabytes = Format(dblAns, "###,###,##0.00")
  
End Function

Public Function GetCPUUsage() As Long
    Dim rc As Long
    Dim hKey As Long
    Dim KeyData As String
    Dim KeyValType As Long
    Dim lAns As Long
    Dim KeyValSize As Long
    rc = RegOpenKeyEx(HKEY_DYN_DATA, "PerfStats\StatData", 0, KEY_ALL_ACCESS, hKey)
    If (rc <> 0) Then GoTo GetKeyError
    KeyData = String$(1024, 0)
    KeyValSize = 1024
    rc = RegQueryValueEx(hKey, "KERNEL\CPUUsage", 0, KeyValType, KeyData, Len(KeyData))
    If (rc <> 0) Then GoTo GetKeyError
    
    If (Asc(Mid(KeyData, Len(KeyData), 1)) = 0) Then
        KeyData = Left(KeyData, Len(KeyData) - 1)
    Else
        KeyData = Left(KeyData, Len(KeyData))
    End If
    
    lAns = Asc(KeyData) * -1 + 100
    GetCPUUsage = 100 - lAns
Exit Function
GetKeyError:
    GetCPUUsage = -1
    rc = RegCloseKey(hKey)
End Function


