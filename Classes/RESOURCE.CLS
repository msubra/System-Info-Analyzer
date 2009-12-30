VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ResourceMeter"
Attribute VB_Creatable = True
Attribute VB_Exposed = True
'----------------------------------------------------
' MUST be compiled as a 32-bit executable
' otherwise an error is raised at compile-time
'----------------------------------------------------

#If Win32 Then

Private Type MemoryStatus
    dwSize As Long
    ' size of structure in byte
    dwMemoryLoad As Long
    ' percentage of memory used
    dwTotalPhys As Long
    ' amount of total physical memory in bytes
    dwAvailPhys As Long
    ' amount of free physical memory in bytes
    dwTotalPageFile As Long
    ' length of paging (swap) file in bytes
    dwAvailPageFile As Long
    ' free bytes within the paging file
    dwTotalVirtual As Long
    ' amount of total virtual memory in bytes
    dwAvailVirtual As Long
    ' amount of free virtual memory in bytes
End Type
Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MemoryStatus)

#End If

' this is the object used to communicate with
' the RESSVR server
Dim res16 As New ResSvr16.SystemHeapInfo
' this is used when calling GlobalMemoryStatus
Dim memInfo32 As MemoryStatus

Option Explicit


Private Sub Class_Initialize()

    ' initialize the length of global memInfo structures
    memInfo32.dwSize = LenB(memInfo32)

End Sub


Public Property Get PhysicalMemory() As Long
Attribute PhysicalMemory.VB_Description = "Amount of total physical memory (bytes)"

    ' return size of physical memory in bytes
    GlobalMemoryStatus memInfo32
    PhysicalMemory = memInfo32.dwTotalPhys

End Property


Public Property Get VirtualMemory() As Long
Attribute VirtualMemory.VB_Description = "Amount of virtual memory (bytes)"

    ' return size of virtual memory in bytes
    GlobalMemoryStatus memInfo32
    VirtualMemory = memInfo32.dwTotalVirtual

End Property


Property Get FreePhysicalMemory() As Long
Attribute FreePhysicalMemory.VB_Description = "Amount of free physical memory (bytes)"

    ' return size of free physical memory in bytes
    GlobalMemoryStatus memInfo32
    FreePhysicalMemory = memInfo32.dwAvailPhys

End Property

Property Get FreeVirtualMemory() As Long
Attribute FreeVirtualMemory.VB_Description = "Amount of free virtual memory (bytes)"

    ' return size of free virtual memory in bytes
    GlobalMemoryStatus memInfo32
    FreeVirtualMemory = memInfo32.dwAvailVirtual

End Property


Public Property Get SwapMemory() As Long
Attribute SwapMemory.VB_Description = "Length of system page file (bytes)"

    ' return size of page file in byte
    GlobalMemoryStatus memInfo32
    SwapMemory = memInfo32.dwTotalPageFile

End Property

Public Property Get FreeSwapMemory() As Long

    ' return size of free page file in byte
    GlobalMemoryStatus memInfo32
    FreeSwapMemory = memInfo32.dwAvailPageFile

End Property

Public Property Get FreeUserResources() As Integer
Attribute FreeUserResources.VB_Description = "Percentage of free User resources"

    ' return percentage of free resources in User heap
    FreeUserResources = res16.FreeUserResources

End Property



Public Property Get FreeGDIResources() As Integer
Attribute FreeGDIResources.VB_Description = "Percentage of free GDI resources"

    ' return percentage of free resources in User heap
    FreeGDIResources = res16.FreeGDIResources

End Property




