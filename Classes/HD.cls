VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HardDisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim h As New clsReg
Dim Key As String
Dim Values As cRegValues
Dim Keys As cRegKeys

Private Type HardDiskinfo
    HardDiskName As String
    HardDiskId As String
    Manufacturer As String
    Driver As String
    DriveLetters As String
End Type
Dim HdInfo As HardDiskinfo
Public Function GetAllInfo(Index As Integer)
Dim P As String
Dim k As cRegKeys
Dim v1  As New clsValues
Dim t As New clsReg

v1.hKey = HKEY_LOCAL_MACHINE
P = Key & "\" & Keys(Index).Key
Set k = t.ListKeys(HKEY_LOCAL_MACHINE, P)
P = P & "\" & k.Item(1).Key

v1.Path = P
v1.Refresh

With HdInfo
    .HardDiskId = v1.Item("ProductId").Value
    .Manufacturer = v1.Item("Manufacturer").Value
    .HardDiskName = v1.Item("DeviceDesc").Value
    .Driver = v1.Item("Driver").Value
    .DriveLetters = v1.Item("CurrentDriveLetterAssignment").Value
End With

End Function

Public Property Get CountHardDisk() As Byte
CountHardDisk = Keys.Count
End Property
Public Property Get HardDiskName() As String
    HardDiskName = HdInfo.HardDiskName
End Property
Public Property Get HardDiskId() As String
    HardDiskId = HdInfo.HardDiskId
End Property
Public Property Get DrivesAllocated() As String
DrivesAllocated = HdInfo.DriveLetters
End Property
Public Property Get Manufacturer() As String
Manufacturer = HdInfo.Manufacturer
End Property
Public Property Get Drivers() As String
Drivers = HdInfo.Driver
End Property
Public Property Get TotalSize() As String
Dim s As String
Dim i As Integer
Dim Sz As Double
Dim d As New Disk

s = DrivesAllocated
Sz = 0

For i = 1 To Len(s)
    d.GetDriveInfo (Mid(s, i, 1))
    Sz = Sz + d.TotalSpace
Next i
TotalSize = CStr(Sz)
End Property
Private Sub Class_Initialize()
Key = "Enum\ESDI"
Set Keys = h.ListKeys(HKEY_LOCAL_MACHINE, Key)
End Sub
