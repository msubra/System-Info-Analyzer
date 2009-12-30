VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FloppyDisk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim h As New clsReg
Dim Key As String
Dim Values As cRegValues
Dim Keys As cRegKeys

Private Type FloppyDriveinfo
    FDName As String
    FDId As String
    Manufacturer As String
    Driver As String
    DriveLetter As String
End Type
Dim FDInfo As FloppyDriveinfo

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

With FDInfo
    .FDId = v1.Item("ProductId").Value
    .Manufacturer = v1.Item("Manufacturer").Value
    .FDName = v1.Item("DeviceDesc").Value
    .Driver = v1.Item("Driver").Value
    .DriveLetter = v1.Item("CurrentDriveLetterAssignment").Value
End With

End Function

Public Property Get CountFloppyDrives() As Byte
CountFloppyDrives = Keys.Count
End Property
Public Property Get FloppyDriveName() As String
    FloppyDriveName = FDInfo.FDName
End Property
Public Property Get FloppyDriveId() As String
    FloppyDriveId = FDInfo.FDId
End Property
Public Property Get DrivesAllocated() As String
DrivesAllocated = FDInfo.DriveLetter
End Property
Public Property Get Manufacturer() As String
Manufacturer = FDInfo.Manufacturer
End Property
Public Property Get Drivers() As String
Drivers = FDInfo.Driver
End Property

Private Sub Class_Initialize()
Key = "Enum\FLOP"
Set Keys = h.ListKeys(HKEY_LOCAL_MACHINE, Key)
End Sub

