VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get BiosName() As String
BiosName = RegVs("BIOSName").Value
End Property
Public Property Get BiosVersion() As String
BiosVersion = RegVs("BIOSVersion").Value
End Property
Public Property Get BiosDate() As String
BiosDate = RegVs("BIOSDate").Value
End Property
Public Property Get BusType() As String
BusType = RegVs("BusType").Value
End Property
Public Property Get CPUType() As String
CPUType = RegVs("CPU").Value
End Property
Public Property Get DeviceDescription() As String
DeviceDescription = RegVs("DeviceDesc").Value
End Property
Public Property Get MachineType() As String
MachineType = RegVs("MachineType").Value
End Property
Private Sub Class_Initialize()

Set RegK = Reg.RegKeyFromHKey(HKEY_LOCAL_MACHINE)
Set RegK = RegK.ParseKeyName("Enum\Root\*PNP0C01\0000")
Set RegVs = RegK.Values

End Sub
