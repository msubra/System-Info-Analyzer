VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDROM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim r As New clsKeys

Private Type CD_INFO
    DriveLetter As String
    Manufacturer As String
    ProductId As String
    Driver As String
End Type
Dim tempV As RegValues
Dim tempK As RegKeys

Dim Info As CD_INFO
Public Property Get CDROM_Present() As Boolean

r.Path = "\System\CurrentControlSet\Services\Class"
r.Refresh


Set tempK = Reg.RegKeyFromHKey(HKEY_LOCAL_MACHINE).SubKeys
Set tempK = tempK.RegKey.ParseKeyName("System\CurrentControlSet\Services\Class").SubKeys

    If tempK.Item("CDROM") <> "" Then
        CDROM_Present = True
    Else
        CDROM_Present = False
    End If
End Property

Private Function Metrics(sMet As SystemMetricsIndexes) As Integer
Metrics = GetSystemMetrics(sMet)
End Function
Public Property Get CDROM_Count() As Integer

Set tempK = Reg.RegKeyFromHKey(HKEY_LOCAL_MACHINE).SubKeys
Set tempK = tempK.RegKey.ParseKeyName("Enum\SCSI").SubKeys

CDROM_Count = tempK.Count
End Property
Public Property Get CDROM_Name() As String
CDROM_Name = Info.Manufacturer
End Property
Private Sub Class_Initialize()
Set RegK = Reg.RegKeyFromHKey(HKEY_LOCAL_MACHINE)
Set RegK = RegK.ParseKeyName("Enum\SCSI")
End Sub

Function LoadInfo(Index As Integer)

Set RegK = RegK.SubKeys(Index).SubKeys(1)
Set RegVs = RegK.Values

With Info
    .DriveLetter = RegVs.Item("CurrentDriveLetterAssignment").Value
    .Driver = RegVs.Item("Driver").Value
    .Manufacturer = RegVs.Item("Manufacturer").Value
    .ProductId = RegVs.Item("ProductId").Value
End With

End Function
