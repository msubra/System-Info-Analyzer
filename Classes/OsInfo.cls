VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OsInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim k As New clsKey
Dim v As New clsValues

Public Property Get ProductName() As String
ProductName = RegVs("ProductName").Value
End Property

Public Property Get VersionNumber() As String
VersionNumber = RegVs("VersionNumber").Value
End Property
Public Property Get SubVersionNumber() As String
SubVersionNumber = UCase(RegVs("SubVersionNumber").Value)
End Property
Public Property Get RegisteredOwner() As String
RegisteredOwner = RegVs("RegisteredOwner").Value
End Property
Public Property Get RegisteredOrganization() As String
RegisteredOrganization = RegVs("RegisteredOrganization").Value
End Property
Public Property Get ProductKey() As String
ProductKey = RegVs("ProductKey").Value
End Property
Public Property Get ProductId() As String
'ProductId = v("ProductId").Value
ProductId = RegVs("ProductId").Value
End Property
Public Property Get PlusVersionNumber() As String
PlusVersionNumber = RegVs("Plus! VersionNumber").Value
End Property
Public Property Get LicensingInfo() As String
If RegVs("LicensingInfo").Value = "" Then
    LicensingInfo = "[Not Available]"
Else
    LicensingInfo = RegVs("LicensingInfo").Value
End If
End Property
Public Property Get DoneRegistration() As String

If RegVs("RegDone").Value = "" Then
    DoneRegistration = "[Not Available]"
Else
    DoneRegistration = RegVs("RegDone").Value
End If
End Property
Public Property Get SystemRoot() As String
SystemRoot = RegVs("SystemRoot").Value
End Property
Private Sub Class_Initialize()
Dim Path As String

Set RegKs = Reg.RegKeyFromHKey(HKEY_LOCAL_MACHINE).SubKeys
Set RegVs = Reg.RegKeyFromHKey(HKEY_LOCAL_MACHINE).Values

Path = "SOFTWARE\Microsoft\Windows\CurrentVersion"

Set RegKs = RegKs.RegKey.ParseKeyName(Path).SubKeys
Set RegVs = RegVs.RegKey.ParseKeyName(Path).Values

End Sub

