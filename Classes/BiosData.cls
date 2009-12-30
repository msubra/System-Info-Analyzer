VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BiosData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Dim isNT As Boolean

Public Property Get VideoBiosDate() As String
Attribute VideoBiosDate.VB_Description = "Get VideoBiosDate (read-only)"
    If isNT Then
       VideoBiosDate = GetRegistryValue(HKLM, "Hardware\Description\System", "VideoBiosDate", "")
    Else
'       VideoBiosDate = Mid(StrFromPtrA(&HC00A8), 1, 8) '-Date build
       VideoBiosDate = Mid(StrFromPtrA(&HC00A8), 9, 8) '-Date revision
    End If
End Property

Public Property Get VideoBiosVersion() As String
    Dim s As String
    If isNT Then
       s = GetRegistryValue(HKLM, "Hardware\Description\System", "VideoBiosVersion", "")
    Else
       s = StrFromPtrA(&HC0048)
       s = Left(s, InStr(1, s, vbCrLf) - 1)
       s = s & vbCrLf & "ChipType: " & GetRegistryValue(HKLM, "System\CurrentControlSet\Services\Class\Display\0000\INFO", "ChipType", "")
    End If
    VideoBiosVersion = s
End Property

Public Property Get VideoBiosCopyRight() As String
    Dim s As String
    If isNT Then
       s = "Unavailable on NT"
    Else
       s = StrFromPtrA(&HC0048)
       s = Mid$(s, InStr(1, s, vbCrLf) + 2)
    End If
    VideoBiosCopyRight = s
End Property

Public Property Get SystemBiosDate() As String
    If isNT Then
       SystemBiosDate = GetRegistryValue(HKLM, "Hardware\Description\System", "SystemBiosDate", "")
    Else
       SystemBiosDate = StrFromPtrA(&HFFFF5)
    End If
End Property

Public Property Get SystemBiosCopyRight() As String
    If isNT Then
       SystemBiosCopyRight = "Unvailable on NT"
    Else
       SystemBiosCopyRight = StrFromPtrA(&HFE091)
    End If
End Property

Public Property Get SystemBiosVersion() As String
    Dim vAns As Variant
    
    If isNT Then
        On Error Resume Next
      SystemBiosVersion = CDate(GetRegistryValue(HKLM, "Hardware\Description\System", "SystemBiosVersion", ""))
        If Err.Number > 0 Then SystemBiosVersion = "Unavailable"
        
    Else
      SystemBiosVersion = StrFromPtrA(&HFE061)
    End If
End Property

Public Property Get SystemBiosExtraInfo() As String
Attribute SystemBiosExtraInfo.VB_Description = "Get SystemBiosExtraInfo (read-only, unavailable on NT)"
    If isNT Then
       SystemBiosExtraInfo = "Unvailable on NT"
    Else
       SystemBiosExtraInfo = StrFromPtrA(&HFEC71)
    End If
End Property

Private Sub Class_Initialize()
   isNT = IsWindowsNT
End Sub
