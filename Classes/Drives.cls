VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Drives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim d As New Disk

Public Property Get CDROM() As String
Dim i As Integer
Dim P As String

For i = 1 To 26
    P = Chr(Asc("a") + i) & ":\"
    d.GetDriveInfo (P)
    
    If d.DriveType = "CDROM" Then
        CDROM = P
        Exit For
    End If
Next i

End Property

