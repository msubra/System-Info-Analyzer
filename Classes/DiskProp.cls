VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Disk"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Dim Fso As New FileSystemObject
Dim Drv As Drive
Dim DrvProp(10) As String
Dim DrvName As String
Dim VolName, FSNameBuffer As String
Dim VolNameSize, VolSerialNum, MaxComLen, SysFlags, FSNameSize As Long

Public Sub GetDriveInfo(DriveName As String)
Set Drv = Fso.GetDrive(DriveName)

GetVolumeInformation DrvName, VolName, VolNameSize, VolSerialNum, _
                    MaxComLen, SysFlags, FSNameBuffer, FSNameSize
End Sub

Property Get DriveName() As String
DriveName = Drv.VolumeName
End Property

Property Get TotalSpace() As Double
TotalSpace = Drv.TotalSize
End Property

Property Get FreeSpace() As Double
FreeSpace = Drv.FreeSpace
End Property

Property Get UsedSpace() As Double
UsedSpace = Str(Val(TotalSpace) - Val(FreeSpace))
End Property

Property Get VolumeName() As String
VolumeName = Drv.VolumeName
End Property

Property Get VolumeSerialNumber() As String
VolumeSerialNumber = Hex(Drv.SerialNumber)
End Property
Property Get IsDiskReady() As Boolean
IsDiskReady = Drv.IsReady
End Property
Property Get FileSystem() As String
FileSystem = Drv.FileSystem
End Property
Property Get DriveType() As String
Dim DType As DriveTypeConst
Dim TypeName As String

TypeName = "Unknown"
DType = Drv.DriveType

Select Case DType
    Case CDRom
        TypeName = "CD-Rom"
    Case Fixed
        TypeName = "Fixed"
    Case RamDisk
        TypeName = "RamDisk"
    Case Remote
        TypeName = "Remote"
    Case Removable
        TypeName = "Removable"
    Case UnknownType
        TypeName = "Unknown"
End Select
DriveType = TypeName
End Property
Public Function FormatValue(Value As Long) As String
FormatValue = Format(Value, "###,###,###,###")
End Function
