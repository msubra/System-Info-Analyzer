VERSION 5.00
Begin VB.Form Build 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Build System Detail"
   ClientHeight    =   585
   ClientLeft      =   1230
   ClientTop       =   690
   ClientWidth     =   6270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin System.XPButton XPButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Building Your System Details"
   End
End
Attribute VB_Name = "Build"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
Dim s As String
Dim CompName As String

XPButton1.HoverColor = RGB(91, 200, 140)

Details = Details & "Created At:" & Now & vbCrLf
LoadMemDevicedetails
LoadProcessorDetails
Details = Details & "ComputerName:" & ComputerName & vbCrLf

Details = EncryptorDecrypt(Details)
Ts.Write Details & vbCrLf

Unload Me
End Sub

Sub LoadMemDevicedetails()
Dim hd As New HardDisk
Dim cd As New CDROM
Dim fd As New FloppyDisk
Dim m As New Memory

Dim i As Integer

With hd
For i = 1 To .CountHardDisk
    .GetAllInfo (i)
    Details = Details & "[HARD DISK INFO " & i & "]" & vbCrLf
    Details = Details & "COUNT:" & .CountHardDisk & vbCrLf
    Details = Details & "DRIVERS:" & .Drivers & vbCrLf
    Details = Details & "DRIVESALLOCATED:" & .DrivesAllocated & vbCrLf
    Details = Details & "HARDDISKID:" & .HardDiskId & vbCrLf
    Details = Details & "HARDDISKNAME:" & .HardDiskName & vbCrLf
    Details = Details & "MANUFACTURER:" & .Manufacturer & vbCrLf
    Details = Details & "TOTALSIZE:" & .TotalSize & " Bytes" & vbCrLf & vbCrLf
Next i
End With

For i = 1 To cd.CDROM_Count
    Details = Details & "[CD INFO " & i & "]" & vbCrLf
    cd.LoadInfo (1)
    Details = Details & "CDROM AVAILABLE:" & cd.CDROM_Present & vbCrLf
    Details = Details & "CDROM NAME:" & cd.CDROM_Name & vbCrLf & vbCrLf
Next i

For i = 1 To fd.CountFloppyDrives
    Details = Details & "[FD INFO " & i & "]" & vbCrLf
    fd.GetAllInfo (i)
    
    Details = Details & "CountFloppyDrives:" & fd.CountFloppyDrives & vbCrLf
    Details = Details & "Drivers:" & fd.Drivers & vbCrLf
    Details = Details & "DrivesAllocated:" & fd.DrivesAllocated & vbCrLf
    Details = Details & "FloppyDriveId:" & fd.FloppyDriveId & vbCrLf
    Details = Details & "FloppyDriveName:" & fd.FloppyDriveName & vbCrLf
    Details = Details & "Manufacturer:" & fd.Manufacturer & vbCrLf & vbCrLf
Next i

With m
    Details = Details & "TotalMemory :" & m.TotalMemory & vbCrLf
    Details = Details & "TotalPhysicalMemory:" & m.TotalPhysicalMemory & vbCrLf
    Details = Details & "PageFileSize:" & m.PageFileSize & vbCrLf
End With

End Sub

Sub LoadProcessorDetails()
Dim p As New Processor
Dim i As Integer

For i = 1 To p.CountProcessors
With p
    Details = Details & "[PROCESSOR INFORMATION " & i & "]" & vbCrLf
    Details = Details & "ActiveProcessor:" & .ActiveProcessor & vbCrLf
    Details = Details & "AlphaInstructions:" & .AlphaInstructions & vbCrLf
    Details = Details & "CheckMathProcessor:" & .CheckMathProcessor & vbCrLf
    Details = Details & "CompareExchangeDouble:" & .CompareExchangeDouble & vbCrLf
    Details = Details & "CountProcessors:" & .CountProcessors & vbCrLf
    Details = Details & "FloatingPointEmulated:" & .FloatingPointEmulated & vbCrLf
    Details = Details & "FloatingPointPrecision:" & .FloatingPointPrecision & vbCrLf
    Details = Details & "MaximumApplicationAddress:" & .MaximumApplicationAddress & vbCrLf
    Details = Details & "MemMoveBit64:" & .MemMoveBit64 & vbCrLf
    Details = Details & "MemoryPageSize:" & .MemoryPageSize & vbCrLf
    Details = Details & "MinimumApplicationAddress:" & .MinimumApplicationAddress & vbCrLf
    Details = Details & "MMX_Instruction:" & .MMX_Instruction & vbCrLf
    Details = Details & "ProcessorName:" & .ProcessorName & vbCrLf
    Details = Details & "ProcessorType:" & .ProcessorType & vbCrLf
    Details = Details & "PerformanceCounter :" & .PerformanceCounter & vbCrLf
    Details = Details & "PerformanceFrequency :" & .PerformanceFrequency & vbCrLf
End With
Next i
End Sub

Function EncryptorDecrypt(Data As String) As String
Dim l, l1 As Long
Dim Key(0 To 5) As String * 1
Dim EncData As String
Dim i As Long



l = Len(Data)

Key(0) = "M"
Key(1) = "A"
Key(2) = "H"
Key(3) = "E"
Key(4) = "S"
Key(5) = "H"

l1 = 6

For i = 1 To l
    EncData = EncData & Chr(Asc(Mid(Data, i, 1)) Xor Asc(Key(i Mod l1)))
Next i

EncryptorDecrypt = EncData
End Function

