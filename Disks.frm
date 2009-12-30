VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Disks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Disk And Memory Status"
   ClientHeight    =   5415
   ClientLeft      =   1680
   ClientTop       =   2250
   ClientWidth     =   8670
   Icon            =   "Disks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5415
   ScaleWidth      =   8670
   Begin System.XPButton XPButton1 
      Height          =   495
      Left            =   0
      TabIndex        =   41
      Top             =   4920
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Refresh System Status"
      Picture         =   "Disks.frx":0ECA
      IsPressed       =   -1  'True
      PictureAlignH   =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3000
      Top             =   2640
   End
   Begin VB.DriveListBox Drive 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   3255
   End
   Begin MSComctlLib.ProgressBar PUsed 
      Height          =   195
      Left            =   720
      TabIndex        =   21
      Top             =   1260
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar PFree 
      Height          =   195
      Left            =   720
      TabIndex        =   22
      Top             =   1680
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar TotalMem 
      Height          =   195
      Left            =   4680
      TabIndex        =   33
      Top             =   720
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar FreeMem 
      Height          =   195
      Left            =   4680
      TabIndex        =   34
      Top             =   1140
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar PhyTot 
      Height          =   195
      Left            =   4800
      TabIndex        =   37
      Top             =   3240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar PhyAvail 
      Height          =   195
      Left            =   4800
      TabIndex        =   38
      Top             =   3660
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Available:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3960
      TabIndex        =   40
      Top             =   3630
      Width           =   825
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4200
      TabIndex        =   39
      Top             =   3270
      Width           =   480
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4080
      TabIndex        =   36
      Top             =   750
      Width           =   480
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Available:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3840
      TabIndex        =   35
      Top             =   1110
      Width           =   825
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   3600
      X2              =   8640
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Physical Memory Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   32
      Top             =   2640
      Width           =   3585
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Total Memory Resources:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4110
      TabIndex        =   31
      Top             =   1440
      Width           =   2160
   End
   Begin VB.Label TotMem 
      AutoSize        =   -1  'True
      Caption         =   "Disk Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6360
      TabIndex        =   30
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label asdf 
      AutoSize        =   -1  'True
      Caption         =   "Available Memory Resources:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3765
      TabIndex        =   29
      Top             =   1800
      Width           =   2505
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Total Physical Memory(RAM):"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3825
      TabIndex        =   28
      Top             =   4080
      Width           =   2505
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Available Physical Memory:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4020
      TabIndex        =   27
      Top             =   4440
      Width           =   2310
   End
   Begin VB.Label PhyTotMem 
      AutoSize        =   -1  'True
      Caption         =   "Disk Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6480
      TabIndex        =   26
      Top             =   4080
      Width           =   795
   End
   Begin VB.Label AvailMem 
      AutoSize        =   -1  'True
      Caption         =   "Disk Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6360
      TabIndex        =   25
      Top             =   1800
      Width           =   795
   End
   Begin VB.Label PhyAvailMem 
      AutoSize        =   -1  'True
      Caption         =   "Starting Volume:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   6480
      TabIndex        =   24
      Top             =   4440
      Width           =   1410
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Memory Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   23
      Top             =   120
      Width           =   3465
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   3600
      X2              =   3600
      Y1              =   0
      Y2              =   4920
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Free:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   1650
      Width           =   420
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Used:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   1290
      Width           =   465
   End
   Begin VB.Label Tip 
      AutoSize        =   -1  'True
      BackColor       =   &H00D6FEFD&
      Caption         =   "Tip"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "System Disk Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      TabIndex        =   16
      Top             =   120
      Width           =   2925
   End
   Begin VB.Label Status 
      AutoSize        =   -1  'True
      Caption         =   "Ending Volume:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   15
      Top             =   4440
      Width           =   1275
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Disk Status:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   465
      TabIndex        =   14
      Top             =   4440
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Disk Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   675
      TabIndex        =   13
      Top             =   2160
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Volume Label:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   300
      TabIndex        =   12
      Top             =   3135
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Free Space:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   495
      TabIndex        =   11
      Top             =   2490
      Width           =   975
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Space Occupied:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   10
      Top             =   2805
      Width           =   1365
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Disk Type:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   9
      Top             =   3780
      Width           =   870
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "File System:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   450
      TabIndex        =   8
      Top             =   4110
      Width           =   1020
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Serial Number:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   7
      Top             =   3465
      Width           =   1230
   End
   Begin VB.Label SerialNo 
      AutoSize        =   -1  'True
      Caption         =   "Starting Volume:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   6
      Top             =   3435
      Width           =   1410
   End
   Begin VB.Label evol 
      AutoSize        =   -1  'True
      Caption         =   "Space Occupied:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   5
      Top             =   4080
      Width           =   1365
   End
   Begin VB.Label svol 
      AutoSize        =   -1  'True
      Caption         =   "Space Occupied:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   4
      Top             =   3765
      Width           =   1365
   End
   Begin VB.Label occu 
      AutoSize        =   -1  'True
      Caption         =   "Disk Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   2805
      Width           =   795
   End
   Begin VB.Label free 
      AutoSize        =   -1  'True
      Caption         =   "Disk Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   2475
      Width           =   795
   End
   Begin VB.Label vol 
      AutoSize        =   -1  'True
      Caption         =   "Disk Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   1
      Top             =   3120
      Width           =   795
   End
   Begin VB.Label Total 
      AutoSize        =   -1  'True
      Caption         =   "Disk Size:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   0
      Top             =   2160
      Width           =   795
   End
End
Attribute VB_Name = "Disks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Drive_Change()
LoadDiskDetails (GetDiskName(Drive.Drive))
End Sub

Private Sub Form_Activate()

LoadDiskDetails ("C:\")
LoadMemoryDetails

Me.SetFocus
XPButton1.HoverColor = RGB(91, 200, 140)
End Sub

Function GetDiskName(Name As String) As String
GetDiskName = Left(Name, 1)
End Function

Sub LoadDiskDetails(Name As String)
Dim d As New Disk
Dim k As New KeyBoard

d.GetDriveInfo (Name)

If d.IsDiskReady = True Then
    free = Format(d.FreeSpace, "####, ###, ###, ###") & " Bytes"
    Total = Format(d.TotalSpace, "####, ###, ###, ###") & " Bytes"
    occu = Format(d.UsedSpace, "####, ###, ###, ###") & " Bytes"
    vol = d.VolumeName
    SerialNo = d.VolumeSerialNumber
    svol = d.DriveType
    evol = d.FileSystem
    Status = "Disk Ready"
    
    'progress bar
    
    PFree.Value = (d.FreeSpace / d.TotalSpace) * 100
    PUsed.Value = (d.UsedSpace / d.TotalSpace) * 100
Else
    free = "Not Available"
    Total = "Not Available"
    occu = "Not Available"
    vol = "Not Available"
    SerialNo = "Not Available"
    svol = "Not Available"
    evol = "Not Available"
    Status = "Disk Not Ready"
    
    PFree.Value = 0
    PUsed.Value = 0
    
End If

End Sub

Sub LoadMemoryDetails()
Dim m As New Memory

TotMem = m.TotalMemory & " MegaBytes"
AvailMem = m.AvailableMemory & " MegaBytes"
PhyTotMem = m.TotalPhysicalMemory & " MegaBytes"
PhyAvailMem = m.AvailablePhysicalMemory & " MegaBytes"



TotalMem.Value = 100
FreeMem.Value = (m.AvailableMemory / m.TotalMemory) * 100

PhyTot.Value = 100
PhyAvail.Value = (m.AvailablePhysicalMemory / m.TotalPhysicalMemory) * 100

End Sub

Private Sub XPButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Drive.ListIndex = 1
Call Drive_Change
LoadMemoryDetails
End Sub
