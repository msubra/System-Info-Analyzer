VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3069B1D2-18C7-11D6-887C-BDB0E2C3E80C}#1.0#0"; "LEDPROJ.OCX"
Begin VB.MDIForm SysMain 
   BackColor       =   &H8000000C&
   Caption         =   "System Builder"
   ClientHeight    =   7485
   ClientLeft      =   1275
   ClientTop       =   1395
   ClientWidth     =   9675
   Icon            =   "SysBuild.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   Picture         =   "SysBuild.frx":27A2
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   9615
      TabIndex        =   16
      Top             =   0
      Width           =   9675
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System Analyser By S.MaheshWaran(B.E CSE I)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000C&
         Height          =   450
         Left            =   120
         TabIndex        =   17
         Top             =   0
         Width           =   8895
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H8000000A&
      Height          =   6870
      Left            =   0
      ScaleHeight     =   6810
      ScaleWidth      =   2070
      TabIndex        =   0
      Top             =   615
      Width           =   2130
      Begin VB.PictureBox picGraph 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   120
         ScaleHeight     =   34.426
         ScaleMode       =   0  'User
         ScaleWidth      =   8.861
         TabIndex        =   12
         Top             =   7920
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox picUsage 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H0000C000&
         Height          =   975
         Left            =   0
         ScaleHeight     =   61
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   62
         TabIndex        =   10
         Top             =   6240
         Width           =   990
         Begin VB.Label lblCpuUsage 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0%"
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   0
            TabIndex        =   11
            Top             =   600
            Width           =   930
         End
      End
      Begin VB.Timer tmrRefresh 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   1920
         Top             =   6480
      End
      Begin System.XPButton XPButton1 
         Height          =   495
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Hard Disk"
         Picture         =   "SysBuild.frx":17AA8
         PictureAlignH   =   1
      End
      Begin System.XPButton XPButton2 
         Height          =   495
         Left            =   0
         TabIndex        =   2
         Top             =   1320
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Processor"
         Picture         =   "SysBuild.frx":1D29A
         PictureAlignH   =   1
      End
      Begin System.XPButton XPButton4 
         Height          =   495
         Left            =   0
         TabIndex        =   3
         Top             =   2040
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "OS"
         Picture         =   "SysBuild.frx":1EFA4
         PictureAlignH   =   1
      End
      Begin System.XPButton XPButton5 
         Height          =   495
         Left            =   0
         TabIndex        =   4
         Top             =   2760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Extensions"
         Picture         =   "SysBuild.frx":20736
         PictureAlignH   =   1
      End
      Begin System.XPButton XPButton6 
         Height          =   495
         Left            =   0
         TabIndex        =   5
         Top             =   3480
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "CLSID"
         Picture         =   "SysBuild.frx":22440
         PictureAlignH   =   1
      End
      Begin System.XPButton XPButton7 
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   4200
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Garbage"
         Picture         =   "SysBuild.frx":2414A
         PictureAlignH   =   1
      End
      Begin System.XPButton XPButton8 
         Height          =   495
         Left            =   0
         TabIndex        =   7
         Top             =   5640
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Exit Me"
         Picture         =   "SysBuild.frx":26284
         PictureAlignH   =   1
      End
      Begin System.XPButton XPButton3 
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   4920
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Collect"
         Picture         =   "SysBuild.frx":2C51E
         PictureAlignH   =   1
      End
      Begin Insert_Project_Name.LED LED1 
         Height          =   975
         Index           =   0
         Left            =   960
         TabIndex        =   9
         Top             =   6240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1720
         LEDColorOff     =   -2147483638
      End
      Begin Insert_Project_Name.LED LED1 
         Height          =   975
         Index           =   1
         Left            =   1320
         TabIndex        =   13
         Top             =   6240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1720
         LEDColorOff     =   -2147483638
      End
      Begin Insert_Project_Name.LED LED1 
         Height          =   975
         Index           =   2
         Left            =   1680
         TabIndex        =   14
         Top             =   6240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   1720
         LEDColorOff     =   -2147483638
      End
      Begin System.XPButton XPButton10 
         Height          =   495
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Click the Below"
         Picture         =   "SysBuild.frx":2E5A0
         NoDown          =   -1  'True
         IsPressed       =   -1  'True
         PictureAlignH   =   1
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3360
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SysBuild.frx":3142A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SysMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private QueryObject As Object
Private Sub Command1_Click()
Load Build
Unload Build
End Sub
Private Sub MDIForm_Terminate()
Ts.Close
End Sub

Private Sub XPButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Disks.Show
Disks.ZOrder vbBringToFront
End Sub

Private Sub XPButton2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Pro.Show
Pro.ZOrder vbBringToFront
End Sub

Private Sub XPButton3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Build.Show
End Sub

Private Sub XPButton4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Os.Show
Os.ZOrder vbBringToFront
End Sub

Private Sub XPButton5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Ext.Show
Ext.ZOrder vbBringToFront
End Sub

Private Sub XPButton6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
CLSIDPROGID.Show
CLSIDPROGID.ZOrder vbBringToFront
End Sub

Private Sub XPButton7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
WasteFiles.Show
WasteFiles.ZOrder vbBringToFront
End Sub

Private Sub XPButton8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
End
End Sub
Private Sub MDIForm_Activate()
Dim Fso As New FileSystemObject

    'set this form always on top
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
Set Ts = Fso.CreateTextFile("SysInfo")

    XPButton1.HoverColor = RGB(91, 200, 140)
    XPButton2.HoverColor = RGB(91, 200, 140)
    XPButton3.HoverColor = RGB(91, 200, 140)
    XPButton4.HoverColor = RGB(91, 200, 140)
    XPButton5.HoverColor = RGB(91, 200, 140)
    XPButton6.HoverColor = RGB(91, 200, 140)
    XPButton7.HoverColor = RGB(91, 200, 140)
    XPButton8.HoverColor = RGB(91, 200, 140)
    XPButton10.HoverColor = RGB(91, 200, 140)
    
End Sub
Private Sub MDIForm_Load()
    'set the Priority of this process to 'High'
    'this makes sure our program gets updated, even when
    'another process is consuming lots of CPU cycles
    SetThreadPriority GetCurrentThread, THREAD_BASE_PRIORITY_MAX
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    'Initialize our Query object
    If IsWinNT Then
        Set QueryObject = New clsCPUUsageNT
    Else
        Set QueryObject = New clsCPUUsage
    End If
    'Initializing is necesarry for the correct values to be retrieved
    QueryObject.Initialize
    'start the timer
    tmrRefresh.Enabled = True
    'don't wait for the first interval to elapse
    tmrRefresh_Timer
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    'stop the timer
    tmrRefresh.Enabled = False
    'clean up
    QueryObject.Terminate
    Set QueryObject = Nothing
End Sub

Private Sub tmrRefresh_Timer()
    Dim Ret As Long
    'query the CPU usage
    Ret = QueryObject.Query
    If Ret = -1 Then
        tmrRefresh.Enabled = False
        lblCpuUsage.Caption = ":("
        MsgBox "Error while retrieving CPU usage"
    Else
        DrawUsage Ret, picUsage, picGraph
        lblCpuUsage.Caption = CStr(Ret) + "%"
    End If
    
    If Ret = 100 Then
        LED1(0).Value = 1
        LED1(1).Value = 0
        LED1(2).Value = 0

    Else
        LED1(0).Value = 0
        LED1(1).Value = CInt(Ret - (Ret Mod 10)) / 10
        LED1(2).Value = CInt(Ret Mod 10)
    End If
    

End Sub


