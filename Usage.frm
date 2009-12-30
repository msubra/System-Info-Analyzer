VERSION 5.00
Object = "{3069B1D2-18C7-11D6-887C-BDB0E2C3E80C}#1.0#0"; "LEDPROJ.OCX"
Begin VB.Form Use 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   1455
   ClientLeft      =   2595
   ClientTop       =   4020
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1455
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   Begin Insert_Project_Name.LED LED1 
      Height          =   975
      Index           =   0
      Left            =   4680
      TabIndex        =   4
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1720
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5520
      Top             =   240
   End
   Begin System.XPButton XPButton1 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   960
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Your System CPU Usage"
      Picture         =   "Usage.frx":0000
      NoDown          =   -1  'True
      IsPressed       =   -1  'True
      PictureAlignH   =   1
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
      TabIndex        =   1
      Top             =   0
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
         TabIndex        =   2
         Top             =   600
         Width           =   930
      End
   End
   Begin VB.PictureBox picGraph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00008000&
      Height          =   975
      Left            =   1080
      ScaleHeight     =   100
      ScaleMode       =   0  'User
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   0
      Width           =   3615
   End
   Begin Insert_Project_Name.LED LED1 
      Height          =   975
      Index           =   1
      Left            =   5160
      TabIndex        =   5
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1720
   End
   Begin Insert_Project_Name.LED LED1 
      Height          =   975
      Index           =   2
      Left            =   5640
      TabIndex        =   6
      Top             =   0
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1720
   End
End
Attribute VB_Name = "Use"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'frmMonitor - copyright © 2001, The KPD-Team
'Visit our site at http://www.allapi.net
'or email us at KPDTeam@allapi.net
Option Explicit
Private QueryObject As Object
Private Sub Form_Activate()
    'set this form always on top
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
End Sub
Private Sub Form_Load()
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
Private Sub Form_Unload(Cancel As Integer)
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

