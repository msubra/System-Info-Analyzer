VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Pro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processor Information"
   ClientHeight    =   2985
   ClientLeft      =   3495
   ClientTop       =   3570
   ClientWidth     =   4620
   Icon            =   "Processor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4620
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8160
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Processor.frx":1CFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Processor.frx":25D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Processor.frx":87B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Processor.frx":DFB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin System.XPButton Pros 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   2400
      Width           =   4455
      _ExtentX        =   7858
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
      Caption         =   "Refresh Processor Information"
      Picture         =   "Processor.frx":1424E
      PictureAlignH   =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Processor Information:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   1950
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000006&
      Height          =   1815
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   360
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Os Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   1080
      TabIndex        =   9
      Top             =   480
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Win"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   0
      Left            =   1920
      TabIndex        =   8
      Top             =   480
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Win"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   1
      Left            =   1920
      TabIndex        =   7
      Top             =   840
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Win"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   2
      Left            =   1920
      TabIndex        =   6
      Top             =   1200
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Win"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   3
      Left            =   1920
      TabIndex        =   5
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Win"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   210
      Index           =   4
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Os Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   1080
      TabIndex        =   3
      Top             =   840
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Os Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   1200
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Os Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   1080
      TabIndex        =   1
      Top             =   1560
      Width           =   780
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Os Name:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   780
   End
End
Attribute VB_Name = "Pro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Label1(0) = "Processor Name="
Label1(1) = "No.Of Processors="
Label1(2) = "Processor Family="
Label1(3) = "Math Processor="
Label1(4) = "Memory PageSize:"

Label2(0) = "N/A"
Label2(1) = "N/A"
Label2(2) = "N/A"
Label2(3) = "N/A"
Label2(4) = "N/A"
Call Pros_Click
Pros.HoverColor = RGB(91, 200, 140)
End Sub

Private Sub Pros_Click()
Dim p As New Processor

Label1(0) = "Processor Name="
Label1(1) = "No.Of Processors="
Label1(2) = "Processor Family="
Label1(3) = "Math Processor="
Label1(4) = "Memory PageSize:"


Label2(0) = p.ProcessorName
Label2(1) = p.CountProcessors
Label2(2) = p.ProcessorType
Label2(3) = p.CheckMathProcessor
Label2(4) = p.MemoryPageSize & " MB"

End Sub

Private Sub Proc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub
