VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3069B1D2-18C7-11D6-887C-BDB0E2C3E80C}#1.0#0"; "LEDPROJ.OCX"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6300
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Start.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Start.frx":000C
   ScaleHeight     =   4245
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   135
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
   End
   Begin Insert_Project_Name.LED LED1 
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2640
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      Value           =   1
      ShowDecimal     =   -1  'True
   End
   Begin Insert_Project_Name.LED LED2 
      Height          =   495
      Left            =   5460
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2640
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      ShowDecimal     =   -1  'True
   End
   Begin Insert_Project_Name.LED LED3 
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2640
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      ShowDecimal     =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   735
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Activate()
Dim t1, t2
Dim t As Integer
t1 = Timer
ProgressBar1.Max = 201
While t <= 2
    t2 = Timer
    DoEvents
    t = CInt(val(Format(t2 - t1, "##.#")))
    
    If t * 100 <= 200 Then ProgressBar1.Value = t * 100
Wend

Me.Hide
Load SysMain
SysMain.Show
Unload Me

End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub lblCompanyProduct_Click()

End Sub
