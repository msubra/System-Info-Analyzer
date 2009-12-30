VERSION 5.00
Object = "{11998CBD-30CA-11D5-AFAD-0000B43618D7}#32.0#0"; "DIGITBOX.OCX"
Begin VB.Form Process 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Processing Data"
   ClientHeight    =   420
   ClientLeft      =   4710
   ClientTop       =   4005
   ClientWidth     =   2700
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   420
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   720
      Top             =   0
   End
   Begin WSDigitbox.DigitBox DigitBox1 
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   767
      DigitDisplay    =   "12345"
      DigitPlaceHolders=   10
      DigitSize       =   1
   End
End
Attribute VB_Name = "process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
DoEvents
Me.Show
End Sub

Private Sub Timer1_Timer()
DoEvents
DigitBox1.DigitDisplay = Rnd * (1 ^ 10)
End Sub
