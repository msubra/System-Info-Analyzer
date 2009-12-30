VERSION 5.00
Begin VB.Form Ext 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Available File Extensions In The System"
   ClientHeight    =   3300
   ClientLeft      =   2550
   ClientTop       =   3300
   ClientWidth     =   6450
   Icon            =   "Extensions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   6450
   Begin System.XPButton XPButton1 
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Get Extensions"
      Picture         =   "Extensions.frx":1CFA
      IsPressed       =   -1  'True
      PictureAlignH   =   1
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3150
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin System.XPButton XPButton2 
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "Stop Process"
      Enabled         =   0   'False
      Picture         =   "Extensions.frx":3A04
      IsPressed       =   -1  'True
      PictureAlignH   =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   3840
      TabIndex        =   2
      Top             =   240
      Width           =   630
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Ext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StopProcess As Boolean

Private Sub Form_Load()
StopProcess = False
XPButton1.HoverColor = RGB(91, 200, 140)
XPButton2.HoverColor = RGB(91, 200, 140)

End Sub

Private Sub XPButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim r As New clsReg
Dim k As cRegKeys
Dim v As cRegValue
Dim i As Integer
Dim E As String
Dim v1 As New clsValues

XPButton1.Enabled = False
XPButton2.Enabled = True
List1.Clear
Set k = r.ListKeys(HKEY_CLASSES_ROOT, "")
v1.hKey = HKEY_CLASSES_ROOT

For i = 1 To k.Count
    E = k(i).Key
    If StopProcess = True Then Exit For
    If Left(E, 1) = "." Then
        Set v = r.ListValues(HKEY_CLASSES_ROOT, E).Item(1)
        v1.Path = E
        v1.Refresh
        If v1.Count > 0 Then
            List1.AddItem E & "-----" & v1(1).Value
        End If
    End If
    DoEvents
Next i
Label1.Caption = "Total Extensions Supported: " & List1.ListCount
XPButton1.Enabled = True
XPButton2.Enabled = False
StopProcess = False
End Sub

Private Sub XPButton2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
XPButton1.Enabled = True
StopProcess = True
XPButton2.Enabled = False

End Sub
