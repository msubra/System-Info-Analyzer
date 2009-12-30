VERSION 5.00
Begin VB.Form Test2 
   Caption         =   "Form1"
   ClientHeight    =   4095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   ScaleHeight     =   4095
   ScaleWidth      =   5910
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   360
      MultiLine       =   -1  'True
      TabIndex        =   2
      Text            =   "test2.frx":0000
      Top             =   1200
      Width           =   5415
   End
   Begin System.XPButton XPButton1 
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
   End
End
Attribute VB_Name = "Test2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Dim o As New OSInfo

MsgBox o.PlusVersionNumber

End Sub

Private Sub XPButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Fso As New FileSystemObject
Dim Data As String
Data = ""
Open "Sysinfo" For Binary As #1

While Not EOF(1)    ' Loop until end of file.
   Data = Data & Input(1, #1)  ' Get one character.
Wend

Close #1

Data = Decrypt(Data)
Text1.Text = Data
End Sub
Function Decrypt(Data As String) As String
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

Decrypt = EncData
End Function


