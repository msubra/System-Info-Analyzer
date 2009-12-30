VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form CLSIDPROGID 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProgID and CLSID"
   ClientHeight    =   4500
   ClientLeft      =   1080
   ClientTop       =   2775
   ClientWidth     =   9675
   Icon            =   "Reg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   9675
   Begin System.XPButton XPButton1 
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   3960
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
      Caption         =   "Load Details"
      Picture         =   "Reg.frx":4E12
      IsPressed       =   -1  'True
      PictureAlignH   =   1
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   6588
      SortKey         =   2
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CLSID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "PROGID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "PATH"
         Object.Width           =   2540
      EndProperty
   End
   Begin System.XPButton StopProc 
      Height          =   495
      Left            =   4800
      TabIndex        =   3
      Top             =   3960
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "Stop"
      Enabled         =   0   'False
      Picture         =   "Reg.frx":9C34
      IsPressed       =   -1  'True
      PictureAlignH   =   1
   End
   Begin System.XPButton Shut 
      Height          =   495
      Left            =   7320
      TabIndex        =   4
      Top             =   3960
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "Exit"
      Picture         =   "Reg.frx":A50E
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
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   60
   End
End
Attribute VB_Name = "CLSIDPROGID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fso As New FileSystemObject
Dim StopSearch As Boolean

Private Sub Command1_Click()
Dim p As New OleData
Dim r As New clsReg
Dim k As cRegKeys


Dim kCount, i As Integer
Dim ProgId As String
Dim CLSId As String
Dim FileName, FilePath As String

Dim l As ListItem

Set k = r.ListKeys(HKEY_CLASSES_ROOT, "CLSID")

kCount = k.Count
SpFolder = SystemFolder
WinDir = fso.GetSpecialFolder(SpFolder)

For i = 1 To kCount
DoEvents
    CLSId = k.Item(i).Key
    
    Label1.Caption = "Processing:" & i
    
    If p.IsCLSID(CLSId) Then
        ProgId = p.CLSIDToProgID(CLSId)
        FilePath = p.GetFileFromCLSID(CLSId)
    
'    DoEvents

       Set l = ListView1.ListItems.Add(, , CLSId)

       l.SubItems(1) = Trim(ProgId)
       l.SubItems(2) = FilePath
    End If
    
    
Next i
End Sub

Private Sub Form_Load()
StopSearch = False
XPButton1.HoverColor = RGB(91, 200, 140)
StopProc.HoverColor = RGB(91, 200, 140)
Shut.HoverColor = RGB(91, 200, 140)
End Sub

Private Sub Shut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub StopProc_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
StopSearch = True
End Sub

Private Sub XPButton1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim p As New OleData
Dim r As New clsReg
Dim k As cRegKeys


Dim kCount, i As Integer
Dim ProgId As String
Dim CLSId As String
Dim FileName, FilePath As String

Dim l As ListItem

StopProc.Enabled = True
Shut.Enabled = False

Set k = r.ListKeys(HKEY_CLASSES_ROOT, "CLSID")

kCount = k.Count
SpFolder = SystemFolder
WinDir = fso.GetSpecialFolder(SpFolder)

For i = 1 To kCount
DoEvents
    CLSId = k.Item(i).Key
    
    If StopSearch = True Then Exit For
    Label1.Caption = "Processing:" & i
    
    If p.IsCLSID(CLSId) Then
        ProgId = p.CLSIDToProgID(CLSId)
        FilePath = p.GetFileFromCLSID(CLSId)
    
'    DoEvents

       Set l = ListView1.ListItems.Add(, , CLSId)

       l.SubItems(1) = Trim(ProgId)
       l.SubItems(2) = FilePath
    End If
    
    
Next i
StopProc.Enabled = False
Shut.Enabled = True
End Sub
