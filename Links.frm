VERSION 5.00
Object = "{3F5E1A26-01BF-11D4-B3E7-7C1807C10000}#1.0#0"; "SHELLLINK.OCX"
Begin VB.Form Links 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5625
   ClientLeft      =   2205
   ClientTop       =   1905
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7005
   StartUpPosition =   1  'CenterOwner
   Begin ShellLinkLib.ShellLink lnk 
      Left            =   5040
      Top             =   2880
      _Version        =   65536
      _ExtentX        =   873
      _ExtentY        =   873
      _StockProps     =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5400
      Top             =   3960
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   5160
      Width           =   2055
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2370
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   6855
   End
   Begin VB.CommandButton Status 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Bad Links:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   3360
      Width           =   840
   End
   Begin VB.Label Bad 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   3360
      Width           =   105
   End
   Begin VB.Label LinkCount 
      AutoSize        =   -1  'True
      Caption         =   "0"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   3000
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Links In Current Folders:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   45
   End
End
Attribute VB_Name = "Links"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fso As New FileSystemObject
Dim fold As Folder, Folds As Folders
Dim I As Integer
Dim InitFolder As String
Dim Ready As Boolean
Sub Check()
Dim Path As String

If File1.ListCount > 0 Then
LinkCount.Caption = File1.ListCount
For I = 0 To File1.ListCount - 1
    If Not Ready Then Exit For
    Path = File1.Path & "\" & File1.List(I)
    Label1.Caption = Path
    Bad = List1.ListCount
    DoEvents
    If UCase(fso.GetExtensionName(Path)) = "LNK" Then
         lnk.LoadFromFile (Path)
             DoEvents
         If lnk.HasTargetPath = True Then
            If fso.FileExists(lnk.TargetPath) = False Then
                 
                 List1.AddItem Path
            End If
        End If
    End If
Next I
End If
End Sub
Sub ScanFolders(Path As String)

Set fold = fso.GetFolder(Path)
Set Folds = fold.SubFolders

For Each fold In Folds
        
        File1.Path = fold.Path
        File1.Pattern = "*.LNK"
        If Not Ready Then Exit For
        If File1.ListCount > 0 Then Check
        DoEvents
        Bad = List1.ListCount
        Label1.Caption = File1.Path
        ScanFolders (fold.Path)
Next

End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_Activate()
SpFolder = WindowsFolder
WinDir = fso.GetSpecialFolder(SpFolder)

End Sub

Private Sub Status_Click()
Dim I As Integer
Dim P As String

Select Case UCase(Status.Caption)
    Case "START"
        List1.Clear
        Status.Caption = "Stop"
        Ready = True
        While I <= fso.Drives.Count
            P = Chr(Asc("A") + I) & ":\"
            If fso.DriveExists(P) Then
                If fso.GetDrive(P).IsReady Then
                    InitFolder = P
                    ChDir InitFolder
                    ScanFolders (InitFolder)
                End If
            End If
            I = I + 1
        Wend
       Label1.Caption = "Finished Searching"
    Case "STOP"
        Status.Caption = "Start"
        Label1.Caption = "Intreuppted By User"
        Ready = False
        Exit Sub
End Select
End Sub

Private Sub Timer1_Timer()
Bad = List1.ListCount
End Sub
