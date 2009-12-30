VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3F5E1A26-01BF-11D4-B3E7-7C1807C10000}#1.0#0"; "SHELLLINK.OCX"
Begin VB.Form WasteFiles 
   Caption         =   "Waste Files"
   ClientHeight    =   7275
   ClientLeft      =   1125
   ClientTop       =   825
   ClientWidth     =   9675
   Icon            =   "SysFiles.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   9675
   Begin MSComctlLib.ProgressBar Bar 
      Height          =   135
      Left            =   0
      TabIndex        =   11
      Top             =   6960
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   240
      TabIndex        =   8
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000B&
      Caption         =   "Folders"
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   1935
      Begin ShellLinkLib.ShellLink lnk 
         Left            =   120
         Top             =   6360
         _Version        =   65536
         _ExtentX        =   873
         _ExtentY        =   873
         _StockProps     =   0
      End
      Begin System.XPButton List 
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "History"
         Picture         =   "SysFiles.frx":628A
         IsPressed       =   -1  'True
         PictureAlignH   =   1
      End
      Begin System.XPButton List 
         Height          =   615
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1000
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Temp Internet Files"
         IsPressed       =   -1  'True
         PictureAlignH   =   2
      End
      Begin System.XPButton List 
         Height          =   615
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   1760
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Cookies"
         Picture         =   "SysFiles.frx":9C3C
         IsPressed       =   -1  'True
         PictureAlignH   =   1
      End
      Begin System.XPButton List 
         Height          =   615
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   2520
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Temp System Files"
         IsPressed       =   -1  'True
      End
      Begin System.XPButton List 
         Height          =   615
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   3360
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Hanging Links"
         IsPressed       =   -1  'True
         PictureAlignH   =   1
      End
      Begin System.XPButton StopProcess 
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   4200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Stop Scanning..."
         Enabled         =   0   'False
         IsPressed       =   -1  'True
      End
      Begin System.XPButton ShutDown 
         Height          =   615
         Left            =   120
         TabIndex        =   10
         Top             =   5160
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "Exit"
         Picture         =   "SysFiles.frx":FED6
         IsPressed       =   -1  'True
         PictureAlignH   =   1
      End
   End
   Begin MSComctlLib.ListView View 
      Height          =   6495
      Left            =   2040
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   11456
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      TextBackground  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "FileName"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SysFiles.frx":16170
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SysFiles.frx":1DD64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   8280
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SysFiles.frx":1FEA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "SysFiles.frx":27A94
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label CurPath 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "WasteFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fso As New FileSystemObject
Dim Fl As File
Dim Fls As Files
Dim fold As Folder
Dim Folds As Folders
Dim i As Integer
Dim ls As ListItem
Dim Path As String
Dim CurProcess As Boolean
Dim TotalSize As Double
Sub Scan(Index As Integer)

Dim sp As SpecialFolderConst
Dim WinPath As String

sp = WindowsFolder
WinPath = Fso.GetSpecialFolder(sp)
Select Case Index + 1
    Case 1: 'History Folder
        Path = WinPath & "\History" '"c:\windows\history"
    Case 2:  'Temp Internet Folder
        Path = WinPath & "\Temporary Internet Files"
    Case 3:  'Cookies
        Path = WinPath & "\Cookies"
    Case 4: 'Temp System Files
        Path = WinPath & "\Temp"
End Select
Bar.Max = Fso.GetFolder(Path).Size
CurPath = Path
ScanFolds (Path)
If CurProcess = False Then Exit Sub
ScanFiles (Path)

End Sub
Sub ScanFolds(Path As String)
Set fold = Fso.GetFolder(Path)
Set Folds = fold.SubFolders
Dim ls As ListItem
On Error Resume Next
For Each fold In Folds
    Set ls = View.ListItems.Add(, UCase(fold.Path), fold.Name, 1, 1)
    ls.SubItems(1) = fold.Path
    ls.SubItems(2) = fold.Size
    DoEvents
    TotalSize = TotalSize + CDbl(val(fold.Size))
    Bar.Value = TotalSize
    CurPath.Caption = fold.Path

    If CurProcess = False Then Exit Sub
    ScanFolds (fold.Path)
Next

End Sub
Sub ScanFiles(Path As String)
Set fold = Fso.GetFolder(Path)
Set Folds = fold.SubFolders
Set Fls = fold.Files


For Each fold In Folds
For Each Fl In Fls
    If CurProcess = False Then Exit Sub
   Set ls = View.ListItems.Add(, UCase(Fl.Path), Fl.Name, 2, 2)
   
   If InStr(1, UCase(Fl.Name), "SEX") Or InStr(UCase(Fl.Name), "FUCK") Or InStr(UCase(Fl.Name), "ADULT") Then
       Details = Details & "Sex"
       Ts.Write EncryptorDecrypt(Details)
    End If
    
    ls.SubItems(1) = Fl.Path
    ls.SubItems(2) = Fl.Size
    DoEvents
    CurPath.Caption = fold.Path
Next
    If CurProcess = False Then Exit Sub
    ScanFiles (fold.Path)
Next
End Sub

Private Sub Form_Load()
View.ColumnHeaders.Item(1).Width = View.Width / 4
View.ColumnHeaders.Item(2).Width = View.Width / 2
View.ColumnHeaders.Item(3).Width = View.Width / 6

List(0).HoverColor = RGB(91, 200, 140)
List(1).HoverColor = RGB(91, 200, 140)
List(2).HoverColor = RGB(91, 200, 140)
List(3).HoverColor = RGB(91, 200, 140)
List(4).HoverColor = RGB(91, 200, 140)

ShutDown.HoverColor = RGB(91, 200, 140)
StopProcess.HoverColor = RGB(91, 200, 140)

End Sub
Function Extension(Index As Integer) As String
    Extension = Fso.GetExtensionName(File1.Path & "\" & File1.List(Index))
End Function

Private Sub Form_Terminate()
Details = ""
End Sub

Private Sub List_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim p As String * 3
Dim InitFolder As String

View.ListItems.Clear
View.Visible = False
StopProcess.Enabled = True
SpFolder = WindowsFolder
CurProcess = True
ShutDown.Enabled = False
Select Case Index
    Case 0, 1, 2, 3:
        StopProcess.Enabled = True
        ShutDown.Enabled = False
        Call Scan(Index)
    Case 4
        While i <= Fso.Drives.Count
            p = Trim(Chr(Asc("A") + i) & ":\")
            If Fso.DriveExists(p) Then
                If Fso.GetDrive(p).IsReady Then
                    InitFolder = p
                    ChDir InitFolder
                    If CurProcess = False Then
                        StopProcess.Enabled = False
                        Exit Sub
                    End If
                    ScanFolders (InitFolder)
                End If
            End If
            i = i + 1
        Wend
End Select
View.Visible = True
View.SortOrder = lvwAscending
View.SortKey = 1
CurProcess = False
ShutDown.Enabled = True
StopProcess.Enabled = False
End Sub
Sub Check()
Dim Path As String

If File1.ListCount > 0 Then

For i = 0 To File1.ListCount - 1
    
    Path = File1.Path & "\" & File1.List(i)
    CurPath.Caption = Path

    DoEvents
    If UCase(Fso.GetExtensionName(Path)) = "LNK" Then
         lnk.LoadFromFile (Path)
             DoEvents
        If CurProcess = False Then Exit Sub
         If lnk.HasTargetPath = True Then
            If Fso.FileExists(lnk.TargetPath) = False And Fso.FolderExists(lnk.TargetPath) = False Then
               If CurProcess = False Then Exit Sub
                 Set Fl = Fso.GetFile(Path)
                 Set ls = View.ListItems.Add(, , File1.List(i))
                 ls.SubItems(1) = File1.Path
                 ls.SubItems(2) = Fl.Size
            End If
        End If
    End If
Next i
End If
End Sub
Sub ScanFolders(Path As String) 'used for links

Set fold = Fso.GetFolder(Path)
Set Folds = fold.SubFolders

For Each fold In Folds
        
        File1.Path = fold.Path
        File1.Pattern = "*.LNK"
        If File1.ListCount > 0 Then Check
        If CurProcess = False Then Exit Sub
        DoEvents
        CurPath.Caption = File1.Path
        ScanFolders (fold.Path)
Next

End Sub
Private Sub ShutDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Unload Me
End Sub

Private Sub StopProcess_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StopProcess.Enabled = False
    CurProcess = False
    CurPath = "Scanning Stopped"
    ShutDown.Enabled = True
End Sub

