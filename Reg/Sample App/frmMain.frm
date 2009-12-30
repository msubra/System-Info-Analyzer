VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Registry Class Demo"
   ClientHeight    =   6285
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   6285
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReadBinary 
      Caption         =   "ReadBinary"
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmdWriteBinary 
      Caption         =   "WriteBinary"
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   4800
      Width           =   1455
   End
   Begin VB.TextBox txtBinary 
      Height          =   375
      Index           =   4
      Left            =   3240
      TabIndex        =   15
      Text            =   "97"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtBinary 
      Height          =   375
      Index           =   3
      Left            =   2520
      TabIndex        =   14
      Text            =   "237"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtBinary 
      Height          =   375
      Index           =   2
      Left            =   1800
      TabIndex        =   13
      Text            =   "67"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtBinary 
      Height          =   375
      Index           =   1
      Left            =   1080
      TabIndex        =   12
      Text            =   "43"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtBinary 
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   11
      Text            =   "175"
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton cmdReadDWord 
      Caption         =   "ReadDWord"
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   3480
      Width           =   1455
   End
   Begin VB.CommandButton cmdWriteDWord 
      Caption         =   "WriteDWord"
      Height          =   375
      Left            =   360
      TabIndex        =   9
      Top             =   3480
      Width           =   1455
   End
   Begin VB.TextBox txtWriteDWord 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Text            =   "2578"
      Top             =   3000
      Width           =   6255
   End
   Begin VB.TextBox txtWriteString 
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Text            =   "Put Some Text in here"
      Top             =   1800
      Width           =   6255
   End
   Begin VB.CommandButton cmdReadString 
      Caption         =   "ReadString"
      Height          =   375
      Left            =   1920
      TabIndex        =   6
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdWriteString 
      Caption         =   "WriteString"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   2280
      Width           =   1455
   End
   Begin VB.CommandButton cmdDeleteKey 
      Cancel          =   -1  'True
      Caption         =   "DeleteKey"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   840
      Width           =   1575
   End
   Begin VB.CommandButton cmdCreateKey 
      Caption         =   "CreateKey"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdKeyExist 
      Caption         =   "KeyExist"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
   Begin VB.TextBox txtRegSubKey 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Text            =   "SOFTWARE\RegDemo"
      Top             =   240
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Sub Key:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************
'* RegClass                                                             *
'* ActiveX object for reading and writing to the Registry.              *
'*                                                                      *
'* Writen by David Wheater, Ackworth.Computing                          *
'* Copyright © 2001, David Wheater                                      *
'*                                                                      *
'* You may freely use the object and/or the code contained within it    *
'* in your own personal projects. The code is not to be used for any    *
'* commercial venture, without the permission of the author.            *
'* http://www.ackworth.com (no spam please, I won't buy anything)       *
'*                                                                      *
'* Special thanks to the people at www.freevbcode.com and www.vbapi.com *
'* without whom, I wouldn't have known how to create this class.        *
'*                                                                      *
'* You must first register the RegDemo.dll file, to do this open the    *
'* RegDemo project and complie it, or use regsvr32                      *
'*                                                                      *
'************************************************************************



Dim mReg As clsReg

Private Sub Form_Load()
    'On loading the form we create an instance of the class
    Set mReg = New clsReg

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'On Unloading the form we set the instance of the class
    'to nothing to close it.
    Set mReg = Nothing
End Sub

Private Sub cmdKeyExist_Click()
    If mReg.KeyExist(HKEY_LOCAL_MACHINE, txtRegSubKey.Text) Then
        MsgBox "The key " & txtRegSubKey & " does exist", vbOKOnly + vbInformation, "KeyExist"
    Else
        MsgBox "The key " & txtRegSubKey & " does not exist", vbOKOnly + vbCritical, "KeyExist"
    End If
End Sub

Private Sub cmdCreateKey_Click()
    If mReg.CreateKey(HKEY_LOCAL_MACHINE, txtRegSubKey.Text) Then
        MsgBox txtRegSubKey & " has been created.", vbOKOnly + vbInformation, "CreateKey"
    Else
        MsgBox txtRegSubKey & " has not been created.", vbOKOnly + vbCritical, "CreateKey"
    End If
End Sub

Private Sub cmdDeleteKey_Click()
    If mReg.DeleteKey(HKEY_LOCAL_MACHINE, txtRegSubKey.Text) = True Then
        MsgBox txtRegSubKey & " has been deleted.", vbOKOnly + vbInformation, "DeleteKey"
    Else
        MsgBox txtRegSubKey & " has not been deleted.", vbOKOnly + vbCritical, "DeleteKey"
    End If
End Sub


Private Sub cmdWriteString_Click()
    mReg.WriteString HKEY_LOCAL_MACHINE, txtRegSubKey, "WriteString", txtWriteString
End Sub

Private Sub cmdReadString_Click()
    MsgBox mReg.ReadString(HKEY_LOCAL_MACHINE, txtRegSubKey.Text, "WriteString", "Default String"), vbOKOnly, "WriteString / ReadString"
End Sub

Private Sub cmdWriteDWord_Click()
    mReg.WriteDWord HKEY_LOCAL_MACHINE, txtRegSubKey.Text, "WriteDWord", CLng(txtWriteDWord.Text)
End Sub

Private Sub cmdReadDWord_Click()
    Dim MyLong As Long
    
    MyLong = mReg.ReadDWord(HKEY_LOCAL_MACHINE, txtRegSubKey.Text, "WriteDWord", 0)
    MsgBox "The Long Variable you stored is: " & CStr(MyLong)
End Sub

Private Sub cmdWriteBinary_Click()

    Dim byArray(4) As Byte
    Dim x As Integer
    
    For x = 0 To 4
        byArray(x) = CByte(txtBinary(x).Text)
    Next

    mReg.WriteBinary HKEY_LOCAL_MACHINE, txtRegSubKey.Text, "WriteBinary", byArray

End Sub


Private Sub cmdReadBinary_Click()

    Dim MyVar As Variant
    Dim msg As String
    Dim x As Integer

    MyVar = mReg.ReadBinary(HKEY_LOCAL_MACHINE, txtRegSubKey.Text, "WriteBinary")
    msg = ""
    For x = LBound(MyVar) To UBound(MyVar)
        msg = msg & CStr(MyVar(x)) & " "
    Next x
    MsgBox "The numbers you stored are: " & msg
End Sub
















