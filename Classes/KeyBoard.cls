VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "KeyBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function GetKBCodePage Lib "user32" () As Long
Private Declare Function GetOEMCP Lib "kernel32" () As Long
Dim Handle As Long

Public Property Get GetKeyboardType() As String
Dim s As String

Handle = GetOEMCP()

Select Case Handle
Case 437
    s = "MS-DOS United States"
Case 708
    s = "Arabic (ASMO 708)"
Case 709
    s = "Arabic (ASMO 449+, BCON V4)"
Case 710
    s = "Arabic (Transparent Arabic)"
Case 720
    s = "Arabic (Transparent ASMO)"
Case 737
    s = "Greek (formerly 437G)"
Case 775
    s = "Baltic"
Case 850
    s = "MS-DOS Multilingual (Latin I)"
Case 852
    s = "MS-DOS Slavic (Latin II)"
Case 855
    s = "IBM Cyrillic (primarily Russian)"
Case 857
    s = "IBM Turkish"
Case 860
    s = "MS-DOS Portuguese"
Case 861
    s = "MS-DOS Icelandic"
Case 862
    s = "Hebrew"
Case 863
    s = "MS-DOS Canadian-French"
Case 864
    s = "Arabic"
Case 865
    s = "MS-DOS Nordic"
Case 866
    s = "MS-DOS Russian (former USSR)"
Case 869
    s = "IBM Modern Greek"
Case 874
    s = "Thai"
Case 932
    s = "Japan"
Case 936
    s = "Chinese (PRC, Singapore)"
Case 949
    s = "Korean"
Case 950
    s = "Chinese (Taiwan, Hong Kong)"
Case 1361
    s = "Korean (Johab)"
End Select

GetKeyboardType = s
End Property



