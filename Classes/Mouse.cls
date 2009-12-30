VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim v As New clsValues
Public Property Get MouseButtonsCount() As Integer
MouseButtonsCount = SystemMetrics(SM_CMOUSEBUTTONS)
End Property

Public Property Get ButtonSwapped() As Integer
ButtonSwapped = SystemMetrics(SM_SWAPBUTTON)
End Property
Public Property Get MousePresent() As Boolean
MousePresent = SystemMetrics(SM_MOUSEPRESENT)

End Property
Public Property Get MouseWheelPresent() As Boolean
MouseWheelPresent = SystemMetrics(SM_MOUSEPRESENT)

End Property
Public Property Get DblClickTime() As Integer
DblClickTime = GetDoubleClickTime
End Property
Public Property Get SetDblClickTime(ClickTime As Long) As Integer
SetDblClickTime = SetDoubleClickTime(ClickTime)
End Property
Public Property Get MouseName() As String
MouseName = v("DriverDesc").Value
End Property
Public Property Get MouseType() As String
MouseType = v("MouseType").Value
End Property

Private Sub Class_Initialize()

v.hKey = HKEY_LOCAL_MACHINE
v.Path = "System\CurrentControlSet\Services\Class\Mouse\0000"
v.Refresh

End Sub
