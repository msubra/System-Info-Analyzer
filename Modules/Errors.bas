Attribute VB_Name = "Errors"
Option Explicit

Function SystemErrorDescription(ByVal ErrCode As Long) As String
    Dim buffer As String * 1024
    Dim Ret As Long
    ' return value is the length of the result message
    Ret = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, ErrCode, 0, buffer, _
        Len(buffer), 0)
    SystemErrorDescription = Left$(buffer, Ret)
End Function


