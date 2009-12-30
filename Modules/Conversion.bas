Attribute VB_Name = "Conversion"
Option Explicit

Public Function hiByte(ByVal w As Integer) As Byte
   If w And &H8000 Then
      hiByte = &H80 Or ((w And &H7FFF) \ &HFF)
   Else
      hiByte = w \ 256
    End If
End Function

Public Function HiWord(dw As Long) As Integer
 If dw And &H80000000 Then
      HiWord = (dw \ 65535) - 1
 Else
    HiWord = dw \ 65535
 End If
End Function

Public Function LoByte(w As Integer) As Byte
 LoByte = w And &HFF
End Function

Public Function LoWord(dw As Long) As Integer
  If dw And &H8000& Then
      LoWord = &H8000 Or (dw And &H7FFF&)
   Else
      LoWord = dw And &HFFFF&
   End If
End Function

Public Function MakeInt(ByVal LoByte As Byte, _
   ByVal hiByte As Byte) As Integer

MakeInt = ((hiByte * &H100) + LoByte)

End Function

Public Function MakeLong(ByVal LoWord As Integer, _
  ByVal HiWord As Integer) As Long

MakeLong = ((HiWord * &H10000) + LoWord)

End Function
