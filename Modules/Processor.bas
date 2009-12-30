Attribute VB_Name = "Pro"
Option Explicit

Function Processor(i As Long) As String
Dim p As String
Select Case i
    Case PROCESSOR_INTEL_386
        p = "Intel 386 Family"
    Case PROCESSOR_INTEL_486
        p = "Intel 486 Family"
    Case PROCESSOR_INTEL_PENTIUM
        p = "Intel Pentinum 586 Family"
    Case PROCESSOR_MIPS_R4000
        p = "MIPS R4000 Family"
    Case Else
        p = "unknown"
End Select
Processor = p
End Function

Function ProcessorName(i As Long) As String
If i = PROCESSOR_ARCHITECTURE_INTEL Then ProcessorName = "Genuine Intel Family"
End Function

Function CheckMathProcessor() As String
Dim v As New clsValues
Dim k As New clsKey

k.hKey = HKEY_LOCAL_MACHINE

k.Path = "Hardware\Description\System\FloatingPointProcessor"
Set v = k.Values

If v.Count <= 0 Then
    CheckMathProcessor = "Not Found"
ElseIf v.Count > 0 Then
    CheckMathProcessor = "Found"
End If
End Function

