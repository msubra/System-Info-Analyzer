VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OleData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


' Convert a string representation of a CLSID, including the
' surrounding brace brackets, into the corresponding ProgID.

Public Property Get CLSIDToProgID(ByVal CLSId As String) As String
    Dim pResult As Long, pChar As Long
    Dim Char As Integer, Length As Long
    ' No need to use a special UDT
    Dim guid(15) As Byte

    ' convert from string to a binary CLSID
    CLSIDFromString StrPtr(CLSId), guid(0)
    ' convert to a string, get pointer to result
    ProgIDFromClsID guid(0), pResult
    ' find the terminating null char
    pChar = pResult - 2
    Do
        pChar = pChar + 2
        CopyMemory Char, ByVal pChar, 2
    Loop While Char
    ' now get the entire string in one operation
    Length = pChar - pResult
    ' no need for a temporary string
    CLSIDToProgID = Space$(Length \ 2)
    CopyMemory ByVal StrPtr(CLSIDToProgID), ByVal pResult, Length
End Property
' Get the name of the DLL or EXE that contains
' the class whose CLSID is passed as an argument
'
' Note: Remote servers aren't taken into account
Public Property Get ProgIDToCLSID(ProgId As String) As String
Dim r As New clsReg
Dim CLSId As String

CLSId = r.ListValues(HKEY_CLASSES_ROOT, ProgId & "\CLSID").item(1).Value
ProgIDToCLSID = CLSId
End Property

Public Property Get GetFileFromCLSID(ByVal CLSId As String) As String
Dim GetFile 'FromProgID

    Const HKEY_CLASSES_ROOT = &H80000000

    ' try to read the HKEY_CLASSES_ROOT\CLSID\{...}\InProcServer32 value
    GetFile = GetRegistryValue(HKEY_CLASSES_ROOT, _
        "CLSID\" & CLSId & "\InProcServer32", "")
    
    If Len(GetFile) = 0 Then
        ' try to read the HKEY_CLASSES_ROOT\CLSID\{...}\LocalServer32 value
        GetFile = GetRegistryValue(HKEY_CLASSES_ROOT, _
            "CLSID\" & CLSId & "\LocalServer32", "")
    End If
    GetFileFromCLSID = GetFile
End Property
Public Property Get GetFileFromProgID(ByVal ProgId As String) As String


    Dim CLSId As String
    Const HKEY_CLASSES_ROOT = &H80000000

    ' get the CLSID from the registry, exit if not found
    CLSId = GetRegistryValue(HKEY_CLASSES_ROOT, ProgId & "\CLSID", "")
    If Len(CLSId) = 0 Then Exit Property
    
    ' try to read the HKEY_CLASSES_ROOT\CLSID\{...}\InProcServer32 value
    GetFileFromProgID = GetRegistryValue(HKEY_CLASSES_ROOT, _
        "CLSID\" & CLSId & "\InProcServer32", "")
    ' exit if found
    If Len(GetFileFromProgID) Then Exit Property
    
    ' try to read the HKEY_CLASSES_ROOT\CLSID\{...}\LocalServer32 value
    GetFileFromProgID = GetRegistryValue(HKEY_CLASSES_ROOT, _
        "CLSID\" & CLSId & "\LocalServer32", "")
    
End Property
Public Property Get IsProgID(ProgId As String) As Boolean
Dim r As New clsReg

IsProgID = r.KeyExist(HKEY_CLASSES_ROOT, ProgId)

End Property
Public Property Get IsCLSID(CLSId As String) As Boolean
Dim r As New clsReg

IsCLSID = r.KeyExist(HKEY_CLASSES_ROOT, "CLSID" & "\" & CLSId)

End Property
' Read a Registry value
'
' Use KeyName = "" for the default value
' If the value isn't there, it returns the DefaultValue
' argument, or Empty if the argument has been omitted
'
' Supports DWORD, REG_SZ, REG_EXPAND_SZ, REG_BINARY and REG_MULTI_SZ
' REG_MULTI_SZ values are returned as a null-delimited stream of strings
' (VB6 users can use SPlit to convert to an array of string)

Private Function GetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, Optional DefaultValue As Variant) As Variant
    Dim handle As Long
    Dim resLong As Long
    Dim resString As String
    Dim resBinary() As Byte
    Dim Length As Long
    Dim retVal As Long
    Dim valueType As Long
    
    ' Prepare the default result
    GetRegistryValue = IIf(IsMissing(DefaultValue), Empty, DefaultValue)
    
    ' Open the key, exit if not found.
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
        Exit Function
    End If
    
    ' prepare a 1K receiving resBinary
    Length = 1024
    ReDim resBinary(0 To Length - 1) As Byte
    
    ' read the registry key
    retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
        Length)
    ' if resBinary was too small, try again
    If retVal = ERROR_MORE_DATA Then
        ' enlarge the resBinary, and read the value again
        ReDim resBinary(0 To Length - 1) As Byte
        retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
            Length)
    End If
    
    ' return a value corresponding to the value type
    Select Case valueType
        Case REG_DWORD
            CopyMemory resLong, resBinary(0), 4
            GetRegistryValue = resLong
        Case REG_SZ, REG_EXPAND_SZ
            ' copy everything but the trailing null char
            resString = Space$(Length - 1)
            CopyMemory ByVal resString, resBinary(0), Length - 1
            GetRegistryValue = resString
        Case REG_BINARY
            ' resize the result resBinary
            If Length <> UBound(resBinary) + 1 Then
                ReDim Preserve resBinary(0 To Length - 1) As Byte
            End If
            GetRegistryValue = resBinary()
        Case REG_MULTI_SZ
            ' copy everything but the 2 trailing null chars
            resString = Space$(Length - 2)
            CopyMemory ByVal resString, resBinary(0), Length - 2
            GetRegistryValue = resString
        Case Else
            RegCloseKey handle
            err.Raise 1001, , "Unsupported value type"
    End Select
    
    ' close the registry key
    RegCloseKey handle
End Function

