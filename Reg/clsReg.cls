VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
'************************************************************************

Private Type FILETIME
  dwLowDateTime As Long
  dwHighDateTime As Long
End Type

Private Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Private mRegKey As New cRegKeys
Private mRegValues As New cRegValues

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
                (ByVal hKey As Long, _
                 ByVal lpSubKey As String, _
                 ByVal ulOptions As Long, _
                 ByVal samDesired As Long, _
                 phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
                (ByVal hKey As Long, _
                 ByVal lpValueName As String, _
                 ByVal lpReserved As Long, _
                 lpType As Long, lpData As Any, _
                 lpcbData As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
                (ByVal hKey As Long, _
                 ByVal lpSubKey As String) As Long
                 
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
                (ByVal hKey As Long, _
                 ByVal lpValueName As String) As Long
                 
Private Declare Function RegCreatekey Lib "advapi32.dll" Alias "RegCreateKeyA" _
                (ByVal hKey As Long, _
                 ByVal lpSubKey As String, _
                 phkResult As Long) As Long
                 
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
                (ByVal hKey As Long, _
                 ByVal lpSubKey As String, _
                 ByVal Reserved As Long, _
                 ByVal lpClass As String, _
                 ByVal dwOptions As Long, _
                 ByVal samDesired As Long, _
                 lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                 phkResult As Long, _
                 lpdwDisposition As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
                (ByVal hKey As Long, _
                 ByVal lpValueName As String, _
                 ByVal Reserved As Long, _
                 ByVal dwType As Long, _
                 lpData As Any, _
                 ByVal cbData As Long) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
                (ByVal hKey As Long, _
                 ByVal dwIndex As Long, _
                 ByVal lpName As String, _
                 lpcbName As Long, _
                 lpReserved As Long, _
                 ByVal lpClass As String, _
                 lpcbClass As Long, _
                 lpftLastWriteTime As FILETIME) As Long
                 
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
                (ByVal hKey As Long, _
                 ByVal dwIndex As Long, _
                 ByVal lpValueName As String, _
                 lpcbValueName As Long, _
                 ByVal lpReserved As Long, _
                 lpType As Long, _
                 lpData As Byte, _
                 lpcbData As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
                (Destination As Any, _
                 Source As Any, _
                 ByVal Length As Long)


'Main Registry Keys
'------------------------------------------------------------------------
Public Enum HKEY_Type
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

'Registry Entry Types
'------------------------------------------------------------------------
Public Enum Reg_Type
    REG_NONE = 0                  'No data type.
    REG_SZ = 1                    'A string terminated by a null character.
    REG_EXPAND_SZ = 2             'A null-terminated string which contains unexpanded environment variables.
    REG_BINARY = 3                'A non-text sequence of bytes.
    REG_DWORD = 4                 'Same as REG_DWORD_LITTLE_ENDIAN.
    REG_DWORD_LITTLE_ENDIAN = 4   'A 32-bit integer stored in little-endian format. This is the way Intel-based computers normally store numbers.
    REG_DWORD_BIG_ENDIAN = 5      'A 32-bit integer stored in big-endian format. This is the opposite of the way Intel-based computers normally store numbers -- the word order is reversed.
    REG_LINK = 6                  'A Unicode symbolic link.
    REG_MULTI_SZ = 7              'A series of strings, each separated by a null character and the entire set terminated by a two null characters.
    REG_RESOURCE_LIST = 8         'A list of resources in the resource map.
End Enum
'Secuirty Constants
'------------------------------------------------------------------------
Const KEY_ALL_ACCESS = &HF003F      'Permission for all types of access.
Const KEY_CREATE_LINK = &H20        'Permission to create symbolic links.
Const KEY_CREATE_SUB_KEY = &H4      'Permission to create subkeys.
Const KEY_ENUMERATE_SUB_KEYS = &H8  'Permission to enumerate subkeys.
Const KEY_EXECUTE = &H20019         'Same as KEY_READ.
Const KEY_NOTIFY = &H10             'Permission to give change notification.
Const KEY_QUERY_VALUE = &H1         'Permission to query subkey data.
Const KEY_READ = &H20019            'Permission for general read access.
Const KEY_SET_VALUE = &H2           'Permission to set subkey data.
Const KEY_WRITE = &H20006           'Permission for general write access.


Const REG_OPTION_NON_VOLATILE = 0

'Error Numbers
'------------------------------------------------------------------------
Const REG_ERR_OK = 0                'No Problems
Const REG_ERR_NOT_EXIST = 1         'Key does not exist
Const REG_ERR_NOT_STRING = 2        'Value is not a string
Const REG_ERR_NOT_DWORD = 4         'Value not DWORD
'
Const ERROR_NONE = 0
Const ERROR_BADDB = 1
Const ERROR_BADKEY = 2
Const ERROR_CANTOPEN = 3
Const ERROR_CANTREAD = 4
Const ERROR_CANTWRITE = 5
Const ERROR_OUTOFMEMORY = 6
Const ERROR_ARENA_TRASHED = 7
Const ERROR_ACCESS_DENIED = 8
Const ERROR_INVALID_PARAMETERS = 87
Const ERROR_NO_MORE_ITEMS = 259



Public Function KeyExist(Key As HKEY_Type, sSubKey As String) As Boolean
    Dim hKey As Long
    Dim retval As Long

    retval = RegOpenKeyEx(Key, sSubKey, 0, KEY_QUERY_VALUE, hKey)
    Debug.Print hKey
    If retval = ERROR_NONE Then
        KeyExist = True
    Else
        KeyExist = False
    End If
    RegCloseKey hKey
End Function

Public Function CreateKey(Key As HKEY_Type, sSubKey As String) As Boolean
Attribute CreateKey.VB_HelpID = 1002
    Dim hKey As Long
    Dim retval As Long
    
    retval = RegCreatekey(Key, sSubKey, hKey)
    If retval <> ERROR_NONE Then
        CreateKey = False
    Else
        retval = RegCloseKey(hKey)
        CreateKey = True
    End If

End Function

Public Function DeleteKey(Key As HKEY_Type, sSubKey As String) As Boolean
Attribute DeleteKey.VB_HelpID = 1003
    Dim retval As Long
    
    retval = RegDeleteKey(Key, sSubKey)
    If retval = ERROR_NONE Then
        DeleteKey = True
    Else
        DeleteKey = False
    End If
End Function


Public Function WriteString(Key As HKEY_Type, SubKey As String, sName As String, sData As String) As Boolean
Attribute WriteString.VB_HelpID = 1004

    Dim hKey As Long
    Dim retval As Long
    Dim deposit As Long
    Dim secattr As SECURITY_ATTRIBUTES
    
    secattr.nLength = Len(secattr)
    secattr.lpSecurityDescriptor = 0
    secattr.bInheritHandle = 1
    
    retval = RegCreateKeyEx(Key, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If retval <> ERROR_NONE Then
        WriteString = False
        Exit Function
    End If

    retval = RegSetValueEx(hKey, sName, 0, REG_SZ, ByVal sData, Len(sData))
    
    If retval <> ERROR_NONE Then
        WriteString = False
        Exit Function
    End If
    
    retval = RegCloseKey(hKey)
    WriteString = True
End Function


Public Function ReadString(Key As HKEY_Type, SubKey As String, sName As String, Optional sDefault As String = "") As String
Attribute ReadString.VB_HelpID = 1005
    
    Dim hKey As Long
    Dim retval As Long
    Dim sBuffer As String
    Dim slength As Long
    Dim DataType As Long
    
    retval = RegOpenKeyEx(Key, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If retval <> ERROR_NONE Then
        ReadString = sDefault
        Exit Function
    End If
    
    sBuffer = Space(255)
    slength = 255
    retval = RegQueryValueEx(hKey, sName, 0, DataType, ByVal sBuffer, slength)
    
    If retval = ERROR_NONE Then
        If DataType = REG_SZ Or DataType = REG_EXPAND_SZ Then
            sBuffer = Left(sBuffer, slength - 1)
            ReadString = sBuffer
        Else
            ReadString = sDefault
        End If
    Else
        ReadString = sDefault
    End If
    
    retval = RegCloseKey(hKey)
End Function



Public Function WriteDWord(Key As HKEY_Type, SubKey As String, sName As String, lData As Long) As Boolean
Attribute WriteDWord.VB_HelpID = 1006

    Dim hKey As Long
    Dim retval As Long
    Dim deposit As Long
    Dim secattr As SECURITY_ATTRIBUTES
    
    secattr.nLength = Len(secattr)
    secattr.lpSecurityDescriptor = 0
    secattr.bInheritHandle = 1
    
    retval = RegCreateKeyEx(Key, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If retval <> ERROR_NONE Then
        WriteDWord = False
        Exit Function
    End If
    
    retval = RegSetValueEx(hKey, sName, 0&, REG_DWORD, lData, 4)
    If retval = ERROR_NONE Then
        WriteDWord = True
    Else
        WriteDWord = False
    End If
    
    retval = RegCloseKey(hKey)
End Function


Public Function ReadDWord(Key As HKEY_Type, SubKey As String, sName As String, Optional lDefault As Long = 0) As Long
Attribute ReadDWord.VB_HelpID = 1007

    Dim hKey As Long
    Dim retval As Long
    Dim DataType As Long
    Dim lBuffer As Long
    
    retval = RegOpenKeyEx(Key, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If retval <> ERROR_NONE Then
        ReadDWord = lDefault
        Exit Function
    End If
    
    retval = RegQueryValueEx(hKey, sName, 0, DataType, lBuffer, 4)
    
    If retval = ERROR_NONE Then
        If DataType = REG_DWORD Then
            ReadDWord = lBuffer
        Else
            ReadDWord = lDefault
        End If
    Else
        ReadDWord = lDefault
    End If
    retval = RegCloseKey(hKey)
End Function


Public Function WriteBinary(Key As HKEY_Type, SubKey As String, sName As String, byData() As Byte) As Boolean
Attribute WriteBinary.VB_HelpID = 1009

    Dim hKey As Long
    Dim retval As Long
    Dim deposit As Long
    Dim secattr As SECURITY_ATTRIBUTES
    
    secattr.nLength = Len(secattr)
    secattr.lpSecurityDescriptor = 0
    secattr.bInheritHandle = 1
    
    retval = RegCreateKeyEx(Key, SubKey, 0, "", 0, KEY_WRITE, secattr, hKey, deposit)
    If retval <> ERROR_NONE Then
        WriteBinary = False
        Exit Function
    End If

    retval = RegSetValueEx(hKey, sName, 0&, REG_BINARY, byData(0), UBound(byData) + 1)
    
    If retval = ERROR_NONE Then
        WriteBinary = True
    Else
        WriteBinary = False
    End If
    retval = RegCloseKey(hKey)
End Function


Public Function ReadBinary(Key As HKEY_Type, SubKey As String, sName As String, Optional byDefault As Variant = 0) As Variant
Attribute ReadBinary.VB_HelpID = 1008

    Dim hKey As Long
    Dim retval As Long
    Dim DataType As Long
    Dim byBuffer() As Byte
    Dim lbuffersize As Long
    Dim byDefArray(0) As Byte

    byDefArray(0) = 0


    retval = RegOpenKeyEx(Key, SubKey, 0, KEY_ALL_ACCESS, hKey)
    If retval <> ERROR_NONE Then
        If VarType(byDefault) = vbArray + vbByte Then
            ReadBinary = byDefault
        Else
            ReadBinary = 0
        End If
        Exit Function
    End If

    retval = RegQueryValueEx(hKey, sName, 0&, DataType, ByVal 0&, lbuffersize)
    
    If retval = ERROR_NONE Then
        If DataType = REG_BINARY Then
            ReDim byBuffer(lbuffersize - 1)
            retval = RegQueryValueEx(hKey, sName, 0&, DataType, byBuffer(0), lbuffersize)
            ReadBinary = byBuffer
        Else
            If VarType(byDefault) = vbArray + vbByte Then
                ReadBinary = byDefault
            Else
                ReadBinary = byDefArray
            End If
        End If
     Else
        If VarType(byDefault) = vbArray + vbByte Then
            ReadBinary = byDefault
        Else
            ReadBinary = byDefArray
        End If
    End If
    retval = RegCloseKey(hKey)
        
End Function


Public Function ListKeys(Key As HKEY_Type, SubKey As String) As cRegKeys
Attribute ListKeys.VB_HelpID = 1010

    Dim keyname As String
    Dim keylen As Long
    Dim classname As String
    Dim classlen As Long
    Dim lastwrite As FILETIME
    Dim hKey As Long
    Dim retval As Long
    Dim Index As Long
    
    retval = RegOpenKeyEx(Key, SubKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey)
    If retval <> ERROR_NONE Then
        mRegKey.Add ""
        Set ListKeys = mRegKey
        Exit Function
    End If
    
    Index = 0
    
    While retval = 0
    
        keyname = Space(255)
        keylen = 255
        classname = Space(255)
        classlen = 255
        
        retval = RegEnumKeyEx(hKey, Index, keyname, keylen, ByVal 0, classname, classlen, lastwrite)
        If retval = ERROR_NONE Then
            keyname = Left(keyname, keylen)
            mRegKey.Add keyname
        Else
      
        End If
        Index = Index + 1
    Wend
    Set ListKeys = mRegKey
    retval = RegCloseKey(hKey)
End Function


Public Function ListValues(Key As HKEY_Type, SubKey As String) As cRegValues
Attribute ListValues.VB_HelpID = 1011

    Dim ValueName As String
    Dim ValueLen As Long
    Dim DataType As Long
    Dim Data(0 To 254) As Byte
    Dim datalen As Long
    Dim datastring As String
    Dim hKey As Long
    Dim Index As Long
    Dim retval As Long
    
    Dim sBuffer As String
    Dim lBuffer As Long
    Dim BufferLen As Long
    Dim byBuffer() As Byte

    
    
    retval = RegOpenKeyEx(Key, SubKey, 0, KEY_QUERY_VALUE, hKey)
    If retval <> ERROR_NONE Then
        mRegValues.Add "", "", ""
        Set ListValues = mRegValues
        Exit Function
    End If

    Index = 0
    
    While retval = 0
    
        datalen = 255
        ValueName = Space(255)
        ValueLen = 255
        BinaryString = ""
        
        retval = RegEnumValue(hKey, Index, ValueName, ValueLen, 0, DataType, Data(0), datalen)
        If retval = ERROR_NONE Then
            ValueName = Left(ValueName, ValueLen)
            
            Select Case DataType
            
                Case REG_SZ, REG_EXPAND_SZ
                    mRegValues.Add ValueName, DataType, ReadString(Key, SubKey, ValueName, "")
            
                Case REG_DWORD, REG_DWORD_LITTLE_ENDIAN
                    mRegValues.Add ValueName, REG_DWORD, ReadDWord(Key, SubKey, ValueName, 0)

                Case REG_BINARY
                    mRegValues.Add ValueName, DataType, ReadBinary(Key, SubKey, ValueName)
                
                Case Else
                    mRegValues.Add ValueName, DataType, Nothing
                
            End Select

        End If
        Index = Index + 1
    Wend
    
    Set ListValues = mRegValues
    retval = RegCloseKey(hKey)
    
End Function



