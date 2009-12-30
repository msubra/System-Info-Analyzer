Attribute VB_Name = "GDIObjectsManagement"
Option Explicit

'\\ --[GDIObjectsManagement]-----------------------------------------
'\\ Keeps track of all the GDI objects that have been created
'\\ so that they can be deleted when they are no longer needed.
'\\ -----------------------------------------------------------------


'\\ When an object's reference count = 0 then delete it...
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long


Private Declare Function EnumObjects Lib "gdi32" (ByVal hdc As Long, ByVal n As Long, ByVal lpGOBJEnumProc As Long, ByVal lpVoid As Long) As Long

Private CurrentObjectType As GDIObjectTypes
Private colGDIObjects As Collection
Dim foo As Long

Public Function GetDCPenCollection(ByVal hdc As Long) As Collection
Dim lRet As Long

Set colGDIObjects = New Collection
CurrentObjectType = OBJ_PEN

lRet = EnumObjects(hdc, CurrentObjectType, AddressOf ENUMOBJECTSPROC, VarPtr(foo))

Set GetDCPenCollection = colGDIObjects

End Function

Public Function GetDCBrushCollection(ByVal hdc As Long) As Collection

Dim lRet As Long

Set colGDIObjects = New Collection
CurrentObjectType = OBJ_BRUSH

lRet = EnumObjects(hdc, CurrentObjectType, AddressOf ENUMOBJECTSPROC, VarPtr(foo))

Set GetDCBrushCollection = colGDIObjects

End Function
'\\ --[ENUMOBJECTSPROC]----------------------------------------------
'\\ Callback function used to enumerate the existing pens or brushes
'\\ in a Device Context
'\\ Prototype:
'\\   int CALLBACK EnumObjectsProc(
'\\          LPVOID lpLogObject,  // object attributes
'\\          LPARAM lpData        // application-defined data
'\\ -----------------------------------------------------------------
'\\ (c) 2001 Merrion Computing ltd
Public Function ENUMOBJECTSPROC(ByVal lpObject As Long, ByVal lpData As Long) As Long

Select Case CurrentObjectType
Case OBJ_BRUSH
    '\\ We are enumerating the DCs brushes
    Dim brshThis As New ApiLogBrush
    brshThis.CreateFromPointer lpObject
    colGDIObjects.Add brshThis
    
Case OBJ_PEN
    '\\ We are enumerating the DCs pens
    Dim penThis As New ApiLogPen
    penThis.CreateFromPointer lpObject
    colGDIObjects.Add penThis
    
End Select

If Err.LastDllError Then
    ReportError Err.LastDllError, "ApiDeviceContext:EnumObjects", GetLastSystemError
Else
    ENUMOBJECTSPROC = 1
End If

End Function
