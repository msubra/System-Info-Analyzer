VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"cRegKey"
Attribute VB_Ext_KEY = "Member0" ,"cRegKey"
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

Private mCol As Collection

Public Function Add(Key As String, Optional sKey As String) As cRegKey
    Dim objNewMember As cRegKey
    Set objNewMember = New cRegKey
    
    objNewMember.Key = Key
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If

    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As cRegKey
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property


Public Property Get Count() As Long
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

