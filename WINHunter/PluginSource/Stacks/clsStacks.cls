VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStacks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'written by Andrew Reed 1996-2001
'WinHunter, a lottery prediction and statistical analysis toolkit
'Copyright (C)
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful, but
'WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA
'
'Contact via snailmail:
'
'Andrew Reed
'7870 Almarante Place
'Laurel Hill, FL    32567
'
'email via:
'winhunter@winhunter.freeservers.com
'
'Please see the LICENSE.TXT file for the complete license


Private mCol As Collection      'the collection of available filters
Private mDrawings As Object        'local reference to Drawings Object


'set the local copy of the drawings to use here
Public Property Set Drawings(ByVal vData As Object)
Dim mLocalStack As clsStack

    Set mDrawings = vData
    For Each mLocalStack In mCol
        Set mLocalStack.Drawings = mDrawings
    Next

End Property


Public Sub Clear()

    'destroy the collection
    Set mCol = Nothing
    'create a new instance of the collection
    Set mCol = New Collection

End Sub

Public Function Add(sKey As String, Optional objStack As clsStack) As clsStack

    
    'must have a drawings object to proceed
    If mDrawings Is Nothing Then Exit Function
    
    'set the properties passed into the method
    Dim objNewMember As clsStack
    Set objNewMember = New clsStack
    'Pass the drawings collection down the pipe for reference
    Set objNewMember.Drawings = mDrawings
    objNewMember.Key = sKey

    mCol.Add objNewMember, sKey
    'return the object created
    Set Add = objNewMember

    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsStack
Attribute Item.VB_UserMemId = 0

    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)

End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    GetObjects
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
    Set mDrawings = Nothing
End Sub





