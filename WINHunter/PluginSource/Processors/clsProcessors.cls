VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessors"
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


'This collection is populated upon initialization
'As any new filtering methods are written, they are added here
'so as to be available to the user for use as a filtering method
'Programmatically, no new filters may be added at runtime
'so there is no need to expose any methods other than Count & Item

Public Enum bRunType   'boolean only!
    Calculate = False
    Shift = True
End Enum

Private mCol As Collection      'the collection of available filters
Private mDrawings As Object        'local reference to Drawings Object

'set the local copy of the drawings to use here
Public Property Set Drawings(ByVal vData As Object)

    Set mDrawings = vData

End Property


Private Sub Clear()

    'destroy the collection
    Set mCol = Nothing
    'create a new instance of the collection
    Set mCol = New Collection

End Sub

Private Function Add(objProcessor As Object) As Object

    'set the properties passed into the method
    mCol.Add objProcessor, objProcessor.Key

    'return the object created
    Set Add = objProcessor

End Function

Public Property Get Item(vntIndexKey As Variant) As Object
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

Private Sub Remove(vntIndexKey As Variant)
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
    'creates the collection when this class is created
    Set mCol = New Collection
    
    'Add the Processor class items here!!!!
    'Add them according to the following four line syntax
    'Dim mProcessor As *ClassName*
    'Set mProcessor = New *ClassName*
    'mCol.Add mProcessor
    'Set mProcessor = Nothing
    
    GetObjects

    Dim mProcessor1 As clsProcessor1
    Set mProcessor1 = New clsProcessor1
    Add mProcessor1
    Set mProcessor1 = Nothing
    
    Dim mProcessor2 As clsProcessor2
    Set mProcessor2 = New clsProcessor2
    Add mProcessor2
    Set mProcessor2 = Nothing
    
    Dim mProcessor3 As clsProcessor3
    Set mProcessor3 = New clsProcessor3
    Add mProcessor3
    Set mProcessor3 = Nothing
    
    Dim mProcessor4 As clsProcessor4
    Set mProcessor4 = New clsProcessor4
    Add mProcessor4
    Set mProcessor4 = Nothing
    
    Dim mProcessor5 As clsProcessor5
    Set mProcessor5 = New clsProcessor5
    Add mProcessor5
    Set mProcessor5 = Nothing
    
    Dim mProcessor6 As clsProcessor6
    Set mProcessor6 = New clsProcessor6
    Add mProcessor6
    Set mProcessor6 = Nothing
    
    Dim mProcessor7 As clsProcessor7
    Set mProcessor7 = New clsProcessor7
    Add mProcessor7
    Set mProcessor7 = Nothing
    
    Dim mProcessor8 As clsProcessor8
    Set mProcessor8 = New clsProcessor8
    Add mProcessor8
    Set mProcessor8 = Nothing
    
    Dim mProcessor9 As clsProcessor9
    Set mProcessor9 = New clsProcessor9
    Add mProcessor9
    Set mProcessor9 = Nothing

End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub



