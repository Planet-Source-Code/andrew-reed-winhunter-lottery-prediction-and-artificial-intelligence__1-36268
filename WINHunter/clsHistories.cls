VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Event ThisStackComplete(ByVal sName As String, ByVal sMachineKey As String, ByVal sStackKey As String)

Private mCol As Collection      'the collection of available filters
Private WithEvents mHistory As clsHistory
Attribute mHistory.VB_VarHelpID = -1


Public Property Set RunningHistory(cHistory As clsHistory)
    Set mHistory = cHistory
End Property

Public Sub Clear()

    'destroy the collection
    Set mCol = Nothing
    'create a new instance of the collection
    Set mCol = New Collection

End Sub

Public Function Add(sKey As String, Optional objHistory As clsHistory) As clsHistory

    'set the properties passed into the method
    Dim objNewMember As clsHistory
    Set objNewMember = New clsHistory

    mCol.Add objNewMember, sKey
    'return the object created
    Set Add = objNewMember

    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As clsHistory
Attribute Item.VB_UserMemId = 0
Attribute Item.VB_MemberFlags = "200"

On Error GoTo HistoriesItemError
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  
  
    'apparently, this code will break the event loop that was experienced
    'on 12-3-01...
    If Not mHistory Is Nothing Then
        If mHistory.Name <> mCol(vntIndexKey).Name Then
            Set mHistory = mCol(vntIndexKey)
        End If
    Else
        Set mHistory = mCol(vntIndexKey)
    End If
  
  
  
  Set Item = mHistory
  Exit Property

HistoriesItemError:

    If Err.Number = 5 Then
        Set Item = Nothing
        Err.Clear
        Exit Property
    End If

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

    If Not mHistory Is Nothing Then
        If mHistory.Name = vntIndexKey Then
            'destroy the only other connection to the class created
            Set mHistory = Nothing
        End If
    End If
        mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
    
    'Add the Processor class items here!!!!
End Sub

Private Sub Class_Terminate()
    Set mHistory = Nothing
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Private Sub mHistory_MachineStackComplete(ByVal sName As String, ByVal sMachineKey As String, ByVal sStackKey As String)

    RaiseEvent ThisStackComplete(sName, sMachineKey, sStackKey)

End Sub
