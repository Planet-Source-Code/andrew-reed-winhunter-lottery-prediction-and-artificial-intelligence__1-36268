VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGroup"
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


Private mUseGroup As Boolean

Private mFilters As Collection          'the collection of filters to run
                                        'Group.Filter() how it will be exposed?

Private mSelections As Object    'The collection of Selections
                                        'Group.Selections() how it will be exposed

Private mSelection As Object            'the selection method to use
Private mDrawings As Object        'local reference to Drawings Object
Private mKey As String                  'Collection Key
Private mName As String                 'full group name
Private mScore() As Long                'Group Score, which gets passed
                                        'to the selector
Private mPredictDrawing As Long


Public Property Let PredictDrawing(lValue As Long)
    mPredictDrawing = lValue
End Property


Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(sName As String)
    mName = sName
End Property


Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(sKey As String)
    mKey = sKey
End Property


'set the local copy of the drawings to use here
Public Property Set Drawings(ByVal vData As Object)
Dim mLocalFilter As clsFilter

    Set mDrawings = vData
    For Each mLocalFilter In mFilters
        Set mLocalFilter.Drawings = mDrawings
    Next
    ReDim mScore(mDrawings.BallCount)
    mSelection.MinimumBallNumber = mDrawings.MinimumBallNumber

End Property

Public Sub Reset()
Dim mLocalFilter As clsFilter

    For Each mLocalFilter In mFilters
        mLocalFilter.Reset
    Next
    'reset the group object here
    Set mLocalFilter = Nothing

End Sub


Public Function AddFilter(sKey As Variant) As clsFilter

    'create a new object
    Dim objNewMember As clsFilter
    Set objNewMember = New clsFilter
    
    'Pass the drawings collection down the pipe for reference
    Set objNewMember.Drawings = mDrawings
    objNewMember.Key = sKey
    
    'set the properties passed into the method
    mFilters.Add objNewMember, sKey

    'return the object created
    Set AddFilter = objNewMember

    Set objNewMember = Nothing

End Function

Public Property Get FilterItem(vntIndexKey As Variant) As clsFilter
Attribute FilterItem.VB_UserMemId = 0

    Set FilterItem = mFilters(vntIndexKey)

End Property

Public Property Get FilterCount() As Integer

    FilterCount = mFilters.Count

End Property

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mFilters.Count
End Property

Public Sub RemoveFilter(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mFilters.Remove vntIndexKey
End Sub

Public Sub ClearFilters()
    
    'destroy the collection
    Set mFilters = Nothing
    'create a new instance of the collection
    Set mFilters = New Collection

End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mFilters.[_NewEnum]
End Property



Public Sub SetSelector(vntIndexKey As Variant)

    Set mSelection = mSelections.Item(vntIndexKey)

End Sub

Public Property Get Selector() As Object

    Set Selector = mSelection

End Property

Public Property Get SelectorCount() As Long
    
    SelectorCount = mSelections.Count

End Property


Public Property Get SelectorName() As String

    Selector = mSelection.Name
    
End Property

Public Sub RunGroup(mExcludedArray As Variant, Optional vntStackScore As Variant)
Dim oFilter As clsFilter

    If IsArray(mExcludedArray) Then
        ReDim mScore(mDrawings.BallCount)
        If mFilters.Count > 0 And Not mSelection Is Nothing Then
            For Each oFilter In mFilters
                oFilter.PredictDrawing = mPredictDrawing
                'Debug.Print mFilters.Count;
                oFilter.RunFilter mScore
                DoEvents
            Next
            
            'When done, we need to run the selection method.
            mSelection.ResetSelector UBound(mScore)
            mSelection.RunSelect mExcludedArray, mScore
            If Not IsMissing(vntStackScore) Then
                If IsArray(vntStackScore) Then
                    For i = 1 To UBound(vntStackScore)
                        'tally the scores here
                        vntStackScore(i) = vntStackScore(i) + mScore(i)
                        'get the selected balls
                        'mExcludedArray(i) = mSelection.Excluded(i)
                    Next i
                    mExcludedArray = mSelection.Excluded
                    'j = 0
                    'For i = 1 To 53
                    '    If Not mExcludedArray(i) Then
                    '        j = j + 1
                    '    Else
                    '        'Stop
                    '    End If
                    'Next i
                    'Debug.Print j
                End If
            End If
        End If
    End If

End Sub


Public Property Get UseGroup() As Boolean
    UseGroup = mUseGroup
End Property
Public Property Let UseGroup(bUseGroup As Boolean)
    mUseGroup = bUseGroup
End Property


Private Sub Class_Initialize()
    GetObjects
    'creates the collection when this class is created
    Set mFilters = New Collection
    Set mSelections = CreateObject(sObjSelections)
    Set mSelection = mSelections("highavglow")  'set the default selector
    mUseGroup = True
    
End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mFilters = Nothing
    Set mSelections = Nothing
    Set mSelection = Nothing
    Set mDrawings = Nothing
End Sub


