VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStack"
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

Public Event Complete(ByVal sKey As String)           'Let the parent object know we are done
Public Event Plot(ByVal larrayScores As Variant)

Public Groups As clsGroups      'This is the collection of Groups
Attribute Groups.VB_VarUserMemId = 0
Public Output As clsOutput      'This is the Output Object
                                'it calculates the output for each prediction

Private mOutput() As Long    'This is the final output score
Private mExcluded As Variant 'This is the final set of predicted balls
Private mUseStack As Boolean    'process this stack?
Private mDrawings As Object        'local reference to Drawings Object
Private mKey As String                  'Collection Key
Private mName As String         'full stack name
Private mStop As Boolean        'stop stack testing

Private mPredictDrawing As Long


Public Property Let PredictDrawing(lValue As Long)
    mPredictDrawing = lValue
End Property
Public Property Get PredictDrawing() As Long
    PredictDrawing = mPredictDrawing
End Property


Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(sKey As String)
    mKey = sKey
End Property


Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(sName As String)
    mName = sName
End Property


'set the local copy of the drawings to use here
Public Property Set Drawings(ByVal vData As Object)

    Set mDrawings = vData
    'Pass the drawings collection down the pipe for reference
    Set Groups.Drawings = mDrawings
    Set Output.Drawings = mDrawings
    'adjust the arrays accordingly
    ReDim mPredicted(mDrawings.BallCount)
    ReDim mOutput(mDrawings.BallCount)

End Property

Public Property Get Drawings() As Object

    Set Drawings = mDrawings

End Property

Public Sub Reset()
Dim mLocalGroup As clsGroup

    'reset the filter object here
    ReDim mExcluded(mDrawings.BallCount)
    ReDim mOutput(mDrawings.BallCount)
    For Each mLocalGroup In Groups
        mLocalGroup.Reset
    Next
    Output.Reset
    Set mLocalGroup = Nothing

End Sub


Public Property Get OutputScore(Optional ByVal iIndex As Variant) As Variant
    
    'IF we have an index#, then pass the value back
    'no index, then pass the array object back!
    If IsMissing(iIndex) Then
        Output = mOutput
    Else
        Output = mOutput(iIndex)
    End If

End Property
Public Property Let OutputScore(ByVal iIndex As Variant, ByVal iOutput As Variant)
  
    If IsNumeric(iOutput) Then
        mOutput(iIndex) = iOutput
    End If

End Property



Public Property Get Excluded(Optional ByVal iIndex As Variant) As Variant
    
    'IF we have an index#, then pass the value back
    'no index, then pass the array object back!
    If IsMissing(iIndex) Then
        Predicted = mExcluded
    Else
        Predicted = mExcluded(iIndex)
    End If

End Property
Public Property Let Excluded(ByVal iIndex As Variant, ByVal bPredicted As Variant)
  
    If IsNumeric(bPredicted) Then
        If bPredicted = 0 Or bPredicted = -1 Then
            mExcluded(iIndex) = bPredicted
        End If
    End If

End Property


Public Property Get UseStack() As Boolean
    UseStack = mUseStack
End Property
Public Property Let UseStack(bUseStack As Boolean)
    mUseStack = bUseStack
End Property


Private Sub Class_Initialize()
    GetObjects
    Set Groups = New clsGroups
    Set Output = New clsOutput
    mUseStack = True
End Sub

Private Sub Class_Terminate()
    Set Groups = Nothing
    Set Output = Nothing
    Set mDrawings = Nothing
End Sub

Public Sub RunStack()
Dim oGroup As clsGroup
Dim addstring As String
Dim bTemp() As Boolean
Dim iGroup As Integer

    Output.Reset
    Set mExcluded = Nothing
    ReDim bTemp(mDrawings.BallCount)
    mExcluded = bTemp
    If Groups.Count > 0 Then
        'ReDim mPredicted(mDrawings.BallCount)
        'ReDim mOutput(mDrawings.BallCount)
        For Each oGroup In Groups
            If oGroup.UseGroup Then
                oGroup.PredictDrawing = mPredictDrawing
                oGroup.RunGroup mExcluded, mOutput
                'get the results back from the Group
                'So we can pass the comined results
                'back through the Groups
                mExcluded = oGroup.Selector.Excluded
            End If
            DoEvents
        Next
        iGroup = Groups.Count
        Do Until Groups.Item(iGroup).UseGroup
            iGroup = iGroup - 1
        Loop
        RaiseEvent Plot(Groups.Item(iGroup).Selector.Score)
        'deal with the reasults
        'If mPredictDrawing > 0 Then
            Output.Count_Matches mExcluded, mPredictDrawing
        'End If
        'let the upstream parents know we are done
        RaiseEvent Complete(mKey)
    End If

End Sub

Public Sub TestStack()
Dim oGroup As clsGroup
Dim addstring As String
Dim bTemp() As Boolean

    mStop = False
    Output.Reset
    If Groups.Count > 0 Then
        'ReDim mPredicted(mDrawings.BallCount)
        'ReDim mOutput(mDrawings.BallCount)
        For mPredictDrawing = 1 To mDrawings.Count
            'stop testing NOW
            If mStop Then Exit For
            
            'reset the score before we test again
            ReDim mOutput(mDrawings.BallCount)
            Set mExcluded = Nothing
            ReDim bTemp(mDrawings.BallCount)
            mExcluded = bTemp
            For Each oGroup In Groups
                'stop testing NOW
                If mStop Then Exit For
                'oGroup.Reset
                If oGroup.UseGroup Then
                    oGroup.PredictDrawing = mPredictDrawing
                    oGroup.RunGroup mExcluded, mOutput
                    'get the results back from the Group
                    'So we can pass the combined results
                    'back through the Groups
                    mExcluded = oGroup.Selector.Excluded
                End If
                'j = 0
                'For i = 1 To 53
                '    If Not mExcluded(i) Then
                '        j = j + 1
                '    Else
                '        'Stop
                '    End If
                'Next i
                'Debug.Print j
            Next
            'RaiseEvent Plot(Groups.Item(Groups.Count).Selector.Score)
            'deal with the results
            Output.Count_Matches mExcluded, mPredictDrawing
            'let the upstream parents know we are done
            RaiseEvent Complete(mKey)
            DoEvents
        Next
    End If

End Sub

Private Sub StopTesting()

    mStop = True

End Sub

