VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessor4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'not supported
Private mCalculated As bRunType

Private mUseProcessor As Boolean



Private mStart As Long
Private mStop As Long
Private mLastStop As Long
Private mLastStart As Long
Private mPredictDrawing As Long

Private mPrevDwgsWeight As Long            'score weight
Private mPrevDwgsTestValue As Long         'Adjustment Value

Private mKey As String                  'processor key, set during initialization
Private mName As String                 'processor name, set during initialization
Private mProperties As Object

Private mScore() As Long                'overall processor score
Private mDrawings As Object        'local reference to Drawings Object

'set the local copy of the drawings to use here
Public Property Set Drawings(ByVal vData As Object)

    Set mDrawings = vData
    ReDim mScore(mDrawings.BallCount)

End Property


Public Property Get Calculated() As bRunType
'not supported, but still needs to exist
End Property
Public Property Let Calculated(bRun As bRunType)
'not supported, but still needs to exist
End Property

Public Property Let StartDrawing(lValue As Long)
    If lValue <> mStart + 1 Then
        'if we have not shifted one, then recalculate
        mCalculated = Calculate
    End If
    mLastStart = mStart
    mStart = lValue
End Property
Public Property Let StopDrawing(lValue As Long)
    mStop = lValue
End Property
Public Property Let PredictDrawing(lValue As Long)
    If lValue <> mPredictDrawing + 1 Then
        'if we have not shifted one, then recalculate
        mCalculated = Calculate
    End If
    mPredictDrawing = lValue
End Property

Public Property Get PropertyValues() As Object

    'set the current local values to the collection?
    'Write_Properties
    
    
    
    'pass back the collection so the parent object
    'can have access to the collection
    Set PropertyValues = mProperties

End Property


Public Property Get Key() As String
    Key = mKey
End Property
Public Property Get Name() As String
    Name = mName
End Property



Private Sub Rule0()
Dim lDigit As Long
    
    If mPrevDwgsWeight = 0 Then Exit Sub
    For i = mStart To mStart + mPrevDwgsTestValue
        If mDrawings.Count < i Then Exit For
        For j = 0 To mDrawings.Drawn - 1
            lDigit = mDrawings.Item(i).Numbers(j)
            mScore(lDigit) = mScore(lDigit) - mPrevDwgsWeight
        Next j
    Next i
  
End Sub


Public Sub RunProcessor(vntFilterScore As Variant)

    If IsArray(vntFilterScore) Then
        Read_Properties
        ReDim mScore(mDrawings.BallCount)
        
        'run the rule here!
        'No preprocessing Calcs/Shifts are needed for this processor
        Rule0
        'Add the filter scores to the group scores here
        For i = mDrawings.MinimumBallNumber To UBound(vntFilterScore)
            vntFilterScore(i) = vntFilterScore(i) + mScore(i)
        Next i
    End If

End Sub

Private Sub Class_Initialize()
Dim tempProperty As Object

    mKey = "previousdwgs"
    mName = "PreviousDwgs"

    Set mProperties = CreateObject(sObjPropertyValues)
    Set tempProperty = mProperties.Add("prevdwgtestvalue")
    With tempProperty
        .Name = "Previous Drawings"
        .Value = 0
        .Max = 10000
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("prevdwgweight")
    With tempProperty
        .Name = "Previous Weight (-)"
        .Value = 0
        .Max = 10
        .Group = 1
        .Huntable = True
    End With
    mUseProcessor = True

End Sub

Private Sub Read_Properties()

    'Get Properties from the local property group
    mPrevDwgsWeight = mProperties("prevdwgweight").Value
    mPrevDwgsTestValue = mProperties("prevdwgtestvalue").Value

End Sub

Private Sub Write_Properties()

    'Set Properties to the local property group
    mProperties("prevdwgweight").Value = mPrevDwgsWeight
    mProperties("prevdwgtestvalue").Value = mPrevDwgsTestValue

End Sub


Public Property Get UseProcessor() As Boolean
    UseProcessor = mUseProcessor
End Property
Public Property Let UseProcessor(bUseProcessor As Boolean)
    mUseProcessor = bUseProcessor
End Property


Private Sub Class_Terminate()

    Set mProperties = Nothing

End Sub

