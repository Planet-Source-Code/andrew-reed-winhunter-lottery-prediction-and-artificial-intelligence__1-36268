VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrigger1"
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

Private mCalculated As bRunType

Private mSingleHit() As Long            'stored hit counts
Private mStart As Long
Private mStop As Long
Private mLastStop As Long
Private mRule0Test As Long
Private mRule0Method As Long

Private mDrawings As Object        'local reference to Drawings Object
Private mKey As String                  'trigger key, set during initialization
Private mName As String                 'trigger name, set during initialization
Private mProperties As Object
Private mUseTrigger As Boolean  'Use Trigger
Private mRun As Boolean         'What is the result from the trigger

'set the local copy of the drawings to use here
Public Property Set Drawings(ByVal vData As Object)

    Set mDrawings = vData
    ReDim mSingleHit(mDrawings.BallCount)

End Property

Public Property Get Run() As Boolean
    Run = mRun
End Property

Public Property Get Key() As String
    Key = mKey
End Property
Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(sName As String)
    mName = sName
End Property


Public Property Get UseTrigger() As Boolean
    UseTrigger = mUseTrigger
End Property
Public Property Let UseTrigger(bUseTrigger As Boolean)
    mUseTrigger = bUseTrigger
End Property

Public Property Let StartDrawing(lValue As Long)
    If lValue <> mStart + 1 Then
        'if we have not shifted one, then recalculate
        mCalculated = Calculate
    End If
    mStart = lValue
End Property
Public Property Let StopDrawing(lValue As Long)
    mStop = lValue
End Property

'This routine simply totals the number of times
'a single ball has been drawn for the number of drawings
'specified
Public Sub Calc_Single_Scans()
Dim mDigit As Long
    
    'We must have a Drawings Object to continue!
    If mDrawings Is Nothing Then Exit Sub
    'Now we must have a start and stop!
    If mStart = 0 Then
        mStart = 1
    End If
    If mStop = 0 Then
        mStop = mDrawings.Count
    End If
    
    
    For i = mStart To mStop
        mSum = 0
        For n = 0 To mDrawings.Drawn - 1
            'DoEvents
            mDigit = mDrawings.Item(i).Numbers(n)
            mSingleHit(mDigit) = mSingleHit(mDigit) + 1
        Next n
    Next i
    mLastStop = mStop

End Sub

'Private mLastStartDrawing As Integer
'Private mLastStopDrawing As Integer
Public Sub Shift_Single_Scans(Optional vntShift As Variant)
Dim mDigit As Long
    
    'We must have a Drawings Object to continue!
    If mDrawings Is Nothing Then Exit Sub
    
    'Now we must have a start and stop!
    If IsMissing(vntShift) Then
        vntShift = 1
    End If
    
    
    'OK, we have to subtract for the drawing that was dropped
    'and add for the new drawing
    'so we gotta account for the dropped on (mStartDrawing - 1)
    'and we gotta account for the new drawing (mStopDrawing)
    
    'Old Start Drawing
    For n = 0 To mDrawings.Drawn - 1
        'DoEvents
        mDigit = mDrawings.Item(mStart - 1).Numbers(n)
        mSingleHit(mDigit) = mSingleHit(mDigit) - 1
    Next n
    
    If mStop <> mLastStop Then
        'New Stop Drawing
        For n = 0 To mDrawings.Drawn - 1
            'DoEvents
            mDigit = mDrawings.Item(mStop).Numbers(n)
            mSingleHit(mDigit) = mSingleHit(mDigit) + 1
        Next n
    End If

End Sub

Private Sub Rule0()

    For i = mDrawings.MinimumBallNumber To UBound(mSingleHit)
        If mSingleHit(i) < mRule0Test Then
            mRun = True
        End If
    Next i

End Sub

Public Sub RunTrigger()

    Read_Properties
    If mCalculated = Calculate Then
        Calc_Single_Scans
        'next time around, be sure to SHIFT instead...
        mCalculated = Shift
    Else
        Shift_Single_Scans
    End If
    
    'run the rules here
    Rule0
        

End Sub


Private Sub Class_Initialize()
Dim tempProperty As Object
    
    mKey = "singletrigger"
    mName = "SingleTrigger"
    
    Set mProperties = CreateObject(sObjPropertyValues)
    Set tempProperty = mProperties.Add("singletest")
    With tempProperty
        .Name = "Hit Count"
        .Value = 10
        .Group = 1
        .Max = 50
        .Min = 0
    End With
    Set tempProperty = mProperties.Add("singlemethod")
    With tempProperty
        .Name = "Trigger Method"
        .Value = 1
        .Group = 100
    End With
    Set tempProperty = mProperties.Add("greaterthan")
    With tempProperty
        .Name = "Greater Than"
        .Value = 1
        .Group = 101
    End With
    Set tempProperty = mProperties.Add("lessthan")
    With tempProperty
        .Name = "Less Than"
        .Value = 2
        .Group = 102
    End With

End Sub

Public Property Get PropertyValues() As Object

    'pass back the collection so the parent object
    'can have access to the collection
    Set PropertyValues = mProperties

End Property

Private Sub Read_Properties()

    'Get Properties from the local property group
    mRule0Test = mProperties("singletest").Value
    mRule0Method = mProperties("singlemethod").Value
    
End Sub

Private Sub Write_Properties()

    'Set Properties to the local property group
    mProperties("singletest").Value = mRule0Test
    mProperties("singlemethod").Value = mRule0Method

End Sub

