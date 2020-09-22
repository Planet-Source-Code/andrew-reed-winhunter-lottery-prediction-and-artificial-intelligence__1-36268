VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessor1"
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

Private mCalculated As bRunType

Private mUseProcessor As Boolean

Private mStart As Long
Private mStop As Long
Private mLastStop As Long
Private mLastStart As Long
Private mPredictDrawing As Long

Private mSingleHit() As Long            'stored hit counts
Private mRule0LowWeight As Long         'score weight
Private mRule0AvgWeight As Long         'score weight
Private mRule0HighWeight As Long        'score weight
Private mSingleHitScore() As Long       'Final Score after rule applied

Private mRule1Weight As Long
Private mMaxHitsTestValue As Long         'Adjustment Value
Private mMaxHitsScore() As Long

Private mKey As String         'processor key, set during initialization
Private mName As String         'processor name, set during initialization
Private mProperties As Object

Private mScore() As Long                'overall processor score
                                        'which is a total of the
                                        'rules within the processor
                                        'which are then added to the
                                        'filter score
Private mDrawings As Object        'local reference to Drawings Object

'set the local copy of the drawings to use here
Public Property Set Drawings(ByVal vData As Object)

    Set mDrawings = vData
    ReDim mScore(mDrawings.BallCount)
    ReDim mSingleHit(mDrawings.BallCount)
    ReDim mSingleHitScore(mDrawings.BallCount)
    ReDim mMaxHitsScore(mDrawings.BallCount)

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



'This routine simply totals the number of times
'a single ball has been drawn for the number of drawings
'specified
Public Sub Calc_Single_Scans()
Dim mDigit As Long
Dim mSum As Long
    
    'We must have a Drawings Object to continue!
    If mDrawings Is Nothing Then Exit Sub
    'Now we must have a start and stop!
    If mStart = 0 Then
        mStart = 1
    End If
    If mStop = 0 Then
        mStop = mDrawings.Count
    End If
    
    
    'Reset Values Before Calculating!
    mMinSum = 10000
    mMaxSum = 0
    mAvgSum = 0
    
    For i = mStart To mStop
        mSum = 0
        For n = 0 To mDrawings.Drawn - 1
            mDigit = mDrawings.Item(i).Numbers(n)
            mSum = mSum + mDigit
            mSingleHit(mDigit) = mSingleHit(mDigit) + 1
        Next n
        mAvgSum = mAvgSum + mSum
        If mSum < mMinSum Then mMinSum = mSum
        If mSum > mMaxSum Then mMaxSum = mSum
    Next i
    If mStart < mStop > 0 Then
        mAvgSum = mAvgSum / (mStop - mStart)
    End If
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
        mDigit = mDrawings.Item(mStart - 1).Numbers(n)
        mSingleHit(mDigit) = mSingleHit(mDigit) - 1
    Next n
    
    If mStop <> mLastStop Then
        'New Stop Drawing
        For n = 0 To mDrawings.Drawn - 1
            mDigit = mDrawings.Item(mStop).Numbers(n)
            mSingleHit(mDigit) = mSingleHit(mDigit) + 1
        Next n
    End If

End Sub

Private Sub Rule0()
Dim iBottom As Long
Dim iTop As Long
Dim dRange As Double

    If mRule0LowWeight = 0 And mRule0AvgWeight = 0 And mRule0HighWeight = 0 Then Exit Sub
    
    
    iBottom = 10000         'set an out of range value
    For i = mDrawings.MinimumBallNumber To UBound(mSingleHit)
      If mSingleHit(i) < iBottom Then iBottom = mSingleHit(i)
      If mSingleHit(i) > iTop Then iTop = mSingleHit(i)
    Next i
    
    dRange = (Fix(((iTop - iBottom) / 7) * 6) / 2) - 1
    
    For i = mDrawings.MinimumBallNumber To UBound(mSingleHit)
      If mSingleHit(i) < iBottom + dRange Then
          mSingleHitScore(i) = mSingleHitScore(i) + mRule0LowWeight
      ElseIf mSingleHit(i) > iTop - dRange Then
          mSingleHitScore(i) = mSingleHitScore(i) + mRule0AvgWeight
      Else
          mSingleHitScore(i) = mSingleHitScore(i) + mRule0HighWeight
      End If
    Next i

End Sub

Private Sub Rule1()

    If mRule1Weight = 0 Then Exit Sub
    
    For i = mDrawings.MinimumBallNumber To UBound(mSingleHit)
      If mSingleHit(i) < mMaxHitsTestValue Then
          mMaxHitsScore(i) = mMaxHitsScore(i) + mRule1Weight
      End If
    Next i

End Sub

Public Sub RunProcessor(vntFilterScore As Variant)

    If IsArray(vntFilterScore) Then
        ReDim mSingleHitScore(mDrawings.BallCount)
        Read_Properties
        If mCalculated = Calculate Then
            Calc_Single_Scans
            'next time around, be sure to SHIFT instead...
            mCalculated = Shift
        Else
            If mStart = mLastStart + 1 Then
                Shift_Single_Scans
            Else
                'we have not shifted correctly
                'so we must recalculate
                Calc_Single_Scans
            End If
        End If
        
        'run the rules here
        Rule0
        For i = mDrawings.MinimumBallNumber To UBound(mScore)
            'tally the scores from the rules here
            mScore(i) = mSingleHitScore(i) + mMaxHitsScore(i)
        Next i
        
        'Add the filter scores to the group scores here
        For i = mDrawings.MinimumBallNumber To UBound(vntFilterScore)
            vntFilterScore(i) = vntFilterScore(i) + mScore(i)
        Next i
    End If

End Sub

Public Property Get Calculated() As bRunType

    Calculated = mCalculated

End Property

Public Property Let Calculated(bRun As bRunType)

    mCalculated = bRun
    If Not bRun Then
        'The Processor is being reset...
        'so we must reset the calculated values here
        ReDim mSingleHit(mDrawings.BallCount)
    End If


End Property


Private Sub Class_Initialize()
Dim tempProperty As Object

    mKey = "singlehit"
    mName = "SingleHit"

    Set mProperties = CreateObject(sObjPropertyValues)
    Set tempProperty = mProperties.Add("lowweight")
    With tempProperty
        .Name = "Low Weight"
        .Value = 0
        .Max = 10
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("averageweight")
    With tempProperty
        .Name = "Average Weight"
        .Value = 0
        .Max = 10
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("highweight")
    With tempProperty
        .Name = "High Weight"
        .Value = 0
        .Max = 10
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("maxhitsweight")
    With tempProperty
        .Name = "Max Hits Weight"
        .Value = 0
        .Max = 10
        .Group = 2
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("maxhits")
    With tempProperty
        .Name = "Max Hits"
        .Value = 0
        .Max = 1000
        .Group = 2
        .Huntable = True
    End With
    mUseProcessor = True

End Sub

Private Sub Read_Properties()

    'Get Properties from the local property group
    mRule0LowWeight = mProperties("lowweight").Value
    mRule0AvgWeight = mProperties("averageweight").Value
    mRule0HighWeight = mProperties("highweight").Value
    mRule1Weight = mProperties("maxhitsweight").Value
    mMaxHitsTestValue = mProperties("maxhits").Value

End Sub

Private Sub Write_Properties()

    'Set Properties to the local property group
    mProperties("lowweight").Value = mRule0LowWeight
    mProperties("averageweight").Value = mRule0AvgWeight
    mProperties("highweight").Value = mRule0HighWeight
    mProperties("maxhitsweight").Value = mRule1Weight
    mProperties("maxhits").Value = mMaxHitsTestValue

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
