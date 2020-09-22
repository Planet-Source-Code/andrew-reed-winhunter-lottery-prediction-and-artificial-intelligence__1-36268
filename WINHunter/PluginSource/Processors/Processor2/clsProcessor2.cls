VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessor2"
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

Private mCalculated As bRunType

Private mUseProcessor As Boolean



Private mStart As Long
Private mStop As Long
Private mLastStop As Long
Private mLastStart As Long
Private mPredictDrawing As Long

Private mPairHit() As Long              'stored hit counts
Private mNeighborHit() As Long            'stored neighbor hits
Private mRule0LowWeight As Long         'score weight
Private mRule0AvgWeight As Long         'score weight
Private mRule0HighWeight As Long        'score weight
Private mPairHitScore() As Long           'Final Score after rule applied
Private mNeighborHitScore() As Long   'Final Score after rule applied

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
    ReDim mPairHit(mDrawings.BallCount - 1, mDrawings.BallCount)
    ReDim mNeighborHit(mDrawings.BallCount)
    ReDim mPairHitScore(mDrawings.BallCount)
    ReDim mNeighborHitScore(mDrawings.BallCount)

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



'This routine totals the number of times a pair of balls
'has been drawn for the number of drawings specified
Public Sub Calc_Pair_Scans()
Dim mDigit1 As Long
Dim mDigit2 As Long

    'We must have a Drawings Object to continue!
    If mDrawings Is Nothing Then Exit Sub
    'Now we must have a start and stop!
    If mStart = 0 Then
        mStart = 1
    End If
    If mStop = 0 Then
        mStop = mDrawings.Count
    End If
    
    'Note: the + here is to remove the drawing we are trying to predict
    For i = mStart To mStop
        'loop through the first 5 balls
        For j = 0 To mDrawings.Drawn - 2
            'loop through the last 5 balls
            'being sure not to use the balls from the first loop
            For k = j + 1 To mDrawings.Drawn - 1
                mDigit1 = mDrawings.Item(i).Numbers(j)
                'Do we want to use this ball?
                'If Not mExclude(mDigit1) Then
                    mDigit2 = mDrawings.Item(i).Numbers(k)
                    'Do we want to use this ball?
                    'If Not mExclude(mDigit2) Then
                        'Make sure we sequence the pairs (lower,higher)
                        If mDigit1 < mDigit2 Then
                            mPairHit(mDigit1, mDigit2) = mPairHit(mDigit1, mDigit2) + 1
                            
                            'This counts the number of times a pair
                            '(1/2, 3/4, n/n+1) has hit
                            If mDigit2 = mDigit1 + 1 Then
                                mNeighborHit(mDigit1) = mNeighborHit(mDigit1) + 1
                                mNeighborHit(mDigit2) = mNeighborHit(mDigit2) + 1
                            End If
                        Else
                            mPairHit(mDigit2, mDigit1) = mPairHit(mDigit2, mDigit1) + 1
                            
                            'This counts the number of times a pair
                            '(1/2, 3/4, n/n+1) has hit
                            If mDigit1 = mDigit2 + 1 Then
                                mNeighborHit(mDigit1) = mNeighborHit(mDigit1) + 1
                                mNeighborHit(mDigit2) = mNeighborHit(mDigit2) + 1
                            End If
                        End If
                    'End If
                'End If
            Next k
        Next j
    Next i

End Sub


'This routine totals the number of times a pair of balls
'has been drawn for the number of drawings specified
Public Sub Shift_Pair_Scans(Optional vntShift As Variant)
Dim mDigit1 As Long
Dim mDigit2 As Long
    
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
    For j = 0 To mDrawings.Drawn - 2
        'loop through the last 5 balls
        'being sure not to use the balls from the first loop
        For k = j + 1 To mDrawings.Drawn - 1
            mDigit1 = mDrawings.Item(mStart - 1).Numbers(j)
            'Do we want to use this ball?
            'If Not mExclude(mDigit1) Then
                mDigit2 = mDrawings.Item(mStart - 1).Numbers(k)
                If mDigit1 < mDigit2 Then
                    mPairHit(mDigit1, mDigit2) = mPairHit(mDigit1, mDigit2) - 1
                End If
                If mDigit2 = mDigit1 + 1 Then
                    mNeighborHit(mDigit1) = mNeighborHit(mDigit1) - 1
                    mNeighborHit(mDigit2) = mNeighborHit(mDigit2) - 1
                End If
            'End If
        Next k
    Next j
    
    If mStop <> mLastStop Then
        'New Stop Drawing
        For j = 0 To mDrawings.Drawn - 2
            'loop through the last 5 balls
            'being sure not to use the balls from the first loop
            For k = j + 1 To mDrawings.Drawn - 1
                 mDigit1 = mDrawings.Item(mStop).Numbers(j)
                'Do we want to use this ball?
                'If Not mExclude(mDigit1) Then
                    mDigit2 = mDrawings.Item(mStop).Numbers(k)
                    If mDigit1 < mDigit2 Then
                        mPairHit(mDigit1, mDigit2) = mPairHit(mDigit1, mDigit2) - 1
                    End If
                    If mDigit2 = mDigit1 + 1 Then
                        mNeighborHit(mDigit1) = mNeighborHit(mDigit1) - 1
                        mNeighborHit(mDigit2) = mNeighborHit(mDigit2) - 1
                    End If
                'End If
            Next k
        Next j
    End If

    'shift complete


End Sub



Private Sub Rule0()
Dim iBottom As Long
Dim iTop As Long
Dim dRange As Double
Dim ScaleMe() As Long

    If mRule0LowWeight = 0 And mRule0AvgWeight = 0 And mRule0HighWeight = 0 Then Exit Sub
    
    
    iBottom = 10000
    ReDim ScaleMe(mDrawings.BallCount)
    For i = 1 To mDrawings.BallCount - 1
        For j = i + 1 To mDrawings.BallCount
            If mPairHit(i, j) > iTop Then
                iTop = mPairHit(i, j)
            End If
            If mPairHit(i, j) < iBottom Then
                iBottom = mPairHit(i, j)
            End If
        Next j
    Next i
    dRange = (Fix(((iTop - iBottom) / 7) * 6) / 2)
    For i = mDrawings.MinimumBallNumber To mDrawings.BallCount - 1
      For j = i + mDrawings.MinimumBallNumber To mDrawings.BallCount
        If Not i = j Then
          If mPairHit(i, j) < iBottom + dRange Then
            ScaleMe(i) = ScaleMe(i) + 1
            ScaleMe(j) = ScaleMe(j) + 1
          ElseIf mPairHit(i, j) > iTop - dRange Then
            ScaleMe(i) = ScaleMe(i) + 3
            ScaleMe(j) = ScaleMe(j) + 3
          Else
            ScaleMe(i) = ScaleMe(i) + 2
            ScaleMe(j) = ScaleMe(j) + 2
          End If
        End If
      Next j
    Next i
    
    'now scale the eliminate2() array down to 1/2/3 scale
    iBottom = 10000
    For i = mDrawings.MinimumBallNumber To mDrawings.BallCount
        If ScaleMe(i) < iBottom Then
            iBottom = ScaleMe(i)
        End If
        If ScaleMe(i) > iTop Then
            iTop = ScaleMe(i)
        End If
    Next i
    dRange = (Fix(((iTop - iBottom) / 7) * 6) / 2)
    For i = mDrawings.MinimumBallNumber To mDrawings.BallCount
      If ScaleMe(i) < iBottom + dRange Then
          mPairHitScore(i) = mPairHitScore(i) + mRule0LowWeight
      ElseIf ScaleMe(i) > iTop - dRange Then
          mPairHitScore(i) = mPairHitScore(i) + mRule0AvgWeight
      Else
          mPairHitScore(i) = mPairHitScore(i) + mRule0HighWeight
      End If
    Next i

End Sub

Public Sub RunProcessor(vntFilterScore As Variant)

    If mUseProcessor Then
        If IsArray(vntFilterScore) Then
            Read_Properties
            ReDim mPairHitScore(mDrawings.BallCount)
            ReDim mNeighborHitScore(mDrawings.BallCount)
            If mCalculated = Calculate Then
                Calc_Pair_Scans
                'next time around, be sure to SHIFT instead...
                mCalculated = Shift
            Else
                If mStart = mLastStart + 1 Then
                    Shift_Pair_Scans
                Else
                    'we have not shifted correctly
                    'so we must recalculate
                    Calc_Pair_Scans
                End If
            End If
            
            'run the rules here
            Rule0
            'Rule1
            For i = mDrawings.MinimumBallNumber To UBound(mScore)
                'tally the scores from the rules here
                mScore(i) = mPairHitScore(i) + mNeighborHitScore(i)
            Next i
            
            'Add the filter scores to the group scores here
            For i = mDrawings.MinimumBallNumber To UBound(vntFilterScore)
                vntFilterScore(i) = vntFilterScore(i) + mScore(i)
            Next i
        End If
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
        ReDim mScore(mDrawings.BallCount)
        ReDim mPairHit(mDrawings.BallCount - 1, mDrawings.BallCount)
        ReDim mNeighborHit(mDrawings.BallCount)
        ReDim mPairHitScore(mDrawings.BallCount)
        ReDim mNeighborHitScore(mDrawings.BallCount)
    End If


End Property


Private Sub Class_Initialize()
Dim tempProperty As Object

    mKey = "pairhit"
    mName = "PairHit"

    Set mProperties = CreateObject(sObjPropertyValues)
    Set tempProperty = mProperties.Add("pairlowweight")
    With tempProperty
        .Name = "Pair Low Weight"
        .Value = 0
        .Max = 10
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("pairaverageweight")
    With tempProperty
        .Name = "Pair Average Weight"
        .Value = 0
        .Max = 10
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("pairhighweight")
    With tempProperty
        .Name = "Pair High Weight"
        .Value = 0
        .Max = 10
        .Group = 1
        .Huntable = True
    End With
    'Uncomment below when a sufficient rule is devised for this statistic
    'Set tempProperty = mProperties.Add("neighborhit")
    'With tempProperty
    '    .Name = "Neighbor Hit Weight"
    '    .Value = 0
    '    .max = 10
    '    .Group = 2
    'End With
    mUseProcessor = True

End Sub

Private Sub Read_Properties()

    'Get Properties from the local property group
    mRule0LowWeight = mProperties("pairlowweight").Value
    mRule0AvgWeight = mProperties("pairaverageweight").Value
    mRule0HighWeight = mProperties("pairhighweight").Value

End Sub

Private Sub Write_Properties()

    'Set Properties to the local property group
    mProperties("pairlowweight").Value = mRule0LowWeight
    mProperties("pairaverageweight").Value = mRule0AvgWeight
    mProperties("pairhighweight").Value = mRule0HighWeight

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

