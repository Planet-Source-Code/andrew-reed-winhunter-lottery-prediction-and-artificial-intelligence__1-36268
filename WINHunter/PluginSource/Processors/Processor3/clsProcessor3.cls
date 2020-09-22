VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessor3"
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




Private mSingleHit() As Long                'stored hits
Private mHowRecentHit() As Long             'stored hits
Private mLongScan() As Long                 'stored hits
Private mRepeatability() As Long            'stored hits
Private mRepeatScan() As Long               'stored hits

Private mOverdueScore() As Long             'Final Score after rule applied
Private mNotRecentScore() As Long           'Final Score after rule applied
Private mRepeatabilityScore() As Long       'Final Score after rule applied
Private mUnderRepeatsScore() As Long        'Final Score after rule applied

Private mNotRecentTestValue As Long         'Adjustment Value
Private mRepeatabilityTestValue As Long     'Adjustment Value
Private mUnderRepeatTest As Long            'Adjustment Value

Private mOverdueWeight As Long              'score weight
Private mNotRecentWeight As Long            'score weight
Private mRepeatabilityWeight As Long        'score weight
Private mUnderRepeatsWeight As Long         'score weight





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
    ReDim mHowRecentHit(mDrawings.BallCount)
    ReDim mLongScan(mDrawings.BallCount)
    ReDim mRepeatability(mDrawings.BallCount)
    ReDim mRepeatScan(mDrawings.BallCount)
    ReDim mOverdueScore(mDrawings.BallCount)
    ReDim mNotRecentScore(mDrawings.BallCount)
    ReDim mRepeatabilityScore(mDrawings.BallCount)
    ReDim mUnderRepeatsScore(mDrawings.BallCount)

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



'This routine finds the longest peroid betwen hits of each individual ball (LongScan)
'Also finds how recent the ball has been drawn (HowRecentHit)
'Also counts the total times a ball was drawn multiple weeks in a row (Repeatability)
Public Sub Calc_Individual_Scans()
Dim bRepeatability() As Boolean
Dim iRepeatability() As Long
Dim iDrawnBall As Long
Dim iHitCount() As Long
Dim iHitSet() As Boolean
    
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
    'rlsSetup.ResetLongScan
    'rlsSetup.ResetHowRecentHit
    'rlsSetup.ResetRepeatability
    'rlsSetup.ResetRepeatScan
    ReDim bRepeatability(mDrawings.BallCount)    'Local repeat flag
    ReDim iRepeatability(mDrawings.BallCount)    'local repeat count
    ReDim iHitSet(mDrawings.BallCount)
    ReDim iHitCount(mDrawings.BallCount)
    
    'loop through the drawings
    For j = mStart To mStop
        'loop through the balls
        'INCREASE RECENT COUNT HERE
        For i = mDrawings.MinimumBallNumber To mDrawings.BallCount
            'If Not mExclude(i) Then
                iHitCount(i) = iHitCount(i) + 1
                If iHitCount(i) > mLongScan(i) Then
                    mLongScan(i) = iHitCount(i)
                End If

                'mHowRecentHit(i) = mHowRecentHit(i) + 1
            'End If
        Next i
        'loop through the drawn balls
        For n = 0 To mDrawings.Drawn - 1
            'get the drawn ball
            iDrawnBall = mDrawings.Item(j).Numbers(n)
            
            'Count the single hits for reference later!
            mSingleHit(mDigit) = mSingleHit(mDigit) + 1
            
            'If iDrawnBall = 1 Then
            '    Debug.Print ;
            'End If
            'Do we want to use this ball?
            'If Not mExclude(iDrawnBall) Then
                If bRepeatability(iDrawnBall) Then
                    'if ball was not hit recently then reset the repeatability count
                    'recentcount = 1 means the ball was drawn the previous drawing
                    'recentcout > 1 means the ball was not drawn recently
                    If iHitCount(iDrawnBall) > 1 Then
                        iRepeatability(iDrawnBall) = 0
                    Else
                        If iRepeatability(iDrawnBall) = 0 Then iRepeatability(iDrawnBall) = 1
                    End If
                    
                    
                    'Count the repeats
                    iRepeatability(iDrawnBall) = iRepeatability(iDrawnBall) + 1
                    'if more than 1 repeat, then process it!
                    If iRepeatability(iDrawnBall) > 1 Then
                        'if the repeat count > Max repeats, then save the max repeats
                        If iRepeatability(iDrawnBall) - 1 > mRepeatScan(iDrawnBall) Then
                            'SAVE THE MAX REPEATS HERE
                            mRepeatScan(iDrawnBall) = iRepeatability(iDrawnBall) - 1
                        End If
                        'SAVE THE TOTAL REPEATS HERE
                        mRepeatability(iDrawnBall) = mRepeatability(iDrawnBall) + 1
                    End If
                Else
                    iRepeatability(iDrawnBall) = 0
                End If
                'test to see if the recenthit was the longest
                'then reset the recent count
                If iHitCount(iDrawnBall) > mLongScan(iDrawnBall) Then
                    'SAVE LONGEST HIT INTERVAL HERE
                    mLongScan(iDrawnBall) = iHitCount(iDrawnBall)
                End If
                If j = mStart Then
                    iHitCount(iDrawnBall) = 0
                End If
                'RESET RECENT COUNT HERE
                If Not iHitSet(iDrawnBall) Then
                    mHowRecentHit(iDrawnBall) = iHitCount(iDrawnBall)
                    iHitSet(iDrawnBall) = True
                End If
                iHitCount(iDrawnBall) = 0
            'End If
        Next n
        
        If j > 1 Then
            For n = 0 To mDrawings.Drawn - 1
                iDrawnBall = mDrawings.Item(j - 1).Numbers(n)
                'If Not mExclude(iDrawnBall) Then
                    'reset the old drawnball
                    bRepeatability(iDrawnBall) = False
                'End If
                iDrawnBall = mDrawings.Item(j).Numbers(n)
                'If Not mExclude(iDrawnBall) Then
                    'set the new drawnball
                    bRepeatability(iDrawnBall) = True
                'End If
            Next n
        End If
    Next j
    For i = mDrawings.MinimumBallNumber To mDrawings.BallCount
        'If Not mExclude(i) Then
            If mSingleHit(i) = 0 Then
                mHowRecentHit(i) = iHitCount(i)
            End If
        'End If
    Next i

End Sub

'This routine finds the longest peroid betwen hits of each individual ball (LongScan)
'Also finds how recent the ball has been drawn (HowRecentHit)
'Also counts the total times a ball was drawn multiple weeks in a row (Repeatability)
Public Sub Shift_Individual_Scans(Optional vntShift As Variant)
Dim iDrawnBall As Long
Dim iRepeats As Long
Dim iHits As Long
Dim iGap As Long
Dim bReScan As Boolean

    
    'We must have a Drawings Object to continue!
    If mDrawings Is Nothing Then Exit Sub
    'Now we must have a start and stop!
    If IsMissing(vntShift) Then
        vntShift = 1
    End If
    
    iGap = (mStop - mStart) + 1
    
    
    'OK, we have to account for some major things in here
    'first, we have the dropped drawing to account for
    'and the new drawing added
    'we have to decrement the balls by one for the indv scans
    'and adjust the max scans accordingly as well
    
    
    
    'Calculate the NEEDED single hit counts here
    For n = 0 To mDrawings.Drawn - 1
        iDrawnBall = mDrawings.Item(mStart - 1).Numbers(n)
        mSingleHit(iDrawnBall) = mSingleHit(iDrawnBall) - 1
    Next n
    If mStop <> mLastStop Then
        'New Stop Drawing
        For n = 0 To mDrawings.Drawn - 1
            iDrawnBall = mDrawings.Item(mStop).Numbers(n)
            mSingleHit(iDrawnBall) = mSingleHit(iDrawnBall) + 1
        Next n
    End If
    
    'decrement recent count here
    For i = mDrawings.MinimumBallNumber To mDrawings.BallCount
        'testing purposes only
        
        'If the number is not being excluded
        'If Not mExclude(i) Then
            iHits = 0
            iRepeats = 0
            'check to see if the recent hit is equal to the gap
            'if so, then the number has not been drawn yet
            'so let's check to see if it has been drawn with the newly added drawing
            If mHowRecentHit(i) = iGap And mSingleHit(i) < 2 Then
                If mSingleHit(i) = 1 Then
                    If mDrawings.Item(mStop).WasDrawn(i) Then
                        'the number was drawn on the newly added drawing
                        mHowRecentHit(i) = iGap
                        mLongScan(i) = iGap
                        mRepeatability(i) = 0
                        mRepeatScan(i) = 0
                    Else
                        mHowRecentHit(i) = mHowRecentHit(i) - 1
                        mLongScan(i) = mHowRecentHit(i)
                        mRepeatability(i) = 0
                        mRepeatScan(i) = 0
                    End If
                    If mDrawings.Item(mStart).WasDrawn(i) Then
                        mHowRecentHit(i) = 0
                    End If
                    '************OUTPUT CORRECT
                    'Debug.Print i & "N  ";
                    GoTo wholeskip
                ElseIf mSingleHit(i) = 0 Then
                    'still not drawn yet
                    'we really shouldnt have to reset these values here
                    'mHowRecentHit(i) = iGap
                    'mLongScan(i) = iGap
                    'mRepeatability(i) = 0
                    'mRepeatScan(i) = 0
                    '************OUTPUT CORRECT
                    'Debug.Print i & "X  ";
                    GoTo wholeskip
                End If
            ElseIf mHowRecentHit(i) = 0 Then
                'just rebuild the stats here
                'because the dropped drawing contained the number
                bReScan = True
            Else
                'If we have more than 1 hit for this number then we will to check further
                If mSingleHit(i) > 1 Then
                    'do we have the longest hit here?
                    'If the longscan is greater than (or equal to) half the gap
                    'we have the LONGEST hitgap
                    'so only rescan if it is less than half
                    If mHowRecentHit(i) = mLongScan(i) And mLongScan(i) < (Fix(iGap / 2) + 1) Then
                        bReScan = True
                        GoTo Skip2ReScan
                    End If
                Else
                    mHowRecentHit(i) = mHowRecentHit(i) - 1
                    mLongScan(i) = mHowRecentHit(i)
                    If mHowRecentHit(i) = 1 Then
                        mHowRecentHit(i) = 0
                        mLongScan(i) = iGap
                    ElseIf mHowRecentHit(i) < (Fix(iGap / 2) + 1) Then
                        mLongScan(i) = iGap - mHowRecentHit(i)
                    End If
                    '************OUTPUT CORRECT
                    'Debug.Print i & "W  ";
                    GoTo wholeskip
                End If
                
                'decrementing Howrecenthit
                mHowRecentHit(i) = mHowRecentHit(i) - 1
                'if the number hasn't shown up at the newly added one, just continue
                If mDrawings.Item(mStop).WasDrawn(i) Then
                    If (mHowRecentHit(i) + 1) = mLongScan(i) And mLongScan(i) > (Fix(iGap / 2) - 1) Then
                        'if the recent scan reveals a greater than half recent hit
                        'then adjust the longscan along with the recent hit
                        mLongScan(i) = mHowRecentHit(i)
                    End If
                    If mDrawings.Item(mStart - 1).WasDrawn(i) Then
                        'gotta make sure we account for the dropped drawing
                        iRepeats = 1
                        For j = (mStart) To mStop
                            If mDrawings.Item(j).WasDrawn(i) Then
                                iRepeats = iRepeats + 1
                            Else
                                Exit For
                            End If
                        Next j
                        If iRepeats > 1 Then
                            'only increment the repeatability by 1, since
                            'any other previous repeats have already been accounted for
                            mRepeatability(i) = mRepeatability(i) - 1
                            If (iRepeats - 1) = mRepeatScan(i) Then
                                'This might have been the max repeats here
                                'should we just rescan?
                                mRepeatScan(i) = iRepeats - 2
                                'we'll just assume this is the max repeats here
                                'and adjust accordingly
                            End If
                        End If
                    End If
                    iRepeats = 1
                    'we will need to check the maxrepeatability here
                    'now lets check to see if we have a new repeatscan
                    'walk backwards from the newest drawing that was added
                    For j = (mStop - 1) To mStart Step -1
                        If mDrawings.Item(j).WasDrawn(i) Then
                            iRepeats = iRepeats + 1
                        Else
                            Exit For
                        End If
                    Next j
                    If iRepeats > 1 Then
                        'only increment the repeatability by 1, since
                        'any other previous repeats have already been accounted for
                        mRepeatability(i) = mRepeatability(i) + 1
                        If (iRepeats - 1) > mRepeatScan(i) Then
                            'SAVE THE MAX REPEATS HERE
                            mRepeatScan(i) = iRepeats - 1
                        End If
                    Else
                        'no more numbers found
                        'If (rlsSetup.HowRecentHit(i) + 1) <> rlsSetup.LongScan(i) And rlsSetup.LongScan(i) > (Fix(iGap / 2) - 1) Then
                            'recent hit and longscan dont match
                            'so the last drawing here must be the new longscan
                            iHits = 1
                            For j = (mStop - 1) To mStart Step -1
                                If mDrawings.Item(j).WasDrawn(i) Then
                                    Exit For
                                Else
                                    iHits = iHits + 1
                                End If
                            Next j
                            If iHits = mLongScan(i) + 1 Then
                                mLongScan(i) = mLongScan(i) + 1
                            End If
                        'End If
                    End If
                    If mDrawings.Item(mStart).WasDrawn(i) Then
                        mHowRecentHit(i) = 0
                    End If
                    '************OUTPUT CORRECT
                    'Debug.Print i & "R  ";
                Else
                    
                    
                    'we will need to check the longscan here
                    If iGap - mHowRecentHit(i) > mLongScan(i) Then
                        'we could have a new longscan here
                        'so let's check and see
                        iHits = 0
                        'start after the most recent hit
                        For j = (mStart + mHowRecentHit(i)) To mStop
                            If Not mDrawings.Item(j).WasDrawn(i) Then
                                iHits = iHits + 1
                            Else
                                'dont exit yet if the number repeated itself
                                If Not iHits = 0 Then
                                    If Not (iGap - (j - mStart)) > mLongScan(i) Then
                                        Exit For
                                    End If
                                    iHits = 0
                                End If
                            End If
                        Next j
                        If iHits > mLongScan(i) Then mLongScan(i) = iHits
                    Else
                        mLongScan(i) = mHowRecentHit(i)
                    End If
                    If mDrawings.Item(mStart).WasDrawn(i) Then
                        mHowRecentHit(i) = 0
                    End If
                    '************OUTPUT CORRECT
                    'Debug.Print i & "H  ";
                End If
            End If
            
Skip2ReScan:

            If bReScan Then
                If mSingleHit(i) = 0 Then
                    mHowRecentHit(i) = iGap
                    mLongScan(i) = iGap
                    mRepeatability(i) = 0
                    mRepeatScan(i) = 0
                    'Debug.Print i & "S  ";
                    GoTo wholeskip
                End If
                mRepeatability(i) = 0
                mRepeatScan(i) = 0
                mLongScan(i) = 0
                mHowRecentHit(i) = 0
                iHits = 0
                iRepeats = 0
                'now we have to recount for this number
                For j = mStart To mStop
                    iHits = iHits + 1   'increment the hit count
                    'SAVE LONGEST HIT INTERVAL HERE
                    If iHits > mLongScan(i) Then
                        mLongScan(i) = iHits
                    End If
                    
                    'only save the recentcount if we finaly run into another hit
                    If mDrawings.Item(j).WasDrawn(i) Then
                        'save the hitcount here
                        If mHowRecentHit(i) = 0 Then
                            mHowRecentHit(i) = iHits
                        End If
                        iHits = 0
                        iRepeats = iRepeats + 1
                        'save the recent repeatability here
                        If iRepeats > 1 Then
                            If (iRepeats - 1) > mRepeatScan(i) Then
                                'SAVE THE MAX REPEATS HERE
                                mRepeatScan(i) = iRepeats - 1
                            End If
                            'SAVE THE TOTAL REPEATS HERE
                            mRepeatability(i) = mRepeatability(i) + 1
                        End If
                    Else
                        'count until we cant find the number
                        iRepeats = 0
                    End If
                Next j
                If mDrawings.Item(mStart).WasDrawn(i) Then
                    'keep the lead count to zero
                    mHowRecentHit(i) = 0
                End If
                If mLongScan(i) = iGap And Not mDrawings.Item(mStop).WasDrawn(i) Then
                    'the ball was never drawn!
                    mHowRecentHit(i) = mLongScan(i)
                End If
                
                '************OUTPUT CORRECT
                'Debug.Print i & "S  ";
            End If
        'End If
wholeskip:
        bReScan = False
        'Debug.Print rlsSetup.HowRecentHit(i) & "," & m2HowRecentHit(i);
        'Debug.Print " - " & rlsSetup.LongScan(i) & "," & m2LongScan(i);
        'Debug.Print " - " & rlsSetup.Repeatability(i) & "," & m2Repeatability(i);
        'Debug.Print " - " & rlsSetup.RepeatScan(i) & "," & m2RepeatScan(i)
    Next i
    'Debug.Print
    
End Sub


Private Sub Rule0()

    If mOverdueWeight = 0 Then Exit Sub
    For i = mDrawings.MinimumBallNumber To mDrawings.BallCount
        If mLongScan(i) > mHowRecentHit(i) + 1 Then
            mOverdueScore(i) = mOverdueScore(i) + mOverdueWeight
        End If
    Next i

End Sub

Private Sub Rule1()

    If mNotRecentWeight = 0 Then Exit Sub
    For i = mDrawings.MinimumBallNumber To mDrawings.BallCount
        'how long is too long for a hit NOT to come in?
        'Is it the average of the longscan intermediate hits?
        If mHowRecentHit(i) > mNotRecentTestValue Then
            mNotRecentScore(i) = mNotRecentScore(i) + mNotRecentWeight
        End If
    Next i

End Sub

Private Sub Rule2()
Dim lDrawnBall As Long

    If mRepeatabilityWeight = 0 Then Exit Sub
    'loop through the available drawings
    For k = mStart To mStop
        'loop through the drawn balls for each drawing
        For i = 0 To mDrawings.Drawn - 1
            'pull out each ball
            lDrawnBall = mDrawings.Item(k).Numbers(i)
            'repeatability is how many times total the number EVER repeated
            If Not mRepeatability(lDrawnBall) < mRepeatabilityTestValue Then
                mRepeatabilityScore(lDrawnBall) = mRepeatabilityScore(lDrawnBall) + mRepeatabilityWeight
            End If
        Next i
    Next k

End Sub

Private Sub Rule3()
Dim lRecentHit As Long

    
    '!!!!!!!!!!!!!!!!!!!!!!!!!!LOOK AT THIS ROUTINE AGAIN!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'This Routine only looks to see how many times the ball
    'has repeated in the last 3 drawings
    'if it is under the maximum repeats, then we give it a score
    
    If mUnderRepeatsWeight = 0 Then Exit Sub
    If mStart = mStop Then Exit Sub
    
    If mUnderRepeatTest = 0 Then
        For i = 0 To mDrawings.Drawn - 1
            j = mDrawings.Item(mStart).Numbers(i)
            For k = mStart To mStart + 2    'maybe we should add a slide adjustment here
                If mDrawings.Count < k Then Exit For
                For l = 0 To mDrawings.Drawn - 1
                    If mDrawings.Item(k).Numbers(l) = j Then
                        lRecentHit = lRecentHit + 1
                    End If
                Next l
    
                'very similar to repeatability scan
                'except this only saves the Max number of hits
                If lRecentHit < mRepeatScan(j) Then
                    mUnderRepeatsScore(j) = mUnderRepeatsScore(j) + mUnderRepeatsWeight
                End If
                If lRecentHit = mRepeatScan(j) Then
                    mUnderRepeatsScore(j) = 0
                End If
            Next k
        Next i
    Else
        For i = mDrawings.MinimumBallNumber To mDrawings.BallCount
            If mUnderRepeatTest > mRepeatScan(i) Then
                mUnderRepeatsScore(i) = mUnderRepeatsScore(i) + mUnderRepeatsWeight
            End If
        Next i
    End If


End Sub

Public Sub RunProcessor(vntFilterScore As Variant)

    If mUseProcessor Then
        If IsArray(vntFilterScore) Then
            Read_Properties
            ReDim mOverdueScore(mDrawings.BallCount)
            ReDim mNotRecentScore(mDrawings.BallCount)
            ReDim mRepeatabilityScore(mDrawings.BallCount)
            ReDim mUnderRepeatsScore(mDrawings.BallCount)
            If mCalculated = Calculate Then
                Calc_Individual_Scans
                'next time around, be sure to SHIFT instead...
                mCalculated = Shift
            Else
                If mStart = mLastStart + 1 Then
                    Shift_Individual_Scans
                Else
                    'we have not shifted correctly
                    'so we must recalculate
                    Calc_Individual_Scans
                End If
            End If
            'run the rules here
            Rule0
            Rule1
            Rule2
            Rule3
            For i = mDrawings.MinimumBallNumber To UBound(mScore)
                'tally the scores from the rules here
                mScore(i) = mOverdueScore(i) + mNotRecentScore(i) + mRepeatabilityScore(i) + mUnderRepeatsScore(i)
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
        ReDim mSingleHit(mDrawings.BallCount)
        ReDim mHowRecentHit(mDrawings.BallCount)
        ReDim mLongScan(mDrawings.BallCount)
        ReDim mRepeatability(mDrawings.BallCount)
        ReDim mRepeatScan(mDrawings.BallCount)
        ReDim mOverdueScore(mDrawings.BallCount)
        ReDim mNotRecentScore(mDrawings.BallCount)
        ReDim mRepeatabilityScore(mDrawings.BallCount)
        ReDim mUnderRepeatsScore(mDrawings.BallCount)
    End If


End Property


Private Sub Class_Initialize()
Dim tempProperty As Object

    mKey = "compoundhit"
    mName = "CompoundHit"

    
    Set mProperties = CreateObject(sObjPropertyValues)
    Set tempProperty = mProperties.Add("overdueweight")
    With tempProperty
        .Name = "Overdue Weight"
        .Value = 0
        .Max = 10
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("notrecentweight")
    With tempProperty
        .Name = "Not Recent Weight"
        .Value = 0
        .Max = 10
        .Group = 2
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("hitspan")
    With tempProperty
        .Name = "Hit Span"
        .Value = 3
        .Max = 1000
        .Group = 2
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("repeatabilityweight")
    With tempProperty
        .Name = "Repeatability Weight"
        .Value = 0
        .Max = 6
        .Group = 3
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("undertotalrepeats")
    With tempProperty
        .Name = "Under Total Repeats"
        .Value = 1
        .Max = 100
        .Group = 3
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("underrepeattest")
    With tempProperty
        .Name = "Max Repeats"
        .Value = 0
        .Max = 10
        .Group = 4
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("undermaxrepeatsweight")
    With tempProperty
        .Name = "Max Repeats Weight"
        .Value = 0
        .Max = 10
        .Group = 4
        .Huntable = True
    End With
    mUseProcessor = True

End Sub

Private Sub Read_Properties()

    'Get Properties from the local property group
    mOverdueWeight = mProperties("overdueweight").Value
    mNotRecentWeight = mProperties("notrecentweight").Value
    mNotRecentTestValue = mProperties("hitspan").Value
    mRepeatabilityWeight = mProperties("repeatabilityweight").Value
    mRepeatabilityTestValue = mProperties("undertotalrepeats").Value
    mUnderRepeatsWeight = mProperties("undermaxrepeatsweight").Value
    mUnderRepeatTest = mProperties("underrepeattest").Value

End Sub

Private Sub Write_Properties()

    'Set Properties to the local property group
    mProperties("overdueweight").Value = mOverdueWeight
    mProperties("notrecentweight").Value = mNotRecentWeight
    mProperties("hitspan").Value = mNotRecentTestValue
    mProperties("repeatabilityweight").Value = mRepeatabilityWeight
    mProperties("undertotalrepeats").Value = mRepeatabilityTestValue
    mProperties("undermaxrepeatsweight").Value = mUnderRepeatsWeight
    mProperties("underrepeattest").Value = mUnderRepeatTest

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


