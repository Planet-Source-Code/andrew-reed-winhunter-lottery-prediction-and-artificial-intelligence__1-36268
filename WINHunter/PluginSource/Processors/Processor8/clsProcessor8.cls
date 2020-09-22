VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessor8"
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

Private mUseProcessor As Boolean

Private mStart As Long
Private mStop As Long
Private mLastStop As Long
Private mLastStart As Long
Private mPredictDrawing As Long

Private mMinPositionValue() As Long
Private mMaxPositionValue() As Long
Private mMeanPositionValue() As Long
Private mGrossPositionValue() As Long
Private mBasePositionGap() As Long            'stored hit counts
Private mPositionGap() As Long            'stored hit counts
Private mRule0GapWeight As Long         'score weight
Private mRule0MaximumGap As Long         'score weight
Private mPositionGapScore() As Long         'Final Score after rule applied

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
    ReDim mPositionGap(mDrawings.Drawn)
    ReDim mBasePositionGap(mDrawings.Drawn)
    ReDim mMeanPositionValue(mDrawings.Drawn)
    ReDim mMinPositionValue(mDrawings.Drawn)
    ReDim mMaxPositionValue(mDrawings.Drawn)
    ReDim mGrossPositionValue(mDrawings.Drawn)
    ReDim mScore(mDrawings.BallCount)
    ReDim mPositionGapScore(mDrawings.BallCount)

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



'This routine calculates the distance (gap) from the last drawing
'to the next drawing, and so on, adding the totals together
'and dividing that by the total number of gaps collected
'
'It basically yields the span of numbers the position has been
'hitting for any particular draw period
Public Sub Calc_Position_Gap()
Dim mThisDigit() As Long
Dim mLastDigit() As Long
Dim mCount As Long
    
    'We must have a Drawings Object to continue!
    If mDrawings Is Nothing Then Exit Sub
    'Now we must have a start and stop!
    If mStart = 0 Then
        mStart = 1
    End If
    If mStop = 0 Then
        mStop = mDrawings.Count
    End If
    
    
    ReDim mThisDigit(mDrawings.Drawn)
    ReDim mLastDigit(mDrawings.Drawn)
    
    mCount = mStop - mStart + 1
    For i = mStart To mStop
        For n = 0 To mDrawings.Drawn - 1
            mThisDigit(n) = mDrawings.Item(i).Numbers(n)
            If mLastDigit(n) > 0 Then
                mGrossPositionValue(n) = mGrossPositionValue(n) + mThisDigit(n)
                mBasePositionGap(n) = mBasePositionGap(n) + Abs(mThisDigit(n) - mLastDigit(n))
            Else
                mLastDigit(n) = mThisDigit(n)
                mGrossPositionValue(n) = mGrossPositionValue(n) + mThisDigit(n)
            End If
        Next n
    Next i
    If mCount > 1 Then
        For n = 0 To mDrawings.Drawn - 1
            mPositionGap(n) = mBasePositionGap(n) / (mCount - 1)
            mMeanPositionValue(n) = mGrossPositionValue(n) / mCount
        Next n
    End If
    
    mLastStop = mStop

End Sub


Public Sub Shift_Position_Gap(Optional vntShift As Variant)
Dim mThisDigit() As Long
Dim mLastDigit() As Long
Dim mCount As Long
    
    'We must have a Drawings Object to continue!
    If mDrawings Is Nothing Then Exit Sub
    
    'Now we must have a start and stop!
    If IsMissing(vntShift) Then
        vntShift = 1
    End If
    
    
    ReDim mThisDigit(mDrawings.Drawn)
    ReDim mLastDigit(mDrawings.Drawn)
    
    'OK, we have to subtract for the drawing that was dropped
    'and add for the new drawing
    'so we gotta account for the dropped on (mStartDrawing - 1)
    'and we gotta account for the new drawing (mStopDrawing)
    
    mCount = mStop - mStart + 1
    If mCount = 1 Then Exit Sub
    'Old Start Drawing
    For i = mStart - 1 To mStart
        For n = 0 To mDrawings.Drawn - 1
            mThisDigit(n) = mDrawings.Item(i).Numbers(n)
            If mLastDigit(n) > 0 Then
                'mMeanPositionValue(N) = mMeanPositionValue(N) - mThisDigit(N)
                mBasePositionGap(mDrawings.Drawn) = mBasePositionGap(n) - Abs(mThisDigit(n) - mLastDigit(n))
            Else
                'this is the drawing being dropped
                mLastDigit(n) = mThisDigit(n)
                mGrossPositionValue(n) = mGrossPositionValue(n) - mThisDigit(n)
            End If
        Next n
    Next i
    
    If mStop <> mLastStop Then
        'New Stop Drawing
        ReDim mLastDigit(mDrawings.Drawn)
        For i = mStop - 1 To mStop
            For n = 0 To mDrawings.Drawn - 1
                mThisDigit(n) = mDrawings.Item(i).Numbers(n)
                If mLastDigit(n) > 0 Then
                    mBasePositionGap(mDrawings.Drawn) = mBasePositionGap(n) + Abs(mThisDigit(n) - mLastDigit(n))
                Else
                    mLastDigit(n) = mThisDigit(n)
                    mGrossPositionValue(n) = mGrossPositionValue(n) + mThisDigit(n)
                End If
            Next n
        Next i
        mLastStop = mStop
    End If
    
    'recalculate here
    For n = 0 To mDrawings.Drawn - 1
        mPositionGap(n) = mBasePositionGap(n) / (mCount - 1)
        mMeanPositionValue(n) = mGrossPositionValue(n) / mCount
    Next n

End Sub

Private Sub Rule0()
Dim lLowDigit As Long
Dim lHighDigit As Long

    If mRule0GapWeight = 0 Then Exit Sub
    
    'the broader the range, the more likely the numbers at the ends
    'of the range are what should be played
    'For Example:
    '3's and 5's keep being drawn in the first position
    'the meanposition value would be 4
    'and the position gap would be 3
    'so the span would be 3,4,5
    
    For n = 0 To mDrawings.Drawn - 1
        lLowDigit = mMeanPositionValue(n) - mPositionGap(n)
        lHighDigit = mMeanPositionValue(n) + mPositionGap(n)
        If lLowDigit < 0 Or lHighDigit < 0 Then Exit Sub
        If lLowDigit > 0 Then mPositionGapScore(lLowDigit) = mPositionGapScore(lLowDigit) + mRule0GapWeight
        If lHighDigit < mDrawings.BallCount + 1 Then mPositionGapScore(lHighDigit) = mPositionGapScore(lHighDigit) + mRule0GapWeight
        If mPositionGap(n) > 6 Then
            lLowDigit = lLowDigit + 1
            lHighDigit = lHighDigit - 1
            If lLowDigit > 0 Then mPositionGapScore(lLowDigit) = mPositionGapScore(lLowDigit) + mRule0GapWeight
            If lHighDigit < mDrawings.BallCount + 1 Then mPositionGapScore(lHighDigit) = mPositionGapScore(lHighDigit) + mRule0GapWeight
        End If
        If mPositionGap(n) > 12 Then
            lLowDigit = lLowDigit + 1
            lHighDigit = lHighDigit - 1
            If lLowDigit > 0 Then mPositionGapScore(lLowDigit) = mPositionGapScore(lLowDigit) + mRule0GapWeight
            If lHighDigit < mDrawings.BallCount + 1 Then mPositionGapScore(lHighDigit) = mPositionGapScore(lHighDigit) + mRule0GapWeight
        End If
    Next n
    
    
End Sub

Public Sub RunProcessor(vntFilterScore As Variant)

    If IsArray(vntFilterScore) Then
        ReDim mPositionGapScore(mDrawings.BallCount)
        ReDim mGrossPositionValue(mDrawings.BallCount)
        ReDim mBasePositionGap(mDrawings.BallCount)
        Read_Properties

        If mCalculated = Calculate Then
            Calc_Position_Gap
            'next time around, be sure to SHIFT instead...
            mCalculated = Shift
        Else
            If mStart = mLastStart + 1 Then
                Shift_Position_Gap
            Else
                'we have not shifted correctly
                'so we must recalculate
                Calc_Position_Gap
            End If
        End If
        
        'run the rules here
        Rule0
        
        For i = mDrawings.MinimumBallNumber To UBound(mScore)
            'tally the scores from the rules here
            mScore(i) = mPositionGapScore(i)
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
        ReDim mPositionGap(mDrawings.Drawn)
        ReDim mBasePositionGap(mDrawings.Drawn)
        ReDim mMeanPositionValue(mDrawings.Drawn)
        ReDim mGrossPositionValue(mDrawings.Drawn)
        ReDim mMinPositionValue(mDrawings.Drawn)
        ReDim mMaxPositionValue(mDrawings.Drawn)
    End If


End Property


Private Sub Class_Initialize()
Dim tempProperty As Object

    mKey = "positiongap"
    mName = "PositionGap"

    Set mProperties = CreateObject(sObjPropertyValues)
    Set tempProperty = mProperties.Add("gapweight")
    With tempProperty
        .Name = "Gap Weight"
        .Value = 0
        .Max = 10
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("maximumgap")
    With tempProperty
        .Name = "Maximum Gap"
        .Value = 0
        .Max = 15
        .Group = 1
        .Huntable = True
    End With
    mUseProcessor = True

End Sub

Private Sub Read_Properties()

    'Get Properties from the local property group
    mRule0GapWeight = mProperties("gapweight").Value
    mRule0MaximumGap = mProperties("maximumgap").Value

End Sub

Private Sub Write_Properties()

    'Set Properties to the local property group
    mProperties("gapweight").Value = mRule0GapWeight
    mProperties("maximumgap").Value = mRule0MaximumGap

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
