VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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


Private mPredictedCount As Long
Private mExclude As Variant
Private mPredictDrawing As Long
Private mStopAt As Long
Private mMatchCount As Long
Private mPickCount As Long
Private mOutput() As Long
Private mHits() As Long
Private mJackpotOutputs() As Long
Private mAverageOutput As Long
Private mAverageHits As Long
Private mJackpotHits As Long
Private mJackpotOutput As Long
Private mDrawings As Object        'local reference to Drawings Object


'set the local copy of the drawings to use here
Public Property Set Drawings(ByVal vData As Object)

    Set mDrawings = vData
    ReDim mOutput(mDrawings.Count)
    ReDim mHits(mDrawings.Count)

End Property

Public Sub Count_Matches(vntExclude As Variant, iPredicted As Long)
'Dim iCheckHit As Integer

    mPredictedCount = 0
    'If mPredict = 1 Then mMatchCount = 0
    Set mExclude = Nothing
    mExclude = vntExclude
    mPredictDrawing = iPredicted
    mMatchCount = 0
    If iPredicted = 0 Then Exit Sub
    'For n = 0 To mDrawings.Drawn - 1
        'If mPredict = 1 Then iCheckHit = mDrawings.Item(iPredicted).Numbers(n)
        For i = 1 To UBound(vntExclude)
            If Not vntExclude(i) Then
                'Count the predicted numbers on the first loop only
                mPredictedCount = mPredictedCount + 1
                'If mPredicted = 1 Then
                    If mDrawings.Item(iPredicted).WasDrawn(i) Then
                        mMatchCount = mMatchCount + 1
                    End If
                'End If
            End If
        Next i
        'If mPredict = 0 Then Exit For
    'Next n
    mOutput(iPredicted) = mPredictedCount
    mHits(iPredicted) = mMatchCount
    
    mAverageOutput = 0
    mAverageHits = 0
    For i = 1 To iPredicted
        mAverageOutput = mAverageOutput + mOutput(i)
        mAverageHits = mAverageHits + mHits(i)
    Next i
    mAverageOutput = mAverageOutput / iPredicted
    mAverageHits = mAverageHits / iPredicted
    
    
    
    'Dont count this drawing in the jackpot count,
    'because it wasnt a true jackpot hit
    'especially it will adveresly affect the hit counts
    If mPredictedCount = mDrawings.BallCount Then Exit Sub
    mJackpotOutput = 0
    
    
    
    If mMatchCount = mDrawings.Drawn Then
        mJackpotHits = mJackpotHits + 1
        'add JackpotOutput Here
        ReDim Preserve mJackpotOutputs(mJackpotHits)
        mJackpotOutputs(mJackpotHits) = mPredictedCount
    End If
    If mJackpotHits > 0 Then
        For i = 1 To mJackpotHits
            mJackpotOutput = mJackpotOutput + mJackpotOutputs(i)
        Next i
        mJackpotOutput = mJackpotOutput / mJackpotHits
    End If


End Sub

Public Property Get matchcount() As Integer

    matchcount = mMatchCount

End Property

Public Property Get Excluded(ByVal iIndex As Integer) As Boolean
  
    'IF we have an index#, then pass the value back
    'no index, then pass the array object back!
    'If IsMissing(iIndex) Then
    '    Predicted = mScore
    'Else
        Excluded = mExclude(iIndex)
    'End If

End Property


Public Property Get JackpotOutput() As Integer

    JackpotOutput = mJackpotOutput

End Property

Public Property Get JackpotHits() As Integer

    JackpotHits = mJackpotHits

End Property

Public Property Get BallCount() As Integer

    BallCount = mDrawings.BallCount

End Property

Public Property Get DrawCount() As Integer

    DrawCount = mDrawings.Drawn

End Property

Public Property Get AverageHits() As Integer

    AverageHits = mAverageHits

End Property

Public Property Get AverageOutput() As Integer

    AverageOutput = mAverageOutput

End Property

Public Property Get PredictDrawing() As Integer

    PredictDrawing = mPredictDrawing

End Property


Public Property Get PredictedCount() As Integer

    PredictedCount = mPredictedCount

End Property

Public Sub Reset()

    ReDim mOutput(mDrawings.Count)
    ReDim mHits(mDrawings.Count)
    ReDim mJackpotOutputs(mDrawings.Count)
    mAverageOutput = 0
    mAverageHits = 0
    mJackpotHits = 0
    mJackpotOutput = 0

End Sub

