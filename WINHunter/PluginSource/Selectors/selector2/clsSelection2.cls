VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSelection2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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



Private mCrestToTrough As Long              'Amplitude
Private mCenterLine As Long                 'Verticle Offset
Private mPeakToPeak As Long                 'Frequency
Private mOffset As Long                     'Horizontal Offset

Private mMinimumBallNumber As Integer       'lowest ball number
Private mKey As String                      'selector key
Private mName As String                     'selector name, set during initialization
Private mKeeps As String
Private mKeep As Variant
Private mDeletes As String
Private mDelete As Variant
Private mMinimumToSelect As Integer         'Minimum # of Draws to select
Private mSelectorMethod As iSelectorMethod  'selector method
Private mSelectorType As bType              'does it select balls to include, or exclude?
Private mExcluded As Variant                'Balls selected during selection phase
                                            '(does not include previously selected balls)
Private mScore() As Long

Private mProperties As Object

Public Property Get MinimumBallNumber() As Integer
    MinimumBallNumber = mMinimumBallNumber
End Property
Public Property Let MinimumBallNumber(iMinimumBallNumber As Integer)
    mMinimumBallNumber = iMinimumBallNumber
End Property

Public Property Get PropertyValues() As Object

    'set the current local values to the collection?
    'Write_Properties
    
    
    
    'pass back the collection so the parent object
    'can have access to the collection
    Set PropertyValues = mProperties

End Property


Public Sub ResetSelector(ByVal iBallCount As Integer)
    'the Selector array must be reset(redim'd) before use
    ReDim mExcluded(iBallCount)
End Sub
Public Property Get Excluded(Optional ByVal lIndex As Variant) As Variant
    'IF we have an index#, then pass the value back
    'no index, then pass the array object back!
    If IsMissing(lIndex) Then
        Excluded = mExcluded
    Else
        Excluded = mExcluded(lIndex)
    End If
End Property
Public Property Let Selected(ByVal lIndex As Variant, ByVal bSelected As Variant)
    mExcluded(lIndex) = bSelected
End Property

Public Property Get Score() As Variant

    'pass the entire array
    Score = mScore

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




Private Sub Class_Initialize()
Dim tempProperty As Object

    mKey = "ocilliscope"
    mName = "O-scope"
    Set mProperties = CreateObject(sObjPropertyValues)
    
    
    Set tempProperty = mProperties.Add("cresttotrough")
    With tempProperty
        .Name = "Crest To Trough"
        .Value = 1
        .Group = 1
    End With
    Set tempProperty = mProperties.Add("peaktopeak")
    With tempProperty
        .Name = "Peak to Peak"
        .Value = 1
        .Group = 1
    End With
    Set tempProperty = mProperties.Add("centerline")
    With tempProperty
        .Name = "Centerline"
        .Value = 0
        .Group = 1
    End With
    Set tempProperty = mProperties.Add("offset")
    With tempProperty
        .Name = "Offset"
        .Value = 0
        .Group = 1
    End With
    
    Set tempProperty = mProperties.Add("selector")
    With tempProperty
        .Name = "Selection Method"
        .Value = 1
        .Group = 200
    End With
    Set tempProperty = mProperties.Add("selecthigh")
    With tempProperty
        .Name = "High"
        .Value = 1
        .Group = 201
    End With
    Set tempProperty = mProperties.Add("selectaverage")
    With tempProperty
        .Name = "Average"
        .Value = 2
        .Group = 202
    End With
    Set tempProperty = mProperties.Add("selectlow")
    With tempProperty
        .Name = "Low"
        .Value = 3
        .Group = 203
    End With
    Set tempProperty = mProperties.Add("selectinvertedhigh")
    With tempProperty
        .Name = "Inverted High"
        .Value = 4
        .Group = 204
    End With
    Set tempProperty = mProperties.Add("selectinvertedaverage")
    With tempProperty
        .Name = "Inverted Average"
        .Value = 5
        .Group = 205
    End With
    Set tempProperty = mProperties.Add("selectinvertedlow")
    With tempProperty
        .Name = "Inverted Low"
        .Value = 6
        .Group = 206
    End With
 

 

 

 
    
    
    
    Set tempProperty = mProperties.Add("type")
    With tempProperty
        .Name = "Selection Type"
        .Value = 0
        .Group = 300
    End With
    Set tempProperty = mProperties.Add("inclusion")
    With tempProperty
        .Name = "Inclusion"
        .Value = True
        .Group = 301
    End With
    Set tempProperty = mProperties.Add("exclusion")
    With tempProperty
        .Name = "Exclusion"
        .Value = False
        .Group = 302
    End With
    
    
    
    Set tempProperty = mProperties.Add("selectmin")
    With tempProperty
        .Name = "Minimum To Select"
        .Value = 10
        .Group = 4
    End With

    Set tempProperty = mProperties.Add("keeps")
    With tempProperty
        .Name = "Keeps"
        .Value = ""
        .Group = 5
    End With
    
    Set tempProperty = mProperties.Add("deletes")
    With tempProperty
        .Name = "Deletes"
        .Value = ""
        .Group = 5
    End With

End Sub

Private Sub Read_Properties()

    'Get Properties from the local property group
    mCrestToTrough = mProperties("cresttotrough").Value
    mCenterLine = mProperties("centerline").Value
    mPeakToPeak = mProperties("peaktopeak").Value
    mOffset = mProperties("offset").Value
    
    mSelectorMethod = mProperties("selector").Value
    mMinimumToSelect = mProperties("selectmin").Value
    mSelectorType = mProperties("type").Value
    mKeeps = mProperties("keeps").Value
    mDeletes = mProperties("deletes").Value

End Sub

Private Sub Write_Properties()

    'Set Properties to the local property group
    mProperties("cresttotrough").Value = mCrestToTrough
    mProperties("centerline").Value = mCenterLine
    mProperties("peaktopeak").Value = mPeakToPeak
    mProperties("offset").Value = mOffset
    
    mProperties("selectmin").Value = mMinimumToSelect
    mProperties("selector").Value = mSelectorMethod
    mProperties("type").Value = mSelectorType
    mProperties("keeps").Value = mKeeps
    mProperties("deletes").Value = mDeletes

End Sub


Private Sub Class_Terminate()

    Set mProperties = Nothing

End Sub


Public Sub RunSelect(mExclude As Variant, selScore As Variant)
Dim lHigh As Long
Dim lLow As Long
Dim lAverage As Long
Dim iBallNum As Integer
Dim UpperWaterMark As Long
Dim LowerWaterMark As Long
Dim iSelectedCount As Integer
Dim bMovedHigh As Boolean
Dim lDrawnCount As Long
Dim lBallCount As Long
Dim bTemp() As Boolean


    
    'Ok, a valid sine wave is a plotted arc which flows through 360 degrees of motion
    
    
    
    
    'Dim F, Pi As Single
    'Dim SR, A, L, Hz, O As Integer
    '
    'Hz = 440 Frequency in Hertz (cycles per second)
    'L = 1000 Length of sample (in samples)
    'A = 32760 Amplitude of Operator
    'SR = 22000 Sampling Rate in Hertz
    'T = Samples/Ticks (time)
    'F = Frequency of Operator
    '
    'Pi = 3.1415
    'F = (2 * Pi * Hz) / SR
    'For T = 0 To L
    '    O = Int(A * (sIn(F * T)))
    'Next T
    'Equation 2 - Additive Waveform
    'O3 = (O1 + O2) / 2
    'Equation 3 - Feedback Sine Wave
    'O = A ´ (sin((F * T) + (FB * O)))
    'Equation 4 - FM Waveform
    'O2 = A2 ´ (sin((F2 ´ T) + (FI ´ O1)))

    
    
    
    
    
    
    If Not IsArray(mExclude) And Not IsArray(selScore) Then
        'We NEED the previously exluded balls to show up here
        MsgBox "Error Initializing Selector"
    End If
    ReDim bTemp(UBound(mExclude))
    ReDim mScore(UBound(selScore))
    mScore = selScore
    mExcluded = bTemp
    Read_Properties
    lDrawnCount = mMinimumToSelect
    lBallCount = UBound(mExclude)
    
    'Ok, let's get the high/low/avg vars
    iSelectedCount = 0
    lLow = 10000
    
    If Len(mKeeps) > 0 Then
        If InStr(mKeeps, ",") Then
            mKeep = Split(mKeeps, ",")
        ElseIf InStr(mKeeps, " ") Then
            mKeep = Split(mKeeps, " ")
        ElseIf InStr(mKeeps, ";") Then
            mKeep = Split(mKeeps, ";")
        End If
    Else
        mKeep = Empty
    End If
    If Len(mDeletes) > 0 Then
        If InStr(mDeletes, ",") Then
            mDelete = Split(mDeletes, ",")
        ElseIf InStr(mDeletes, " ") Then
            mDelete = Split(mDeletes, " ")
        ElseIf InStr(mDeletes, ";") Then
            mDelete = Split(mDeletes, ";")
        End If
    Else
        mDelete = Empty
    End If
    
    
    For iBallNum = mMinimumBallNumber To lBallCount
        'don't count previously excluded balls here!
        If Not mExclude(iBallNum) Then
            If selScore(iBallNum) < lLow Then
                lLow = selScore(iBallNum)
            End If
            If selScore(iBallNum) > lHigh Then
                lHigh = selScore(iBallNum)
            End If
            If iBallNum > 1 Then
                lAverage = (lAverage + selScore(iBallNum)) / 2
            Else
                lAverage = selScore(iBallNum)
            End If
            iSelectedCount = iSelectedCount + 1
        End If
    Next iBallNum
    
    'Added select minimum checking...
    'So if we have already selected the minimum
    'or if we have eliminated too many, then exit
    If iSelectedCount = mMinimumToSelect Or iSelectedCount < mMinimumToSelect Then
        For iBallNum = mMinimumBallNumber To lBallCount
            mExcluded(iBallNum) = mExclude(iBallNum)
        Next iBallNum
        Exit Sub
    End If
    
    If lLow = 0 And lHigh = 0 And lAverage = 0 Then
        'Hmmm, no score
        'we must be in some sort of filter cycle
        'so just pass the excluded balls through
        For iBallNum = 1 To lBallCount
            mExcluded(iBallNum) = mExclude(iBallNum)
        Next iBallNum
        Exit Sub
    End If
    
    If mSelectorMethod = selInvertedHigh Or mSelectorMethod = selInvertedAverage Or mSelectorMethod = selInvertedLow Then
        'Hmmm, first we have to INVERT the scores
        'then we get the high/low/avg scores
        'MsgBox "You need to revamp the inverted selection methods!)"
        'ReDim selScore(mDrawings.BallCount)
        For iBallNum = mMinimumBallNumber To lBallCount
            'dont count previously excluded balls here!
            If Not mExclude(iBallNum) Then
                If selScore(iBallNum) > lAverage Then
                    'score is above the average line
                    selScore(iBallNum) = lAverage - (selScore(iBallNum) - lAverage)
                ElseIf selScore(iBallNum) < lAverage Then
                    'score is below the average line
                    selScore(iBallNum) = lAverage + (lAverage - selScore(iBallNum))
                End If
            End If
        Next iBallNum
        'Ok, let's get the high/low/avg vars
        lLow = 10000
        For iBallNum = mMinimumBallNumber To lBallCount
            'dont count previously excluded balls here!
            If Not mExclude(iBallNum) Then
                If selScore(iBallNum) < lLow Then
                    lLow = selScore(iBallNum)
                End If
                If selScore(iBallNum) > lHigh Then
                    lHigh = selScore(iBallNum)
                End If
                If iBallNum > 1 Then
                    lAverage = (lAverage + selScore(iBallNum)) / 2
                Else
                    lAverage = selScore(iBallNum)
                End If
            End If
        Next iBallNum
    End If
    
    Select Case mSelectorMethod
        Case selHigh
            UpperWaterMark = lHigh
            Do
                For iBallNum = mMinimumBallNumber To lBallCount
                    'dont count previously excluded balls here!
                    If Not mExclude(iBallNum) Then
                        If selScore(iBallNum) > UpperWaterMark Then
                            If Not (mExcluded(iBallNum) = mSelectorType) Then
                                iSelectedCount = iSelectedCount + 1
                                'Typical Include with type set to exclusion
                                mExcluded(iBallNum) = mSelectorType
                            End If
                        Else
                            'dont count the ball twice
                            If Not mExcluded(iBallNum) Then
                                'Typical Exclude with type set to exclusion
                                mExcluded(iBallNum) = Not mSelectorType
                                'will subtract one if false
                                iSelectedCount = iSelectedCount + (Not mSelectorType)
                                'Debug.Print iSelected
                                'If iSelectedCount = lDrawnCount Then
                                '    Exit Sub
                                'End If
                            End If
                        End If
                    Else
                        'pass on the Excluded balls
                        mExcluded(iBallNum) = True
                    End If
                Next iBallNum
                If iSelectedCount < mMinimumToSelect Or iSelectedCount = lBallCount Then
                    UpperWaterMark = UpperWaterMark - 1
                End If
                'Safety Mechanism, exit the loop at some point...
                If UpperWaterMark < lLow Then Exit Do
            Loop While iSelectedCount < mMinimumToSelect Or iSelectedCount = lBallCount
        Case selInvertedHigh
            UpperWaterMark = lHigh
            Do
                For iBallNum = mMinimumBallNumber To lBallCount
                    'dont count previously excluded balls here!
                    If Not mExclude(iBallNum) Then
                        If selScore(iBallNum) = UpperWaterMark Then
                            iSelectedCount = iSelectedCount + 1
                            'Typical Include with type set to exclusion
                            mExcluded(iBallNum) = mSelectorType
                        Else
                            'dont count the ball twice
                            If Not mExcluded(iBallNum) Then
                                'Typical Exclude with type set to exclusion
                                mExcluded(iBallNum) = Not mSelectorType
                                'will subtract one if false
                                iSelectedCount = iSelectedCount + (Not mSelectorType)
                                'Debug.Print iSelected
                                'If iSelectedCount = lDrawnCount Then
                                '    Exit Sub
                                'End If
                            End If
                        End If
                    Else
                        'pass on the Excluded balls
                        mExcluded(iBallNum) = True
                    End If
                Next iBallNum
                If iSelectedCount < mMinimumToSelect Or iSelectedCount = lBallCount Then
                    UpperWaterMark = UpperWaterMark - 1
                End If
                If UpperWaterMark < lLow Then Exit Do
            Loop While iSelectedCount < mMinimumToSelect Or iSelectedCount = lBallCount
        Case selLow
            LowerWaterMark = lLow
            Do
                For iBallNum = mMinimumBallNumber To lBallCount
                    'dont count previously excluded balls here!
                    If Not mExclude(iBallNum) Then
                        If selScore(iBallNum) < LowerWaterMark Then
                            iSelectedCount = iSelectedCount + 1
                            'Typical Include with type set to exclusion
                            mExcluded(iBallNum) = mSelectorType
                        Else
                            'dont count the ball twice
                            If Not mExcluded(iBallNum) Then
                                'Typical Exclude with type set to exclusion
                                mExcluded(iBallNum) = Not mSelectorType
                                'will subtract one if false
                                iSelectedCount = iSelectedCount + (Not mSelectorType)
                                'Debug.Print iSelected
                                'If iSelectedCount = lDrawnCount Then
                                '    Exit Sub
                                'End If
                            End If
                        End If
                    Else
                        'pass on the Excluded balls
                        mExcluded(iBallNum) = True
                    End If
                Next iBallNum
                If iSelectedCount < mMinimumToSelect Or iSelectedCount = lBallCount Then
                    LowerWaterMark = LowerWaterMark + 1
                End If
                If LowerWaterMark > lHigh Then Exit Do
            Loop While iSelectedCount < mMinimumToSelect Or iSelectedCount = lBallCount
        Case selInvertedLow
            LowerWaterMark = lLow
            Do
                For iBallNum = mMinimumBallNumber To lBallCount
                    'dont count previously excluded balls here!
                    If Not mExclude(iBallNum) Then
                        If selScore(iBallNum) = LowerWaterMark Then
                            iSelectedCount = iSelectedCount + 1
                            'Typical Include with type set to exclusion
                            mExcluded(iBallNum) = mSelectorType
                        Else
                            'dont count the ball twice
                            If Not mExcluded(iBallNum) Then
                                'Typical Exclude with type set to exclusion
                                mExcluded(iBallNum) = Not mSelectorType
                                'will subtract one if false
                                iSelectedCount = iSelectedCount + (Not mSelectorType)
                                'Debug.Print iSelected
                                'If iSelectedCount = lDrawnCount Then
                                '    Exit Sub
                                'End If
                            End If
                        End If
                    Else
                        'pass on the Excluded balls
                        mExcluded(iBallNum) = True
                    End If
                Next iBallNum
                If iSelectedCount < mMinimumToSelect Or iSelectedCount = lBallCount Then
                    LowerWaterMark = LowerWaterMark + 1
                End If
                If LowerWaterMark > lHigh Then Exit Do
            Loop While iSelectedCount < mMinimumToSelect Or iSelectedCount = lBallCount
        Case selAverage, selInvertedAverage
            UpperWaterMark = lAverage
            LowerWaterMark = lAverage
            Do
                For iBallNum = mMinimumBallNumber To lBallCount
                    'dont count previously excluded balls here!
                    If Not mExclude(iBallNum) Then
                        If selScore(iBallNum) > LowerWaterMark And selScore(iBallNum) < UpperWaterMark Then
                            iSelectedCount = iSelectedCount + 1
                            'Typical Include with type set to exclusion
                            mExcluded(iBallNum) = mSelectorType
                        Else
                            'dont count the ball twice
                            If Not mExcluded(iBallNum) Then
                                'Typical Exclude with type set to exclusion
                                mExcluded(iBallNum) = Not mSelectorType
                                'will subtract one if false
                                iSelectedCount = iSelectedCount + (Not mSelectorType)
                                'Debug.Print iSelected
                                'If iSelectedCount = lDrawnCount Then
                                '    Exit Sub
                                'End If
                            End If
                        End If
                    Else
                        'pass on the Excluded balls
                        mExcluded(iBallNum) = True
                    End If
                Next iBallNum
                If iSelectedCount < mMinimumToSelect Or iSelectedCount = lBallCount Then
                    If bMovedHigh Then
                        LowerWaterMark = LowerWaterMark - 1
                        bMovedHigh = False
                    Else
                        bMovedHigh = True
                        UpperWaterMark = UpperWaterMark + 1
                    End If
                End If
                If LowerWaterMark < lLow Then Exit Do
                If UpperWaterMark > lHigh Then Exit Do
            Loop While iSelectedCount < mMinimumToSelect Or iSelectedCount = lBallCount
    End Select
    
    If Not IsEmpty(mKeep) Then
        For iSelectedCount = 0 To UBound(mKeep)
            'un-exclude the Keeps
            If IsNumeric(mKeep(iSelectedCount)) Then
                If Val(mKeep(iSelectedCount)) < UBound(mExclude) Then
                    mExclude(Val(mKeep(iSelectedCount))) = False
                End If
            End If
        Next iSelectedCount
    End If
    If Not IsEmpty(mDelete) Then
        For iSelectedCount = 0 To UBound(mDelete)
            'un-exclude the Keeps
            If IsNumeric(mDelete(iSelectedCount)) Then
                If Val(mDelete(iSelectedCount)) < UBound(mExclude) Then
                    mExclude(Val(mDelete(iSelectedCount))) = True
                End If
            End If
        Next iSelectedCount
    End If
    
    
    'For i = 1 To 53
    '    Debug.Print mExcluded(i)
    'Next i

End Sub


