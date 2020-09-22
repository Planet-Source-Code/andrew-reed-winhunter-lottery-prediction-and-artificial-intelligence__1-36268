VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessor9"
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

Private mPerkisized() As Long
Private mPerkisizeScore() As Long         'Final Score after rule applied
Private mRule0GroupASize As Long
Private mRule0GroupBSize As Long
Private mRule0GroupCSize As Long
Private mRule0GroupDSize As Long
Private mRule0GroupESize As Long
Private mRule0PerkisizeWeight As Long

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
    ReDim mPerkisizeScore(mDrawings.BallCount)
    ReDim mPerkisized(mDrawings.BallCount)

End Property


Public Property Let StartDrawing(lValue As Long)
    mLastStart = mStart
    mStart = lValue
End Property
Public Property Let StopDrawing(lValue As Long)
    mStop = lValue
End Property
Public Property Let PredictDrawing(lValue As Long)
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
Public Sub Calc_Perkisize()
Dim lCount As Long
Dim lDigit As Long

    'We must have a Drawings Object to continue!
    If mDrawings Is Nothing Then Exit Sub
    'Now we must have a start and stop!
    If mStart = 0 Then
        mStart = 1
    End If
    'mStop = mStart + 16
    If mStop > mDrawings.Count Then
        'we cant perkisize less than 16 drawings
        Exit Sub
    End If
    
    
    If mRule0GroupASize + mRule0GroupBSize + _
        mRule0GroupCSize + mRule0GroupDSize + mRule0GroupESize < mStop - mStart Then Exit Sub

    ReDim mPerkisized(mDrawings.BallCount)
    lCount = 1
    For i = mStart To mStop
        For n = 0 To mDrawings.Drawn - 1
            If lCount < mRule0GroupASize + 1 Then
                'A Group
                'increment count
                lDigit = mDrawings.Item(i).Numbers(n)
                mPerkisized(lDigit) = mPerkisized(lDigit) + 1
            End If
            If lCount > mRule0GroupASize + mRule0GroupBSize And lCount < mRule0GroupASize + mRule0GroupBSize + mRule0GroupCSize + 1 Then
                'C Group
                'increment count
                lDigit = mDrawings.Item(i).Numbers(n)
                mPerkisized(lDigit) = mPerkisized(lDigit) + 1
            End If
        Next n
        lCount = lCount + 1
    Next i
    lCount = 1
    For i = mStart To mStop
        For n = 0 To mDrawings.Drawn - 1
            If lCount > mRule0GroupASize And lCount < mRule0GroupASize + mRule0GroupBSize + 1 Then
                'B Group
                'increment count
                lDigit = mDrawings.Item(i).Numbers(n)
                If mPerkisized(lDigit) = 0 Then
                    'mark the unscored numbers
                    mPerkisized(lDigit) = -1
                ElseIf mPerkisized(lDigit) > 0 Then
                    'increment the count here
                    mPerkisized(lDigit) = mPerkisized(lDigit) + 1
                End If
            End If
            If lCount > mRule0GroupASize + mRule0GroupBSize + mRule0GroupCSize + mRule0GroupDSize And lCount < mRule0GroupASize + mRule0GroupBSize + mRule0GroupCSize + mRule0GroupDSize + mRule0GroupESize + 1 Then
                'E Group
                'increment count
                lDigit = mDrawings.Item(i).Numbers(n)
                If mPerkisized(lDigit) = 0 Then
                    'mark the unscored numbers
                    mPerkisized(lDigit) = -1
                End If
            End If
        Next n
        lCount = lCount + 1
    Next i
    
    mLastStop = mStop

End Sub

Private Sub Rule0()

    If mRule0PerkisizeWeight = 0 Then Exit Sub
    
    For i = mDrawings.MinimumBallNumber To mDrawings.BallCount
        If mPerkisized(i) Then
            mPerkisizeScore(i) = mRule0PerkisizeWeight
        End If
    Next i
    
    
End Sub

Public Sub RunProcessor(vntFilterScore As Variant)

    If IsArray(vntFilterScore) Then
        ReDim mPerkisizeScore(mDrawings.BallCount)
        ReDim mScore(mDrawings.BallCount)
        Read_Properties
        'If mCalculated = Calculate Then
            Calc_Perkisize
            'next time around, be sure to SHIFT instead...
            'mCalculated = Shift
        'Else
            'Shift_Position_Gap
        'End If
        
        'run the rules here
        Rule0
        
        For i = mDrawings.MinimumBallNumber To UBound(mScore)
            'tally the scores from the rules here
            mScore(i) = mPerkisizeScore(i)
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
        ReDim mPerkisized(mDrawings.BallCount)
        ReDim mPerkisizeScore(mDrawings.BallCount)
    End If


End Property


Private Sub Class_Initialize()
Dim tempProperty As Object

    mKey = "perkisize"
    mName = "Perkisize"

    Set mProperties = CreateObject(sObjPropertyValues)
    Set tempProperty = mProperties.Add("groupa")
    With tempProperty
        .Name = "Count"
        .Value = 3
        .Max = 25
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("groupb")
    With tempProperty
        .Name = "Count/Eliminate"
        .Value = 3
        .Max = 25
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("groupc")
    With tempProperty
        .Name = "Count"
        .Value = 4
        .Max = 25
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("groupd")
    With tempProperty
        .Name = "Skip"
        .Value = 3
        .Max = 25
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("groupe")
    With tempProperty
        .Name = "Eliminate"
        .Value = 3
        .Max = 25
        .Group = 1
        .Huntable = True
    End With
    
    Set tempProperty = mProperties.Add("perkisizeweight")
    With tempProperty
        .Name = "Perkisize Weight"
        .Value = 0
        .Max = 10
        .Group = 1
        .Huntable = True
    End With
    mUseProcessor = True

End Sub

Private Sub Read_Properties()

    'Get Properties from the local property group
    mRule0PerkisizeWeight = mProperties("perkisizeweight").Value
    mRule0GroupASize = mProperties("groupa").Value
    mRule0GroupBSize = mProperties("groupb").Value
    mRule0GroupCSize = mProperties("groupc").Value
    mRule0GroupDSize = mProperties("groupd").Value
    mRule0GroupESize = mProperties("groupe").Value

End Sub

Private Sub Write_Properties()

    'Set Properties to the local property group
    mProperties("perkisizeweight").Value = mRule0PerkisizeWeight
    mProperties("groupa").Value = mRule0GroupASize
    mProperties("groupb").Value = mRule0GroupBSize
    mProperties("groupc").Value = mRule0GroupCSize
    mProperties("groupd").Value = mRule0GroupDSize
    mProperties("groupe").Value = mRule0GroupESize

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
