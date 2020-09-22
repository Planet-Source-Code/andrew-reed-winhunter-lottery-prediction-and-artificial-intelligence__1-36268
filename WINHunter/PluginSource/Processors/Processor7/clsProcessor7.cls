VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcessor7"
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

Private mUseProcessor As Boolean

Private mStart As Long
Private mStop As Long
Private mLastStop As Long
Private mLastStart As Long
Private mPredictDrawing As Long
Private mUseDriver() As Integer

Private mRedBusScore() As Long              'Final Score after rule applied

Private mDriverPosition As String             'score weight
Private mConductorCount As Long             'score weight
Private mPassengerSelection As String       'score weight
'Private mBusWrap As Boolean                 'score weight



Private mKey As String                      'processor key, set during initialization
Private mName As String                     'processor name, set during initialization
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
    ReDim mRedBusScore(mDrawings.BallCount)
    mProperties("driverposition").Max = mDrawings.Drawn

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
Dim l As Long
Dim iConductorCount As Integer
Dim iDriver As Integer
Dim xShift As Integer
Dim yShift As Integer
Dim PassengerScore As Integer
Dim CurrentPassenger As Integer
Dim bRet As Boolean
Dim iSelectedPassenger As Integer
Dim xPosition As Integer
'There are 7 buses.  There are 7 drivers.  There are 3 conductors.  There are 9
'passengers in each bus.
'
'In the UK National the latest draw was 667: and these balls were drawn.
'
'4,6,10,19,40,49 BB 10  these are the 'Drivers' of the Bus.  The drivers are always the
'latest balls drawn in the uknational.txt file
'
'The conductor is the next occurance of a ball number (each driver) as you look down the
'UK Lottery results table.  the uknational.txt file.
'
'There are three conductors for each driver.
'Each conductor is responsible for three passengers.
'These passenger are
'\ and | and /
'Left diagonal, vertical and right diagonal.
'
'A Red Bus can now be visualised as consisting of a driver, three conductors and 9
'passengers.



    'get the driver
    Get_Drivers mDriverPosition
    For j = 0 To UBound(mUseDriver)
        iDriver = mDrawings.Item(mStart).Numbers(mUseDriver(j))
        iConductorCount = 0
        For l = mStart + 1 To mStop
            'loop through and find the first conductor
            If mDrawings.Item(l).Numbers(mUseDriver(j)) = iDriver Then
                iConductorCount = iConductorCount + 1
                'since it was drawn, we have to get the
                'position so we know where we need to
                'get the passengers from next
                
                'Process Passenger commands here
                bRet = GetPassengerInfo(CurrentPassenger, xShift, yShift, PassengerScore)
                'we cant select past the last drawing...
                If l + yShift < mStop + 1 Then
                    xPosition = mUseDriver(j) + xShift
                    'wrap around
                    If xPosition > (mDrawings.Drawn - 1) Then xPosition = xPosition - (mDrawings.Drawn - 1)
                    If xPosition < 0 Then xPosition = (mDrawings.Drawn - 1) + (xPosition + 1)
                    iSelectedPassenger = mDrawings.Item(l + yShift).Numbers(xPosition)
                    mRedBusScore(iSelectedPassenger) = mRedBusScore(iSelectedPassenger) + PassengerScore
                End If
                If bRet Then Exit For
            End If
            If iConductorCount = mConductorCount Then Exit For
        Next l
    Next j



End Sub

Private Function GetPassengerInfo(iPassenger, iX As Integer, iY As Integer, ReturnScore As Integer) As Boolean
Dim itemp As Integer
Dim sTemp As String
Dim sLast As String
Dim bFoundStart As Boolean
Dim bFoundStop As Boolean
Dim i As Integer
Dim iSelectionStringUsed As Integer
Dim iStartAt As Integer

'O<^X1 O^X2 O^>X3
'Here is the action legend:
'O - start from Conductor position
'^ - Up one drawing in the history
'<>- move one position left or right in the drawing
'X# - save current position, with # score weight

    ReturnScore = 0
    iStartAt = 1
    Do Until itemp = iPassenger + 1
        itemp = itemp + 1
        For i = iStartAt To Len(mPassengerSelection)
            sTemp = Mid$(mPassengerSelection, i, 1)
            Select Case sTemp
                Case "O", "o"
                    bFoundStart = True
                    bFoundStop = False
                    iX = 0
                    iY = 0
                Case "X", "x"
                    bFoundStop = True
                Case "<"
                    iX = iX - 1
                Case ">"
                    iX = iX + 1
                Case "^"
                    iY = iY + 1
                Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
                    If bFoundStop Then
                        'we have to make sure we have located the passenger first!
                        ReturnScore = Val(sTemp)
                        'If IsNumeric(sLast) Then
                        '    'handle scores up to 99
                        '    ReturnScore = (Val(sLast) * 10) + Val(sTemp)
                        'End If
                        Exit For
                    End If
            End Select
            'sLast = sTemp
        Next i
        iStartAt = i
        If i >= Len(mPassengerSelection) Then
            GetPassengerInfo = True
        End If
    Loop
    iPassenger = itemp

End Function

Private Sub Get_Drivers(vIn As Variant)
Dim iUseCount As Integer
Dim sUse As String
Dim sDriver As String
Dim i As Integer

    If VarType(vIn) = vbString Then
        sUse = vIn
        Do While Len(sUse) > 0
            ReDim Preserve mUseDriver(iUseCount)
            If InStr(sUse, ",") > 0 Then
                sDriver = Left$(sUse, InStr(sUse, ",") - 1)
            Else
                sDriver = sUse
            End If
            'position array is 0 based
            mUseDriver(iUseCount) = Val(RTrim(LTrim(sDriver))) - 1
            If Len(sUse) > Len(sDriver) Then
                sUse = Right(sUse, Len(sUse) - (Len(sDriver) + 1))
            Else
                sUse = ""
            End If
            iUseCount = iUseCount + 1
        Loop
    Else
        'make sure we have not exceeded the drawn qty
        If Not vIn > mDrawings.Drawn Then
            For i = 0 To vIn - 1
                'build/populate the array
                mUseDriver(i) = i + 1
            Next i
        End If
    End If

End Sub

Public Sub RunProcessor(vntFilterScore As Variant)

    If mUseProcessor Then
        If IsArray(vntFilterScore) Then
            Read_Properties
            ReDim mRedBusScore(mDrawings.BallCount)
            'run the rules here
            Rule0
            For i = mDrawings.MinimumBallNumber To UBound(mScore)
                'tally the scores from the rules here
                mScore(i) = mRedBusScore(i)
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
        ReDim mRedBusScore(mDrawings.BallCount)
    End If


End Property


Private Sub Class_Initialize()
Dim tempProperty As Object

    mKey = "redbus"
    mName = "RedBus"
    Set mProperties = CreateObject(sObjPropertyValues)
    
    Set tempProperty = mProperties.Add("driverposition")
    With tempProperty
        .Name = "Driver Positions"
        .Value = "1"
        .Group = 1
        .Min = 1
        .Max = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("conductorcount")
    With tempProperty
        .Name = "Conductor Count"
        .Value = 1
        .Group = 1
        .Min = 1
        .Max = 10
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("passengerselection")
    With tempProperty
        .Name = "Selection String"
        .Value = "O<^X1O^X2O^>X3"
        .Group = 1
        .Huntable = False
    End With
    mUseProcessor = True

End Sub

Private Sub Read_Properties()

    'Get Properties from the local property group
    mDriverPosition = mProperties("driverposition").Value
    mConductorCount = mProperties("conductorcount").Value
    mPassengerSelection = mProperties("passengerselection").Value

End Sub

Private Sub Write_Properties()

    'Set Properties to the local property group
    mProperties("driverposition").Value = mDriverPosition
    mProperties("conductorcount").Value = mConductorCount
    mProperties("passengerselection").Value = mPassengerSelection

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



