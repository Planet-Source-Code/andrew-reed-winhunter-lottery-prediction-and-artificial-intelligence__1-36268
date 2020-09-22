VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFilter"
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


'Most routines here use the following loop
'to flow through the drawings
'EXAMPLE:
'For i = (mDrawingsToBypass + Abs(mPredicting)) To mDrawingsToUse
'GRAPHICAL EXAMPLE
'   ...........X000000000--------------++++++++++++
'
'-----LEGEND-----
'. = drawings being bypassed                                    mDrawingsToBypass
'X = drawing to predict (INCLUDED IN STATS IF NOT PREDICTING)   mPredictDrawing
'0 = drawings skipped (unused in statistics)                    mDrawingsToSkip
'- = drawings used for statistics                               mDrawingsToUse
'+ = drawings left over (unused in statistics)




'All processors in this filter will run with the same settings
'assigned to this filter ONLY!


'the following two collections are added to
'via another method outside of the normal collection
'methods usually used (add/remove/clear)
'the addtrigger/addprocessor methods will have to be key based
Public Enum bStartCycle
    StartOn = -1
    StartOff = 0
End Enum

Private mProcessors As Collection       'the collection of filters to run
Private mTriggers As Collection         'the collection of triggers to use


Private mAllTriggers As Object     'local reference to the Triggers collection
Private mAllProcessors As Object 'local reference to the Processors collection
                                        'this is basically a read only collection
                                        'to pull the processors out of and put into
                                        'the local collection

Private mUseFilter As Boolean
Private mStartCycle As bStartCycle      'When to start the cycle
Private mOnCycle As Integer             'Filter cycle on length
Private mOffCycle As Integer            'Filter cycle off length
Private mCycleCounter As Integer        'Filter cycle counter

Private mLastStartDrawing As Integer    'Drawing from last run
Private mLastStopDrawing As Integer     'Drawing from last run
Private mStartDrawing As Integer        'Calculated Start Drawing
Private mStopDrawing As Integer         'Calculated Stop Drawing
Private mDrawingsToBypass As Integer    'Number of Drawings to skip
                                        'start from most recent and
                                        'skips preceding drawings
Private mDrawingsToUse As Integer       'number of drawings to use

Event FilterComplete(ByVal iCurrentdrawing As Integer)
'Event FilterStarted()
Event Plot(ByVal larrayScores As Variant)

Private mEndProcess As Boolean          'Stop the processor

Private mDrawings As Object        'local reference to Drawings Object
Private mKey As String                  'filter key, set during initialization
Private mName As String                 'filter name, set during initialization
Private mProperties As Object
Private mScore() As Long                'filter score, which is a total
                                        'of the processor scores,
                                        'which is then added to the group score
Private mPredictDrawing As Long
Private sObjPropertyValues As String
Private sObjProcessors As String
Private sObjTriggers As String

Private Sub GetObjects()
Dim SXML        As New CGoXML
Dim i           As Integer
Dim sObjName    As String
Dim sClsName    As String

    SXML.Initialize (pavAUTO)
    'START INITIAL FILE TEMPLATE
    Call SXML.OpenFromFile(App.Path & "\plugin.xml")
    If SXML.NodeCount("/PLUGINS/PLUGIN") > 0 Then
        For i = 0 To SXML.NodeCount("/PLUGINS/PLUGIN") - 1
            'PluginXML.OpenFromString SXML.ReadNodeXML("/PLUGINS/PLUGIN[" & i & "]")
            'use late binding to create objects
            'this way, it doesnt matter when the object was created
            'or what the object's registry key is
            sObjName = SXML.ReadNode("/PLUGINS/PLUGIN[" & i & "]/OBJECT_NAME")
            sClsName = SXML.ReadNode("/PLUGINS/PLUGIN[" & i & "]/CLASS_NAME")
            Select Case SXML.ReadNode("/PLUGINS/PLUGIN[" & i & "]/TYPE")
                Case "PROPERTYVALUES"
                    'CreateObject(sObjPropertyValues)
                    sObjPropertyValues = sObjName & "." & sClsName
                Case "PROCESSORS"
                    'CreateObject(sObjProcessors)
                    sObjProcessors = sObjName & "." & sClsName
                Case "TRIGGERS"
                    'CreateObject(sObjTriggers)
                    sObjTriggers = sObjName & "." & sClsName
            End Select
        Next
    End If

End Sub


Public Property Let PredictDrawing(lValue As Long)
    mPredictDrawing = lValue
End Property
Public Property Get PredictDrawing() As Long
    PredictDrawing = mPredictDrawing
End Property


Public Property Get UseFilter() As Boolean
    UseFilter = mUseFilter
End Property
Public Property Let UseFilter(bUseFilter As Boolean)
    mUseFilter = bUseFilter
End Property


Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(sName As String)
    mName = sName
End Property

Public Property Get Key() As String
    Key = mKey
End Property
Public Property Let Key(sKey As String)
    mKey = sKey
End Property


'set the local copy of the drawings to use here
Public Property Set Drawings(ByVal vData As Object)
Dim mLocalProcessor As Object
Dim mLocalTrigger As Object

    Set mDrawings = vData
    For Each mLocalProcessor In mProcessors
        Set mLocalProcessor.Drawings = mDrawings
    Next
    For Each mLocalTrigger In mTriggers
        Set mLocalTrigger.Drawings = mDrawings
    Next
    mProperties("use").Max = mDrawings.Count
    mProperties("skip").Max = mDrawings.Count - 1
    mProperties("oncycle").Max = mDrawings.Count
    mProperties("offcycle").Max = mDrawings.Count
    ReDim mScore(mDrawings.BallCount)

End Property

Public Property Get Drawings() As Object

    Set Drawings = mDrawings

End Property

Public Sub Reset()
Dim mLocalProcessor As Object
Dim mLocalTrigger As Object

    'reset the filter object here
    For Each mLocalProcessor In mProcessors
        mLocalProcessor.Calculated = False
    Next
    For Each mLocalTrigger In mTriggers
        mLocalTrigger.Reset
    Next
    Set mLocalProcessor = Nothing
    Set mLocalTrigger = Nothing
    'cause a reset of the cycler when the filter is run
    mCycleCounter = 0

End Sub

Public Sub ClearProcessors()
    'destroy the collection
    Set mProcessors = Nothing
    'create a new instance of the collection
    Set mProcessors = New Collection
End Sub

Public Function AddProcessor(sKey As Variant, sProcessorKey As Variant) As Object
    'create a new object
    Dim objNewMember As Object
    Set objNewMember = ProcessorItem(sProcessorKey)
    
    If objNewMember Is Nothing Then
        Set objNewMember = mAllProcessors(sProcessorKey)
        'Pass the drawings object down to the processors
        Set objNewMember.Drawings = mDrawings
        'set the properties passed into the method
        mProcessors.Add objNewMember, sProcessorKey
        'mProcessors.Add objNewMember, sProcessorKey
        'return the object created
    Else
        Set objNewMember = Nothing
    End If
    Set AddProcessor = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get ProcessorItem(vntIndexKey As Variant) As Object
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
    On Error GoTo ProcessorItemError
    Set ProcessorItem = mProcessors(vntIndexKey)
    'For Each objProcessor In mProcessors
    '    If objProcessor.Key = vntIndexKey Then
    '        Set ProcessorItem = objProcessor
    '        Exit For
    '    ElseIf objProcessor.Name = vntIndexKey Then
    '        Set ProcessorItem = mProcessors(vntIndexKey)
    '        Exit For
    '    End If
    'Next
    Exit Property

ProcessorItemError:

    If Err.Number = 5 Then
        Set ProcessorItem = Nothing
        Err.Clear
        Resume Next
    End If

End Property

Public Property Get ProcessorCount() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    ProcessorCount = mProcessors.Count
End Property

Public Sub RemoveProcessor(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    mProcessors.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mProcessors.[_NewEnum]
End Property




Public Sub ClearTriggers()
    'destroy the collection
    Set mTriggers = Nothing
    'create a new instance of the collection
    Set mTriggers = New Collection
End Sub

Public Function AddTrigger(sKey As Variant, sTriggerKey As Variant) As Object
    'create a new object
    Dim objNewMember As Object
    
    Set objNewMember = TriggerItem(sTriggerKey)
    If objNewMember Is Nothing Then
        Set objNewMember = mAllTriggers(sTriggerKey)
        'Pass the drawings object down to the processors
        Set objNewMember.Drawings = mDrawings
        'set the properties passed into the method
        mTriggers.Add objNewMember, sTriggerKey
        'mProcessors.Add objNewMember, sProcessorKey
        'return the object created
    Else
        Set objNewMember = Nothing
    End If
    'return the object created
    Set AddTrigger = objNewMember
    Set objNewMember = Nothing
End Function

Public Property Get TriggerItem(vntIndexKey As Variant) As Object
    On Error GoTo TriggerItemError
    Set TriggerItem = mTriggers(vntIndexKey)
    Exit Property

TriggerItemError:

    If Err.Number = 5 Then
        Set TriggerItem = Nothing
        Err.Clear
        Resume Next
    End If
End Property

Public Property Get TriggerCount() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    TriggerCount = mTriggers.Count
End Property

Public Sub RemoveTrigger(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)
    If mTriggers.Count > 0 Then
        mTriggers.Remove vntIndexKey
    End If
End Sub




Private Sub Class_Initialize()
Dim tempProperty As Object
    
    GetObjects
    'creates the collection when this class is created
    Set mProcessors = New Collection
    
    'Add the Available Processors here
    Set mAllProcessors = CreateObject(sObjProcessors)
    
    
    Set mTriggers = New Collection
    
    'Add the Available Triggers here
    Set mAllTriggers = CreateObject(sObjTriggers)
    
    Set mProperties = CreateObject(sObjPropertyValues)
    Set tempProperty = mProperties.Add("use")
    With tempProperty
        .Name = "Drawings To Use"
        .Value = 50
        .Min = 1
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("skip")
    With tempProperty
        .Name = "Drawings To Bypass"
        .Value = 1
        .Group = 1
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("oncycle")
    With tempProperty
        .Name = "Cycle On Count"
        .Value = 0
        .Group = 2
        .Max = 100
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("offcycle")
    With tempProperty
        .Name = "Cycle Off Count"
        .Value = 0
        .Group = 2
        .Max = 100
        .Huntable = True
    End With
    
    Set tempProperty = mProperties.Add("startcycle")
    With tempProperty
        .Name = "Cycle Start Point"
        .Value = 0
        .Group = 300
        .Min = 301
        .Max = 302
        .Huntable = True
    End With
    Set tempProperty = mProperties.Add("starton")
    With tempProperty
        .Name = "Start On"
        .Value = True
        .Group = 301
        .Huntable = False
    End With
    Set tempProperty = mProperties.Add("startoff")
    With tempProperty
        .Name = "Start Off"
        .Value = False
        .Group = 302
        .Huntable = False
    End With
    mUseFilter = True

End Sub

Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mProcessors = Nothing
    Set mAllProcessors = Nothing
    Set mTriggers = Nothing
    Set mAllTriggers = Nothing
    Set mProperties = Nothing
    Set mDrawings = Nothing
End Sub

Public Sub RunFilter(Optional vntGroupScore As Variant)
Dim oTrigger As Object
Dim oProcessor As Object
Dim bRun As Boolean

    
    'reset the filter object here
    ReDim mScore(mDrawings.BallCount)
    
    Read_Properties
    
    If mCycleCounter = 0 Then
        'this should only be run ONCE during a
        Set_Cycler
    End If
    'Here we will need to run several functions
    'First we need to run the triggers
    
    'Debug.Print mCycleCounter;
    
    If mOnCycle > 0 And mOffCycle > 0 Then
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

'NOTICE!
'NOTICE!
'THE CYCLER MUST BE BACK PROGRAMMED TO START FROM THE BEGINNING DRAWING
'WHICH IS THE OLDEST DRAWING IN THE DATABASE!
'THIS WAY, EACH TIME A PREDICTION IS RUN, THE CYCLER CAN EASILY DETERMINE
'IF IT SHOULD BE ON OR OFF!
'OTHERWISE, THE CYCLER WILL NOT BEHAVE AS IT SHOULD


        'The cycler takes precedence over the trigger
        'but when the cycler is off, and a trigger is available
        'then the trigger can also be run
        Select Case mCycleCounter
            Case Is < 0
                mCycleCounter = mCycleCounter + 1
                If mCycleCounter = 0 Then mCycleCounter = mOnCycle
            Case Is > 0
                mCycleCounter = mCycleCounter - 1
                If mCycleCounter = 0 Then mCycleCounter = mOffCycle * -1
                bRun = True
            Case Else
                MsgBox "Cycler Error!"
                If mStartCycle Then
                    'gonna run this time
                    'be sure to decrease by one for the count
                    mCycleCounter = mOnCycle - 1
                    bRun = True
                Else
                    mCycleCounter = (mOffCycle * -1) + 1
                End If
        End Select
        'Debug.Print bRun
        If Not bRun Then
            For Each oTrigger In mTriggers
                If Not oTrigger.Run Then
                    bRun = False
                    Exit For
                Else
                    bRun = True
                End If
            Next
        End If
    Else
        If mTriggers.Count > 0 Then
            For Each oTrigger In mTriggers
                If Not oTrigger.Run Then
                    bRun = False
                    Exit For
                Else
                    bRun = True
                End If
            Next
        Else
            'Nothing else is set, so we must base running
            'on pure user input
            bRun = mUseFilter
        End If
    End If
    
    If mDrawingsToBypass > Drawings.Count - 2 Then
        mDrawingsToBypass = mDrawings.Count - 2
    End If
    mStartDrawing = (mPredictDrawing + 1) + mDrawingsToBypass
    mStopDrawing = (mStartDrawing + mDrawingsToUse) - 1
    If mStopDrawing > mDrawings.Count Then
        'we can't go past the last drawing!
        mStopDrawing = mDrawings.Count
    End If
    If mStartDrawing > mDrawings.Count Then
        'we can't go past the last drawing!
        mStartDrawing = mDrawings.Count
    End If
    If mStartDrawing = mStopDrawing And mStopDrawing = mPredictDrawing Then
        'we cant use the same drawing to predict itself
        'so we must exit here
        Exit Sub
    End If
    
    If mProcessors.Count > 0 Then
        'If IsArray(vntGroupScore) Then
        For Each oProcessor In mProcessors
            If oProcessor.UseProcessor Then
                oProcessor.PredictDrawing = mPredictDrawing
                oProcessor.StartDrawing = mStartDrawing
                oProcessor.StopDrawing = mStopDrawing
                oProcessor.RunProcessor mScore
                DoEvents
            'Else
            '    MsgBox "skipping processor"
            End If
        Next
        'End If
    End If
    RaiseEvent Plot(mScore)
    
    'if we can run, then total the processors here
    'we MUST have an array for the score to be passed
    'back to the group object
    If Not IsMissing(vntGroupScore) Then
        If IsArray(vntGroupScore) Then
            If bRun Then
                'Add the filter scores to the group scores here
                For i = 1 To UBound(vntGroupScore)
                    vntGroupScore(i) = vntGroupScore(i) + mScore(i)
                Next i
            'Else
            '    If bRun Then
            '    End If
            End If
        End If
    End If

End Sub

Public Property Get PropertyValues() As Object

    'set the current local values to the collection?
    'Write_Properties
    
    
    
    'pass back the collection so the parent object
    'can have access to the collection
    Set PropertyValues = mProperties

End Property

Private Sub Read_Properties()

    'Get Properties from the local property group
    mDrawingsToUse = mProperties("use").Value
    mDrawingsToBypass = mProperties("skip").Value
    mStartCycle = mProperties("startcycle").Value
    mOnCycle = mProperties("oncycle").Value
    mOffCycle = mProperties("offcycle").Value
    
End Sub

Private Sub Set_Cycler()

    'no need to setup the cycler if we dont have a true cycle setup
    If mOnCycle > 0 And mOffCycle > 0 Then
        'We must setup the cycle counter here!
        mCycleCounter = 0
        For i = 1 To mDrawings.Count - mPredictDrawing
            Select Case mCycleCounter
                Case Is > 0
                    'we're in an on cycle
                    mCycleCounter = mCycleCounter + 1
                    If mCycleCounter > mOnCycle Then
                        'we've already reached the end of the on cycle,
                        'so we must reset to the beginning of the off cycle
                        mCycleCounter = -1
                    End If
                Case Is < 0
                    mCycleCounter = mCycleCounter - 1
                    If mCycleCounter < mOffCycle Then
                        'we've already reached the end of the off cycle,
                        'so we must reset to the beginning of the on cycle
                        mCycleCounter = 1
                    End If
                Case Else
                    'WE SHOULD ONLY START HERE ONCE!
                    'AND NEVER COME BACK DURING THE LOOPING CYCLE
                    If mStartCycle Then
                        'start the on cycle here
                        mCycleCounter = 1
                    Else
                        'start the off cycle
                        mCycleCounter = -1
                    End If
            End Select
        Next
    End If

End Sub

Private Sub Write_Properties()

    'Set Properties to the local property group
    mProperties("use").Value = mDrawingsToUse
    mProperties("skip").Value = mDrawingsToBypass
    mProperties("startcycle").Value = mStartCycle
    mProperties("oncycle").Value = mOnCycle
    mProperties("offcycle").Value = mOffCycle

End Sub

