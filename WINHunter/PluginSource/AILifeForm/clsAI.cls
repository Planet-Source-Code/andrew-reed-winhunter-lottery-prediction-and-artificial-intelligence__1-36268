VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAI"
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
Public Event status(ByVal lConfidenceLevel As Long, ByVal dSuccess As Double, ByVal lDecisionsPerSecond As Long)

Private mStack          As clsStack         'this is the current stack that we are tinkering with
Private mGroup          As clsGroup         'this is the current group that we are tinkering with
Private mFilter         As clsFilter        'this is the current filter that we are tinkering with
Private mProcessor      As Object           'this is the current processor that we are tinkering with
Private mTrigger        As Object           'this is the current trigger that we are tinkering with
Private mSelector       As Object           'this is the current selector that we are tinkering with
Private mLifeSign       As Boolean          'is the LifeForm still alive?
Private mKillLifeForm   As Boolean          'end the Lifeform's cycle

Private lPreviousPredictionCount As Long
Private mDrawings As Object        'local reference to Drawings Object

Public Property Get LifeSign() As Boolean

    LifeSign = mLifeSign

End Property

Public Property Let KillLifeForm(bTerminateLife As Boolean)

    mKillLifeForm = bTerminateLife

End Property

Public Sub BeginLife()

    'frmMonitorAI.Show
    'DoEvents
    Call AI_Life_Cycle
    DoEvents

End Sub

Private Sub AI_Life_Cycle()
Dim bNewPath        As Boolean
Dim lThinking       As Long
Dim lChooseItem     As Long
Dim lReflections    As Long
Dim oProcessors     As Object
Dim oTriggers       As Object
Dim oSelectors      As Object
    
    'only run one instance
    If mLifeSign Then Exit Sub
    DoEvents
    
    Set oProcessors = CreateObject(sObjProcessors)
    Set oSelectors = CreateObject(sObjSelections)
    Set oTriggers = CreateObject(sObjTriggers)
    'initialize Stack for AI life cycle
    mStack.PredictDrawing = 1
    'reset stored value
    lPreviousPredictionCount = 0
    
    'Start a whole new path,
    'Unless we have saved paths
    bNewPath = True
    
    'ensure the LifeForm will live
    mKillLifeForm = False
    
    'This life form exists all on it's own.
    'It does not exist within the Machine's Own Stack collection
    Set mStack.Drawings = mDrawings
    
    'Show Signs of Life
    mLifeSign = True
    'let's start out with some Very good confidence
    lReflections = 50
    'Birth LifeForm into existence
    Do
        DoEvents
        'Start a new path here.
        'We will start a new path for many reasons:
        '1> We are just starting our life
        '2> We didnt like the path we chose
        '3> We are curious about a new path, and save the old path for later
        If bNewPath Then
            'Start_New_Path
            'Make sure we dont start a new path again,
            'unless we decide to later on
            bNewPath = False
            GoTo Start_New_Path
        End If
        
        'Analyze the Path we are going down.
        'Is it a good path?
        'Do we want to wait to see if it is a good path?
        'Is it just a very bad path?
        'Should we choose a new path?
        lThinking = Analyze_Path(lThinking, lReflections)

        'execute the decision
        Select Case lThinking
            Case 0  'Choose a new path
                'let's start out with some Very good confidence
                lReflections = 50
Start_New_Path:
                'bNewPath = True
                mStack.Groups.Clear
                'Now we need to add a Brand new group
                GoTo AddGroup
            
            Case 1  'Journey down the path
AddGroup:
                'Add a new Group to the Stack
                Set mGroup = mStack.Groups.Add(GenerateKey)
                'We will have to choose from the available selectors
                'We need to make sure our choice does not already exist
                lChooseItem = Random_Choice(1, mGroup.SelectorCount)
                mGroup.SetSelector (lChooseItem)
                Adjust_Properties mGroup.Selector.PropertyValues
                'First we need to add a Filter
                GoTo AddFilter
            
            Case 2  'Walk down path
AddFilter:
                'Add a new Filter to the Group
                Set mFilter = mGroup.AddFilter(GenerateKey)
                'now change some settings to the Filter
                Adjust_Properties mFilter.PropertyValues
                
                'After adding a new filter, we need to decide what else
                'we want to add to the filter, otherwise just a fliter will
                'do nothing at all
                'First we need to add a Processor
                GoTo AddProcessor
                
            Case 3  'Take a step down the path
AddProcessor:
                'Add a new Processor to the Filter
                'Have we used all the processors for this filter?
                If mFilter.ProcessorCount < oProcessors.Count Then
                    'We will have to choose from the available processors
                    'We need to make sure our choice does not already exist
                    Set mProcessor = Nothing
                    Do Until Not mProcessor Is Nothing
                        lChooseItem = Random_Choice(1, oProcessors.Count)
                        Set mProcessor = mFilter.AddProcessor(GenerateKey, oProcessors.Item(lChooseItem).Key)
                    Loop
                    Adjust_Properties mProcessor.PropertyValues
                Else
                    'No more available processors for this filter
                    'should we make another decision?????????????????????????
                    'MsgBox "No more processors available."
                End If
                lThinking = 3
                
            Case 4  'Change Pace
AddTrigger:
                'Add a Trigger to the Filter
                'Have we used all the triggers for this filter?
                If mFilter.TriggerCount < oTriggers.Count Then
                    'We will have to choose from the available processors
                    'We need to make sure our choice does not already exist
                    Set mTrigger = Nothing
                    Do Until Not mTrigger Is Nothing
                        lChooseItem = Random_Choice(1, oTriggers.Count)
                        Set mTrigger = mFilter.AddTrigger(GenerateKey, oTriggers.Item(lChooseItem).Key)
                    Loop
                    Adjust_Properties mTrigger.PropertyValues
                Else
                    'No more available triggers for this filter
                    'should we make another decision?????????????????????????
                    'MsgBox "No more triggers available."
                End If
            
            Case 5  'Change Glasses
                'Change Selector
                'We will have to choose from the available selectors
                'We need to make sure our choice does not already exist
                lChooseItem = Random_Choice(1, mGroup.SelectorCount)
                mGroup.SetSelector (lChooseItem)
                Adjust_Properties mGroup.Selector.PropertyValues
                
            Case 6  'Attempt to clean up the path
                'Work with current Filter settings
                Adjust_Properties mFilter.PropertyValues
                
            Case 7  'Attempt to clean up the path
                'Work with current Processor settings
                Adjust_Properties mProcessor.PropertyValues
                
            Case 8  'Attempt to clean up the path
                'Work with current Trigger settings
                Adjust_Properties mTrigger.PropertyValues
                
            Case 9  'Wipe our eyes
                'Work with current Selector settings
                Adjust_Properties mGroup.Selector.PropertyValues
                
                
                
                
            Case 10  'Change Path
                'Save Current path for later testing,
                'and start a new path
                SaveStack mStack, GenerateFileName & ".xml"
                
                'let's start out with some Very good confidence
                lReflections = 50
                'nNewPath = True
                mStack.Groups.Clear
                Set mGroup = Nothing
                Set mFilter = Nothing
                Set mProcessor = Nothing
                Set mTrigger = Nothing
                'Now we need to add a Brand new group
                GoTo AddGroup
                
            Case 11  'Resume old path
                'stop(save?) current path, and resume an old path
                SaveStack mStack, GenerateFileName & ".xml"
                'let's start out with some Very good confidence
                lReflections = 50
                'nNewPath = True
                mStack.Groups.Clear
                Set mGroup = Nothing
                Set mFilter = Nothing
                Set mProcessor = Nothing
                Set mTrigger = Nothing
                'Now we need to load an old stack
                LoadStack mStack, FindStackFile
                
                
            Case 12  'Journey down the path, again
                'Remove Previous Group & Add a new Group to the Stack
                mStack.Groups.Remove mGroup.Key
                Set mGroup = Nothing
                Set mFilter = Nothing
                Set mProcessor = Nothing
                Set mTrigger = Nothing
                'let's start out with some Very good confidence
                lReflections = 25
                GoTo AddGroup
                
            Case 13  'Walk down path, again
                'Remove Previous Filter & Add a new Filter to the Group
                mStack.Groups(mGroup.Key).RemoveFilter mFilter.Key
                Set mFilter = Nothing
                Set mProcessor = Nothing
                Set mTrigger = Nothing
                'let's start out with some Very good confidence
                lReflections = 15
                GoTo AddFilter
                
            Case 14  'Take a step down the path, again
                'Remove Previous Processor & Add a new Processor to the Filter
                mStack.Groups(mGroup.Key).FilterItem(mFilter.Key).RemoveProcessor mProcessor.Key
                Set mProcessor = Nothing
                'let's start out with some Very good confidence
                lReflections = 10
                GoTo AddProcessor
                
            Case 15  'Change Pace, again
                'Remove Previous Trigger & Add a Trigger to the Filter
                mStack.Groups(mGroup.Key).FilterItem(mFilter.Key).RemoveTrigger mTrigger.Key
                Set mTrigger = Nothing
                GoTo AddTrigger
                
        End Select
        DoEvents
        'If lReflections > 50 Then MsgBox "Got more confidence."
    Loop Until mKillLifeForm
    mLifeSign = False
    mKillLifeForm = False

End Sub

Private Sub Adjust_Properties(ItemProperties As Object)
Dim ItemProperty As Object
Dim l As Long

    'loop through the properties
    For l = 1 To ItemProperties.Count
        'get the property item here
        Set ItemProperty = ItemProperties(l)
        With ItemProperty
            'can we make choices for this item?
            If .Huntable Then
                'make a random decision first to change the property or not
                If (Random_Choice(0, 1) * -1) Then
                    'ok, we have decided to change this property
                    If .Group < 99 Then
                        'just calculate the value
                        .Value = Random_Choice(.Min, .Max)
                    Else
                        'calculate a combo box entry
                        'based on the group values
                        'once we have a group value, then get the value from the group item
                        .Value = ItemProperties(Random_Choice(.Min, .Max) - .Group).Value
                    End If
                End If
            End If
        End With
    Next

End Sub

Private Sub Count_Decisions()
Static dSTime As Double
Static lDecisions As Long

    If Timer > dSTime + 1 Then
        dSTime = Timer
        lDecisions = 0
    End If
    lDecisions = lDecisions + 1

End Sub


Private Function Analyze_Path(ByVal lPreviousDecision As Long, lConfidence As Long) As Long
Dim lDecision           As Long
Dim dRawSuccessRate     As Double
Dim dSuccessRate        As Double
Dim bBadChoice          As Double
Static dSTime As Double
Static lDecisions As Long
Static lD As Long


    'During the analysis phase, WINHunter will only test the most recent lottery drawing
    'To see how well it can predict it.
    'This method should be blazing fast.
    'Thus, it will allow many itterations of code, and allow many different tests to be performed.
    
    'Run the stack.
    mStack.RunStack
    'DoEvents
    'Once done, we can pull out the prediction results
    'mStack.Output.matchcount         'number of winning balls matched
    'mStack.Output.PredictedCount     'number of balls predicted
    
    'Raw success rate is simply the number of matches divided by the number of balls drawn
    dRawSuccessRate = (mStack.Output.matchcount / mStack.Drawings.Drawn) * 100
    dSuccessRate = dRawSuccessRate * ((mStack.Drawings.BallCount - mStack.Output.PredictedCount) / (mStack.Drawings.BallCount - mStack.Drawings.Drawn))
    
    
    
    'Make the first logical choice based on these facts!!!!!!!!!!!!!!!!!!!!
    '----------------------------------------------------------------------
    
    'fact #1
    'If the predicted count is the same as the total number of balls in the machine
    'then we didnt make ANY predictions at all.
    'fact #2
    'If we didnt predict any/enough numbers, then whatever path decision we made
    'must have been a wrong one and the prediction isnt valid.
    'fact #3
    'If we did predict a valid range of numbers, and the rawsucess rate is <100%
    'then we have set a bad setting
    '
    'Cause: Invalid settings, due to previous choice of 6,7,8,9
    'Fix: Change Settings.
    'Choice Values: 6,7,8,9
    '
    'Cause: Invalid settings, due to previous choice of 1,2,3,4,5
    'Fix: Change Settings.
    'Choice Values: 12,13,14,15
    If mStack.Output.PredictedCount = mStack.Drawings.BallCount _
        Or mStack.Output.PredictedCount < mStack.Drawings.Drawn _
        Or (dRawSuccessRate < 100 And mStack.Output.PredictedCount < lPreviousPredictionCount) Then
        'should we make a logical choice here,
        'or should we allow the AI to make a random choice?
        'Maybe we should allow it to make a random decision every once in a while
        If lConfidence >= 5 Then
            If Random_Choice(1, 8) < 5 Then
                Select Case lPreviousDecision
                    Case 2, 3, 4, 5
                        'added a new function,
                        'so let's work with it a little bit first
                        lDecision = lPreviousDecision + 4
                    Case 6, 7, 8, 9
                        'still tinkering
                        lDecision = lPreviousDecision
                    Case 10, 11
                        'still tinkering
                        'only we are tinkering with a re-loaded file
                        lDecision = lPreviousDecision
                End Select
                lConfidence = lConfidence - 1   'lost confidence
            Else
                'still tinkering
                lDecision = Random_Choice(6, 9)
                
                'Can we tinker the selected item?
                'Does the selected item have objects to tinker with?
                If lDecision = 8 Then
                    bBadChoice = True
                    Do
                        If lDecision = 8 And mTrigger Is Nothing Then
                            lDecision = Random_Choice(6, 9)
                        Else
                            bBadChoice = False
                        End If
                    Loop Until Not bBadChoice
                End If
               lConfidence = lConfidence - 1   'lost confidence
            End If
        Else
            Select Case lPreviousDecision
                Case 1, 2, 3, 4
                    'added a new function,
                    'so let's remove it
                    lDecision = lPreviousDecision + 11
                Case 5
                    'changed selector,
                    'so let's change again?
                    If Random_Choice(1, 8) < 5 Then
                        'change the settings
                        lDecision = 9
                    Else
                        'change selector
                        lDecision = 5
                    End If
                Case 6, 7, 8
                    'tinkered with a function, so let's remove it
                    lDecision = lPreviousDecision + 7
            End Select
            'we need to have a little confidence in the decisions we make
            'so let's regain some confidence
            lConfidence = lConfidence + 5
        End If
    End If
    
    'fact #4
    'If we did predict a valid range of numbers, and the rawsuccess rate is 100%
    'and the qty predicted is lesser than the previous qty predicted
    'then we should explore further.
    '
    'Cause: Valid settings.
    'Choice Values: 1,2,3,4 AND 9,10(save & move on)
    If dRawSuccessRate = 100 Then
        If mStack.Output.PredictedCount < lPreviousPredictionCount Then
            'randomly choose a function to perform
            Select Case Random_Choice(1, 12)
                Case 1, 2
                    lDecision = 1
                Case 3, 4
                    lDecision = 2
                Case 5, 6
                    lDecision = 3
                Case 7, 8
                    lDecision = 4
                Case 9
                    lDecision = 9
                Case 10
                    lDecision = 10
                Case 11, 12
                    lDecision = 11
           End Select
            lConfidence = lConfidence + 10      'increase in confidence
        End If
    ElseIf dRawSuccessRate < 50 And lDecision = 0 Then
        If Random_Choice(1, 8) < 5 Then
            Select Case lPreviousDecision
                Case 2, 3, 4, 5
                    'added a new function,
                    'so let's work with it a little bit first
                    lDecision = lPreviousDecision + 4
                Case 6, 7, 8, 9
                    'still tinkering
                    lDecision = lPreviousDecision
            End Select
            lConfidence = lConfidence - 1   'lost confidence
        Else
            'still tinkering
            lDecision = Random_Choice(6, 9)
            
            'Can we tinker the selected item?
            'Does the selected item have objects to tinker with?
            If lDecision = 8 Then
                bBadChoice = True
                Do
                    If lDecision = 8 And mTrigger Is Nothing Then
                        lDecision = Random_Choice(6, 9)
                    Else
                        bBadChoice = False
                    End If
                Loop Until Not bBadChoice
            End If
           lConfidence = lConfidence - 1   'lost confidence
        End If
    End If
    
    'If dRawSuccessRate > 80 And dRawSuccessRate < 100 Then
    '    MsgBox "partial sucess"
    'End If
    'If lDecision = 0 Then
    '    MsgBox "AI was unable to make a decision, not enough facts."
    'End If
    
    'store previous predicted count
    lPreviousPredictionCount = mStack.Output.PredictedCount
    
    'Limit the confidence level from spiraling negatively
    If lConfidence < 1 Then
        'we have hit rock bottom in our confidence
        'We might feel better sometimes more than other times...
        'so we need to regain some confidence in order to move forward
        lConfidence = Random_Choice(1, 10)
    End If
    If Timer < 10 Then
        'Timer was reset at midnight
        dSTime = 0
    End If
    If Timer > dSTime + 1 Then
        dSTime = Timer
        lDecisions = lD
        lD = 0
    End If
    lD = lD + 1
    
    'frmMonitorAI.txtConfidence = lConfidence
    'frmMonitorAI.txtSuccess = dSuccessRate
    'frmMonitorAI.txtDecisions = lDecisions
    DoEvents

    RaiseEvent status(lConfidence, dSuccessRate, lDecisions)
    'Choice has been made, now act on it
    'If lDecision = 8 Then MsgBox "trigger"
    Analyze_Path = lDecision
    
End Function

Private Function Random_Choice(lMinValue As Long, lMaxValue As Long) As Long
Dim lTrueMax As Long
Dim lTrueMin As Long

    'Make a random choice based upon minimum and maximum values.
    'If lMinValue = lMaxValue And lMinValue = 1 Then
    '    'not much of a choice here...
    '    Random_Choice = 1
    'ElseIf lMaxValue = 1 And lMinValue = 0 Then
    '    Random_Choice = CLng(CLng(((lMaxValue * 4) * Rnd) + lMinValue) / 4)
    'Else
    '    'randomize the random number generator seed
    '    Randomize
    '    'make the choice
    '    Random_Choice = CLng((lMaxValue * Rnd) + lMinValue)
    'End If
    lTrueMax = lMaxValue - lMinValue
    'Make a random choice based upon minimum and maximum values.
    If lTrueMax = 0 And lMinValue = 1 Then
        'not much of a choice here...
        Random_Choice = 1
    ElseIf lTrueMax = 1 Then
        Random_Choice = CLng(CLng(((lTrueMax * 4) * Rnd) / 4) + lMinValue)
    Else
        'randomize the random number generator seed
        Randomize
        'make the choice
        Random_Choice = CLng((lTrueMax * Rnd) + lMinValue)
    End If
    
End Function

Private Sub LoadStack(oStack As clsStack, sFileName As String)
Dim StackXML    As New CGoXML
Dim GroupXML    As New CGoXML
Dim FilterXML   As New CGoXML
Dim TempXML     As New CGoXML
Dim PropVals    As Object
Dim iPropCount  As Integer
Dim sGroupKey   As String
Dim sFilterKey  As String
Dim sProcessorKey As String
Dim sTriggerKey As String

    StackXML.Initialize (pavAUTO)
    GroupXML.Initialize (pavAUTO)
    FilterXML.Initialize (pavAUTO)
    TempXML.Initialize (pavAUTO)
    Call StackXML.OpenFromFile(App.Path & "\AI\" & sFileName)
    If StackXML.NodeCount("/STACK/GROUPS") > 0 Then
        Do
            GroupXML.OpenFromString StackXML.ReadNodeXML("/STACK/GROUPS/GROUP[" & oStack.Groups.Count & "]")
            'add the group to the stack
            sGroupKey = "grou" & GenerateKey
            Set mGroup = oStack.Groups.Add(sGroupKey)
            
            
            
            With oStack.Groups(sGroupKey)
                'get the boolean value
                .UseGroup = Str2Bool(GroupXML.ReadAttribute("/GROUP", "usegroup"))
                If GroupXML.NodeCount("/GROUP/FILTERS/FILTER") > 0 Then
                    Do
                        FilterXML.OpenFromString GroupXML.ReadNodeXML("/GROUP/FILTERS/FILTER[" & .FilterCount & "]")
                        sFilterKey = "filter" & GenerateKey
                        Set mFilter = .AddFilter(sFilterKey)
                        
                        
                        
                        
                        With .FilterItem(sFilterKey)
                            .UseFilter = Str2Bool(FilterXML.ReadAttribute("/FILTER", "usefilter"))
                            Set PropVals = .PropertyValues
                            iPropCount = 0
                            Do
                                PropVals.Item(FilterXML.ReadAttribute("/FILTER/PROPERTIES/PROPERTY[" & iPropCount & "]", "keyname")).Value = FilterXML.ReadNode("/FILTER/PROPERTIES/PROPERTY[" & iPropCount & "]")
                                iPropCount = iPropCount + 1
                            Loop Until FilterXML.NodeCount("/FILTER/PROPERTIES/PROPERTY") = iPropCount
                            
                            If FilterXML.NodeCount("/FILTER/PROCESSORS/PROCESSOR") > 0 Then
                                Do
                                    TempXML.OpenFromString FilterXML.ReadNodeXML("/FILTER/PROCESSORS/PROCESSOR[" & .ProcessorCount & "]")
                                    sProcessorKey = "processor" & GenerateKey
                                    Set mProcessor = .AddProcessor(sProcessorKey, TempXML.ReadAttribute("/PROCESSOR", "keyname"))
                                    With .ProcessorItem(TempXML.ReadAttribute("/PROCESSOR", "keyname"))
                                        .UseProcessor = Str2Bool(TempXML.ReadAttribute("/PROCESSOR", "useprocessor"))
                                        Set PropVals = .PropertyValues
                                        iPropCount = 0
                                        Do
                                            PropVals.Item(TempXML.ReadAttribute("/PROCESSOR/PROPERTIES/PROPERTY[" & iPropCount & "]", "keyname")).Value = TempXML.ReadNode("/PROCESSOR/PROPERTIES/PROPERTY[" & iPropCount & "]")
                                            iPropCount = iPropCount + 1
                                        Loop Until TempXML.NodeCount("/PROCESSOR/PROPERTIES/PROPERTY") = iPropCount
                                    End With
                                Loop Until .ProcessorCount = FilterXML.NodeCount("/FILTER/PROCESSORS/PROCESSOR")
                            End If
                            
                            
                            If FilterXML.NodeCount("/FILTER/TRIGGERS/TRIGGER") > 0 Then
                                Do
                                    TempXML.OpenFromString FilterXML.ReadNodeXML("/FILTER/TRIGGERS/TRIGGER[" & .TriggerCount & "]")
                                    sProcessorKey = "trigger" & GenerateKey
                                    Set mTrigger = .AddTrigger(sTriggerKey, TempXML.ReadAttribute("/TRIGGER", "keyname"))
                                    With .TriggerItem(TempXML.ReadAttribute("/TRIGGER", "keyname"))
                                        .UseTrigger = Str2Bool(TempXML.ReadAttribute("/TRIGGER", "usetrigger"))
                                        Set PropVals = .PropertyValues
                                        iPropCount = 0
                                        Do
                                            PropVals.Item(TempXML.ReadAttribute("/TRIGGER/PROPERTIES/PROPERTY[" & iPropCount & "]", "keyname")).Value = TempXML.ReadNode("/TRIGGER/PROPERTIES/PROPERTY[" & iPropCount & "]")
                                            iPropCount = iPropCount + 1
                                        Loop Until TempXML.NodeCount("/TRIGGER/PROPERTIES/PROPERTY") = iPropCount
                                    End With
                                Loop Until .ProcessorCount = FilterXML.NodeCount("/FILTER/TRIGGERS/TRIGGER")
                            End If
                        End With
                        
                        
                    Loop Until .FilterCount = GroupXML.NodeCount("/GROUP/FILTERS/FILTER")
                End If
                'Add the selector here
                TempXML.OpenFromString GroupXML.ReadNodeXML("/GROUP/SELECTOR")
                .SetSelector TempXML.ReadAttribute("/SELECTOR", "keyname")
                With .Selector
                    Set PropVals = .PropertyValues
                    iPropCount = 0
                    Do
                        PropVals.Item(TempXML.ReadAttribute("/SELECTOR/PROPERTIES/PROPERTY[" & iPropCount & "]", "keyname")).Value = TempXML.ReadNode("/SELECTOR/PROPERTIES/PROPERTY[" & iPropCount & "]")
                        iPropCount = iPropCount + 1
                    Loop Until TempXML.NodeCount("/SELECTOR/PROPERTIES/PROPERTY") = iPropCount
                End With
            End With
        Loop Until oStack.Groups.Count = StackXML.NodeCount("/STACK/GROUPS/GROUP")
    End If

End Sub

Private Function Str2Bool(sData As String) As Boolean

    If Not IsNumeric(sData) Then
        Select Case LCase(sData)
            Case "t", "true", "yes", "inclusion"
                Str2Bool = True
            Case "f", "false", "no", "exclusion"
                Str2Bool = False
            Case Else
                Str2Bool = False
        End Select
    Else
        Select Case LCase(sData)
            Case "-1"
                Str2Bool = True
            Case "0"
                Str2Bool = False
            Case Else
                Str2Bool = False
        End Select
    End If

End Function

Private Sub SaveStack(oStack As clsStack, sSaveName As String)            ', sHistoryName As String, sSaveName As String)
Dim SXML        As New CGoXML
Dim PropVals    As Object
Dim Prop        As Object
Dim iGroup      As Integer
Dim iFilter     As Integer
Dim iProcessor  As Integer

    SXML.Initialize (pavAUTO)
    'START INITIAL FILE TEMPLATE
    Call SXML.OpenFromString("<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " ?>" & vbCrLf & "<STACK>" & vbCrLf & "</STACK>")

    'If Not SXML.InsertNode("/STACK", "INITIAL_HISTORY_FILE", sHistoryName) Then Exit Sub
    If Not SXML.InsertNode("/STACK", "GROUPS") Then Exit Sub
    For iGroup = 0 To oStack.Groups.Count - 1
        With oStack.Groups.Item(iGroup + 1)
            If Not SXML.InsertNode("/STACK/GROUPS", "GROUP", "", "usegroup", .UseGroup) Then Exit Sub
            If Not SXML.InsertNode("/STACK/GROUPS/GROUP", "GROUP_NAME", .Name) Then Exit Sub
            If .FilterCount > 0 Then
                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]", "FILTERS") Then Exit Sub
                For iFilter = 0 To .FilterCount - 1
                    With .FilterItem(iFilter + 1)
                        If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS", "FILTER", "", "usefilter", .UseFilter) Then Exit Sub
                        If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]", "FILTER_NAME", .Name) Then Exit Sub
                        If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]", "PROPERTIES") Then Exit Sub
                        Set PropVals = .PropertyValues
                        For Each Prop In PropVals
                            Select Case Prop.Group
                                Case Is < 100
                                    'textbox input
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Sub
                                Case 100, 200, 300, 400, 500, 600, 700, 800, 900
                                    'combo box input
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Sub
                                Case Else
                                    'MsgBox "invalid Property"
                            End Select
                        Next
                        
                        If .ProcessorCount > 0 Then
                            If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]", "PROCESSORS") Then Exit Sub
                            For iProcessor = 0 To .ProcessorCount - 1
                                With .ProcessorItem(iProcessor + 1)
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS", "PROCESSOR", "", "useprocessor", .UseProcessor) Then Exit Sub
                                    If Not SXML.WriteAttribute("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]", "keyname", .Key) Then Exit Sub
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]", "PROCESSOR_NAME", .Name) Then Exit Sub
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]", "PROPERTIES") Then Exit Sub
                                    Set PropVals = .PropertyValues
                                    For Each Prop In PropVals
                                        Select Case Prop.Group
                                            Case Is < 100
                                                'textbox input
                                                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Sub
                                            Case 100, 200, 300, 400, 500, 600, 700, 800, 900
                                                'combo box input
                                                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Sub
                                            Case Else
                                                'MsgBox "invalid Property"
                                        End Select
                                    Next
                                End With
                            Next iProcessor
                        End If
                    End With
                Next iFilter
                With .Selector
                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]", "SELECTOR", "", "keyname", .Key) Then Exit Sub
                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/SELECTOR", "SELECTOR_NAME", .Name) Then Exit Sub
                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/SELECTOR", "PROPERTIES") Then Exit Sub
                    Set PropVals = .PropertyValues
                    For Each Prop In PropVals
                        Select Case Prop.Group
                            Case Is < 100
                                'textbox input
                                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/SELECTOR/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Sub
                            Case 100, 200, 300, 400, 500, 600, 700, 800, 900
                                'combo box input
                                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/SELECTOR/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Sub
                            Case Else
                                'MsgBox "invalid Property"
                        End Select
                    Next
                End With
            End If
        End With
    Next iGroup
    If Not SXML.Save(App.Path & "\AI\" & sSaveName) Then Exit Sub
    
End Sub

Private Function GenerateKey() As String

    GenerateKey = strScrambled(0, 3, 2)

End Function

Private Function GenerateFileName() As String

    GenerateFileName = strScrambled(0, 4, 4)

End Function

Private Function FindStackFile() As String
Dim FS      As FileSystemObject
Dim fldr    As Folder
Dim fil     As File
Dim FileCount As Long

    Set FS = New FileSystemObject
    Set fldr = FS.GetFolder(App.Path & "\AI")
    
    
    'We need to randomly find a stack file that works
    'we might need to implement a method to make sure this file is truly compatible with the type
    'of machine (6of49, 5of39, etc.) we are pedicting for
    FileCount = fldr.Files.Count
    If FileCount > 0 Then
        For Each fil In fldr.Files
            FileCount = FileCount - 1
            If FileCount = 0 Then
                FindStackFile = fil.Name
                Exit For
           End If
        Next
    End If
    
    FileCount = fldr.Files.Count
    'We can initiate a full blown file test routine here
    If FileCount > 99 Then
        'Let's loop through all the files and test the entire history against the settings found
        TestFiles
    End If


End Function

Private Sub TestFiles()
Dim FS                  As FileSystemObject
Dim fldr                As Folder
Dim fil                 As File
Dim FileCount           As Long
Dim sStackFile          As String
Dim TestStack           As clsStack
Dim dStackSuccess       As Double
Dim dStackSuccessRate   As Double

    Set FS = New FileSystemObject
    Set fldr = FS.GetFolder(App.Path & "\AI")
    For Each fil In fldr.Files
        sStackFile = fil.Name
        Set TestStack = New clsStack
        Set TestStack.Drawings = mStack.Drawings
        LoadStack TestStack, sStackFile
        TestStack.TestStack
        'Stack success is based upon Jackpot hits ONLY
        'so the Stack success rate should always be 100%
        'dStackSuccess = (6 / TestStack.Drawings.Drawn) * 100
        dStackSuccessRate = 100 * ((TestStack.Drawings.BallCount - TestStack.Output.JackpotOutput) / (TestStack.Drawings.BallCount - TestStack.Drawings.Drawn))
        If dStackSuccessRate < 5 Then
            'MsgBox "Deleting Stack File."
            fil.Delete
        End If
        If dStackSuccessRate < 30 And Not dStackSuccessRate < 5 Then
            fil.Move App.Path & "\AI\DUMP\"
        End If
    Next
    
    
    
End Sub

Public Property Set Drawings(ByVal vData As Object)

    'set the local copy of the drawings to use here
    Set mDrawings = vData

End Property

Private Sub Class_Initialize()

    GetObjects
    'Spawn an individual instance of a stack for the lifeform's existence
    Set mStack = New clsStack

End Sub

Private Sub Class_Terminate()

    'destroy any evidence of the entity
    Set mStack = Nothing
    Set mGroup = Nothing
    Set mFilter = Nothing
    Set mProcessor = Nothing
    Set mTrigger = Nothing
    Set mSelector = Nothing
    Set mDrawings = Nothing

End Sub
