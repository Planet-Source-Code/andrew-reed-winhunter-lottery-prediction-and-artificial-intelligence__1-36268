Attribute VB_Name = "modUtility"
Public sObjDrawings As String
Public sObjPropertyValues As String
Public sObjProcessors As String
Public sObjSelections As String
Public sObjTriggers As String

Public Sub GetObjects()
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
                Case "DRAWINGS"
                    'CreateObject(sObjDrawings)
                    sObjDrawings = sObjName & "." & sClsName
                Case "PROCESSORS"
                    'CreateObject(sObjProcessors)
                    sObjProcessors = sObjName & "." & sClsName
                Case "SELECTIONS"
                    'CreateObject(sObjSelections)
                    sObjSelections = sObjName & "." & sClsName
                Case "TRIGGERS"
                    'CreateObject(sObjTriggers)
                    sObjTriggers = sObjName & "." & sClsName
            End Select
        Next
    End If

End Sub

Public Function strScrambled(intCaps%, intLows%, intNums%) As String
Dim intUpperbound%, intLowerbound%, intChar%, strHold$
Dim intCap%, intLow%, intNum%, intPick%


    Do
    'RetryPick:
        Randomize           'randomize the random number generator seed
        If intPick = 0 Then
            'force a lower case letter first!!
            intPick = 2
        Else
            'randomly return numbers 1, 2 or 3 in PICK
            intPick = Int(3 * Rnd + 1)
        End If
        
    'here we will select cap, low or num
    'based on if we have run out of the others yet...
    'you could probably make this loop back and re-randomize again
    'if you don't want it to default to the next item
    'uncomment out the GOTO statement and comment out the extra if-then
    Select Case intPick
            Case 1
                If intCap < intCaps Then
                    intLowerbound = 65
                Else
            'Goto RetryPick
                    If intNum < intNums Then
                        intLowerbound = 0
                    Else
                        intLowerbound = 97
                    End If
                End If
            Case 2
                If intLow < intLows Then
                    intLowerbound = 97
                Else
            'Goto RetryPick
                    If intNum < intNums Then
                        intLowerbound = 0
                    Else
                        intLowerbound = 65
                    End If
                End If
            Case 3
                If intNum < intNums Then
                    intLowerbound = 0
                Else
            'Goto RetryPick
                    If intCap < intCaps Then
                        intLowerbound = 65
                    Else
                        intLowerbound = 97
                    End If
                End If
        End Select

        'now that we know what we've picked
    'we can set the rest of what we need
    'for the selection routine
    If intLowerbound > 10 Then
            intUpperbound = intLowerbound + 25
            If intLowerbound = 65 Then
                intCap = intCap + 1     'increment counter
            Else
                intLow = intLow + 1     'increment counter
            End If
        Else
            intUpperbound = 9
            intNum = intNum + 1     'increment counter
        End If

    'here we return a number based on the UPPERBOUND and LOWERBOUND
    'Lower = 0 and Upper = 9 for Integers
    'Lower = 65 and Upper = 90 for Uppercase
    'Lower = 97 and Upper = 122 for Lowercase
        intChar = Int((intUpperbound - intLowerbound + 1) * Rnd + intLowerbound)
        If intChar < 10 Then
            strHold$ = strHold$ & intChar   'just tack the Character on the end
        Else
            strHold$ = strHold$ & Chr$(intChar) 'just tack the converted Character on the end
        End If
    Loop While Len(strHold$) < (intCaps + intLows + intNums)    'Loop until we are done

    strScrambled = strHold$

End Function






