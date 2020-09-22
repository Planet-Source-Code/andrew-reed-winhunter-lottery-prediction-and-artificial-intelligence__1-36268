Attribute VB_Name = "modUtility"
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






