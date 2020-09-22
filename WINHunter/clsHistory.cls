VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistory"
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


'These are declared as public, since we will not allow the
'drawings to be passed back and forth, or the stacks
Event MachineStackComplete(ByVal sName As String, ByVal sMachineKey As String, ByVal sStackKey As String)

Public Machines As clsMachines      'The collection of lottery machines
Attribute Machines.VB_VarUserMemId = 0

Private mUseHistory As Boolean
Private mName As String             'full history name
Private mFileName As String         'history path/filename
Private mLoadSettings As String     'Load Settings for the history file
Private mBonusBall As Boolean       'Is this a bonus ball lottery?
Private mGrandBall As Boolean       'Does this lottery depend on
                                    'a separately drawn ball to
                                    'win the jackpot?
Private mPredict As ePredict        'predict a drawing set?  Or test the
                                    'prediction against past drawings
Private WithEvents oMachine As clsMachine
Attribute oMachine.VB_VarHelpID = -1

Public Property Set RunningMachine(cMachine As clsMachine)
    Set oMachine = cMachine
End Property

Private Sub ResetStacks()
Dim mLocalMachine As clsMachine
Dim mLocalStack As clsStack

    For Each mLocalMachine In Machines
        For Each mLocalStack In mLocalMachine.Stacks
            'Let's force a reset of the underlying objects
            mLocalStack.Reset
        Next
    Next
    Set mLocalMachine = Nothing
    Set mLocalStack = Nothing

End Sub


Public Property Get Name() As String
    Name = mName
End Property
Public Property Let Name(sName As String)
    mName = sName
End Property


Public Property Get FileName() As String
    FileName = mFileName
End Property
Public Property Let FileName(sFileName As String)
    mFileName = sFileName
End Property


Public Property Get LoadSettings() As String
    LoadSettings = mLoadSettings
End Property
Public Property Let LoadSettings(sLoadSettings As String)
    mLoadSettings = sLoadSettings
End Property


Public Property Get UseHistory() As Boolean
    UseHistory = mUseHistory
End Property
Public Property Let UseHistory(bUseHistory As Boolean)
    mUseHistory = bUseHistory
End Property


Public Property Get BonusBall() As Boolean
    BonusBall = mBonusBall
End Property
Public Property Let BonusBall(bBonusBall As Boolean)
    mBonusBall = bBonusBall
End Property


Public Property Get GrandBall() As Boolean
    GrandBall = mGrandBall
End Property
Public Property Let GrandBall(bGrandBall As Boolean)
    mGrandBall = bGrandBall
End Property

Public Sub RunHistory()
Dim mMachine As clsMachine

    'Reset the variables containing the prediction values from the previous runs
    ResetStacks
    For Each mMachine In Machines
        Set oMachine = mMachine
        mMachine.RunMachine
    Next
    Set oMachine = Nothing
    Set mMachine = Nothing

End Sub

Private Sub Class_Initialize()
    
    Set Machines = New clsMachines
    mUseHistory = True

End Sub

Private Sub Class_Terminate()
    Set oMachine = Nothing
    Set Machines = Nothing
End Sub

Private Sub WriteMachineSettings()
Dim mMachine As clsMachine
Dim i As Long
Dim lBonus As Long
Dim lJackpot As Long

    Set fs = CreateObject("Scripting.FileSystemObject")
    Dim fil1 As Object
    Set fil1 = fs.GetFile(mFileName)
    fil1.Copy (Left$(mFileName, InStr(mFileName, "."))) & "tmp"
    lJackpot = mGrandBall
    
    Open mFileName For Output As #1
    Open (Left$(mFileName, InStr(mFileName, "."))) & "tmp" For Input As #2
        Print #1, "##";
        Print #1, Machines.Count & " " & lJackpot;
        For Each mMachine In Machines
            i = i + 1
            Print #1, " (";
            Print #1, i & " " & mMachine.MachineMaximumBallNumber & " ";
            Print #1, mMachine.MachineMinimumBallNumber & " ";
            Print #1, mMachine.MachineDrawCount & " ";
            lBonus = mMachine.BonusBallMachine
            Print #1, lBonus & ")";
        Next
        Print #1, "##"
    
    
    Do While Not EOF(2)
        Z$ = ""
        Do While Not EOF(2)
            If x$ = "" Then x$ = Input(1, #2)
            If Not (Asc(x$) = 10 Or Asc(x$) = 13) Then
                Z$ = Z$ + x$
                x$ = ""
            Else
                If Asc(x$) = 13 Then
                    x$ = Input(1, #2)
                    If Asc(x$) = 10 Then x$ = ""
                    Exit Do
                End If
            End If
        Loop
        Print #1, Z$
    Loop
    Close #1
    Close #2
    Kill (Left$(mFileName, InStr(mFileName, "."))) & "tmp"
    
End Sub

Private Sub MachineBuilder(ByVal sSetup As String)
Dim lMachineCount As Long
Dim lMachineNumber As Long
Dim lMaxBalls As Long
Dim lMinBalls As Long
Dim lNumDrawn As Long
Dim bBonus As Boolean
Dim bJackpot As Boolean
Dim sTemp As String
Dim sTemp2 As String

'Sample String
'##2 (1 53 1 6 0 0) (2 53 1 6 -1 0) ... (...)##

'the setup string is as follows
'MachineCount
'jackpotmachine 0,-1

'machine number 1-?
'MaxBallNum 2-?
'MinBallNum 0-1
'NumBallsDrawn 1-?
'powerball/bonus 0,-1

    sTemp = Right$(sSetup, Len(sSetup) - 2)
    lMachineCount = Val(Left$(sTemp, InStr(sTemp, " ") - 1))
    sTemp = LTrim(Right$(sTemp, Len(sTemp) - InStr(sTemp, " ")))
    Do
        Machines.Add ("machine" & GenerateKey)
        lMachineCount = lMachineCount - 1
    Loop While lMachineCount > 0
    bJackpot = Val(Left$(sTemp, InStr(sTemp, " ") - 1))
    sTemp = LTrim(Right$(sTemp, Len(sTemp) - InStr(sTemp, " ")))
    mGrandBall = bJackpot
    Do
        sTemp2 = Mid$(sTemp, InStr(sTemp, "("), InStr(sTemp, ")"))
        sTemp = LTrim(Right$(sTemp, Len(sTemp) - Len(sTemp2)))
        sTemp2 = Mid$(sTemp2, 2, Len(sTemp2) - 2)
            lMachineNumber = Val(Left$(sTemp2, InStr(sTemp2, " ") - 1))
            sTemp2 = LTrim(Right$(sTemp2, Len(sTemp2) - Len(Left$(sTemp2, InStr(sTemp2, " ") - 1))))
            
            lMaxBalls = Val(Left$(sTemp2, InStr(sTemp2, " ") - 1))
            sTemp2 = LTrim(Right$(sTemp2, Len(sTemp2) - Len(Left$(sTemp2, InStr(sTemp2, " ") - 1))))
            Machines(lMachineNumber).MachineMaximumBallNumber = lMaxBalls
            Machines(lMachineNumber).Drawings.BallCount = lMaxBalls
            
            lMinBalls = Val(Left$(sTemp2, InStr(sTemp2, " ") - 1))
            sTemp2 = LTrim(Right$(sTemp2, Len(sTemp2) - Len(Left$(sTemp2, InStr(sTemp2, " ") - 1))))
            Machines(lMachineNumber).MachineMinimumBallNumber = lMinBalls
            
            lNumDrawn = Val(Left$(sTemp2, InStr(sTemp2, " ") - 1))
            sTemp2 = LTrim(Right$(sTemp2, Len(sTemp2) - Len(Left$(sTemp2, InStr(sTemp2, " ") - 1))))
            Machines(lMachineNumber).MachineDrawCount = lNumDrawn
            Machines(lMachineNumber).Drawings.Drawn = lNumDrawn
            
            bBonus = sTemp2
            sTemp2 = ""
            Machines(lMachineNumber).BonusBallMachine = bBonus
    Loop While Not InStr(sTemp, "##") < 3

End Sub


Public Sub Load()
Dim sDate As String
Dim sDrawing As String
Dim sDraw As String
Dim sTrash As String
Dim iSpaces As Integer
Dim mMachine As clsMachine
Dim retval As Integer
Dim bTrash As Boolean
Dim bExamined As Boolean
Dim iMachineCount As Integer
Dim lHeader As Long


    On Error GoTo ReadDrawingFileError1
    
    Open mFileName For Input As #1
    Input #1, Z$
    If Left$(Z$, 2) = "##" And Right$(Z$, 2) = "##" Then
        'We have a saved setup to the history header file
        'Pass the string to the MachineBuilder subroutine
        MachineBuilder Z$
        lHeader = 1
        Close #1
   Else
        Close #1
        If Machines.Count = 0 Then
            'Hmmm, no machine count...
            Do
                iMachineCount = Val(InputBox("How many Machines?", "Input Machine Qty", "1"))
            Loop While iMachineCount = 0
            'Ok, now lets add the appropriate machine count
            Do
                Machines.Add ("machine" & GenerateKey)
                iMachineCount = iMachineCount - 1
            Loop While iMachineCount > 0
        End If
        'clear the collections
        For Each mMachine In Machines
            mMachine.ResetDrawings
            If mMachine.MachineMaximumBallNumber = 0 Then
                Do While mMachine.MachineMaximumBallNumber = 0
                    'get the ball count for each machine
                    mMachine.MachineMaximumBallNumber = Val(InputBox("How many balls in this Machine?", "Input Ball Qty", "53"))
                    mMachine.Drawings.BallCount = mMachine.MachineMaximumBallNumber
                Loop
                'what is the first ball?
                If mMachine.MachineMaximumBallNumber > 9 And mMachine.MachineMinimumBallNumber = 0 Then
                    retval = MsgBox("Is the first ball a #1 ball?", vbYesNo, "First Ball Number")
                    If retval = vbYes Then
                        mMachine.MachineMinimumBallNumber = 1
                    End If
                End If
                Do While mMachine.MachineDrawCount = 0
                    'get the ball count for each machine
                    mMachine.MachineDrawCount = Val(InputBox("How many draws from this Machine?", "Input Drawn Qty", "6"))
                    mMachine.Drawings.Drawn = mMachine.MachineDrawCount
                Loop
            End If
        Next
        If Machines(Machines.Count).MachineDrawCount = 1 And Machines.Count > 1 Then
            retval = MsgBox("Is this a power/bonus lottery?", vbYesNo, "Power/Bonus Ball Drawing")
            If retval = vbYes Then
                'Set the last machine to the bonus ball machine
                mBonusBall = True
                Machines(Machines.Count).BonusBallMachine = True
                retval = MsgBox("Is the bonus ball required to win the Jackpot?", vbYesNo, "Jackpot Ball")
                If retval = vbYes Then
                    mGrandBall = True
                End If
            End If
        End If
        WriteMachineSettings
        
    End If
    
    
    'iSpaces = Val(txtSpaces)
    Open mFileName For Input As #1
    
    'if we have a header, then remove it here
    If lHeader > 0 Then
        For i = 1 To lHeader
            If Not EOF(1) Then
                Input #1, Z$
            Else
                MsgBox "Error: End Of File Reached!"
                Exit Sub
            End If
        Next i
    End If
    
    
    x$ = ""
    Do While Not EOF(1)
        
        Z$ = ""
        Do While Not EOF(1)
            If x$ = "" Then x$ = Input(1, #1)
            If Not (Asc(x$) = 10 Or Asc(x$) = 13) Then
                If Asc(x$) = 9 Then
                    x$ = " "
                End If
                Z$ = Z$ + x$
                x$ = ""
            Else
                If Asc(x$) = 13 Then
                    x$ = Input(1, #1)
                    If Asc(x$) = 10 Then x$ = ""
                    Exit Do
                End If
            End If
        Loop
        
        'first, let's trim off the leading spaces
        Z$ = RTrim(LTrim(Z$))
        If Z$ = "" Then GoTo skipload
        
        
        'Analyize the line
        If Not bExamined Then
            bExamined = ExamineLine(Z$, bTrash, iSpaces)
        End If
        
        'next, let's see what comes first
        If bTrash Then
            'Trash/Date/Drawing method used, ok
            'first we must delete the trash
            sTrash = Get_Trash(Z$)
            If IsDate(sTrash) Then
                sDate = sTrash
                If iSpaces > 1 Then
                    sTrash = Get_Spaced_Item(Z$, iSpaces - 1)
                Else
                    sTrash = Get_Trash(Z$)
                End If
            Else
                sDate = Get_Spaced_Item(Z$, iSpaces)
            End If
            'sDraw = Get_Item(Z$)
            sDrawing = Z$
            'If Len(sDraw) < 3 Then
            '    'uhoh, got a space delimited drawing here...
            '    If Len(sDraw) = 1 Or Len(sDraw) = 2 Then
            '        sDrawing = sDraw
            '        Do
            '            sDraw = Get_Item(Z$)
            '            If Len(sDraw) > 2 Then Exit Do
            '            sDrawing = sDrawing & " " & sDraw
            '        Loop While Not Z$ = ""
            '    Else
            '        sDrawing = sDraw
            '    End If
            'Else
            '    sDrawing = sDraw
            'End If
        Else
            'Date/Drawing method used, ok
            'first we must get the date
            sDate = Get_Spaced_Item(Z$, iSpaces)
            'sDrawing = Get_Item(Z$)
            sDrawing = Z$
        End If
        
        
        If Len(sDrawing) = Machines.Count And InStr(sDrawing, " ") = 0 Then
            'we must have a pick3 or cash4 drawing here
            'so we will have to split it up
            sDrawing = SplitDrawing(sDrawing, Machines.Count)
        End If
        
        'load drawing into collection
        For Each mMachine In Machines
            If mMachine.Drawings.Drawn = 0 Then
                mMachine.Drawings.Drawn = mMachine.MachineDrawCount
            End If
            If mMachine.Drawings.Add(sDate, sDrawing) Is Nothing Then
                GoTo ReadDrawingFileError1
            End If
            mMachine.Drawings.MaximumBallNumber = mMachine.MachineMaximumBallNumber
        Next
        'show info
        DoEvents
skipload:
                
    Loop
    'close file
    Close #1
    
    'load drawing count
    'Set MyLotto.Statistics.Drawings = MyLotto.Drawings
    'MsgBox "Load Complete"
    Exit Sub
    
ReadDrawingFileError1:

    Close #1
    MsgBox "Error Reading Drawing File."
    'Resume
    MsgBox "Please check the file and try again."
    Err.Clear

End Sub

Private Function SplitDrawing(sDrawingIn, iSplit As Integer) As String
Dim sTemp As String
Dim itemp As Integer

    itemp = iSplit
    Do While Not itemp = 0
        itemp = itemp - 1
        sTemp = sTemp & Mid$(sDrawingIn, iSplit - itemp, 1) & " "
    Loop
    SplitDrawing = RTrim(sTemp)

End Function

Private Function ExamineLine(sIn As String, bTrash As Boolean, iSpacesInDate As Integer) As Boolean
Dim sTemp As String
Dim iLenTemp As Integer
Dim sTempDate As String
Dim iShift As Integer
Dim iShift1 As Integer
Dim bDay As Boolean
Dim bMonth As Boolean
Dim bYear As Boolean
Dim bWeekday As Boolean

    sTemp = Left$(sIn, InStr(sIn, " ") - 1)
    If IsDate(sTemp) Then
        'is this a true date?
        bTrash = False
        iSpaces = 0
        iShift = Len(sTemp)
        sTemp = LTrim(Right$(sIn, Len(sIn) - InStr(sIn, " ")))
        If InStr(sTemp, " ") > 0 Then
            sTemp = Left$(sTemp, InStr(sTemp, " ") - 1)
        Else
        
        End If
        If IsWeekday(sTemp) Then
            iSpacesInDate = 1
            sTemp = LTrim(Right$(sIn, Len(sIn) - (InStr(sIn, sTemp) + Len(sTemp))))
            sTemp = Left$(sTemp, InStr(sTemp, " ") - 1)
            bTrash = True
            If Not IsNumeric(sTemp) Then
                iSpacesInDate = iSpacesInDate + 1
                'bTrash = True
            End If
        'ElseIf IsNumeric(sTemp) And Val(sTemp) > 31 Then
        
        End If
    Else
        bTrash = True
        iShift = 1
        'now let's scan for the entire date....
        If IsNumeric(sTemp) And Val(sTemp) > 31 Then
            'ok, this must not be part of the date...
            iShift = Len(sTemp) + 2
        End If
        
        iShift1 = 1
        Do
            Do
                sTemp = Mid$(sIn, iShift, iShift1)
                iShift1 = iShift1 + 1
            Loop Until Right$(sTemp, 1) = " "
            iLenTemp = Len(sTemp)
            sTemp = RTrim(LTrim(sTemp))
            If InStr(sTemp, ",") > 0 Then
                sTemp = Left$(sTemp, InStr(sTemp, ",") - 1)
            End If
            If IsNumeric(sTemp) Then
                If bDay Then
                    bYear = True
                    iShift = iShift + iLenTemp
                    sTempDate = sTempDate & " " & sTemp
                    iShift1 = 1
                Else
                    bDay = True
                    iShift = iShift + iLenTemp
                    sTempDate = sTempDate & " " & sTemp
                    iShift1 = 1
                End If
            Else
                If IsWeekday(sTemp) Then
                    iShift = iShift + iLenTemp
                    bWeekday = True
                    iShift1 = 1
                Else
                    If sTemp <> "" Then
                        'might be a month here...
                        bMonth = True
                        iShift = iShift + iLenTemp
                        sTempDate = sTempDate & " " & sTemp
                        iShift1 = 1
                    End If
                End If
            End If
            sTempDate = LTrim(sTempDate)
        Loop While Not IsDate(sTempDate) Or bYear = False
        If bMonth And bDay And bYear Then iSpacesInDate = 2
        If bWeekday And bMonth And bDay And bYear Then iSpacesInDate = 3
    End If
    

    ExamineLine = True

End Function

Private Function IsWeekday(sText As String) As Boolean
Dim sString As String

    If Len(sText) > 2 Then
        sString = Left$(sText, 3)
        Select Case LCase(sString)
            Case "sat", "sun", "mon", "tue", "wed", "thu", "fri"
                IsWeekday = True
        End Select
    End If

End Function


Private Function Get_Spaced_Item(sFull As String, ByVal iSpaceCount As Integer) As String
Dim sItem As String
Dim iCount As Integer
Dim bFoundSpace As Boolean

    If iSpaceCount > 0 Then
        iCount = 0
        Do While iSpaceCount > -1
            iCount = iCount + 1
            If Mid$(sFull$, iCount, 1) = " " And Not bFoundSpace Then
                iSpaceCount = iSpaceCount - 1
                bFoundSpace = True
            ElseIf Mid$(sFull$, iCount, 1) = " " And bFoundSpace Then
                bFoundSpace = True
            Else
                bFoundSpace = False
            End If
        Loop
        iCount = iCount - 1
        sItem = Left$(sFull, iCount)
        sFull = LTrim(Right$(sFull, Len(sFull) - iCount))
    Else
        sItem = Left$(sFull, InStr(sFull, " ") - 1)
        sFull = LTrim(Right$(sFull, Len(sFull) - Len(sItem)))
    End If
    Get_Spaced_Item = sItem

End Function

Private Function Get_Trash(sFull As String) As String
Dim sItem As String
    
    sItem = Left$(sFull, InStr(sFull, " ") - 1)
    sFull = LTrim(Right$(sFull, Len(sFull) - Len(sItem)))
    Get_Trash = sItem

End Function

Private Function Get_Item(sFull As String) As String
Dim sItem As String
Dim iCount As Integer
Dim iSpaceCount As Integer
Dim bSpaceFound As Boolean

    If InStr(sFull, " ") > 0 Then
        sItem = Left$(sFull, InStr(sFull, " ") - 1)
        sFull = sFull & " "
        If Len(sItem) < 3 Then
            'drawings are spaced out here
            iCount = 0
            Do While iSpaceCount < 7
                iCount = iCount + 1
                If Mid$(sFull$, iCount, 1) = " " Then
                    If Not bSpaceFound Then iSpaceCount = iSpaceCount + 1
                    bSpaceFound = True
                Else
                    bSpaceFound = False
                End If
                If iCount = Len(sFull$) Then Exit Do
            Loop
            iCount = iCount - 1
            sItem = Left$(sFull, iCount)
            sFull = LTrim(Right$(sFull, Len(sFull) - iCount))
            Get_Item = sItem
        Else
            sFull = LTrim(Right$(sFull, Len(sFull) - Len(sItem)))
            Get_Item = sItem
        End If
    Else
        sFull = LTrim(Right$(sFull, Len(sFull) - Len(sItem)))
        Get_Item = sFull
    End If

End Function


Private Sub oMachine_StackComplete(ByVal sKey As String, ByVal sStackKey As String)

    'We are raising the stack completed event so that
    'outside sources can process the data however it wants to
    RaiseEvent MachineStackComplete(mName, sKey, sStackKey)

End Sub

Public Function GenerateKey() As String

    GenerateKey = strScrambled(0, 3, 2)

End Function



