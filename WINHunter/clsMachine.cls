VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMachine"
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
Event StackComplete(ByVal sKey As String, ByVal sStackKey As String)

Public Enum ePredict
    PastDrawings = 1
    FutureDrawing = 0
End Enum

Public Drawings As Object      'The actual loaded drawings
Attribute Drawings.VB_VarUserMemId = 0
Public Stacks As clsStacks          'the filter stacks to run
Private mMachineMinimumBallNumber As Integer  'lowest ball number
Private mMachineMaximumBallNumber As Integer  'highest ball number
Private mMachineDrawCount As Integer  'how many balls get drawn?

Private mBonusBallMachine As Boolean  'Is this the bonus ball machine?
Private mKey As String                'Collection Key
Private mName As String         'full stack name

Private mPredictDrawing As Long
Private mPredict As ePredict           'predict a drawing set?  Or test the
                                        'prediction against past drawings
Private WithEvents oStack As clsStack
Attribute oStack.VB_VarHelpID = -1

Public WithEvents AILifeForm As clsAI
Attribute AILifeForm.VB_VarHelpID = -1


Public Sub StartAILifeForm()

    If Not AILifeForm.LifeSign Then
        Set AILifeForm.Drawings = Drawings
        AILifeForm.BeginLife
    End If

End Sub


Public Property Set RunningStack(cStack As clsStack)
    Set oStack = cStack
End Property

Public Property Let PredictDrawing(lValue As Long)
    mPredictDrawing = lValue
End Property


Public Sub ResetDrawings()

    Drawings.Clear
    Set Stacks.Drawings = Drawings

End Sub

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


Public Property Get MachineDrawCount() As Integer
    MachineDrawCount = mMachineDrawCount
End Property
Public Property Let MachineDrawCount(iMachineDrawCount As Integer)
    mMachineDrawCount = iMachineDrawCount
End Property


Public Property Get MachineMaximumBallNumber() As Integer
    MachineMaximumBallNumber = mMachineMaximumBallNumber
End Property
Public Property Let MachineMaximumBallNumber(iMachineMaximumBallNumber As Integer)
    mMachineMaximumBallNumber = iMachineMaximumBallNumber
    Drawings.MinimumBallNumber = iMachineMaximumBallNumber
End Property


Public Property Get MachineMinimumBallNumber() As Integer
    MachineMinimumBallNumber = mMachineMinimumBallNumber
End Property
Public Property Let MachineMinimumBallNumber(iMachineMinimumBallNumber As Integer)
    mMachineMinimumBallNumber = iMachineMinimumBallNumber
    Drawings.MinimumBallNumber = iMachineMinimumBallNumber
End Property


Public Property Get BonusBallMachine() As Boolean
    BonusBallMachine = mBonusBallMachine
End Property
Public Property Let BonusBallMachine(bBonusBallMachine As Boolean)
    mBonusBallMachine = bBonusBallMachine
End Property

Private Sub Class_Initialize()
    Set AILifeForm = New clsAI
    'Create new instances of the classes
    Set Drawings = CreateObject(sObjDrawings)
    Set Stacks = New clsStacks
    'link the drwaings instance to the stacks
    Set Stacks.Drawings = Drawings
End Sub

Private Sub Class_Terminate()
    Set Drawings = Nothing
    Set Stacks = Nothing
    Set oStack = Nothing
    Set AILifeForm = Nothing
End Sub

Public Sub RunMachine()
Dim mStack As clsStack
Dim lDrawingToPredict As Long
    
    ReDim mExclude(Drawings.BallCount)
    If mPredict = PastDrawings Then
        'loop through all the drawings
        For lDrawingToPredict = 1 To Drawings.Count
            For Each mStack In Stacks
                Set oStack = mStack
                mStack.PredictDrawing = lDrawingToPredict
                mStack.RunStack
                'THEN reset the stack
                mStack.Reset
            Next
        Next
    Else
        For Each mStack In Stacks
            Set oStack = mStack
            mStack.PredictDrawing = 0
            mStack.RunStack
            'THEN reset the stack
            mStack.Reset
        Next
    End If
    Set oStack = Nothing
    Set mStack = Nothing

End Sub

Private Sub oStack_Complete(ByVal sKey As String)

    RaiseEvent StackComplete(mKey, sKey)

End Sub

