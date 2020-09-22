Attribute VB_Name = "modMain"
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
Option Explicit

Public MyLotto As clsLotto    'Drawing
'Public Range As New clsRange
Public colWindows As New clsWindowCollection

Dim itemp As Integer
Dim ArrayCount As Integer

Type FormState
    Deleted As Integer
    Dirty As Integer
End Type

Public ScanForm() As New frmScan
Public ScanState() As FormState

Public sObjDrawings As String
Public sObjProcessors As String
Public sObjSelections As String
Public sObjTriggers As String

Public Sub CreateObjects()
Dim SXML        As New CGoXML
Dim i           As Integer
Dim sObjName    As String
Dim sClsName    As String

    SXML.Initialize (pavAUTO)
    'START INITIAL FILE TEMPLATE
    Call SXML.OpenFromFile(App.Path & "\plugins\plugin.xml")
    If SXML.NodeCount("/PLUGINS/PLUGIN") > 0 Then
        For i = 0 To SXML.NodeCount("/PLUGINS/PLUGIN") - 1
            'PluginXML.OpenFromString SXML.ReadNodeXML("/PLUGINS/PLUGIN[" & i & "]")
            'use late binding to create objects
            'this way, it doesnt matter when the object was created
            'or what the object's registry key is
            sObjName = SXML.ReadNode("/PLUGINS/PLUGIN[" & i & "]/OBJECT_NAME")
            sClsName = SXML.ReadNode("/PLUGINS/PLUGIN[" & i & "]/CLASS_NAME")
            Select Case SXML.ReadNode("/PLUGINS/PLUGIN[" & i & "]/TYPE")
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

Function NewMulti(ByVal FormName As String) As Object
Dim fIndex As Integer

    FormName = LCase$(FormName)
    Select Case FormName
        Case "frmscan"
            fIndex = FindFreeScanForm(FormName)
            ScanForm(fIndex).Tag = fIndex
            ScanForm(fIndex).Show
            Set NewMulti = ScanForm(fIndex)
       Case Else
            
    End Select

End Function

Function FindFreeScanForm(ByVal NameIn As String) As Integer

    If colWindows.Exist(NameIn) Then
        ArrayCount = UBound(ScanForm)
        For itemp = 0 To ArrayCount
            If ScanState(itemp).Deleted Then
                FindFreeScanForm = itemp
                ScanState(itemp).Deleted = False
                Exit Function
            End If
        Next
    Else
        ArrayCount = -1
    End If
    ReDim Preserve ScanForm(ArrayCount + 1)
    ReDim Preserve ScanState(ArrayCount + 1)
    ScanState(ArrayCount + 1).Deleted = False
    FindFreeScanForm = UBound(ScanForm)
    
End Function

Sub ActivateMulti(ByVal InForm As Form)
Dim status As Integer

    If colWindows.Exist(InForm.Name) Then
        Dim AnotherForm As Form
        Set AnotherForm = colWindows.Item(InForm.Name)
        If AnotherForm.Tag <> InForm.Tag Then
            colWindows.Remove AnotherForm
            status = colWindows.Add(InForm)
            If Not status Then
                'MsgBox "Form NOT collected!"
            End If
        End If
        Set AnotherForm = Nothing
    Else
        status = colWindows.Add(InForm)
        If Not status Then
            'MsgBox "Form NOT collected!"
        End If
    End If

End Sub

Sub UnloadMulti(ByVal InForm As Form)
        
    If colWindows.Exist(InForm.Name) Then
        Dim PossibleForm As Form
        Set PossibleForm = colWindows.Item(InForm.Name)
        If PossibleForm.Tag = InForm.Tag Then
            colWindows.Remove PossibleForm
        End If
        Set PossibleForm = Nothing
    End If

End Sub

Public Function GetStackFromTree(NodSelected As Node) As clsStack
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String

    If NodSelected.Image = "stack" Then
        sStackKey = NodSelected.Key
        sMachineKey = NodSelected.Parent.Key
        sHistoryKey = NodSelected.Parent.Parent.Key
    Else
        Exit Function
    End If
    Set MyLotto.Histories.RunningHistory = MyLotto.Histories(sHistoryKey)
    Set MyLotto.Histories(sHistoryKey).RunningMachine = MyLotto.Histories(sHistoryKey).Machines(sMachineKey)
    Set MyLotto.Histories(sHistoryKey).Machines(sMachineKey).RunningStack = MyLotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey)
    Set GetStackFromTree = MyLotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey)

End Function
