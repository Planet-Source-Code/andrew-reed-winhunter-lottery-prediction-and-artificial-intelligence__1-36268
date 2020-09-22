VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmScan 
   Caption         =   "Scan Previous Drawings"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   Icon            =   "frmScan.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4380
   ScaleWidth      =   6945
   Begin RichTextLib.RichTextBox rtbOverview 
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1720
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmScan.frx":030A
   End
   Begin MSComctlLib.ListView lvMatches 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   1680
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "matched"
         Text            =   "Matched"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "index"
         Text            =   "Index"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "deltas"
         Text            =   "Deltas"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "results"
         Text            =   "Results"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "date"
         Text            =   "Date"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "indexsort"
         Text            =   "True Draw Index Sort"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Key             =   "datesort"
         Text            =   "True Date Sort"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Key             =   "matchsort"
         Text            =   "True Match Sort"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox txtInput 
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   600
      Width           =   4935
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdScan 
      Caption         =   "Scan"
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   3960
      Width           =   975
   End
   Begin VB.Label lblExample 
      BackStyle       =   0  'Transparent
      Caption         =   "Example: 1,8,14,23"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   8
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblResults 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Match Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   1200
      Width           =   5055
   End
   Begin VB.Label lblOverview 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Match Overview"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label lblInput 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Numbers to Scan for:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmScan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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
Dim mDrawings As Object
Dim mHistory As clsHistory
Dim mScanFor() As Long
Dim mLastIndexClicked As Long



'set the local copy of the drawings to use here
Public Property Set History(ByVal vData As clsHistory)

    Set mHistory = vData

End Property

Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdScan_Click()
Dim lMatchCount As Long
Dim lvItem As ListItem
Dim sInput As String
Dim lMatchCounts() As Long
Dim lMachineMatchCounts() As Long
Dim oMachine As clsMachine
Dim lDigit As Long
Dim lMachineNumber As Long
Dim sTemp As String
Dim sMatchIndex As String
Dim dBuiltIndex As Double
Dim lMaxMatchCounts As Long

    lvMatches.ListItems.Clear
    If txtInput.Text = "" Then
        Get_Numbers mHistory.Machines(1).Drawings.MinimumBallNumber & "-" & mHistory.Machines(1).Drawings.MaximumBallNumber
    Else
        Get_Numbers txtInput.Text
    End If
    rtbOverview.Text = ""
    For Each oMachine In mHistory.Machines
        If lMatchCount < oMachine.MachineDrawCount Then
            lMatchCount = oMachine.MachineDrawCount
        End If
    Next
    
    'make an array for the match counts
    ReDim lMatchCounts(mHistory.Machines.Count, mHistory.Machines(1).Drawings.Drawn)
    ReDim lMachineMatchCounts(mHistory.Machines.Count - 1, mHistory.Machines(1).Drawings.Count)
    
    Screen.MousePointer = 11
    'Loop through the drawings available
    For i = 1 To mHistory.Machines(1).Drawings.Count
        'reset the count
        lMachineNumber = 0
        'Give the User some feedback
        frmMain.sbrInfo.Panels("status").Text = "Scanning " & i & " of " & mHistory.Machines(1).Drawings.Count
        DoEvents
        
        'reset the count
        lMatchCount = 0
        'loop through the machines
        For Each oMachine In mHistory.Machines
            Set mDrawings = oMachine.Drawings
            'loop through the draws
            For k = 0 To oMachine.MachineDrawCount - 1
                'loop through the user input numbers
                For j = 0 To UBound(mScanFor)
                    'get the digit from the draw
                    lDigit = Val(mDrawings.Item(i).Numbers(k))
                    If lDigit = mScanFor(j) Then
                        'increment the overall count
                        lMatchCount = lMatchCount + 1
                        lMachineMatchCounts(lMachineNumber, i) = lMachineMatchCounts(lMachineNumber, i) + 1
                    End If
                Next j
            Next k
            'increment the individual count
            lMatchCounts(lMachineNumber, lMatchCount) = lMatchCounts(lMachineNumber, lMatchCount) + 1
            'increment the machine number
            lMachineNumber = lMachineNumber + 1
        Next    'Get the next machine
        Set mDrawings = mHistory.Machines(1).Drawings
        
        If lMatchCount > lMaxMatchCounts Then
            lMaxMatchCounts = lMatchCount
        End If
        
        If lMatchCount > 2 Then
            Select Case mHistory.Machines.Count
                Case 1
                    dBuiltIndex = mHistory.Machines(1).Drawings.Item(i).HitIndex
                Case 2
                    dBuiltIndex = mHistory.Machines(1).Drawings.Item(i).HitIndex + _
                    (mHistory.Machines(1).Drawings.Item(i).MaxOdds * _
                    (mHistory.Machines(2).Drawings.Item(i).HitIndex - 1))
                Case 3
                    dBuiltIndex = mHistory.Machines(1).Drawings.Item(i).HitIndex + _
                    (mHistory.Machines(1).Drawings.Item(i).MaxOdds * _
                    (mHistory.Machines(2).Drawings.Item(i).HitIndex - 1)) + _
                    ((mHistory.Machines(1).Drawings.Item(i).MaxOdds * _
                    mHistory.Machines(2).Drawings.Item(i).MaxOdds) * _
                    mHistory.Machines(3).Drawings.Item(i).HitIndex - 1)
                Case 4
                    dBuiltIndex = mHistory.Machines(1).Drawings.Item(i).HitIndex + _
                    (mHistory.Machines(1).Drawings.Item(i).MaxOdds * _
                    (mHistory.Machines(2).Drawings.Item(i).HitIndex - 1)) + _
                    ((mHistory.Machines(1).Drawings.Item(i).MaxOdds * _
                    mHistory.Machines(2).Drawings.Item(i).MaxOdds) * _
                    mHistory.Machines(3).Drawings.Item(i).HitIndex - 1) + _
                    ((mHistory.Machines(1).Drawings.Item(i).MaxOdds * _
                    mHistory.Machines(2).Drawings.Item(i).MaxOdds * _
                    mHistory.Machines(3).Drawings.Item(i).MaxOdds) * _
                    mHistory.Machines(4).Drawings.Item(i).HitIndex - 1)
                Case Else
                    MsgBox "WinHunter does not currently support greater than" & vbCrLf & _
                    " 4 machines for calculating index numbers."
            End Select
            
            sTemp = ""
            sMatchIndex = ""
            For j = 0 To mHistory.Machines.Count - 1
                'Calculate the index for multiple machines
                'This is based on the idea that:
                'Index 1st machine + ((max odds 1st machine) * index 2nd machine-1)
                'dBuiltIndex = dBuiltIndex * mHistory.Machines(j + 1).Drawings.Item(i).HitIndex
                
                'build the match string here
                If sTemp <> "" Then sTemp = sTemp & ", "
                sMatchIndex = sMatchIndex & lMachineMatchCounts(j, i)
                sTemp = sTemp & lMachineMatchCounts(j, i) & " of " & mHistory.Machines(j + 1).MachineDrawCount
            Next j
            Set lvItem = lvMatches.ListItems.Add(, "item" & lvMatches.ListItems.Count + 1, sTemp)
            lvItem.SubItems(1) = Format$(dBuiltIndex, "###,###,###,###")
            lvItem.SubItems(2) = Build_Drawing(mDrawings.Item(i).Deltas)
            lvItem.SubItems(3) = mDrawings.Item(i).DrawnOrder
            lvItem.SubItems(4) = mDrawings.Item(i).DrawnOn
            lvItem.SubItems(5) = Get_IndexSort(dBuiltIndex)
            lvItem.SubItems(6) = Get_IndexSort(i)
            lvItem.SubItems(7) = Get_IndexSort(Val(sMatchIndex))
            Set lvItem = Nothing
        End If
    Next i
    
    If lMaxMatchCounts > 2 Then
        For i = 3 To lMaxMatchCounts
            rtbOverview.Text = rtbOverview.Text & lMatchCounts(0, i) & ", " & i & " of " & mDrawings.Drawn & vbCrLf
        Next i
    End If
    Screen.MousePointer = 0
    DoEvents

End Sub

Private Function Build_Drawing(vArray As Variant) As String
Dim temp As String

    If IsArray(vArray) Then
        For i = 0 To UBound(vArray) - 1
            temp = temp & Format$(vArray(i), "0#") & "-"
        Next i
        Build_Drawing = temp & Format$(vArray(i), "0#")
    End If

End Function

Private Function Get_IndexSort(ByVal lIndexVal As Long) As String
Dim sTemp As String
Dim sOutput As String

    sTemp = LTrim(Str(lIndexVal))
    'Do
    '    Select Case Right$(sTemp, Len(sTemp) - (Len(sTemp) - 1))
    '        Case "0"
    '            sOutput = sOutput & "A"
    '        Case "1"
    '            sOutput = sOutput & "B"
    '        Case "2"
    '            sOutput = sOutput & "C"
    '        Case "3"
    '            sOutput = sOutput & "D"
    '        Case "4"
    '            sOutput = sOutput & "E"
    '        Case "5"
    '            sOutput = sOutput & "F"
    '        Case "6"
    '            sOutput = sOutput & "G"
    '        Case "7"
    '            sOutput = sOutput & "H"
    '        Case "8"
    '            sOutput = sOutput & "I"
    '        Case "9"
    '            sOutput = sOutput & "J"
    '    End Select
    '    sTemp = Left$(sTemp, Len(sTemp) - 1)
    'Loop While Len(sTemp) > 0
    
    'Add Characters to make the sort balanced
    'Do
    '    sOutput = sOutput & "A"
    'Loop Until Len(sOutput) = 13
    
    sOutput = sTemp
    Do
        sOutput = "0" & sOutput
    Loop Until Len(sOutput) = 13
    
    'Return the sort string
    Get_IndexSort = sOutput

End Function

Private Sub Form_Activate()

    ActivateMulti Me

End Sub

Private Sub Form_Resize()
    
    If Me.ScaleWidth < 2000 Then Exit Sub
    If Me.ScaleHeight < 2140 Then Exit Sub
    lvMatches.Height = Me.ScaleHeight - 2020
    'rtbOverview.Height = lvMatches.Height
    lvMatches.Width = Me.ScaleWidth - 300
    cmdScan.Top = Me.ScaleHeight - 400
    cmdScan.Left = Me.ScaleWidth - 2340
    cmdExit.Top = Me.ScaleHeight - 400
    cmdExit.Left = Me.ScaleWidth - 1140

End Sub

Private Sub Form_Unload(Cancel As Integer)

    ScanState(Me.Tag).Deleted = True
    UnloadMulti Me

End Sub

Private Sub Get_Numbers(ByVal sText As String)
Dim iUseCount As Integer
Dim sUse As String
Dim sNumber As String
Dim sNumber2 As String
Dim lNumber As Long
Dim sTemp As String

    sUse = sText
    
    Do While Len(sUse) > 0
        ReDim Preserve mScanFor(iUseCount)
        If InStr(sUse, ",") > 0 Then
            sNumber = Left$(sUse, InStr(sUse, ",") - 1)
        Else
            sNumber = sUse
        End If
        
        If Not IsNumeric(sNumber) Then
            If InStr(sNumber, "-") > 0 Then
                sNumber2 = Right$(sNumber, Len(sNumber) - (InStr(sUse, "-")))
                'reset sUse
                If Len(sUse) > Len(sNumber) Then
                    sUse = Right(sUse, Len(sUse) - (Len(sNumber) + 1))
                Else
                    sUse = ""
                End If
                sNumber = Left$(sNumber, InStr(sNumber, "-") - 1)
                lNumber = Val(sNumber)
                'Besure to skip the current sNumber
                sTemp = ""
                For i = lNumber To Val(sNumber2)
                    If sTemp <> "" Then sTemp = sTemp & ","
                    sTemp = sTemp & i
                Next i
                If sUse <> "" Then
                    sUse = sTemp & "," & sUse
                Else
                    sUse = sTemp
                End If
            Else
                Exit Sub
            End If
        End If
        
        lNumber = RTrim(LTrim(sNumber))
        If lNumber = 0 Then
            ReDim Preserve mScanFor(mHistory.Machines(1).MachineMaximumBallNumber)
            For i = 1 To UBound(mScanFor)
                mScanFor(i) = i
            Next i
            Exit Sub
        End If
        For i = 0 To UBound(mScanFor)
             'if the user typed a duplicate,
             'then we exit without adding
             If mScanFor(i) = lNumber Then Exit For
             'if we are here at the end of the loop,
             'then we haven't found a match
             'so let's keep the number!
             If i = UBound(mScanFor) Then
                 mScanFor(iUseCount) = lNumber
             End If
        Next i
        
        If Len(sUse) > Len(sNumber) Then
            sUse = Right(sUse, Len(sUse) - (Len(sNumber) + 1))
        Else
            sUse = ""
        End If
        iUseCount = iUseCount + 1
    Loop

End Sub

Private Sub lvMatches_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    ' When a ColumnHeader object is clicked, the ListView control is
    ' sorted by the subitems of that column.
    ' Set the SortKey to the Index of the ColumnHeader - 1
    Select Case ColumnHeader.Index
        Case 1
            lvMatches.SortKey = ColumnHeader.Index + 6
        Case 2
            lvMatches.SortKey = ColumnHeader.Index + 3
            'lvMatches.SortOrder = lvwDescending
        Case 5
            lvMatches.SortKey = ColumnHeader.Index + 1
        Case Else
            lvMatches.SortKey = ColumnHeader.Index - 1
            'lvMatches.SortOrder = lvwAscending
    End Select
    If mLastIndexClicked = ColumnHeader.Index Then
        'User clicked the same header twice
        'so now let's switch the sort order
        Select Case lvMatches.SortOrder
            Case lvwDescending
                lvMatches.SortOrder = lvwAscending
            Case lvwAscending
                lvMatches.SortOrder = lvwDescending
        End Select
    Else
        mLastIndexClicked = ColumnHeader.Index
    End If
    ' Set Sorted to True to sort the list.
    lvMatches.Sorted = True

End Sub
