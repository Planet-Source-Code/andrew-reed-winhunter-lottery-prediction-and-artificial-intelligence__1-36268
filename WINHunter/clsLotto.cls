VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLotto"
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

Event SingleStackComplete(ByVal sHistoryName As String, ByVal sMachineKey As String, ByVal sStackKey As String)

'Public WithEvents Statistics As clsStatistics
Event FilterComplete(ByVal iCurrentdrawing As Integer)
Event Started()
Event Stopped()
Private mDrawings As Object

Public WithEvents Histories As clsHistories        'Collection of Drawing Histories
Attribute Histories.VB_VarHelpID = -1

Private Sub Histories_ThisStackComplete(ByVal sName As String, ByVal sMachineKey As String, ByVal sStackKey As String)

    'Expose the path to the outside world!
    'to the stack that has fired the -complete- event
    RaiseEvent SingleStackComplete(sName, sMachineKey, sStackKey)

End Sub

Public Property Set Drawings(ByVal vData As Object)

    Set mDrawings = vData

End Property

Public Property Get Drawings() As Object

    Set Drawings = mDrawings

End Property

Private Sub Class_Initialize()
    Set mDrawings = CreateObject(sObjDrawings)
    'Set Statistics = New clsStatistics
    Set Histories = New clsHistories
End Sub

Private Sub Class_Terminate()
    Set mDrawings = Nothing
    Set Histories = Nothing
End Sub

Public Sub RunPrediction()
Dim mHistory As clsHistory

    For Each mHistory In Histories
        Set oHistory = mHistory
        mHistory.RunHistory
    Next

End Sub

Public Function GenerateKey() As String
Dim ooHistory As clsHistory

    Set ooHistory = New clsHistory

    GenerateKey = ooHistory.GenerateKey
    
    Set ooHistory = Nothing

End Function

