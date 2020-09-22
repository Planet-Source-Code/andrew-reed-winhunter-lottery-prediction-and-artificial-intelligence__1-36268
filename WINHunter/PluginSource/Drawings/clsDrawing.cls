VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private mNumbers() As Long
Private mDeltas() As Long
Private mDrawOrder As String
Private mDrawnOn As String
Private iBallNum As Long
Private iDrawCount As Long
Private mIndexNumber As Double
Private mMaxOdds As Double
Private mOddCount As Long

Public Property Get HitIndex() As Double

    HitIndex = mIndexNumber

End Property

Public Property Get MaxOdds() As Double

    MaxOdds = mMaxOdds

End Property


Private Function Combine(ByVal n As Long, ByVal M As Long) As Double
Rem *** This function calculates the total number of combinations ***
Rem *** 'N taken M' at a time
Dim FA As Double
Dim VA As Double

    If M > n Then
        Combine = 0
        Exit Function
    End If
    
    VA = 1
    For j = (n - M + 1) To n
        VA = VA * j
    Next j
    
    FA = 1
    For i = 1 To M
        FA = FA * i
    Next i
    
    Combine = VA / FA

End Function

Public Sub CalculateIndex(ByVal lBallCount As Long)
Dim dIndex As Double

    mMaxOdds = Combine(lBallCount, UBound(mNumbers) + 1)
    dIndex = mMaxOdds
    For i = 1 To (UBound(mNumbers) + 1)
        dIndex = dIndex - Combine(lBallCount - mNumbers(i - 1), (UBound(mNumbers) + 2) - i)
    Next i
    mIndexNumber = dIndex

End Sub



Public Property Get WasDrawn(ByVal iBall As Integer)

    For iBallNum = 0 To iDrawCount - 1
        If iBall < mNumbers(iBallNum) Then
            Exit For
        End If
        If iBall = mNumbers(iBallNum) Then
            WasDrawn = True
            Exit Property
        End If
    Next
    WasDrawn = False

End Property

Public Property Get Numbers(Optional ByVal vntIndexKey As Variant) As Variant
Attribute Numbers.VB_UserMemId = 0

    If IsMissing(vntIndexKey) Then
        Numbers = mNumbers
    Else
        Numbers = mNumbers(vntIndexKey)
    End If

End Property

Public Property Get Deltas(Optional ByVal vntIndexKey As Variant) As Variant

    If IsMissing(vntIndexKey) Then
        Deltas = mDeltas
    Else
        Deltas = mDeltas(vntIndexKey)
    End If

End Property

Public Property Get DrawnOrder() As String

    DrawnOrder = mDrawOrder

End Property

Public Property Get OddCount() As Long

    OddCount = mOddCount

End Property

Public Function Load_DrawnOrder(mOrder As String, ByVal iDrawn As Integer) As Integer
Dim sDrawNumber As String
Dim sDelimiter As String
Dim iUnordered() As Integer
Dim GOo() As Integer
Dim i As Integer
Dim j As Integer
Dim lDelta As Long


    On Error GoTo LoadDrawError
    
    mOrder = LTrim(mOrder)
    mDrawOrder = mOrder
    'mOrder = mDrawOrder
    'iDrawCount = 0
    
    ReDim iUnordered(iDrawn - 1)
    ReDim GOo(iDrawn - 1)
    ReDim mNumbers(iDrawn - 1)
    ReDim mDeltas(iDrawn - 1)
    
    Do While mOrder <> ""
        sDelimiter = "-"
        sDrawNumber = ""
        mOrder = LTrim(mOrder)
        If InStr(mOrder, sDelimiter) > 0 And InStr(mOrder, sDelimiter) < Len(mOrder) Then
            sDrawNumber = Left$(mOrder, InStr(mOrder, sDelimiter) - 1)
        End If
        If Not IsNumeric(sDrawNumber) Then
            sDelimiter = " "
            If InStr(mOrder, sDelimiter) > 0 And InStr(mOrder, sDelimiter) < Len(mOrder) Then
                sDrawNumber = Left$(mOrder, InStr(mOrder, sDelimiter) - 1)
            End If
        End If
        'If sDrawNumber = "" And IsNumeric(mOrder) Then
        '    sDrawNumber = mOrder
        If IsNumeric(sDrawNumber) Then
            iUnordered(iDrawCount) = sDrawNumber
            mOrder = LTrim(Right$(mOrder, Len(mOrder) - (Len(sDrawNumber) + 1)))
        ElseIf sDrawNumber = "" Then
            If IsNumeric(mOrder) And Len(mOrder) < 3 Then
                sDrawNumber = mOrder
                iUnordered(iDrawCount) = sDrawNumber
                mOrder = ""
            End If
        End If
        'If Len(mOrder) > 2 Then
        '    sDrawNumber = Left$(mOrder, InStr(mOrder, sDelimiter) - 1)
        '    iUnordered(iDrawCount) = sDrawNumber
        '    mOrder = LTrim(Right$(mOrder, Len(mOrder) - (Len(sDrawNumber) + 1)))
        'Else
        '    sDrawNumber = mOrder
        '    iUnordered(iDrawCount) = sDrawNumber
        '    mOrder = ""
        'End If
        iDrawCount = iDrawCount + 1
        If iDrawCount = iDrawn Then Exit Do
    Loop
    
    For i = 0 To iDrawCount - 1
        For j = 0 To iDrawCount - 1
            If iUnordered(j) < iUnordered(i) Then GOo(i) = GOo(i) + 1
            If iUnordered(j) = iUnordered(i) Then
                If i = j Then
                    GOo(i) = GOo(i) + 1
                Else
                    MsgBox "Duplicate numbers found in drawing history file!  Please correct and re-load."
                    GoTo LoadDrawError
                End If
            End If
        Next j
    Next i
    For i = 0 To iDrawCount - 1
        For j = 0 To iDrawCount - 1
            If GOo(j) = i + 1 Then
                mNumbers(i) = iUnordered(j)
                'Calculate the Delta Number
                mDeltas(i) = iUnordered(j) - lDelta
                'save the Total to calculate the next delta number
                lDelta = mNumbers(i)
                If IsOdd(mNumbers(i)) Then
                    mOddCount = mOddCount + 1
                End If
                Exit For
            End If
        Next j
    Next i
    
    Load_DrawnOrder = iDrawCount
    Exit Function
    
LoadDrawError:

    MsgBox "Error loading draws."
    mOrder = "BADLOAD"
    'Resume
    Err.Clear

End Function

Private Function IsOdd(lValue As Long) As Boolean
Dim sTemp As String

    sTemp = lValue / 2
    If InStr(sTemp, ".") > 0 Then
        IsOdd = False
        Exit Function
    End If
    IsOdd = True

End Function

Public Property Get DrawnOn() As String

    DrawnOn = mDrawnOn

End Property

Public Property Let DrawnOn(mDate As String)

    mDrawnOn = mDate

End Property

Private Sub Class_Initialize()

    ReDim mNumbers(0)

End Sub


