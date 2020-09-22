VERSION 5.00
Begin VB.Form frmIndividualPlot 
   Caption         =   "Individual"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3405
   ScaleWidth      =   4395
   Begin prjLotto.GraphLite GraphLite1 
      Height          =   2055
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   0
      Max             =   100
      Min             =   1
      TabIndex        =   0
      Top             =   3120
      Value           =   1
      Width           =   4335
   End
End
Attribute VB_Name = "frmIndividualPlot"
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
Private WithEvents mStack As clsStack
Attribute mStack.VB_VarHelpID = -1
Private WithEvents mFilter As clsFilter
Attribute mFilter.VB_VarHelpID = -1

'set the local copy of the drawings to use here
Public Property Set Filter(ByRef cFilter As clsFilter)

    Set mFilter = cFilter
    HScroll1.Max = mFilter.Drawings.Count - 1
    'release the mStack variable
    'otherwise
    'we will have dual plotting going on
    Set mStack = Nothing

End Property

'set the local copy of the drawings to use here
Public Property Set Stack(ByRef cStack As clsStack)

    Set mStack = cStack
    HScroll1.Max = mStack.Drawings.Count - 1
    'release the mFilter variable
    'otherwise
    'we will have dual plotting going on
    Set mFilter = Nothing

End Property

Private Sub Form_Resize()

    'If Me.WindowState <> 2 Then
        'chrtPlot.Width = Me.ScaleWidth
        GraphLite1.Width = Me.ScaleWidth
        HScroll1.Width = Me.ScaleWidth
        'chrtPlot.Height = Me.ScaleHeight - HScroll1.Height
        GraphLite1.Height = Me.ScaleHeight - HScroll1.Height
        'HScroll1.Top = chrtPlot.Height
        HScroll1.Top = GraphLite1.Height
        If GraphLite1.GotData Then
            GraphLite1.Refresh
        End If
    'End If
    
End Sub

Private Sub PlotIndividual(InArray As Variant, Optional ByVal iNext As Integer)
Dim i1low As Integer
Dim i1High As Integer
Dim iHorz() As Integer

    If frmIndividualPlot.WindowState = 1 Then Exit Sub
    'frmIndividualPlot.chrtPlot.Visible = False
    'frmIndividualPlot.chrtPlot.Repaint = False
    'frmIndividualPlot.chrtPlot.RowCount = UBound(InArray)
    i1low = 10000
    i1High = 0
    For i = 1 To UBound(InArray)
        If InArray(i) < i1low Then
            i1low = InArray(i)
        End If
        If InArray(i) > i1High Then
            i1High = InArray(i)
        End If
    Next i
    'For i = 1 To UBound(InArray)
    '    frmIndividualPlot.chrtPlot.Row = i
    '    frmIndividualPlot.chrtPlot.RowLabel = ""
    '    frmIndividualPlot.chrtPlot.Data = InArray(i)
    'Next i

    'frmIndividualPlot.chrtPlot.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = i1High + 5
    'frmIndividualPlot.chrtPlot.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = i1low - 5
    'frmIndividualPlot.chrtPlot.Repaint = True


GraphLite1.ChartType = 1
'For i = 0 To DataPoints
'   MyData(0, i) = Format$(#1/1/1998# + i, "mm/dd")
'   MyData(1, i) = (Rnd - 0.5) * 10 + i
'   MyData(2, i) = (Rnd - 0.5) * i + 10
'Next i
ReDim iHorz(UBound(InArray))
For i = 1 To UBound(iHorz)
    'set horizontal point values
    iHorz(i) = i
Next i
GraphLite1.BackColor = &HFFFFFF 'white
GraphLite1.Columns = UBound(InArray) + 1

GraphLite1.RegisterData 1, iHorz, InArray
'GraphLite1.SetSeriesOptions 0, vbBlue, "Eastern Region"
'GraphLite1.SetSeriesOptions 1, vbRed, "Western Region"
If Not mStack Is Nothing Then
    GraphLite1.Title = "Stack Results"
Else
    GraphLite1.Title = "Filter Results"
End If
GraphLite1.vLowScale = i1low - 1
GraphLite1.HighScale = i1High + 1
GraphLite1.VerticalTickInterval = 1
GraphLite1.HorizontalTickFrequency = 2
GraphLite1.PlotPoints = True

GraphLite1.Refresh










End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mStack = Nothing
    Set mFilter = Nothing

End Sub

Private Sub HScroll1_Change()

    If Not mFilter Is Nothing Then
        Me.Caption = "Filter Result: Drawing #" & HScroll1.Value
        mFilter.PredictDrawing = HScroll1.Value
        mFilter.RunFilter
    ElseIf Not mStack Is Nothing Then
        Me.Caption = "Stack Result: Drawing #" & HScroll1.Value
        mStack.PredictDrawing = HScroll1.Value
        mStack.RunStack
    End If
    

End Sub

Private Sub mFilter_Plot(ByVal larrayScores As Variant)

    PlotIndividual larrayScores
    HScroll1.Value = mFilter.PredictDrawing

End Sub

Private Sub mStack_Plot(ByVal larrayScores As Variant)

    PlotIndividual larrayScores

End Sub
