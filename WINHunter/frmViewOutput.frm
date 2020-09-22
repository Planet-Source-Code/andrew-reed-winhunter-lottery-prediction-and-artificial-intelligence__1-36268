VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmViewOutput 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BackTest Results"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2295
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   2295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Prediction Stats"
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   3840
      Width           =   2055
      Begin VB.Label lblXof 
         Caption         =   "1 0f :"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblHitCounts 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblXof 
         Caption         =   "3 0f :"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   405
      End
      Begin VB.Label lblXof 
         Caption         =   "4 0f :"
         Height          =   255
         Index           =   2
         Left            =   1080
         TabIndex        =   12
         Top             =   240
         Width           =   405
      End
      Begin VB.Label lblXof 
         Caption         =   "5 0f :"
         Height          =   255
         Index           =   3
         Left            =   1080
         TabIndex        =   11
         Top             =   480
         Width           =   405
      End
      Begin VB.Label lblXof 
         Caption         =   "6 0f :"
         Height          =   255
         Index           =   4
         Left            =   1080
         TabIndex        =   10
         Top             =   720
         Width           =   405
      End
      Begin VB.Label lblHitCounts 
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblHitCounts 
         Height          =   255
         Index           =   3
         Left            =   1560
         TabIndex        =   8
         Top             =   240
         Width           =   375
      End
      Begin VB.Label lblHitCounts 
         Height          =   255
         Index           =   4
         Left            =   1560
         TabIndex        =   7
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblHitCounts 
         Height          =   255
         Index           =   5
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   375
      End
      Begin VB.Label lblHitCounts 
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   5
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblXof 
         Caption         =   "2 0f :"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   405
      End
   End
   Begin VB.CheckBox chkVerbose 
      Caption         =   "Verbose"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   5085
      Value           =   1  'Checked
      Width           =   975
   End
   Begin RichTextLib.RichTextBox txtOutput 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   6376
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmViewOutput.frx":0000
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   5040
      Width           =   855
   End
End
Attribute VB_Name = "frmViewOutput"
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
Private mHitCounts(6) As Long

'set the local copy of the Stack to use here
Public Property Set Stack(ByRef cStack As clsStack)

    Set mStack = cStack

End Property

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'txtOutput.Top = Me.ScaleTop
    'txtOutput.Left = Me.ScaleLeft

End Sub

Private Sub Form_Resize()
    
    'If Me.ScaleWidth > 0 Then txtOutput.Width = Me.ScaleWidth
    'If Me.ScaleHeight > 500 Then txtOutput.Height = Me.ScaleHeight - 1000
    'Frame1.Top = txtOutput.Top + txtOutput.Height + 200
    '
    'cmdClose.Top = Me.ScaleHeight - 400
    'cmdClose.Left = Me.ScaleWidth - 1200

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mStack = Nothing

End Sub

Private Sub mStack_Complete(ByVal sKey As String)

    If chkVerbose.Value = 0 Then Exit Sub
    If mStack.Output.PredictDrawing > 0 Then
        If mStack.Output.PredictDrawing = 1 Then
            txtOutput.Text = ""
        End If
        With mStack.Output
            addstring = .PredictDrawing & ": " & .PredictedCount & ", " & .matchcount & " of " & mStack.Output.DrawCount & vbTab & addstring
            addstring = addstring & vbCrLf
            txtOutput.Text = addstring & txtOutput.Text
            If .matchcount > 0 Then
                mHitCounts(.matchcount) = mHitCounts(.matchcount) + 1
                lblHitCounts(.matchcount - 1).Caption = mHitCounts(.matchcount)
            End If
        End With
    End If

End Sub

Private Sub txtOutput_Change()

    If txtOutput.Text = "" Then
        lblHitCounts(0).Caption = ""
        mHitCounts(1) = 0
        lblHitCounts(1).Caption = ""
        mHitCounts(2) = 0
        lblHitCounts(2).Caption = ""
        mHitCounts(3) = 0
        lblHitCounts(3).Caption = ""
        mHitCounts(4) = 0
        lblHitCounts(4).Caption = ""
        mHitCounts(5) = 0
        lblHitCounts(5).Caption = ""
        mHitCounts(6) = 0
    End If

End Sub
