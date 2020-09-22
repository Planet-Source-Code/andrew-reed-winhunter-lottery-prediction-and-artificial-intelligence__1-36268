VERSION 5.00
Begin VB.Form frmViewSettings 
   Caption         =   "Form1"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   4575
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   2520
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1320
      TabIndex        =   1
      Top             =   2520
      Width           =   975
   End
   Begin prjLotto.MultiView mvSettings 
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   4048
   End
End
Attribute VB_Name = "frmViewSettings"
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
Private mObject As Object


Private Sub cmdApply_Click()

    mvSettings.SaveChanges

End Sub

Private Sub cmdCancel_Click()

    Me.Hide
    mvSettings.Clear
    Unload Me

End Sub

Private Sub cmdOK_Click()

    mvSettings.SaveChanges
    Unload Me

End Sub

Private Sub Form_Load()

    mvSettings.Top = 0
    mvSettings.Left = 0
    Me.Height = 3600
    Me.Width = 3900

End Sub

Private Sub Form_Resize()

    If Me.ScaleWidth > 0 Then mvSettings.Width = Me.ScaleWidth
    If Me.ScaleHeight > 500 Then mvSettings.Height = Me.ScaleHeight - 500
    cmdApply.Top = Me.ScaleHeight - 400
    cmdApply.Left = Me.ScaleWidth - 3400
    cmdOk.Top = Me.ScaleHeight - 400
    cmdOk.Left = Me.ScaleWidth - 2200
    cmdCancel.Top = Me.ScaleHeight - 400
    cmdCancel.Left = Me.ScaleWidth - 1000
    

End Sub
Private Sub Form_Unload(Cancel As Integer)

    colWindows.Remove frmViewSettings

End Sub
