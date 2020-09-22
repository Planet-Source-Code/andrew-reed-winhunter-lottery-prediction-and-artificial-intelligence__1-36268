VERSION 5.00
Begin VB.Form frmPickFrom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Item"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.ComboBox cboSelect 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Make your selection"
      Top             =   600
      Width           =   4095
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      Caption         =   "Please Select From The Available Items:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmPickFrom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mLocalIndexer As clsIndexer

Private Sub cboSelect_Change()

    mLocalIndexer.Indexer = cboSelect.ListIndex

End Sub

Private Sub cboSelect_Click()

    mLocalIndexer.Indexer = cboSelect.ListIndex

End Sub

Public Property Set GetIndexer(clsgetIndexer As clsIndexer)

    Set mLocalIndexer = clsgetIndexer

End Property

Private Sub cboSelect_KeyDown(KeyCode As Integer, Shift As Integer)

    'Prevent Keystrokes
    KeyCode = 0

End Sub

Private Sub cboSelect_KeyPress(KeyAscii As Integer)

    'Prevent Keystrokes
    KeyAscii = 0

End Sub

Private Sub cmdCancel_Click()

    mLocalIndexer.Indexer = -1
    Unload Me

End Sub

Private Sub cmdOk_Click()

    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    'destroy local reference to the object
    Set mLocalIndexer = Nothing

End Sub
