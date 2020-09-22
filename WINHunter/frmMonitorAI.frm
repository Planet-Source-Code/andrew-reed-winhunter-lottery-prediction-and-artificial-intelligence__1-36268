VERSION 5.00
Begin VB.Form frmMonitorAI 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AI Monitor"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   3390
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdStop 
      Caption         =   "Kill AI LifeForm"
      Height          =   495
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   2655
   End
   Begin VB.Label txtDecisions 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label txtSuccess 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label txtConfidence 
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblDecisionsPerSecond 
      Caption         =   "Decisions Per Second:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label lblSuccess 
      Caption         =   "Success:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label lblConfidence 
      Caption         =   "Confidence:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frmMonitorAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents mAI As clsAI
Attribute mAI.VB_VarHelpID = -1

Public Property Set Monitor(AIin As clsAI)

    Set mAI = AIin

End Property


Private Sub cmdStop_Click()

    mAI.KillLifeForm = True

End Sub

Private Sub mAI_status(ByVal lConfidenceLevel As Long, ByVal dSuccess As Double, ByVal lDecisionsPerSecond As Long)

    txtConfidence = lConfidenceLevel
    txtSuccess = dSuccess
    txtDecisions = lDecisionsPerSecond
    DoEvents

End Sub
