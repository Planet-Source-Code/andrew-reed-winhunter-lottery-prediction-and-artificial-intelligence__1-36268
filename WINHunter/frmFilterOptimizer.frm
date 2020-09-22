VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmFilterOptimizer 
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8130
   ScaleWidth      =   8400
   Begin prjLotto.LED LEDReset 
      Height          =   210
      Left            =   6120
      TabIndex        =   67
      Top             =   7800
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      State           =   1
      Color           =   9
   End
   Begin prjLotto.LED LEDPause 
      Height          =   210
      Left            =   2160
      TabIndex        =   66
      Top             =   7800
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      State           =   1
      Color           =   6
   End
   Begin prjLotto.LED LEDRun 
      Height          =   210
      Left            =   3240
      TabIndex        =   64
      Top             =   7800
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
      State           =   1
      Color           =   3
   End
   Begin prjLotto.LED LEDStop 
      Height          =   210
      Left            =   5160
      TabIndex        =   63
      Top             =   7800
      Width           =   210
      _ExtentX        =   370
      _ExtentY        =   370
   End
   Begin VB.PictureBox Picture1 
      Height          =   6495
      Left            =   240
      ScaleHeight     =   6435
      ScaleWidth      =   7875
      TabIndex        =   5
      Top             =   840
      Width           =   7935
      Begin VB.Frame fraBestRatio 
         Caption         =   "Best Ratio Results"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   3960
         TabIndex        =   20
         Top             =   2760
         Width           =   3615
         Begin prjLotto.Counter cntBestRatioHits 
            Height          =   375
            Left            =   2880
            TabIndex        =   21
            Top             =   360
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntBestRatioPicks 
            Height          =   375
            Left            =   2880
            TabIndex        =   23
            Top             =   840
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntBestRatio 
            Height          =   375
            Left            =   2880
            TabIndex        =   25
            Top             =   1320
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntBestRatioHitPick 
            Height          =   375
            Left            =   2880
            TabIndex        =   27
            Top             =   1800
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin VB.Label Label10 
            Caption         =   "Hit/Pick Ratio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   28
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label9 
            Caption         =   "Ratio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   26
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label8 
            Caption         =   "Ratio Picks:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label7 
            Caption         =   "Ratio Hits:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   2175
         End
      End
      Begin VB.Frame fraResults 
         Caption         =   "Stored Best Results"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   3840
         TabIndex        =   7
         Top             =   3240
         Width           =   3855
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   375
            Left            =   2760
            TabIndex        =   9
            Top             =   2640
            Width           =   975
         End
         Begin MSComctlLib.ListView lvResults 
            Height          =   2175
            Left            =   120
            TabIndex        =   8
            Top             =   360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   3836
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin MSComDlg.CommonDialog CommonDialog1 
            Left            =   1800
            Top             =   2520
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
      Begin VB.VScrollBar vscrlPrevious 
         Height          =   3495
         Left            =   7560
         TabIndex        =   47
         Top             =   0
         Width           =   255
      End
      Begin VB.Frame fraBestJackpot 
         Caption         =   "Best Jackpot Results"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   3960
         TabIndex        =   38
         Top             =   240
         Width           =   3615
         Begin prjLotto.Counter cntBestJackpotHits 
            Height          =   375
            Left            =   2880
            TabIndex        =   39
            Top             =   360
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntBestJackpotPicks 
            Height          =   375
            Left            =   2880
            TabIndex        =   40
            Top             =   840
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntBestJackpotRatio 
            Height          =   375
            Left            =   2880
            TabIndex        =   41
            Top             =   1320
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntBestJackpotHitPick 
            Height          =   375
            Left            =   2880
            TabIndex        =   42
            Top             =   1800
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin VB.Label Label18 
            Caption         =   "Jackpot Hits:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   46
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label17 
            Caption         =   "Jackpot Picks:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label16 
            Caption         =   "Jackpot Ratio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   44
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label15 
            Caption         =   "Jackpot Hit/Pick:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   1800
            Width           =   2535
         End
      End
      Begin VB.Frame fraBestHitPick 
         Caption         =   "Best Hit/Pick Results"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   120
         TabIndex        =   29
         Top             =   3960
         Width           =   3615
         Begin prjLotto.Counter cntBestHitPickHits 
            Height          =   375
            Left            =   2880
            TabIndex        =   30
            Top             =   360
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntBestHitPickPicks 
            Height          =   375
            Left            =   2880
            TabIndex        =   31
            Top             =   840
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntBestHitPickratio 
            Height          =   375
            Left            =   2880
            TabIndex        =   32
            Top             =   1320
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntBestHitPicks 
            Height          =   375
            Left            =   2880
            TabIndex        =   33
            Top             =   1800
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin VB.Label Label14 
            Caption         =   "Hit/Pick Hits:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   37
            Top             =   360
            Width           =   2175
         End
         Begin VB.Label Label13 
            Caption         =   "Hit/Pick Picks:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   36
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label12 
            Caption         =   "Hit/Pick Ratio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   1320
            Width           =   2175
         End
         Begin VB.Label Label11 
            Caption         =   "HP Hit/Pick:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   1800
            Width           =   2175
         End
      End
      Begin VB.Frame fraStatistics 
         Caption         =   "Performance Stats"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3615
         Begin prjLotto.LED LEDProgress 
            Height          =   210
            Index           =   9
            Left            =   3225
            TabIndex        =   62
            Top             =   480
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   370
            State           =   1
            Color           =   9
         End
         Begin prjLotto.LED LEDProgress 
            Height          =   210
            Index           =   8
            Left            =   3015
            TabIndex        =   61
            Top             =   480
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   370
            State           =   1
            Color           =   3
         End
         Begin prjLotto.LED LEDProgress 
            Height          =   210
            Index           =   7
            Left            =   2805
            TabIndex        =   60
            Top             =   480
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   370
            State           =   1
            Color           =   3
         End
         Begin prjLotto.LED LEDProgress 
            Height          =   210
            Index           =   6
            Left            =   2580
            TabIndex        =   59
            Top             =   480
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   370
            State           =   1
            Color           =   6
         End
         Begin prjLotto.LED LEDProgress 
            Height          =   210
            Index           =   5
            Left            =   2355
            TabIndex        =   58
            Top             =   480
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   370
            State           =   1
            Color           =   6
         End
         Begin prjLotto.LED LEDProgress 
            Height          =   210
            Index           =   4
            Left            =   2145
            TabIndex        =   57
            Top             =   480
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   370
            State           =   1
            Color           =   6
         End
         Begin prjLotto.LED LEDProgress 
            Height          =   210
            Index           =   3
            Left            =   1935
            TabIndex        =   56
            Top             =   480
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   370
            State           =   1
         End
         Begin prjLotto.LED LEDProgress 
            Height          =   210
            Index           =   2
            Left            =   1725
            TabIndex        =   55
            Top             =   480
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   370
            State           =   1
         End
         Begin prjLotto.LED LEDProgress 
            Height          =   210
            Index           =   1
            Left            =   1515
            TabIndex        =   54
            Top             =   480
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   370
            State           =   1
         End
         Begin prjLotto.LED LEDProgress 
            Height          =   210
            Index           =   0
            Left            =   1305
            TabIndex        =   53
            Top             =   480
            Width           =   210
            _ExtentX        =   370
            _ExtentY        =   370
            State           =   1
         End
         Begin prjLotto.Counter cntPerformanceAverageHits 
            Height          =   375
            Left            =   2880
            TabIndex        =   14
            Top             =   2760
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntPerformanceRatio 
            Height          =   375
            Left            =   2880
            TabIndex        =   16
            Top             =   3240
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntPerformanceHitPickRatio 
            Height          =   375
            Left            =   2880
            TabIndex        =   18
            Top             =   3720
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntPerformanceAvgPicks 
            Height          =   375
            Left            =   2880
            TabIndex        =   48
            Top             =   840
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntPerformanceProcessorCount 
            Height          =   375
            Left            =   2880
            TabIndex        =   49
            Top             =   1320
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntPerformanceJackpotHits 
            Height          =   375
            Left            =   2880
            TabIndex        =   50
            Top             =   1800
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin prjLotto.Counter cntPerformanceJackpotPicks 
            Height          =   375
            Left            =   2880
            TabIndex        =   51
            Top             =   2280
            Width           =   630
            _ExtentX        =   1111
            _ExtentY        =   661
            Digits          =   2
         End
         Begin VB.Label lblSuccess 
            BackStyle       =   0  'Transparent
            Caption         =   "Success..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   52
            Top             =   465
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Hit/Pick Ratio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   3720
            Width           =   2655
         End
         Begin VB.Label Label5 
            Caption         =   "Performance Ratio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   3240
            Width           =   2685
         End
         Begin VB.Label Label4 
            Caption         =   "Average Hits:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   2760
            Width           =   2655
         End
         Begin VB.Label Label3 
            Caption         =   "Jackpot Picks:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   2280
            Width           =   2655
         End
         Begin VB.Label Label2 
            Caption         =   "Jackpot Hits:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   12
            Top             =   1800
            Width           =   2655
         End
         Begin VB.Label Label1 
            Caption         =   "Processor Count:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   2655
         End
         Begin VB.Label lblAveragePicks 
            Caption         =   "Average Picks:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Width           =   2175
         End
      End
      Begin MSComctlLib.TreeView tvFilterDisplay 
         Height          =   2175
         Left            =   3840
         TabIndex        =   68
         Top             =   120
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   3836
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7200
      TabIndex        =   4
      Top             =   7680
      Width           =   735
   End
   Begin MSComctlLib.TabStrip tsOverall 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   13150
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Optimizer Setup"
            Key             =   "setup"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Current Statistics"
            Key             =   "statistics"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Best Pass"
            Key             =   "previous"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Saved Results"
            Key             =   "results"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "   Reset"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   7680
      Width           =   735
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "   Stop"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   7680
      Width           =   735
   End
   Begin VB.CommandButton cmdHunt 
      Caption         =   "     Run WIN Hunter"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   7680
      Width           =   1575
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "   Pause"
      Height          =   375
      Left            =   2160
      TabIndex        =   65
      Top             =   7680
      Width           =   855
   End
End
Attribute VB_Name = "frmFilterOptimizer"
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

Dim bStopHunt As Boolean
Dim bHuntFilter As Boolean
Dim iHolder() As Integer
Dim dRatio As Double
Dim dHP As Double
Dim dHPRatio As Double
Dim dJP As Double

Dim dBestRatio As Double
Dim iBestRatioJackpot As Integer
Dim iBestRatioPredicted As Integer
Dim dBestRatioHitPick As Double

Dim iBestJackpotHits As Integer
Dim dBestJackpotRatio As Double
Dim iBestJackpotPredicted As Integer
Dim dBestJackPotHitPick As Double

Dim iBestHPHits As Integer
Dim dBestHPRatio As Double
Dim iBestHPPredicted As Integer
Dim dBestHPHitPick As Double

Dim sBestRatioSettings As String
Dim sBestJackpotSettings As String
Dim sBestHPSettings As String
Dim iHunted As Integer
Dim dPreviousSuccess As Double



Private lBestJackpotTop As Long
Private lBestRatioTop As Long
Private lBestHitPickTop As Long
Private lInitialHeight As Long
Private WithEvents mFilter As clsFilter
Attribute mFilter.VB_VarHelpID = -1
Private WithEvents mStack As clsStack
Attribute mStack.VB_VarHelpID = -1

'set the local copy of the Filter to use here
Public Sub Filter(cFilter As clsFilter, sFilterKey As String)
Dim nodY        As Node
Dim nodT        As Node
Dim PropVals    As Object
Dim Prop        As Object
Dim iProcessor  As Integer

    Set mFilter = cFilter
    
    Me.Caption = "Optimize - " & mFilter.Name
    tvFilterDisplay.Nodes.Clear
    'Load the filter into the Tree Here
    'Be sure to add a Min/Max Range for each Property
    'the tree items have
    With mFilter
        Set nodY = tvFilterDisplay.Nodes.Add(, , sFilterKey, "Filter")
        Set PropVals = .PropertyValues
        For Each Prop In PropVals
            Select Case Prop.Group
                Case Is < 100
                    'textbox input
                    Set nodT = tvFilterDisplay.Nodes.Add(nodY.Key, tvwChild, Prop.Key, Prop.Name & "(" & Prop.Min & "/" & Prop.Max & "):")
                    'Set nodT = tvFilterDisplay.Nodes.Add(Prop.Key, tvwChild, Prop.Key & "min", "MIN(" & Prop.Min & "):")
                    'Set nodT = tvFilterDisplay.Nodes.Add(Prop.Key, tvwChild, Prop.Key & "max", "MAX(" & Prop.Max & "):")
                    'Set nodT = tvFilterDisplay.Nodes.Add(Prop.Key, tvwChild, Prop.Key & "value", "Value(" & Prop.Value & "):")
                Case 100, 200, 300, 400, 500, 600, 700, 800, 900
                    'combo box input
                    Set nodT = tvFilterDisplay.Nodes.Add(nodY.Key, tvwChild, Prop.Key, Prop.Name & ":")
                Case Else
                    'MsgBox "invalid Property"
            End Select
        Next
        
        If .ProcessorCount > 0 Then
            For iProcessor = 0 To .ProcessorCount - 1
                With .ProcessorItem(iProcessor + 1)
                    Set nodY = tvFilterDisplay.Nodes.Add(, , "processor" & MyLotto.GenerateKey, .Name)
                    nodY.Tag = .Key
                    Set PropVals = .PropertyValues
                    For Each Prop In PropVals
                        Select Case Prop.Group
                            Case Is < 100
                                'textbox input
                                Set nodT = tvFilterDisplay.Nodes.Add(nodY.Key, tvwChild, Prop.Key, Prop.Name & "(" & Prop.Min & "/" & Prop.Max & "):")
                                'Set nodT = tvFilterDisplay.Nodes.Add(nodY.Key, tvwChild, Prop.Key, Prop.Name & ":")
                                'Set nodT = tvFilterDisplay.Nodes.Add(Prop.Key, tvwChild, Prop.Key & "min", "MIN(" & Prop.Min & "):")
                                'Set nodT = tvFilterDisplay.Nodes.Add(Prop.Key, tvwChild, Prop.Key & "max", "MAX(" & Prop.Max & "):")
                                'Set nodT = tvFilterDisplay.Nodes.Add(Prop.Key, tvwChild, Prop.Key & "value", "Value(" & Prop.Value & "):")
                            Case 100, 200, 300, 400, 500, 600, 700, 800, 900
                                'combo box input
                                Set nodT = tvFilterDisplay.Nodes.Add(nodY.Key, tvwChild, Prop.Key, Prop.Name & ":")
                            Case Else
                                'MsgBox "invalid Property"
                        End Select
                    Next
                End With
            Next iProcessor
        End If
    End With
    ReDim iHolder(tvFilterDisplay.Nodes.Count)

End Sub

'set the local copy of the Stack to use here
Public Property Set Stack(ByRef cStack As clsStack)

    Set mStack = cStack

End Property




Private Sub cmdExit_Click()

    Unload Me

End Sub

Private Sub cmdHunt_Click()

    If LEDPause.State = LEDOff Then
        LEDRun.State = LEDOn
        LEDStop.State = LEDOff
        tvFilterDisplay.Enabled = False
        bStopHunt = False
        Set frmViewOutput.Stack = mStack
        WINHunt
        LEDStop.State = LEDOn
        LEDRun.State = LEDOff
        LEDPause.State = LEDOff
        LEDReset.State = LEDOff
        tvFilterDisplay.Enabled = True
    End If

End Sub

Private Sub cmdHunt_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If Button = vbLeftButton Then
        LEDRun.Top = cmdHunt.Top + 80
        LEDRun.Left = cmdHunt.Left + 60
    End If

End Sub

Private Sub cmdHunt_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        LEDRun.Top = cmdHunt.Top + 70
        LEDRun.Left = cmdHunt.Left + 50
    End If

End Sub

Private Sub cmdPause_Click()

    If Not LEDRun.State = LEDOff Then
        Select Case LEDPause.State
            Case LEDOn
                'Pause is either pausing, or is paused
                'so let's undo the pause state
                LEDReset.State = LEDOff
                LEDStop.State = LEDOff
                LEDRun.State = LEDOn
                LEDPause.State = LEDOff
                cmdPause.Enabled = False
'            Case LEDOn, LEDOff
'                LEDReset.State = LEDOff
'                LEDStop.State = LEDBlink
'                LEDRun.State = LEDBlink
'                'LEDPause.State = LEDBlink
        End Select
    End If
    If LEDPause.State = LEDBlink And LEDRun.State = LEDOn Then
        'attempting to pause, but now we are canceling the pause attempt
        LEDReset.State = LEDOff
        LEDStop.State = LEDOff
        LEDRun.State = LEDOn
        LEDPause.State = LEDOff
        cmdPause.Enabled = True
    End If

End Sub

Private Sub cmdPause_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        LEDPause.Top = cmdPause.Top + 80
        LEDPause.Left = cmdPause.Left + 60
        If LEDRun.State = LEDOn Then
            Select Case LEDPause.State
                Case LEDOff
                    LEDPause.State = LEDBlink
                Case LEDBlink, LEDOn
            End Select
        End If
    End If

End Sub

Private Sub cmdPause_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        LEDPause.Top = cmdPause.Top + 70
        LEDPause.Left = cmdPause.Left + 50
    End If

End Sub

Private Sub cmdReset_Click()

    If LEDStop.State = LEDOn Then
        LEDReset.State = LEDOn
    End If

End Sub

Private Sub cmdReset_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        LEDReset.Top = cmdReset.Top + 80
        LEDReset.Left = cmdReset.Left + 40
    End If

End Sub

Private Sub cmdReset_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        LEDReset.Top = cmdReset.Top + 70
        LEDReset.Left = cmdReset.Left + 30
    End If

End Sub

Private Sub cmdStop_Click()

    LEDStop.State = LEDOn
    LEDRun.State = LEDOff
    LEDPause.State = LEDOff
    LEDReset.State = LEDOff
    tvFilterDisplay.Enabled = True
    bStopHunt = True

End Sub

Private Sub cmdStop_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        LEDStop.Top = cmdStop.Top + 80
        LEDStop.Left = cmdStop.Left + 40
    End If

End Sub

Private Sub cmdStop_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton Then
        LEDStop.Top = cmdStop.Top + 70
        LEDStop.Left = cmdStop.Left + 30
    End If
End Sub

Private Sub Form_Load()
Dim lLeft As Long
Dim lTop As Long
Dim lHeight As Long

    Me.Top = 0
    Me.Left = 0
    Me.Width = 5800
    Me.Height = 6750
    Picture1.Left = tsOverall.ClientLeft
    Picture1.Top = tsOverall.ClientTop + 40
    lInitialHeight = Me.ScaleHeight
    lLeft = fraStatistics.Left
    lTop = fraStatistics.Top
    lHeight = fraStatistics.Height
    tvFilterDisplay.Left = lLeft
    tvFilterDisplay.Top = lTop
    tvFilterDisplay.Height = lHeight
    'mvSetup.Left = lLeft
    'mvSetup.Top = lTop
    'mvSetup.Height = lHeight
    fraResults.Left = lLeft
    fraResults.Top = lTop
    fraResults.Height = lHeight
    fraResults.Width = fraStatistics.Width
    lvResults.Height = fraResults.Height - 900
    lvResults.Width = fraResults.Width - 280
    cmdSave.Top = fraResults.Height - 460
    cmdSave.Left = fraResults.Width - cmdSave.Width - 160
    
    lBestJackpotTop = lTop
    lBestRatioTop = lBestJackpotTop + fraBestJackpot.Height + 100
    lBestHitPickTop = fraBestRatio.Top + fraBestRatio.Height + 100
    
    fraBestJackpot.Top = lBestJackpotTop
    fraBestJackpot.Left = lLeft
    fraBestRatio.Top = lBestRatioTop
    fraBestRatio.Left = lLeft
    fraBestHitPick.Top = lBestHitPickTop
    fraBestHitPick.Left = lLeft
    vscrlPrevious.Top = 0
    vscrlPrevious.Height = Picture1.Height - 60
    vscrlPrevious.Left = Picture1.Width - vscrlPrevious.Width - 60
    vscrlPrevious.Max = (fraBestHitPick.Height / 2) + 240
    vscrlPrevious.Min = 0
    vscrlPrevious.SmallChange = 120
    vscrlPrevious.LargeChange = 360
    
    'mvSetup.Visible = True
    tvFilterDisplay.Visible = True
    fraStatistics.Visible = False
    fraBestJackpot.Visible = False
    fraBestRatio.Visible = False
    fraBestHitPick.Visible = False
    fraResults.Visible = False
    vscrlPrevious.Visible = False

End Sub

Private Sub Form_Resize()

    If Me.ScaleWidth < 1000 Then
        Exit Sub
    End If
    If tsOverall.Left > Me.ScaleWidth Then Exit Sub
    tsOverall.Width = Me.ScaleWidth - tsOverall.Left - 100
    tsOverall.Height = Me.ScaleHeight - 600
    Picture1.Left = tsOverall.ClientLeft
    Picture1.Top = tsOverall.ClientTop + 40
    'lvFilterDisplay.Height = tsOverall.Height
    Picture1.Height = tsOverall.ClientHeight - 40
    Picture1.Width = tsOverall.ClientWidth
    vscrlPrevious.Height = Picture1.Height - 60
    
    If Picture1.Height > (lBestHitPickTop + fraBestHitPick.Height + 120) Then
        vscrlPrevious.Visible = False
    Else
        If Not tsOverall.SelectedItem Is Nothing Then
            If tsOverall.SelectedItem.Key = "previous" Then
                vscrlPrevious.Visible = True
            End If
        End If
    End If
    vscrlPrevious.Left = Picture1.Width - vscrlPrevious.Width - 60
    cmdExit.Left = Me.ScaleWidth - 835
    cmdExit.Top = Me.ScaleHeight - cmdExit.Height - 40
    cmdReset.Left = Me.ScaleWidth - 2035
    cmdReset.Top = cmdExit.Top
    LEDReset.Top = cmdReset.Top + 70
    LEDReset.Left = cmdReset.Left + 30
    cmdStop.Left = Me.ScaleWidth - 2850
    cmdStop.Top = cmdExit.Top
    LEDStop.Top = LEDReset.Top
    LEDStop.Left = cmdStop.Left + 30
    cmdHunt.Left = Me.ScaleWidth - 5555
    cmdHunt.Top = cmdExit.Top
    LEDRun.Top = LEDReset.Top
    LEDRun.Left = cmdHunt.Left + 50
    cmdPause.Left = Me.ScaleWidth - 3780
    cmdPause.Top = cmdExit.Top
    LEDPause.Top = LEDReset.Top
    LEDPause.Left = cmdPause.Left + 40
    vscrlPrevious.Max = (fraBestHitPick.Height + fraBestHitPick.Top) - Picture1.ScaleHeight + 120

    fraResults.Height = Picture1.ScaleHeight - 240
    fraResults.Width = Picture1.ScaleWidth - 240
    lvResults.Height = fraResults.Height - 900
    lvResults.Width = fraResults.Width - 280
    cmdSave.Top = fraResults.Height - 460
    cmdSave.Left = fraResults.Width - cmdSave.Width - 160
    tvFilterDisplay.Height = Picture1.ScaleHeight - 240
    tvFilterDisplay.Width = Picture1.ScaleWidth - 240
    'mvSetup.Height = Picture1.ScaleHeight - 240
    'mvSetup.Width = Picture1.ScaleWidth - 240

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'cleanup
    Set mFilter = Nothing
    Set mStack = Nothing

End Sub

Private Sub mStack_Complete(ByVal sKey As String)

    If mStack.Output.PredictDrawing = 0 Then Exit Sub
    If mStack.Output.AverageHits > 0 Then
        dRatio = mStack.Output.AverageOutput / mStack.Output.AverageHits
        If mStack.Output.JackpotHits > 0 Then
            dJP = (mStack.Output.PredictDrawing / mStack.Output.JackpotHits) / mStack.Output.AverageHits
        Else
            dJP = dJP / 1000
        End If
    Else
        dRatio = dRatio / 10
        dJP = dJP / 1000
    End If
    If mStack.Output.JackpotHits > 0 And mStack.Output.JackpotOutput > 0 Then
        dHP = (mStack.Output.PredictDrawing / mStack.Output.JackpotHits) / mStack.Output.JackpotOutput
        dHPRatio = mStack.Output.JackpotOutput / mStack.Output.JackpotHits
    Else
        dHP = dHP / 1000
        dHPRatio = dHPRatio / 1000
    End If
    'lblDrawing.Caption = "Hunting: #" & iCurrentdrawing & " - " & Lotto.Drawings.Item(iCurrentdrawing).DrawnOn
    If frmMain.WindowState = 1 Then
        frmMain.Caption = "#" & mStack.Output.PredictDrawing
    End If
    dJP = dJP * 1000
    dHP = dHP * 1000
    dHPRatio = dHPRatio * 1000
    dRatio = dRatio * 10
    cntPerformanceAvgPicks.Value = mStack.Output.AverageOutput
    cntPerformanceAverageHits.Value = mStack.Output.AverageHits
    cntPerformanceProcessorCount.Value = mFilter.ProcessorCount
    If dRatio < 2000000 Then
        cntPerformanceRatio.Value = dRatio
    End If
    cntPerformanceJackpotPicks.Value = mStack.Output.JackpotOutput
    cntPerformanceJackpotHits.Value = mStack.Output.JackpotHits
    If dHP < 1000 Then cntPerformanceHitPickRatio.Value = dHP
    
    Set_Success dRatio
    
    'dJP = dJP * 1000
    'dHP = dHP * 1000
    'dHPRatio = dHPRatio * 1000
    'dBestJackPotHitPick = dBestJackPotHitPick * 1000
    cntBestJackpotHits.Value = iBestJackpotHits
    If dBestJackpotRatio < 1000 Then cntBestJackpotRatio.Value = dBestJackpotRatio * 100
    cntBestJackpotPicks.Value = iBestJackpotPredicted
    If dBestJackPotHitPick < 1000 Then cntBestJackpotRatio.Value = dBestJackPotHitPick * 100
    
   'dBestRatioHitPick = dBestRatioHitPick * 1000
    cntBestRatioHits.Value = iBestRatioJackpot
    If dBestRatio < 1000 Then cntBestRatio.Value = dBestRatio * 100
    cntBestRatioPicks.Value = iBestRatioPredicted
    If dBestRatioHitPick < 1000 Then cntBestRatioHitPick.Value = dBestRatioHitPick * 100

    'dBestHPHitPick = dBestHPHitPick * 1000
    cntBestHitPickHits.Value = iBestHPHits
    If dBestHPRatio < 1000 Then cntBestHitPickratio.Value = dBestHPRatio * 100
    cntBestHitPickPicks.Value = iBestHPPredicted
    If dBestHPHitPick < 1000 Then cntBestHitPicks.Value = dBestHPHitPick * 100

End Sub

Private Sub Set_Success(dValue As Double)
Dim iOn As Integer
Dim i As Integer

    iOn = Fix(dValue / 10)
    For i = 0 To 9
        If i < iOn + 1 Then
            If LEDProgress(i).State = LEDOff Then LEDProgress(i).State = LEDOn
        Else
            If LEDProgress(i).State = LEDOn Then LEDProgress(i).State = LEDOff
        End If
    Next i

End Sub

Private Sub tsOverall_Click()

    If Not tsOverall.SelectedItem Is Nothing Then
        Select Case tsOverall.SelectedItem.Key
            Case "setup"
                'mvSetup.Visible = True
                tvFilterDisplay.Visible = True
                fraStatistics.Visible = False
                fraBestJackpot.Visible = False
                fraBestRatio.Visible = False
                fraBestHitPick.Visible = False
                fraResults.Visible = False
                vscrlPrevious.Visible = False
            Case "statistics"
                fraStatistics.Visible = True
                'mvSetup.Visible = False
                tvFilterDisplay.Visible = False
                fraBestJackpot.Visible = False
                fraBestRatio.Visible = False
                fraBestHitPick.Visible = False
                fraResults.Visible = False
                vscrlPrevious.Visible = False
            Case "previous"
                fraBestJackpot.Visible = True
                fraBestRatio.Visible = True
                fraBestHitPick.Visible = True
                If Picture1.Height > (lBestHitPickTop + fraBestHitPick.Height + 120) Then
                    vscrlPrevious.Visible = False
                Else
                    If Not tsOverall.SelectedItem Is Nothing Then
                        If tsOverall.SelectedItem.Key = "previous" Then
                            vscrlPrevious.Visible = True
                        End If
                    End If
                End If
                'mvSetup.Visible = False
                tvFilterDisplay.Visible = False
                fraStatistics.Visible = False
                fraResults.Visible = False
            Case "results"
                fraResults.Visible = True
                'mvSetup.Visible = False
                tvFilterDisplay.Visible = False
                fraStatistics.Visible = False
                fraBestJackpot.Visible = False
                fraBestRatio.Visible = False
                fraBestHitPick.Visible = False
                vscrlPrevious.Visible = False
        End Select
    End If

End Sub



Private Sub tvFilterDisplay_DblClick()
Dim mSelectedNode As Node

    If Not tvFilterDisplay.SelectedItem.Parent Is Nothing Then
        If tvFilterDisplay.SelectedItem.Children = 0 Then
            Set mSelectedNode = tvFilterDisplay.SelectedItem
            MsgBox mSelectedNode.Key
        End If
    End If

End Sub

Private Sub vscrlPrevious_Change()
Dim lMove As Long

    lMove = vscrlPrevious.Value
    fraBestJackpot.Top = lBestJackpotTop - lMove
    fraBestRatio.Top = lBestRatioTop - lMove
    fraBestHitPick.Top = lBestHitPickTop - lMove

End Sub



















Private Sub Lotto_FilterComplete(ByVal iCurrentdrawing As Integer)
'
'    If Lotto.Statistics.PredictDrawing = 0 Then Exit Sub
'    If Lotto.Statistics.AverageHits > 0 Then
'        dRatio = Lotto.Statistics.AverageOutput / Lotto.Statistics.AverageHits
'        If Lotto.Statistics.JackpotHits > 0 Then
'            dJP = (iHunted / Lotto.Statistics.JackpotHits) / Lotto.Statistics.AverageHits
'        End If
'    End If
'    If Lotto.Statistics.JackpotHits > 0 Then
'        dHP = (iHunted / Lotto.Statistics.JackpotHits) / Lotto.Statistics.JackpotOutput
'        dHPRatio = Lotto.Statistics.JackpotOutput / Lotto.Statistics.JackpotHits
'    End If
'    lblDrawing.Caption = "Hunting: #" & iCurrentdrawing & " - " & Lotto.Drawings.Item(iCurrentdrawing).DrawnOn
'    If frmMain.WindowState = 1 Then
'        frmMain.Caption = "#" & iCurrentdrawing
'    End If
'    lblResults(0).Caption = Lotto.Statistics.AverageOutput
'    lblResults(1).Caption = Lotto.Statistics.AverageHits
'    lblResults(2).Caption = Lotto.Statistics.RuleSets.Count
'    lblResults(3).Caption = Format$(dRatio, "##.#")
'    lblResults(4).Caption = Lotto.Statistics.JackpotHits
'    If dHP < 1000 Then lblResults(5).Caption = Format$(dHP, "##.#")
'
'    lblResults(6).Caption = iBestJackpotHits
'    If dBestJackpotRatio < 1000 Then lblResults(7).Caption = Format$(dBestJackpotRatio, "##.##")
'    lblResults(8).Caption = iBestJackpotPredicted
'    If dBestJackPotHitPick < 1000 Then lblResults(9).Caption = Format$(dBestJackPotHitPick, "##.##")
'
'    lblResults(10).Caption = iBestRatioJackpot
'    If dBestRatio < 1000 Then lblResults(11).Caption = Format$(dBestRatio, "##.##")
'    lblResults(12).Caption = iBestRatioPredicted
'    If dBestRatioHitPick < 1000 Then lblResults(13).Caption = Format$(dBestRatioHitPick, "##.##")
'
'    lblResults(14).Caption = iBestHPHits
'    If dBestHPRatio < 1000 Then lblResults(15).Caption = Format$(dBestHPRatio, "##.##")
'    lblResults(16).Caption = iBestHPPredicted
'    If dBestHPHitPick < 1000 Then lblResults(17).Caption = Format$(dBestHPHitPick, "##.##")
'    lblResults(18) = Lotto.Statistics.JackpotOutput
'
End Sub

Private Sub WINHunt()
Dim bTemp As Boolean
Dim iLastRule As Integer

    'ReDim iHolder(18)
    
    dBestRatio = 10000
    iBestRatioJackpot = 0
    iBestRatioPredicted = 0
    sBestRatioSettings = ""
    dBestRatioHitPick = 10000
    
    dBestJackpotRatio = 10000
    iBestJackpotHits = 0
    iBestJackpotPredicted = 0
    sBestJackpotSettings = ""
    dBestJackPotHitPick = 10000
    
    dBestHPRatio = 10000
    iBestHPHits = 0
    iBestHPPredicted = 0
    sBestHPSettings = ""
    dBestHPHitPick = 10000
    
    'Lotto.Statistics.Predict = 1
    
    'Fill_Rules 0
    
    Do
        'Start the recursive calls here!
        bTemp = Increment(1)
        
        'If we are here, then we probably should add
        'a new filter!
        MsgBox "Hunting complete!"
        Exit Do
        
        'maybe we could cycle through the selection methods!
        If bStopHunt Then Exit Do
    Loop

End Sub

Private Sub PauseWinHunter()

    Do
        If LEDPause.State = LEDOff Then Exit Do
        DoEvents
    Loop
    
End Sub


Private Function Increment(ByVal iSetting As Integer) As Boolean
Dim bTemp As Boolean
Dim iIndex As Integer
Dim iMax As Integer
Dim iMin As Integer
Dim iValue As Integer
Dim sTemp As String

    'get the tree node
    With tvFilterDisplay.Nodes(iSetting)
        If .Checked Then
            'get min/max here!
            If Not .Parent Is Nothing Then
                If .Parent.Text = "Filter" Then
                    iMin = mFilter.PropertyValues(.Key).Min
                    iMax = mFilter.PropertyValues(.Key).Max
                    iHolder(iSetting) = mFilter.PropertyValues(.Key).Value
                Else
                    iMin = mFilter.ProcessorItem(LCase(.Parent.Text)).PropertyValues(.Key).Min
                    iMax = mFilter.ProcessorItem(LCase(.Parent.Text)).PropertyValues(.Key).Max
                    iHolder(iSetting) = mFilter.ProcessorItem(LCase(.Parent.Text)).PropertyValues(.Key).Value
                End If
            Else
                If Not .Child Is Nothing Then
                    'push down into the children nodes
                    'bTemp = Increment(.Next.Index + 1)   'recursive call!
                    bTemp = Increment(iSetting + 1)   'recursive call!
                    GoTo Exit_Increment
                End If
                Increment = False
                GoTo Exit_Increment
            End If
        Else
            If Not .Next Is Nothing Then
                'we are at the last node here
                'so dont use this node
                'and back out of the call
                bTemp = Increment(iSetting + 1)  'recursive call!
                GoTo Exit_Increment
            Else
                With tvFilterDisplay.Nodes(iSetting)
                    If Not .Parent Is Nothing Then
                        If .Parent.Next Is Nothing Then
                            'we are at the last node here
                            'so dont use this node
                            'and back out of the call
                            Increment = False
                            GoTo Exit_Increment
                        Else
                            'Push down through the items
                            bTemp = Increment(.Parent.Next.Index)  'recursive call!
                            If Not .Checked Then GoTo Exit_Increment
                        End If
                    End If
                End With
            End If
        End If
    End With
    
    If iHolder(iSetting) = iMax Then iHolder(iSetting) = iMin
    If iHolder(iSetting) < iMin Then iHolder(iSetting) = iMin

    DoEvents
    If tvFilterDisplay.Nodes(iSetting).Checked Then
        Do While iHolder(iSetting) < iMax + 1
            If Not tvFilterDisplay.Nodes(iSetting).Parent Is Nothing Then
                If tvFilterDisplay.Nodes(iSetting).Parent.Text = "Filter" Then
                    mFilter.PropertyValues(tvFilterDisplay.Nodes(iSetting).Key).Value = iHolder(iSetting)
                Else
                    mFilter.ProcessorItem(LCase(tvFilterDisplay.Nodes(iSetting).Parent.Text)).PropertyValues(tvFilterDisplay.Nodes(iSetting).Key).Value = iHolder(iSetting)
                End If
            End If
            sTemp = tvFilterDisplay.Nodes(tvFilterDisplay.Nodes(iSetting).Key).Text
            sTemp = Left$(sTemp, Len(sTemp) - (Len(sTemp) - InStr(sTemp, ":")))
            tvFilterDisplay.Nodes(tvFilterDisplay.Nodes(iSetting).Key).Text = sTemp & iHolder(iSetting)
            tvFilterDisplay.Refresh
            'push down through the items
            If Not tvFilterDisplay.Nodes(iSetting).Next Is Nothing Then
                If tvFilterDisplay.Nodes(iSetting).Next.Children = 0 Then
                    bTemp = Increment(iSetting + 1)   'recursive call!
                End If
            End If
            If bStopHunt Then Exit Do
            If LEDPause.State = LEDBlink Then
                'Enabling Pause
                'so let's set the pause state
                LEDReset.State = LEDOff
                LEDStop.State = LEDBlink
                LEDRun.State = LEDBlink
                LEDPause.State = LEDOn
                cmdPause.Enabled = True
                PauseWinHunter
            End If
            If Not bTemp Then
                'Run filters!
                
                frmViewOutput.txtOutput.Text = ""
                mStack.Reset
                mStack.TestStack
                
                'Save & check Best Performances here!
                'mStack.Output.
                With mStack.Output
                    If .AverageHits > 0 Then
                        dRatio = .AverageOutput / .AverageHits
                        If .JackpotHits > 0 Then
                            dJP = ((mStack.Drawings.Count - 1) / .JackpotHits) / .AverageHits
                        End If
                    End If
                    If .JackpotHits > 0 Then
                        dHP = ((mStack.Drawings.Count - 1) / .JackpotHits) / .JackpotOutput
                        dHPRatio = .JackpotOutput / .JackpotHits
                    End If
                    If Not (dRatio > dBestRatio) And .JackpotHits > 0 Then
                        If dJP < dBestRatioHitPick And .JackpotHits > iBestRatioPredicted Then
                            dBestRatio = dRatio
                            iBestRatioJackpot = .JackpotHits
                            iBestRatioPredicted = .JackpotOutput    'Lotto.Statistics.AverageOutput
                            sBestRatioSettings = GetStackXML(mStack)
                            dBestRatioHitPick = dJP
                        End If
                    End If
    
                    If Not .JackpotHits < iBestJackpotHits And .JackpotHits > 0 Then
                        If .JackpotHits = iBestJackpotHits Then
                            If .JackpotOutput < iBestJackpotPredicted Then
                                iBestJackpotHits = .JackpotHits
                                dBestJackpotRatio = dRatio
                                iBestJackpotPredicted = .JackpotOutput
                                sBestJackpotSettings = GetStackXML(mStack)
                                dBestJackPotHitPick = dJP
                            End If
                        Else
                            iBestJackpotHits = .JackpotHits
                            dBestJackpotRatio = dRatio
                            iBestJackpotPredicted = .JackpotOutput
                            sBestJackpotSettings = GetStackXML(mStack)
                            dBestJackPotHitPick = dJP
                        End If
                    End If
    
                    If .JackpotHits > 0 And dHP > 1 / .BallCount And dHP < (.BallCount / 1) * .DrawCount Then
                        If .JackpotHits > iBestHPHits Or .JackpotOutput < iBestHPPredicted Then
                            If (dHP < dBestHPHitPick And dHP > 1 / .BallCount) Then
                                iBestHPHits = .JackpotHits
                                dBestHPRatio = dHPRatio
                                iBestHPPredicted = .JackpotOutput
                                sBestHPSettings = GetStackXML(mStack)
                                dBestHPHitPick = dHP
                            End If
                        End If
                    End If
                End With

                bTemp = True
            End If
            'increment
            iHolder(iSetting) = iHolder(iSetting) + 1
            'If iHolder(iSetting) > iMax Then
            '    'We have reached the top of this loop
            '    'so we need to exit out of this loop
            '    iHolder(iSetting) = iMin
            '    Exit Do
            'End If
            'Change Filter setting here
            'Adjust iSetting, iHolder(iSetting)
            DoEvents
        Loop
    'Else
    '    'push down through the items
    '    bTemp = Increment(iSetting + 1)   'recursive call!
    End If
    
Exit_Increment:
Increment = bTemp

End Function

Private Function GetStackXML(oStack As clsStack) As String
Dim SXML        As New CGoXML
Dim PropVals    As Object
Dim Prop        As Object
Dim iGroup      As Integer
Dim iFilter     As Integer
Dim iProcessor  As Integer

    SXML.Initialize (pavAUTO)
    'START INITIAL FILE TEMPLATE
    Call SXML.OpenFromString("<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " ?>" & vbCrLf & "<STACK>" & vbCrLf & "</STACK>")

    If Not SXML.InsertNode("/STACK", "INITIAL_HISTORY_FILE", sHistoryName) Then Exit Function
    If Not SXML.InsertNode("/STACK", "GROUPS") Then Exit Function
    For iGroup = 0 To oStack.Groups.Count - 1
        With oStack.Groups.Item(iGroup + 1)
            If Not SXML.InsertNode("/STACK/GROUPS", "GROUP", "", "usegroup", .UseGroup) Then Exit Function
            If Not SXML.InsertNode("/STACK/GROUPS/GROUP", "GROUP_NAME", .Name) Then Exit Function
            If .FilterCount > 0 Then
                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]", "FILTERS") Then Exit Function
                For iFilter = 0 To .FilterCount - 1
                    With .FilterItem(iFilter + 1)
                        If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS", "FILTER", "", "usefilter", .UseFilter) Then Exit Function
                        If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]", "FILTER_NAME", .Name) Then Exit Function
                        If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]", "PROPERTIES") Then Exit Function
                        Set PropVals = .PropertyValues
                        For Each Prop In PropVals
                            Select Case Prop.Group
                                Case Is < 100
                                    'textbox input
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Function
                                Case 100, 200, 300, 400, 500, 600, 700, 800, 900
                                    'combo box input
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Function
                                Case Else
                                    'MsgBox "invalid Property"
                            End Select
                        Next
                        
                        If .ProcessorCount > 0 Then
                            If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]", "PROCESSORS") Then Exit Function
                            For iProcessor = 0 To .ProcessorCount - 1
                                With .ProcessorItem(iProcessor + 1)
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS", "PROCESSOR", "", "useprocessor", .UseProcessor) Then Exit Function
                                    If Not SXML.WriteAttribute("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]", "keyname", .Key) Then Exit Function
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]", "PROCESSOR_NAME", .Name) Then Exit Function
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]", "PROPERTIES") Then Exit Function
                                    Set PropVals = .PropertyValues
                                    For Each Prop In PropVals
                                        Select Case Prop.Group
                                            Case Is < 100
                                                'textbox input
                                                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Function
                                            Case 100, 200, 300, 400, 500, 600, 700, 800, 900
                                                'combo box input
                                                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Function
                                            Case Else
                                                'MsgBox "invalid Property"
                                        End Select
                                    Next
                                End With
                            Next iProcessor
                        End If
                    End With
                Next iFilter
                With .Selector
                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]", "SELECTOR", "", "keyname", .Key) Then Exit Function
                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/SELECTOR", "SELECTOR_NAME", .Name) Then Exit Function
                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/SELECTOR", "PROPERTIES") Then Exit Function
                    Set PropVals = .PropertyValues
                    For Each Prop In PropVals
                        Select Case Prop.Group
                            Case Is < 100
                                'textbox input
                                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/SELECTOR/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Function
                            Case 100, 200, 300, 400, 500, 600, 700, 800, 900
                                'combo box input
                                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/SELECTOR/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Function
                            Case Else
                                'MsgBox "invalid Property"
                        End Select
                    Next
                End With
            End If
        End With
    Next iGroup
    GetStackXML = SXML.XML
    Set SXML = Nothing
    Set PropVals = Nothing
    Set Prop = Nothing
    
End Function


