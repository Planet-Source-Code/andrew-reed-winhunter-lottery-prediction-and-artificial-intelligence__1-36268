VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "WIN! Hunter"
   ClientHeight    =   7980
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11385
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar sbrInfo 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   7725
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
            Key             =   "status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "info"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   2
            Object.Width           =   1587
            MinWidth        =   1587
            TextSave        =   "9:55 AM"
            Key             =   "time"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstCoolIconsMedium 
      Left            =   4920
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":030A
            Key             =   "output"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":079E
            Key             =   "off"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0B8E
            Key             =   "unloadedhistory"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F86
            Key             =   "machine"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13C6
            Key             =   "powerball"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17E2
            Key             =   "working"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C06
            Key             =   "filter"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":206A
            Key             =   "battery"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24CE
            Key             =   "lasersave"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":290A
            Key             =   "radioactive"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DCA
            Key             =   "communicate"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3212
            Key             =   "thinking"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":364A
            Key             =   "history"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A8E
            Key             =   "processor"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F0A
            Key             =   "export"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4346
            Key             =   "selector"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":47AE
            Key             =   "root"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4C42
            Key             =   "group"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":50D6
            Key             =   "trigger"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5582
            Key             =   "handpick"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5996
            Key             =   "stack"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstCoolIconsSmall 
      Left            =   5280
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5DF6
            Key             =   "output"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":627E
            Key             =   "filter"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":66E2
            Key             =   "thinking"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B1A
            Key             =   "machine"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F5A
            Key             =   "unloadedhistory"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7352
            Key             =   "stack"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":77B2
            Key             =   "processor"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C2E
            Key             =   "export"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8066
            Key             =   "selector"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":84CE
            Key             =   "root"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":895A
            Key             =   "group"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8DE6
            Key             =   "trigger"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9292
            Key             =   "handpick"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":969A
            Key             =   "history"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imglstCoolIcons 
      Left            =   4200
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9ADE
            Key             =   "output"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F66
            Key             =   "stacks"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A3C6
            Key             =   "root"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A852
            Key             =   "processor"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ACCE
            Key             =   "filter"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B106
            Key             =   "selector"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B56E
            Key             =   "group"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B9FA
            Key             =   "trigger"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BEA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C2AE
            Key             =   "history"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picLeft 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7725
      Left            =   0
      ScaleHeight     =   7725
      ScaleWidth      =   3900
      TabIndex        =   0
      Top             =   0
      Width           =   3900
      Begin VB.CommandButton cmdX 
         Height          =   210
         Left            =   1560
         Picture         =   "frmMain.frx":C6F2
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
         Width           =   210
      End
      Begin VB.PictureBox picTreeArea 
         Height          =   4815
         Left            =   120
         ScaleHeight     =   4755
         ScaleWidth      =   2475
         TabIndex        =   1
         Top             =   960
         Width           =   2535
         Begin MSComctlLib.TreeView tvFolders 
            Height          =   3015
            Left            =   480
            TabIndex        =   2
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   5318
            _Version        =   393217
            Indentation     =   141
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "imglstCoolIconsSmall"
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSComctlLib.TabStrip tsOverview 
         Height          =   5775
         Left            =   30
         TabIndex        =   4
         Top             =   360
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   10186
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Layout"
               Key             =   "Layout"
               Object.ToolTipText     =   "Shows Filter Schemes"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H80000002&
         Caption         =   "Overview"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   260
         Left            =   30
         TabIndex        =   5
         Top             =   30
         Width           =   1815
      End
   End
   Begin MSComctlLib.ImageList imglstTreeIcons 
      Left            =   4560
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C854
            Key             =   "closedgroup"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CCB4
            Key             =   "no"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CF48
            Key             =   "output"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D1E0
            Key             =   "filter"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D478
            Key             =   "history"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DF44
            Key             =   "opengroup"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E3A4
            Key             =   "stacks"
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog dlgBrowse 
      Left            =   6120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   ".txt"
      DialogTitle     =   "Browse Drawing Files"
      Filter          =   "*.txt"
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLoad 
         Caption         =   "&Load"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuStructure 
      Caption         =   "Structure"
      Begin VB.Menu mnuWinHunter 
         Caption         =   "WinHunter"
         Begin VB.Menu mnuNewHistory 
            Caption         =   "Add New History"
         End
         Begin VB.Menu mnuLoadHistory 
            Caption         =   "Open History"
         End
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "History"
         Begin VB.Menu mnuLoadHistoryFile 
            Caption         =   "Load File"
         End
         Begin VB.Menu mnuHistoryLoaded 
            Caption         =   "Loaded"
         End
         Begin VB.Menu mnuHistorySep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRemoveHistory 
            Caption         =   "Remove History"
         End
         Begin VB.Menu mnuHistorySep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewHistorySettings 
            Caption         =   "View Settings"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuScanHistory 
            Caption         =   "Scan"
         End
         Begin VB.Menu mnuGetPredictions 
            Caption         =   "*Get Predictions*"
         End
         Begin VB.Menu mnuRunHistory 
            Caption         =   "*Run Tests Now*"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuHistorySep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUseHistory 
            Caption         =   "Use History"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuMachine 
         Caption         =   "Machine"
         Begin VB.Menu mnuAddStack 
            Caption         =   "Add New Stack"
         End
         Begin VB.Menu mnuMachineSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLoadStack 
            Caption         =   "Load Stack"
         End
         Begin VB.Menu mnuMachineSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuStartAI 
            Caption         =   "Start AI Lifeform"
         End
         Begin VB.Menu mnuStatistics 
            Caption         =   "St&atistics"
            Enabled         =   0   'False
            Visible         =   0   'False
            Begin VB.Menu mnuSingleHits 
               Caption         =   "Single Hits"
            End
            Begin VB.Menu mnuSep1 
               Caption         =   "-"
            End
            Begin VB.Menu mnuPairHits 
               Caption         =   "Pair Hits"
            End
         End
      End
      Begin VB.Menu mnuStack 
         Caption         =   "Stack"
         Begin VB.Menu mnuAddGroup 
            Caption         =   "Add New Group"
         End
         Begin VB.Menu mnuStackSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSaveStack 
            Caption         =   "Save Stack"
         End
         Begin VB.Menu mnuStackSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRemoveStack 
            Caption         =   "Remove Stack"
         End
         Begin VB.Menu mnuStackSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlotStack 
            Caption         =   "Plot Stack"
         End
         Begin VB.Menu mnuStackStats 
            Caption         =   "Show Stats"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuGetPrediction 
            Caption         =   "*Get Prediction*"
         End
         Begin VB.Menu mnuRunStack 
            Caption         =   "*Run Tests Now*"
         End
         Begin VB.Menu mnuStopStack 
            Caption         =   "-< STOP TESTING  >-"
         End
         Begin VB.Menu mnuStackSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUseStack 
            Caption         =   "Use Stack"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuGroup 
         Caption         =   "Group"
         Begin VB.Menu mnuAddFilter 
            Caption         =   "Add New Filter"
         End
         Begin VB.Menu mnuGroupSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRemoveGroup 
            Caption         =   "Remove Group"
         End
         Begin VB.Menu mnuGroupSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRunGroup 
            Caption         =   "*Run Now*"
         End
         Begin VB.Menu mnuGroupSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUseGroup 
            Caption         =   "Use Group"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "Filter"
         Begin VB.Menu mnuAddProcessor 
            Caption         =   "Add Processor..."
         End
         Begin VB.Menu mnuAddTrigger 
            Caption         =   "Add Trigger..."
         End
         Begin VB.Menu mnuRemoveFilter 
            Caption         =   "Remove Filter"
         End
         Begin VB.Menu mnuFilterSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewFilterSettings 
            Caption         =   "View Settings"
         End
         Begin VB.Menu mnuPlotFilter 
            Caption         =   "View Plot"
         End
         Begin VB.Menu mnuFilterSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRunFilter 
            Caption         =   "*Run Now*"
         End
         Begin VB.Menu mnuLoadWinHunter 
            Caption         =   "WinHunter"
         End
         Begin VB.Menu mnuFilterSep3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUseFilter 
            Caption         =   "Use Filter"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuProcessor 
         Caption         =   "Processor"
         Begin VB.Menu mnuRemoveProcessor 
            Caption         =   "Remove Processor"
         End
         Begin VB.Menu mnuProcessorSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewProcessorSettings 
            Caption         =   "View Settings"
         End
         Begin VB.Menu mnuProcessorSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRunProcessor 
            Caption         =   "*Run Now*"
         End
         Begin VB.Menu mnuProcessorSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUseProcessor 
            Caption         =   "Use Processor"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuTrigger 
         Caption         =   "Trigger"
         Begin VB.Menu mnuRemoveTrigger 
            Caption         =   "Remove Trigger"
         End
         Begin VB.Menu mnuTriggerSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewTriggerSettings 
            Caption         =   "View Settings"
         End
         Begin VB.Menu mnuTriggerSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuUseTrigger 
            Caption         =   "Use Trigger"
         End
      End
      Begin VB.Menu mnuSelector 
         Caption         =   "Selector"
         Begin VB.Menu mnuChangeSelector 
            Caption         =   "Change Selector"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuSelectorSep 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewSelectorSettings 
            Caption         =   "View Settings"
         End
      End
      Begin VB.Menu mnuAILifeForm 
         Caption         =   "AI LifeForm"
         Begin VB.Menu mnuViewLifeform 
            Caption         =   "Monitor Lifeform"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuAIsep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuKillAILifeform 
            Caption         =   "Kill AI Lifeform"
         End
      End
   End
   Begin VB.Menu mnuOverview 
      Caption         =   "&Overview..."
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
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
Dim WithEvents Lotto As clsLotto
Attribute Lotto.VB_VarHelpID = -1

Private Sub cmdX_Click()

    picLeft.Visible = False

End Sub

Private Sub MDIForm_Load()

    CreateObjects
    frmSplash.Show 1
    'Create instance of the lottery object
    Set MyLotto = New clsLotto
    'get the local copy of the instance
    Set Lotto = MyLotto
    picTreeArea.Left = tsOverview.ClientLeft
    picTreeArea.Top = tsOverview.ClientTop + 300
    picTreeArea.Width = tsOverview.ClientWidth
    picTreeArea.Height = tsOverview.ClientHeight - 300
    Set nodY = tvFolders.Nodes.Add(, , "root", "WINHunter", "root")
    'hide the structure menu, since it only serves to hold the tree popup menus
    mnuStructure.Visible = False
    'FillTree
    'frmLoadDrawings.Show
    


End Sub

Private Sub BrowseHistory()
On Error GoTo Browse_Error

        With dlgBrowse
            .DefaultExt = ".txt"
            .Filter = "Drawing History File (*.txt)|*.txt"
            .Flags = &H1000 Or &H4
            .FileName = "History.txt"
            .InitDir = App.Path & "\History"
            .DialogTitle = "Load History File"
            .ShowOpen
        End With
        Exit Sub

Browse_Error:
    Select Case Err.Number
        Case 32755
            'dont load the filename
            dlgBrowse.FileName = ""
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select



End Sub

Private Sub BrowseStack()
On Error GoTo Browse_Stack_Error

        With dlgBrowse
            .DefaultExt = ".xml"
            .Filter = "Filter Stack File (*.xml)|*.xml"
            .Flags = &H1000 Or &H4
            .FileName = "Stack.xml"
            .InitDir = App.Path & "\stacks"
            .DialogTitle = "Load Stack File"
            .ShowOpen
        End With
        Exit Sub

Browse_Stack_Error:
    Select Case Err.Number
        Case 32755
            'dont load the filename
            dlgBrowse.FileName = ""
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select

End Sub

Private Sub BrowseSaveStack()
On Error GoTo Browse_Save_Stack_Error

        With dlgBrowse
            .DefaultExt = ".xml"
            .Filter = "Filter Stack File (*.xml)|*.xml"
            .Flags = &H4 Or &H2
            .FileName = ""
            .InitDir = App.Path & "\stacks"
            .DialogTitle = "Save Stack File"
            .ShowSave
        End With
        Exit Sub

Browse_Save_Stack_Error:
    Select Case Err.Number
        Case 32755
            'dont load the filename
            dlgBrowse.FileName = ""
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select



End Sub

Private Sub showloadsave(sTitle As String, bShowOpen As Boolean)
On Error GoTo Load_Error

        With dlgBrowse
            .DefaultExt = ".ftr"
            .Filter = "Filter Setup File (*.ftr)|*.ftr"
            .FileName = ".ftr"
            .InitDir = App.Path & "\stacks"
            .DialogTitle = sTitle
            .Flags = &H4 + &H2           'cdlOFNHideReadOnly & cdlOFNOverwritePrompt
            If bShowOpen Then
                .ShowOpen
            Else
                .ShowSave
            End If
        End With
        Exit Sub

Load_Error:
    Select Case Err.Number
        Case 32755
            'dont load the filename
            dlgLoadSave.FileName = ""
        Case Else
            MsgBox Err.Number & " : " & Err.Description
    End Select



End Sub

Private Sub FillTree()
'Dim nodY As Node

'    Set nodY = tvFolders.Nodes.Add(, , "R", "WINHunter", "root")
'    'Do While Not rsTree.EOF
'        Set nodY = tvFolders.Nodes.Add("R", tvwChild, "H1", "FLLotto.txt (History)", "history")
'        Set nodY = tvFolders.Nodes.Add("H1", tvwChild, "M1", "Machine", "machine")
'        Set nodY = tvFolders.Nodes.Add("M1", tvwChild, "S1", "FLLoto.ftr (Stack)", "stack")
'        Set nodY = tvFolders.Nodes.Add("S1", tvwChild, "G1", "Group1", "group")
'        Set nodY = tvFolders.Nodes.Add("G1", tvwChild, "F1", "Filter1", "filter")
'            Set nodY = tvFolders.Nodes.Add("F1", tvwChild, "T1", "Trigger1", "trigger")
'            Set nodY = tvFolders.Nodes.Add("F1", tvwChild, "P1", "Processor1", "processor")
'            Set nodY = tvFolders.Nodes.Add("F1", tvwChild, "P2", "Processor2", "processor")
'        Set nodY = tvFolders.Nodes.Add("G1", tvwChild, "F2", "Filter2", "filter")
'            Set nodY = tvFolders.Nodes.Add("F2", tvwChild, "P3", "Processor3", "processor")
'        Set nodY = tvFolders.Nodes.Add("G1", tvwChild, "SL1", "Selector1", "selector")
'        Set nodY = tvFolders.Nodes.Add("G1", tvwChild, "G2", "Group2", "group")
'        Set nodY = tvFolders.Nodes.Add("G2", tvwChild, "F3", "Filter3", "filter")
'            Set nodY = tvFolders.Nodes.Add("F3", tvwChild, "P4", "Processor4", "processor")
'        Set nodY = tvFolders.Nodes.Add("G2", tvwChild, "O1", "Stack Output", "output")
'    'Loop
'    nodY.EnsureVisible
'    'When filling the tree, The next series Group goes at the bottom of the children
'    'And the Output shows ONLY on the last ACTIVE Group..., Again, at the BOTTOM
'    'of the list

End Sub

Private Sub MDIForm_Resize()

    If Me.WindowState <> 1 Then
        Me.Caption = "WIN! Hunter |)->"
    End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    Screen.MousePointer = 11
        Set Lotto = Nothing
        Set MyLotto = Nothing
        Set colWindows = Nothing
    Screen.MousePointer = 0
    'Set ScanForm() = Nothing

End Sub

Private Sub mnuAbout_Click()

    frmAbout.Show 1

End Sub

Private Sub mnuAddFilter_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim NodSelected As Node
Dim nodY As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "group" Then
        sGroupKey = NodSelected.Key
        Set nodY = NodSelected
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
    Else
        Exit Sub
    End If
    Set nodY = tvFolders.Nodes.Add(NodSelected.Key, tvwChild, "filter" & Lotto.GenerateKey, "Filter", "filter")
    Set nodY.Parent = NodSelected
    Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).AddFilter (nodY.Key)
    nodY.EnsureVisible

End Sub

Private Sub mnuAddGroup_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim NodSelected As Node
Dim nodY As Node
Dim iGroupCount As Integer
Dim nodOutput As Node
Dim nodSelector As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "stack" Then
        sStackKey = NodSelected.Key
        sMachineKey = NodSelected.Parent.Key
        sHistoryKey = NodSelected.Parent.Parent.Key
    Else
        Exit Sub
    End If
    iGroupCount = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups.Count
    If iGroupCount = 0 Then
        Set nodY = tvFolders.Nodes.Add(NodSelected.Key, tvwChild, "group" & Lotto.GenerateKey, "Group", "group")
        Set nodSelector = tvFolders.Nodes.Add(nodY.Key, tvwChild, "selector" & nodY.Key, "Selector", "selector")
        'add the selector to the group here...
        Set nodOutput = tvFolders.Nodes.Add(nodY.Key, tvwChild, "output" & sStackKey, "Output", "output")
    Else
        'ok, groups must be shown at the end of the tree group after the last filter
        'This is simply to show how the filters operate graphically
        'Any tree actions must ensure that the groups ALWAYS end up at the end of
        'the list!
        Set nodOutput = tvFolders.Nodes("output" & tvFolders.SelectedItem.Key)
        'Set nodSelector = tvFolders.Nodes("selector" & tvFolders.SelectedItem.Key)
        Set nodY = tvFolders.Nodes(Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(iGroupCount).Key)
        Set nodY = tvFolders.Nodes.Add(nodY.Key, tvwChild, "group" & Lotto.GenerateKey, "Group", "group")
        Set nodSelector = tvFolders.Nodes.Add(nodY.Key, tvwChild, "selector" & nodY.Key, "Selector", "selector")
        'move the selector and the output to the last
        Set nodOutput.Parent = nodY
        Set nodSelector.Parent = nodY
    End If
    Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups.Add (nodY.Key)
    nodOutput.EnsureVisible

End Sub

Private Sub mnuAddProcessor_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim NodSelected As Node
Dim nodY As Node
Dim oProcessor As Object
Dim oProcessors As Object
Dim iSelectedIndex As New clsIndexer

    Set oProcessors = CreateObject(sObjProcessors)
    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "filter" Then
        sFilterKey = NodSelected.Key
        sGroupKey = NodSelected.Parent.Key
        Set nodY = NodSelected.Parent
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
    Else
        Exit Sub
    End If
    
PickProcessorAgain:
    'Display Processor Selection Window here...
    Load frmPickFrom
    frmPickFrom.Icon = imglstCoolIconsSmall.ListImages.Item("processor").ExtractIcon
    'move the form over the Interface
    If frmMain.Width > frmPickFrom.Width Then
        frmPickFrom.Left = frmMain.Left + (frmMain.Width / 2 - (frmPickFrom.Width / 2))
    End If
    If frmMain.Height > frmPickFrom.Height Then
        frmPickFrom.Top = frmMain.Top + (frmMain.Height / 2 - (frmPickFrom.Height / 2))
    End If
    
    For Each oProcessor In oProcessors
        frmPickFrom.cboSelect.AddItem oProcessor.Name
    Next
    
    'Save the ability to get the index from the form
    Set frmPickFrom.GetIndexer = iSelectedIndex
    
    frmPickFrom.Show 1
    'did the user make a selection?
    If Not iSelectedIndex.Indexer Then
        'What Processor did the user select?
        'Does it already exist in the filter?
        Set oProcessor = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).ProcessorItem(oProcessors(iSelectedIndex.Indexer + 1).Key)
        
        If oProcessor Is Nothing Then
            'add tree node
            Set nodY = tvFolders.Nodes.Add(NodSelected.Key, tvwChild, "processor" & Lotto.GenerateKey, oProcessors(iSelectedIndex.Indexer + 1).Name, "processor")
            'set parent for node
            Set nodY.Parent = NodSelected
            'save the Processor Key for later use
            nodY.Tag = oProcessors(iSelectedIndex.Indexer + 1).Key
            'add processor to filter collection
            Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).AddProcessor nodY.Key, oProcessors(iSelectedIndex.Indexer + 1).Key
            'make sure the node is visible in the treeview
            nodY.EnsureVisible
        Else
            MsgBox "That Processor already selected."
            GoTo PickProcessorAgain
        End If
    End If

End Sub

Private Sub mnuAddStack_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim NodSelected As Node
Dim nodY As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "machine" Then
        sMachineKey = NodSelected.Key
        sHistoryKey = NodSelected.Parent.Key
    Else
        Exit Sub
    End If
    Set nodY = tvFolders.Nodes.Add(NodSelected.Key, tvwChild, "stack" & Lotto.GenerateKey, "Stack", "stack")
    Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks.Add (nodY.Key)
    nodY.EnsureVisible

End Sub

Private Sub mnuAddTrigger_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim NodSelected As Node
Dim nodY As Node
Dim oTrigger As Object
Dim oTriggers As Object
Dim iSelectedIndex As New clsIndexer

    Set oTriggers = CreateObject(sObjTriggers)
    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "filter" Then
        sFilterKey = NodSelected.Key
        sGroupKey = NodSelected.Parent.Key
        Set nodY = NodSelected.Parent
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
    Else
        Exit Sub
    End If
    'Display Processor Selection Window here...
    Load frmPickFrom
    frmPickFrom.Icon = imglstCoolIconsSmall.ListImages.Item("trigger").ExtractIcon
    For Each oTrigger In oTriggers
        frmPickFrom.cboSelect.AddItem oTrigger.Name
    Next
    
    'Save the ability to get the index from the form
    Set frmPickFrom.GetIndexer = iSelectedIndex
    
    frmPickFrom.Show 1
    'did the user make a selection?
    If Not iSelectedIndex.Indexer Then
        'add tree node
        Set nodY = tvFolders.Nodes.Add(NodSelected.Key, tvwChild, "trigger" & Lotto.GenerateKey, oTriggers(iSelectedIndex.Indexer + 1).Name, "trigger")
        nodY.Tag = oTriggers(iSelectedIndex.Indexer + 1).Key
        'set parent for node
        Set nodY.Parent = NodSelected
        'add processor to filter collection
        Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).AddTrigger nodY.Key, oTriggers(iSelectedIndex.Indexer + 1).Key
        'mkae sure the node is visible in the treeview
        nodY.EnsureVisible
    End If

End Sub

Private Sub mnuChangeSelector_Click()

    'frmPickFrom.Icon = imglstCoolIconsSmall.ListImages.Item("selector").ExtractIcon

End Sub

Private Sub mnuExit_Click()

    'Lotto.Statistics.EndProcess = True
    'Do
    '    DoEvents
    'Loop Until Lotto.Statistics.Predicting = False
    Unload Me

End Sub

Private Sub mnuGetPrediction_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim NodSelected As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "stack" Then
        sStackKey = NodSelected.Key
        sMachineKey = NodSelected.Parent.Key
        sHistoryKey = NodSelected.Parent.Parent.Key
    Else
        Exit Sub
    End If
    Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).PredictDrawing = 0
    Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).RunStack
    For i = 1 To Lotto.Histories(sHistoryKey).Machines(sMachineKey).MachineMaximumBallNumber
        With Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Output
            'the prediction routines actually work to EXCLUDE balls
            If Not .Excluded(i) Then
                lPredictedCount = lPredictedCount + 1
                If addstring = "" Then
                    addstring = i
                Else
                    addstring = addstring & ", " & i
                End If
            End If
        End With
    Next i
    addstring = lPredictedCount & " Predictions for the following week: " & addstring & vbCrLf
    frmViewOutput.txtOutput.Text = addstring & frmViewOutput.txtOutput.Text
    DoEvents

End Sub

Private Sub mnuGetPredictions_Click()
Dim sHistoryKey As String

    sHistoryKey = tvFolders.SelectedItem.Key
    Lotto.Histories(sHistoryKey).RunHistory

End Sub

Private Sub mnuLoadHistory_Click()

    NewHistory
    LoadHistory False
    Screen.MousePointer = 0

End Sub

Private Sub mnuLoadHistoryFile_Click()

    LoadHistory True
    Screen.MousePointer = 0

End Sub

Private Sub LoadHistory(bDefaultStack As Boolean)
Dim nodClicked As Node
Dim nodY As Node
Dim mMachine As clsMachine
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String

    Screen.MousePointer = 11
    DoEvents
    Set nodClicked = tvFolders.SelectedItem
    If Not nodClicked Is Nothing Then
        If LCase(nodClicked.Image) = "unloadedhistory" Then
            Lotto.Histories(nodClicked.Key).Load
            nodClicked.Image = "history"
            For Each mMachine In Lotto.Histories(nodClicked.Key).Machines
                Set nodY = tvFolders.Nodes.Add(nodClicked.Key, tvwChild, mMachine.Key, "Machine (" & mMachine.MachineDrawCount & "/" & mMachine.MachineMaximumBallNumber & ")", "machine")
                If mMachine.BonusBallMachine Then
                    nodY.Text = "Bonus " & nodY.Text
                End If
                nodY.EnsureVisible
                If bDefaultStack Then
                    If mMachine.Stacks.Count = 0 Then
                        'brand new machine added here, so let's add a stack
                        Set nodY = tvFolders.Nodes.Add(nodY.Key, tvwChild, "stack" & Lotto.GenerateKey, "Stack", "stack")
                        sStackKey = nodY.Key
                        mMachine.Stacks.Add (sStackKey)
                        'now add a group
                        Set nodY = tvFolders.Nodes.Add(sStackKey, tvwChild, "group" & Lotto.GenerateKey, "Group", "group")
                        sGroupKey = nodY.Key
                        Set nodY = tvFolders.Nodes.Add(sGroupKey, tvwChild, "selector" & sGroupKey, "Selector", "selector")
                        'add the selector to the group here...
                        Set nodY = tvFolders.Nodes.Add(sGroupKey, tvwChild, "output" & sStackKey, "Output", "output")
                        mMachine.Stacks(sStackKey).Groups.Add (sGroupKey)
                        'now add a filter
                        Set nodY = tvFolders.Nodes.Add(sGroupKey, tvwChild, "filter" & Lotto.GenerateKey, "Filter", "filter")
                        mMachine.Stacks(sStackKey).Groups(sGroupKey).AddFilter (nodY.Key)
                        Set nodY.Parent = tvFolders.Nodes(sGroupKey)
                        nodY.EnsureVisible
                    End If
                End If
            Next
        End If
    End If
    Screen.MousePointer = 0

End Sub
Private Sub mnuLoadStack_Click()
Dim mHistory As clsHistory
Dim sFile As String
Dim sHistoryKey As String
Dim sMachineKey As String
Dim NodSelected As Node
Dim nodY As Node

    BrowseStack
    If dlgBrowse.FileName = "" Then
        'MsgBox "No stack file specified."
    Else
        sFile = dlgBrowse.FileName
        Set NodSelected = tvFolders.SelectedItem
        sMachineKey = NodSelected.Key
        sHistoryKey = NodSelected.Parent.Key
        Screen.MousePointer = 11
            Set nodY = tvFolders.Nodes.Add(NodSelected.Key, tvwChild, "stack" & Lotto.GenerateKey, "Stack", "stack")
            nodY.EnsureVisible
            Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks.Add (nodY.Key)
            'Load Stack file here
            LoadStack Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(nodY.Key), nodY, sFile
        Screen.MousePointer = 0
    End If

End Sub

Private Sub LoadStack(oStack As clsStack, StackNode As Node, sFileName As String)
Dim StackXML    As New CGoXML
Dim GroupXML    As New CGoXML
Dim FilterXML   As New CGoXML
Dim TempXML     As New CGoXML
Dim GroupNode   As Node
Dim FilterNode  As Node
Dim tempNode    As Node
Dim PropVals    As Object
Dim iPropCount  As Integer

    StackXML.Initialize (pavAUTO)
    GroupXML.Initialize (pavAUTO)
    FilterXML.Initialize (pavAUTO)
    TempXML.Initialize (pavAUTO)
    Call StackXML.OpenFromFile(sFileName)
    If StackXML.NodeCount("/STACK/GROUPS") > 0 Then
        Do
            GroupXML.OpenFromString StackXML.ReadNodeXML("/STACK/GROUPS/GROUP[" & oStack.Groups.Count & "]")
            If oStack.Groups.Count = 0 Then
                Set GroupNode = tvFolders.Nodes.Add(StackNode.Key, tvwChild, "group" & Lotto.GenerateKey, "Group", "group")
                'add the selector to the group here...
                Set tempNode = tvFolders.Nodes.Add(GroupNode.Key, tvwChild, "selector" & GroupNode.Key, "Selector", "selector")
                sSelectKey = tempNode.Key
                'add the output group here...
                Set tempNode = tvFolders.Nodes.Add(GroupNode.Key, tvwChild, "output" & StackNode.Key, "Output", "output")
            Else
                'ok, groups must be shown at the end of the tree group after the last filter
                Set GroupNode = tvFolders.Nodes.Add(tvFolders.Nodes(oStack.Groups(oStack.Groups.Count).Key).Key, tvwChild, "group" & Lotto.GenerateKey, "Group", "group")
                'add the selector node
                Set tempNode = tvFolders.Nodes.Add(GroupNode.Key, tvwChild, "selector" & GroupNode.Key, "Selector", "selector")
                sSelectKey = tempNode.Key
                'move the selector and the output to the last
                Set tvFolders.Nodes("output" & StackNode.Key).Parent = GroupNode    'Output
                Set tempNode.Parent = GroupNode 'selector
            End If
            'add the group to the stack
            oStack.Groups.Add (GroupNode.Key)
            
            
            
            With oStack.Groups(GroupNode.Key)
                'get the boolean value
                .UseGroup = Str2Bool(GroupXML.ReadAttribute("/GROUP", "usegroup"))
                If GroupXML.NodeCount("/GROUP/FILTERS/FILTER") > 0 Then
                    Do
                        FilterXML.OpenFromString GroupXML.ReadNodeXML("/GROUP/FILTERS/FILTER[" & .FilterCount & "]")
                        Set FilterNode = tvFolders.Nodes.Add(GroupNode.Key, tvwChild, "filter" & Lotto.GenerateKey, "Filter", "filter")
                        Set FilterNode.Parent = GroupNode
                        FilterNode.EnsureVisible
                        .AddFilter (FilterNode.Key)
                        
                        
                        
                        
                        With .FilterItem(FilterNode.Key)
                            .UseFilter = Str2Bool(FilterXML.ReadAttribute("/FILTER", "usefilter"))
                            Set PropVals = .PropertyValues
                            iPropCount = 0
                            Do
                                PropVals.Item(FilterXML.ReadAttribute("/FILTER/PROPERTIES/PROPERTY[" & iPropCount & "]", "keyname")).Value = FilterXML.ReadNode("/FILTER/PROPERTIES/PROPERTY[" & iPropCount & "]")
                                iPropCount = iPropCount + 1
                            Loop Until FilterXML.NodeCount("/FILTER/PROPERTIES/PROPERTY") = iPropCount
                            If FilterXML.NodeCount("/FILTER/PROCESSORS/PROCESSOR") > 0 Then
                                
                                
                                
                                
                                Do
                                    TempXML.OpenFromString FilterXML.ReadNodeXML("/FILTER/PROCESSORS/PROCESSOR[" & .ProcessorCount & "]")
                                    Set tempNode = tvFolders.Nodes.Add(FilterNode.Key, tvwChild, "processor" & Lotto.GenerateKey, TempXML.ReadNode("PROCESSOR_NAME"), "processor")
                                    Set tempNode.Parent = FilterNode
                                    tempNode.Tag = TempXML.ReadAttribute("/PROCESSOR", "keyname")
                                    tempNode.EnsureVisible
                                    Call .AddProcessor(tempNode.Key, tempNode.Tag)
                                    With .ProcessorItem(tempNode.Tag)
                                        .UseProcessor = Str2Bool(TempXML.ReadAttribute("/PROCESSOR", "useprocessor"))
                                        Set PropVals = .PropertyValues
                                        iPropCount = 0
                                        Do
                                            PropVals.Item(TempXML.ReadAttribute("/PROCESSOR/PROPERTIES/PROPERTY[" & iPropCount & "]", "keyname")).Value = TempXML.ReadNode("/PROCESSOR/PROPERTIES/PROPERTY[" & iPropCount & "]")
                                            iPropCount = iPropCount + 1
                                        Loop Until TempXML.NodeCount("/PROCESSOR/PROPERTIES/PROPERTY") = iPropCount
                                    End With
                                Loop Until .ProcessorCount = FilterXML.NodeCount("/FILTER/PROCESSORS/PROCESSOR")
                            End If
                            'If FilterXML.NodeCount("/FILTER/TRIGGERS") > 0 Then
                            '
                            'End If
                        End With
                        
                        
                    Loop Until .FilterCount = GroupXML.NodeCount("/GROUP/FILTERS/FILTER")
                End If
                'Add the selector here
                TempXML.OpenFromString GroupXML.ReadNodeXML("/GROUP/SELECTOR")
                .SetSelector TempXML.ReadAttribute("/SELECTOR", "keyname")
                With .Selector
                    Set PropVals = .PropertyValues
                    iPropCount = 0
                    Do
                        PropVals.Item(TempXML.ReadAttribute("/SELECTOR/PROPERTIES/PROPERTY[" & iPropCount & "]", "keyname")).Value = TempXML.ReadNode("/SELECTOR/PROPERTIES/PROPERTY[" & iPropCount & "]")
                        iPropCount = iPropCount + 1
                    Loop Until TempXML.NodeCount("/SELECTOR/PROPERTIES/PROPERTY") = iPropCount
                End With
            End With
        Loop Until oStack.Groups.Count = StackXML.NodeCount("/STACK/GROUPS/GROUP")
    End If


End Sub

Private Function Str2Bool(sData As String) As Boolean

    If Not IsNumeric(sData) Then
        Select Case LCase(sData)
            Case "t", "true", "yes", "inclusion"
                Str2Bool = True
            Case "f", "false", "no", "exclusion"
                Str2Bool = False
            Case Else
                Str2Bool = False
        End Select
    Else
        Select Case LCase(sData)
            Case "-1"
                Str2Bool = True
            Case "0"
                Str2Bool = False
            Case Else
                Str2Bool = False
        End Select
    End If

End Function


Private Sub mnuLoadWinHunter_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim NodSelected As Node
Dim nodY As Node
Dim oProcessor As Object
Dim oProcessors As Object
Dim iSelectedIndex As New clsIndexer

    Set oProcessors = CreateObject(sObjProcessors)
    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "filter" Then
        sFilterKey = NodSelected.Key
        sGroupKey = NodSelected.Parent.Key
        Set nodY = NodSelected.Parent
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
    Else
        Exit Sub
    End If
    
    Load frmFilterOptimizer
    frmFilterOptimizer.Filter Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey), sFilterKey
    Set frmFilterOptimizer.Stack = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey)
    frmFilterOptimizer.Show
    DoEvents

End Sub

Private Sub mnuNewHistory_Click()

    NewHistory

End Sub

Private Sub NewHistory()
Dim mHistory As clsHistory
Dim nodY As Node
Dim sFile As String

    BrowseHistory
    If dlgBrowse.FileName = "" Then
        'MsgBox "No drawing file found."
    Else
        sFile = dlgBrowse.FileTitle
        'does this history already exist?
        Set mHistory = Lotto.Histories(Left$(sFile, InStr(sFile, ".") - 1))
        If mHistory Is Nothing Then
            Set mHistory = Lotto.Histories.Add(Left$(sFile, InStr(sFile, ".") - 1))
            mHistory.FileName = dlgBrowse.FileName
            mHistory.Name = Left$(sFile, InStr(sFile, ".") - 1)
            Set nodY = tvFolders.Nodes.Add("root", tvwChild, mHistory.Name, sFile & " (History)", "unloadedhistory")
            nodY.EnsureVisible
            Set tvFolders.SelectedItem = nodY
        Else
            MsgBox "History File Already Exists."
        End If
    End If


End Sub

Private Sub mnuOverview_Click()

    If Not picLeft.Visible Then picLeft.Visible = True

End Sub

Private Sub mnuPairHits_Click()

    'If Lotto.Drawings.Count > 0 Then
    '    Lotto.Statistics.Calc_Pair_Scans
    '    frmPairPlot.Show
    'Else
    '    MsgBox "No drawings loaded."
    'End If

End Sub

Private Sub mnuPlotFilter_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim NodSelected As Node
Dim nodY As Node
Dim frmSettings As Form
Dim mProperties As Object

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "filter" Then
        sFilterKey = NodSelected.Key
        sGroupKey = NodSelected.Parent.Key
        Set nodY = NodSelected
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
    Else
        Exit Sub
    End If

    Load frmIndividualPlot
    Set frmIndividualPlot.Filter = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey)
    frmIndividualPlot.Icon = imglstCoolIconsSmall.ListImages.Item("filter").ExtractIcon
    frmIndividualPlot.Show

End Sub

Private Sub mnuResults_Click()

    WHEEL8

End Sub

Private Sub mnuPlotStack_Click()
Dim mHistory As clsHistory
Dim sHistoryKey As String
Dim sMachineKey As String
Dim NodSelected As Node

    Set NodSelected = tvFolders.SelectedItem
    sMachineKey = NodSelected.Parent.Key
    sHistoryKey = NodSelected.Parent.Parent.Key
    Load frmIndividualPlot
    Set frmIndividualPlot.Stack = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks.Item(NodSelected.Key)
    frmIndividualPlot.Icon = imglstCoolIconsSmall.ListImages.Item("stack").ExtractIcon
    frmIndividualPlot.Show

End Sub

Private Sub mnuRemoveFilter_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim NodSelected As Node
Dim sStackKey As String
Dim sGroupKey As String
Dim nodY As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "filter" Then
        sGroupKey = NodSelected.Parent.Key
        Set nodY = NodSelected
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
        Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).RemoveFilter (NodSelected.Key)
        tvFolders.Nodes.Remove (NodSelected.Key)
    Else
        Exit Sub
    End If

End Sub

Private Sub mnuRemoveGroup_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim NodSelected As Node
Dim sStackKey As String
Dim sParentGroupKey As String
Dim nodY As Node
Dim nodOutput As Node
Dim nodSelector As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "group" Then
        Set nodY = NodSelected
        If Get_Node_Type(NodSelected.Parent) = "group" Then
            sParentGroupKey = NodSelected.Parent.Key
        End If
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
        Set nodOutput = tvFolders.Nodes("output" & sStackKey)
        'Set nodSelector = tvFolders.Nodes("selector" & sStackKey)
        If Get_Node_Type(NodSelected.Child.LastSibling) = "group" Then
            'uhoh, we have more groups below us...
            'so we must move this group up one level to the parent of this group
            Set NodSelected.Child.LastSibling.Parent = NodSelected.Parent
            For i = 1 To NodSelected.Parent.Children - 1
                Set NodSelected.Parent.Child.LastSibling.Parent = NodSelected.Parent
            Next i
        End If
        If nodOutput.Parent.Key = NodSelected.Key Then
            'this group is last group of the stack
            If Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups.Count = 1 Then
                'this is the only group...
                'so just remove the nodes from the tree
                tvFolders.Nodes.Remove (nodOutput.Key)
                'tvFolders.Nodes.Remove (nodSelector.Key)
            Else
                'multiple groups...
                'so move the output and selector to the parent group
                Set nodOutput.Parent = NodSelected.Parent
                'Set nodSelector.Parent = nodSelected.Parent
            End If
        End If
        Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups.Remove (NodSelected.Key)
        tvFolders.Nodes.Remove (NodSelected.Key)
        If Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups.Count > 0 Then
            If Get_Node_Type(nodOutput.Parent.Child.LastSibling) = "filter" Or Get_Node_Type(nodOutput.Parent.Child.LastSibling) = "selector" Then
                'uhoh, we have more groups below us...
                'so we must move this group up one level to the parent of this group
                For i = 1 To nodOutput.Parent.Children - 1
                    Set nodOutput.Parent.Child.LastSibling.Parent = nodOutput.Parent
                Next i
            End If
        End If
    Else
        Exit Sub
    End If

End Sub

Private Sub mnuRemoveHistory_Click()
Dim nodClicked As Node

    Set nodClicked = tvFolders.SelectedItem
    If Not nodClicked Is Nothing Then
        If Get_Node_Type(nodClicked) = "history" Then
            Lotto.Histories.Remove (nodClicked.Key)
            tvFolders.Nodes.Remove (nodClicked.Key)
        End If
    End If

End Sub

Private Sub mnuRemoveProcessor_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim NodSelected As Node
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim nodY As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "processor" Then
        sFilterKey = NodSelected.Parent.Key
        sGroupKey = NodSelected.Parent.Parent.Key
        Set nodY = NodSelected
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
        Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).RemoveProcessor (LCase(NodSelected.Text))
        tvFolders.Nodes.Remove (NodSelected.Key)
    Else
        Exit Sub
    End If

End Sub

Private Sub mnuRemoveStack_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim NodSelected As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "stack" Then
        sMachineKey = NodSelected.Parent.Key
        sHistoryKey = NodSelected.Parent.Parent.Key
        Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks.Remove (NodSelected.Key)
        tvFolders.Nodes.Remove (NodSelected.Key)
    Else
        Exit Sub
    End If

End Sub

Private Sub mnuRemoveTrigger_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim NodSelected As Node
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim nodY As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "trigger" Then
        sFilterKey = NodSelected.Parent.Key
        sGroupKey = NodSelected.Parent.Parent.Key
        Set nodY = NodSelected
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
        Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).RemoveTrigger (NodSelected.Key)
        tvFolders.Nodes.Remove (NodSelected.Key)
    Else
        Exit Sub
    End If

End Sub

Private Sub mnuRunHistory_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim NodSelected As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "stack" Then
        sStackKey = NodSelected.Key
        sMachineKey = NodSelected.Parent.Key
        sHistoryKey = NodSelected.Parent.Parent.Key
    Else
        Exit Sub
    End If
    Set Lotto.Histories.RunningHistory = Lotto.Histories(sHistoryKey)
    Set Lotto.Histories(sHistoryKey).RunningMachine = Lotto.Histories(sHistoryKey).Machines(sMachineKey)
    Set Lotto.Histories(sHistoryKey).Machines(sMachineKey).RunningStack = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey)
    Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Reset
    'Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).TestStack

End Sub

Private Sub mnuRunStack_Click()
Dim cStack As clsStack

    Set cStack = GetStackFromTree(tvFolders.SelectedItem)
    If Not cStack Is Nothing Then
        cStack.Reset
        Set frmViewOutput.Stack = cStack
        cStack.TestStack
    End If
    Set cStack = Nothing
    
End Sub

Private Sub mnuSaveStack_Click()
Dim mHistory As clsHistory
Dim sFile As String
Dim sHistoryKey As String
Dim sMachineKey As String
Dim NodSelected As Node
Dim nodY As Node

    BrowseSaveStack
    sFile = dlgBrowse.FileName
    If sFile = "" Then
        MsgBox "No filename specified."
    Else
        Set NodSelected = tvFolders.SelectedItem
        sMachineKey = NodSelected.Parent.Key
        sHistoryKey = NodSelected.Parent.Parent.Key
        Screen.MousePointer = 11
            SaveStack Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(NodSelected.Key), Lotto.Histories(sHistoryKey).FileName, sFile
        Screen.MousePointer = 0
    End If

End Sub

Private Sub SaveStack(oStack As clsStack, sHistoryName As String, sSaveName As String)
Dim SXML        As New CGoXML
Dim PropVals    As Object
Dim Prop        As Object
Dim iGroup      As Integer
Dim iFilter     As Integer
Dim iProcessor  As Integer

    SXML.Initialize (pavAUTO)
    'START INITIAL FILE TEMPLATE
    Call SXML.OpenFromString("<?xml version=" & Chr$(34) & "1.0" & Chr$(34) & " ?>" & vbCrLf & "<STACK>" & vbCrLf & "</STACK>")

    If Not SXML.InsertNode("/STACK", "INITIAL_HISTORY_FILE", sHistoryName) Then Exit Sub
    If Not SXML.InsertNode("/STACK", "GROUPS") Then Exit Sub
    For iGroup = 0 To oStack.Groups.Count - 1
        With oStack.Groups.Item(iGroup + 1)
            If Not SXML.InsertNode("/STACK/GROUPS", "GROUP", "", "usegroup", .UseGroup) Then Exit Sub
            If Not SXML.InsertNode("/STACK/GROUPS/GROUP", "GROUP_NAME", .Name) Then Exit Sub
            If .FilterCount > 0 Then
                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]", "FILTERS") Then Exit Sub
                For iFilter = 0 To .FilterCount - 1
                    With .FilterItem(iFilter + 1)
                        If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS", "FILTER", "", "usefilter", .UseFilter) Then Exit Sub
                        If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]", "FILTER_NAME", .Name) Then Exit Sub
                        If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]", "PROPERTIES") Then Exit Sub
                        Set PropVals = .PropertyValues
                        For Each Prop In PropVals
                            Select Case Prop.Group
                                Case Is < 100
                                    'textbox input
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Sub
                                Case 100, 200, 300, 400, 500, 600, 700, 800, 900
                                    'combo box input
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Sub
                                Case Else
                                    'MsgBox "invalid Property"
                            End Select
                        Next
                        
                        If .ProcessorCount > 0 Then
                            If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]", "PROCESSORS") Then Exit Sub
                            For iProcessor = 0 To .ProcessorCount - 1
                                With .ProcessorItem(iProcessor + 1)
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS", "PROCESSOR", "", "useprocessor", .UseProcessor) Then Exit Sub
                                    If Not SXML.WriteAttribute("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]", "keyname", .Key) Then Exit Sub
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]", "PROCESSOR_NAME", .Name) Then Exit Sub
                                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]", "PROPERTIES") Then Exit Sub
                                    Set PropVals = .PropertyValues
                                    For Each Prop In PropVals
                                        Select Case Prop.Group
                                            Case Is < 100
                                                'textbox input
                                                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Sub
                                            Case 100, 200, 300, 400, 500, 600, 700, 800, 900
                                                'combo box input
                                                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/FILTERS/FILTER[" & iFilter & "]/PROCESSORS/PROCESSOR[" & iProcessor & "]/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Sub
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
                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]", "SELECTOR", "", "keyname", .Key) Then Exit Sub
                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/SELECTOR", "SELECTOR_NAME", .Name) Then Exit Sub
                    If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/SELECTOR", "PROPERTIES") Then Exit Sub
                    Set PropVals = .PropertyValues
                    For Each Prop In PropVals
                        Select Case Prop.Group
                            Case Is < 100
                                'textbox input
                                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/SELECTOR/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Sub
                            Case 100, 200, 300, 400, 500, 600, 700, 800, 900
                                'combo box input
                                If Not SXML.InsertNode("/STACK/GROUPS/GROUP[" & iGroup & "]/SELECTOR/PROPERTIES", "PROPERTY", Prop.Value, "keyname", Prop.Key) Then Exit Sub
                            Case Else
                                'MsgBox "invalid Property"
                        End Select
                    Next
                End With
            End If
        End With
    Next iGroup
    If Not SXML.Save(sSaveName) Then Exit Sub
    
End Sub

Private Sub mnuScanHistory_Click()
Dim frmNewForm As Form
Dim sHistoryKey As String

    sHistoryKey = tvFolders.SelectedItem.Key
    Set frmNewForm = NewMulti("frmscan")
    Set frmNewForm.History = Lotto.Histories(sHistoryKey)
    frmNewForm.Caption = frmNewForm.Caption & " - " & sHistoryKey
    colWindows.Add frmNewForm

End Sub

Private Sub mnuSingleHits_Click()
'Dim iGetNum As Integer
'Dim iHigh As Integer
'Dim iLow As Integer
'Dim iAvg As Long
'
'    If Lotto.Drawings.Count > 0 Then
'        frmChart.Show
'          frmChart.Caption = "Single Hit Statistics"
'          Lotto.Statistics.Calc_Single_Scans
'          iLow = 10000
'          For i = 1 To Lotto.Drawings.BallCount
'            If Lotto.Statistics.RuleSets.Item(1).SingleHit(i) < iLow Then
'              iLow = Lotto.Statistics.RuleSets.Item(1).SingleHit(i)
'            End If
'            If Lotto.Statistics.RuleSets.Item(1).SingleHit(i) > iHigh Then
'              iHigh = Lotto.Statistics.RuleSets.Item(1).SingleHit(i)
'            End If
'            iAvg = iAvg + Lotto.Statistics.RuleSets.Item(1).SingleHit(i)
'          Next i
'          iAvg = iAvg / Lotto.Drawings.BallCount
'          frmChart.chrtPlot.Repaint = False
'          frmChart.chrtPlot.RowCount = Lotto.Drawings.BallCount
'          frmChart.chrtPlot.ColumnCount = 3
'          frmChart.chrtPlot.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = iHigh + 3
'          frmChart.chrtPlot.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = iLow
'          frmChart.chrtPlot.Column = 2
'          For i = 1 To Lotto.Drawings.BallCount
'              frmChart.chrtPlot.Row = i
'              frmChart.chrtPlot.Data = 500
'          Next i
'          For N = 0 To Lotto.Drawings.Drawn - 1
'              iGetNum = Lotto.Drawings.Item(1).Numbers(N)
'              frmChart.chrtPlot.Row = iGetNum
'              frmChart.chrtPlot.Data = Lotto.Statistics.RuleSets.Item(1).SingleHit(iGetNum)
'          Next N
'
'          frmChart.chrtPlot.Column = 3
'
'          For i = 1 To Lotto.Drawings.BallCount
'              frmChart.chrtPlot.Row = i
'              frmChart.chrtPlot.Data = Lotto.Statistics.RuleSets.Item(1).SingleHit(i)
'          Next i
'
'          frmChart.chrtPlot.Column = 1
'
'          For i = 1 To Lotto.Drawings.BallCount
'              frmChart.chrtPlot.Row = i
'              frmChart.chrtPlot.Data = iAvg
'          Next i
'          frmChart.chrtPlot.Repaint = True
'    Else
'        MsgBox "No drawings loaded."
'    End If
'
End Sub

Private Sub WHEEL8()


'full wheel 8!

'   For h = 1 To 7           'ELIMINATE 1
'     For i = h + 1 To 8     'ELIMINATE 2
'       For j = 1 To 8       'PICK
'         'PICK 6 REMAINING
'         If j <> i And j <> h Then
'            k = k + 1
'         End If
'       Next j
'     Next i
'   Next h
'   MsgBox k / 6
   
   For H = 1 To 2           'ELIMINATE 1
     For i = H + 5 To 8     'ELIMINATE 2
       For j = 1 To 8       'PICK
         'PICK 6 REMAINING
         If j <> i And j <> H Then
            k = k + 1
         End If
       Next j
       MsgBox H & i
     Next i
   Next H
   MsgBox k / 6


End Sub

Private Sub mnuStartAI_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim NodSelected As Node
Dim nodY As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "machine" Then
        sMachineKey = NodSelected.Key
        sHistoryKey = NodSelected.Parent.Key
    Else
        Exit Sub
    End If
    Set nodY = tvFolders.Nodes.Add(NodSelected.Key, tvwChild, "ai" & Lotto.GenerateKey, "AI", "thinking")
    Set Lotto.Histories(sHistoryKey).Machines(sMachineKey).AILifeForm.Drawings = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Drawings
    nodY.EnsureVisible
    Load frmMonitorAI
    Set frmMonitorAI.Monitor = Lotto.Histories(sHistoryKey).Machines(sMachineKey).AILifeForm
    frmMonitorAI.Show
    DoEvents
    Call Lotto.Histories(sHistoryKey).Machines(sMachineKey).StartAILifeForm

End Sub

Private Sub mnuKillAILifeform_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sAIKey As String
Dim NodSelected As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "ai" Then
        sAIKey = NodSelected.Key
        sMachineKey = NodSelected.Parent.Key
        sHistoryKey = NodSelected.Parent.Parent.Key
    Else
        Exit Sub
    End If
    Lotto.Histories(sHistoryKey).Machines(sMachineKey).AILifeForm.KillLifeForm = True
    tvFolders.Nodes.Remove (NodSelected.Key)
End Sub

Private Sub mnuUseFilter_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim NodSelected As Node
Dim nodY As Node
    
    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "filter" Then
        sFilterKey = NodSelected.Key
        sGroupKey = NodSelected.Parent.Key
        Set nodY = NodSelected.Parent
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
        mnuUseFilter.Checked = Not mnuUseFilter.Checked
        Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).UseFilter = mnuUseFilter.Checked
    End If

End Sub

Private Sub mnuUseGroup_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim nodY As Node
Dim NodSelected As Node
    
    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "group" Then
        sGroupKey = NodSelected.Key
        Set nodY = NodSelected
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
        mnuUseGroup.Checked = Not mnuUseGroup.Checked
        Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).UseGroup = mnuUseGroup.Checked
    End If

End Sub

Private Sub mnuUseHistory_Click()
Dim sHistoryKey As String
Dim NodSelected As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "history" Then
        sHistoryKey = NodSelected.Key
        mnuUseHistory.Checked = Not mnuUseHistory.Checked
        Lotto.Histories(sHistoryKey).UseHistory = mnuUseHistory.Checked
    End If

End Sub

Private Sub mnuUseProcessor_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim sProcessorKey As String
Dim NodSelected As Node
Dim nodY As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "processor" Then
        sProcessorKey = LCase(NodSelected.Text)
        sFilterKey = NodSelected.Parent.Key
        sGroupKey = NodSelected.Parent.Parent.Key
        Set nodY = NodSelected
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
        mnuUseProcessor.Checked = Not mnuUseProcessor.Checked
        Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).ProcessorItem(sProcessorKey).UseProcessor = mnuUseProcessor.Checked
    End If

End Sub

Private Sub mnuUseStack_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim NodSelected As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "stack" Then
        sStackKey = NodSelected.Key
        sMachineKey = NodSelected.Parent.Key
        sHistoryKey = NodSelected.Parent.Parent.Key
        mnuUseStack.Checked = Not mnuUseStack.Checked
        Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).UseStack = mnuUseStack.Checked
    End If

End Sub

Private Sub mnuUseTrigger_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim sTriggerKey As String
Dim NodSelected As Node
Dim nodY As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "processor" Then
        sProcessorKey = NodSelected.Key
        sFilterKey = NodSelected.Parent.Key
        sGroupKey = NodSelected.Parent.Parent.Key
        Set nodY = NodSelected
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
        mnuUseTrigger.Checked = Not mnuUseTrigger.Checked
        Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).TriggerItem(sTriggerKey).UseTrigger = mnuUseTrigger.Checked
    End If

End Sub

Private Sub mnuViewFilterSettings_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim NodSelected As Node
Dim nodY As Node
Dim frmSettings As Form
Dim mProperties As Object

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "filter" Then
        sFilterKey = NodSelected.Key
        sGroupKey = NodSelected.Parent.Key
        Set nodY = NodSelected
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
    Else
        Exit Sub
    End If
    If colWindows.Exist("frmviewsettings") Then
        Set frmSettings = colWindows.Item("frmviewsettings")
        frmSettings.mvSettings.Clear
        frmSettings.Caption = "View/Edit Settings - " & NodSelected.Text
    Else
        Load frmViewSettings
        colWindows.Add frmViewSettings
        frmViewSettings.Caption = "View/Edit Settings - " & NodSelected.Text
        Set frmSettings = frmViewSettings
    End If
    'get the property values collection from the processor
    'and load them into the control
    frmSettings.mvSettings.InputQty = mvMulti
    frmSettings.mvSettings.InputType = mvNumeric
    Set mProperties = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).PropertyValues
    frmSettings.Icon = imglstCoolIconsSmall.ListImages.Item("filter").ExtractIcon
    frmSettings.mvSettings.LoadViewer mProperties

End Sub

Private Sub mnuViewHistorySettings_Click()
Dim sHistoryKey As String
Dim NodSelected As Node
Dim frmSettings As frmViewSettings
Dim mProperties As Object

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "history" Then
        sHistoryKey = NodSelected.Key
    Else
        Exit Sub
    End If
    If colWindows.Exist("frmviewsettings") Then
        Set frmSettings = colWindows.Item("frmviewsettings")
        frmSettings.mvSettings.Clear
        frmSettings.Caption = "View/Edit Settings - " & NodSelected.Text
    Else
        Load frmViewSettings
        colWindows.Add frmViewSettings
        frmViewSettings.Caption = "View/Edit Settings - " & NodSelected.Text
        Set frmSettings = frmViewSettings
    End If
    'get the property values collection from the processor
    'and load them into the control
    frmSettings.mvSettings.InputQty = mvMulti
    frmSettings.mvSettings.InputType = mvNumeric
    'Set mProperties = Lotto.Histories(sHistoryKey).PropertyValues
    frmSettings.mvSettings.LoadViewer mProperties

End Sub

Private Sub mnuViewLifeform_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sAIKey As String
Dim NodSelected As Node

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "ai" Then
        sAIKey = NodSelected.Key
        sMachineKey = NodSelected.Parent.Key
        sHistoryKey = NodSelected.Parent.Parent.Key
    Else
        Exit Sub
    End If
    Load frmMonitorAI
    Set frmMonitorAI.Monitor = Lotto.Histories(sHistoryKey).Machines(sMachineKey).AILifeForm
    frmMonitorAI.Show

End Sub

Private Sub mnuViewProcessorSettings_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim sProcessorKey As String
Dim NodSelected As Node
Dim nodY As Node
Dim frmSettings As frmViewSettings
Dim mProperties As Object

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "processor" Then
        sProcessorKey = NodSelected.Tag
        sFilterKey = NodSelected.Parent.Key
        sGroupKey = NodSelected.Parent.Parent.Key
        Set nodY = NodSelected
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
    Else
        Exit Sub
    End If
    If colWindows.Exist("frmviewsettings") Then
        Set frmSettings = colWindows.Item("frmviewsettings")
        frmSettings.mvSettings.Clear
        frmSettings.Caption = "View/Edit Settings - " & NodSelected.Text
    Else
        Load frmViewSettings
        colWindows.Add frmViewSettings
        frmViewSettings.Caption = "View/Edit Settings - " & NodSelected.Text
        Set frmSettings = frmViewSettings
    End If
    'get the property values collection from the processor
    'and load them into the control
    frmSettings.mvSettings.InputQty = mvMulti
    frmSettings.mvSettings.InputType = mvNumeric
    Set mProperties = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).ProcessorItem(sProcessorKey).PropertyValues
    frmSettings.Icon = imglstCoolIconsSmall.ListImages.Item("processor").ExtractIcon
    frmSettings.mvSettings.LoadViewer mProperties

End Sub

Private Sub mnuViewSelectorSettings_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim NodSelected As Node
Dim nodY As Node
Dim frmSettings As frmViewSettings
Dim mProperties As Object

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "selector" Then
        sGroupKey = NodSelected.Parent.Key
        Set nodY = NodSelected.Parent
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
    Else
        Exit Sub
    End If
    If colWindows.Exist("frmviewsettings") Then
        Set frmSettings = colWindows.Item("frmviewsettings")
        frmSettings.mvSettings.Clear
        frmSettings.Caption = "View/Edit Settings - " & NodSelected.Text
    Else
        Load frmViewSettings
        colWindows.Add frmViewSettings
        frmViewSettings.Caption = "View/Edit Settings - " & NodSelected.Text
        Set frmSettings = frmViewSettings
    End If
    'get the property values collection from the processor
    'and load them into the control
    frmSettings.mvSettings.InputQty = mvMulti
    frmSettings.mvSettings.InputType = mvNumeric
    Set mProperties = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).Selector.PropertyValues
    frmSettings.Icon = imglstCoolIconsSmall.ListImages.Item("selector").ExtractIcon
    frmSettings.mvSettings.LoadViewer mProperties

End Sub

Private Sub Lotto_SingleStackComplete(ByVal sHistoryName As String, ByVal sMachineKey As String, ByVal sStackKey As String)
Dim addsting As String
Dim lPredictedCount As Long

    DoEvents
    If Lotto.Histories(sHistoryName).Machines(sMachineKey).Stacks(sStackKey).Output.PredictDrawing = 0 Then
        For i = 1 To Lotto.Histories(sHistoryName).Machines(sMachineKey).MachineMaximumBallNumber
            With Lotto.Histories(sHistoryName).Machines(sMachineKey).Stacks(sStackKey).Output
                'the prediction routines actually work to EXCLUDE balls
                If Not .Excluded(i) Then
                    lPredictedCount = lPredictedCount + 1
                    If addstring = "" Then
                        addstring = i
                    Else
                        addstring = addstring & ", " & i
                    End If
                End If
            End With
        Next i
        addstring = lPredictedCount & " Predictions for the following week: " & addstring & vbCrLf
        frmViewOutput.txtOutput.Text = addstring & frmViewOutput.txtOutput.Text
        DoEvents
    End If

End Sub

Private Sub mnuViewTriggerSettings_Click()
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim sTriggerKey As String
Dim NodSelected As Node
Dim nodY As Node
Dim frmSettings As frmViewSettings
Dim mProperties As Object

    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "trigger" Then
        sTriggerKey = NodSelected.Tag
        sFilterKey = NodSelected.Parent.Key
        sGroupKey = NodSelected.Parent.Parent.Key
        Set nodY = NodSelected
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
    Else
        Exit Sub
    End If
    If colWindows.Exist("frmviewsettings") Then
        Set frmSettings = colWindows.Item("frmviewsettings")
        frmSettings.mvSettings.Clear
        frmSettings.Caption = "View/Edit Settings - " & NodSelected.Text
    Else
        Load frmViewSettings
        colWindows.Add frmViewSettings
        frmViewSettings.Caption = "View/Edit Settings - " & NodSelected.Text
        Set frmSettings = frmViewSettings
    End If
    'get the property values collection from the processor
    'and load them into the control
    frmSettings.mvSettings.InputQty = mvMulti
    frmSettings.mvSettings.InputType = mvNumeric
    Set mProperties = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).TriggerItem(sTriggerKey).PropertyValues
    frmSettings.Icon = imglstCoolIconsSmall.ListImages.Item("trigger").ExtractIcon
    frmSettings.mvSettings.LoadViewer mProperties

End Sub

Private Sub picLeft_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If x > (picLeft.Width - 60) And y > 300 Then
        picLeft.MousePointer = 9
    Else
        picLeft.MousePointer = 0
    End If
    If Button = 1 And picLeft.MousePointer = 9 Then
        If x > 1300 Then
            picLeft.Width = x
        End If
    End If

End Sub

Private Sub picLeft_Resize()
Dim iWidth As Integer
Dim iHeight As Integer

    iWidth = picLeft.Width
    iHeight = picLeft.Height
    If iWidth > 1300 Then
        lblTitle.Width = iWidth - 60
        cmdX.Left = iWidth - 280
        tsOverview.Width = iWidth - 90
    End If
    If iHeight > 3000 Then
        tsOverview.Height = picLeft.Height - 380
    End If
    
    picTreeArea.Width = tsOverview.ClientWidth
    picTreeArea.Height = tsOverview.ClientHeight - 300

End Sub

Private Sub picTreeArea_Resize()

    tvFolders.Top = picTreeArea.ScaleTop
    tvFolders.Left = picTreeArea.ScaleLeft
    tvFolders.Width = picTreeArea.ScaleWidth
    tvFolders.Height = picTreeArea.ScaleHeight

End Sub

Private Sub tvFolders_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim NodSelected As Node
Dim nodY As Node
Dim frmSettings As Form

    Set NodSelected = tvFolders.SelectedItem
    Set nodY = tvFolders.SelectedItem
    Do
        Select Case Left$(nodY.Key, 3)
            Case "fil"
                sFilterKey = nodY.Key
            Case "mac"
                sMachineKey = nodY.Key
            Case "sta"
                sStackKey = nodY.Key
            Case "gro"
                If sGroupKey = "" Then
                    sGroupKey = nodY.Key
                End If
            Case "pro"
                sProcessorKey = nodY.Key
            Case "his"
                sHistoryKey = nodY.Key
                Exit Do
            Case "tri"
                sTriggerKey = nodY.Key
        End Select
        Set nodY = nodY.Parent
    Loop

    'We will need to find out where this node is in the tree!
    Select Case Left$(NodSelected.Key, 3)
        Case "his"
            Lotto.Histories(sHistoryKey).Name = NewString
        Case "mac"
            Lotto.Histories(sHistoryKey).Machines(sMachineKey).Name = NewString
        Case "sta"
            Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Name = NewString
        Case "gro"
            Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).Name = NewString
        Case "fil"
            Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).Name = NewString
        Case "pro"
            Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).ProcessorItem(sProcessorKey).Name = NewString
        Case "tri"
            Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).TriggerItem(sTriggerKey).Name = NewString
    End Select

End Sub

Private Sub tvFolders_BeforeLabelEdit(Cancel As Integer)

    Select Case Left$(tvFolders.SelectedItem.Key, 3)
        Case "roo", "out", "sel"
            'don't allow label editing to the
            'root
            'output
            'OR selection nodes
            Cancel = 1
    End Select

End Sub

Private Sub tvFolders_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If picLeft.MousePointer = 9 Then picLeft.MousePointer = 0

End Sub

Private Sub tvFolders_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim NodSelected As Node
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim sFilterKey As String
Dim sProcessorKey As String
Dim sTriggerKey As String
Dim sAIKey As String
Dim nodY As Node

    If Button = vbRightButton Then
        Set NodSelected = tvFolders.HitTest(x, y)
        If Not NodSelected Is Nothing Then
            Select Case Get_Node_Type(NodSelected)
                Case "root"
                    PopupMenu mnuWinHunter
                Case "history"
                    mnuScanHistory.Enabled = True
                    mnuRunHistory.Enabled = True
                    mnuHistoryLoaded.Visible = True
                    mnuHistoryLoaded.Enabled = False
                    mnuHistoryLoaded.Checked = True
                    mnuLoadHistoryFile.Visible = False
                    mnuGetPredictions.Enabled = True
                    sHistoryKey = NodSelected.Key
                    mnuUseHistory.Enabled = True
                    mnuUseHistory.Checked = Lotto.Histories(sHistoryKey).UseHistory
                    mnuScanHistory.Enabled = mnuUseHistory.Checked
                    PopupMenu mnuHistory
                Case "unloadedhistory"
                    mnuScanHistory.Enabled = False
                    mnuRunHistory.Enabled = False
                    mnuHistoryLoaded.Checked = False
                    mnuHistoryLoaded.Visible = False
                    mnuLoadHistoryFile.Visible = True
                    mnuGetPredictions.Enabled = False
                    mnuUseHistory.Enabled = False
                    mnuUseHistory.Checked = False
                    PopupMenu mnuHistory
                Case "machine"
                    sMachineKey = NodSelected.Key
                    sHistoryKey = NodSelected.Parent.Key
                    If Lotto.Histories(sHistoryKey).Machines(sMachineKey).Drawings.Drawn > 1 Then
                        mnuPairHits.Visible = True
                        mnuSep1.Visible = True
                    Else
                        mnuPairHits.Visible = False
                        mnuSep1.Visible = False
                    End If
                    PopupMenu mnuMachine
                Case "stack"
                    sStackKey = NodSelected.Key
                    sMachineKey = NodSelected.Parent.Key
                    sHistoryKey = NodSelected.Parent.Parent.Key
                    mnuUseStack.Checked = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).UseStack
                    PopupMenu mnuStack
                Case "group"
                    sGroupKey = NodSelected.Key
                    Set nodY = NodSelected
                    Do
                        Set nodY = nodY.Parent
                    Loop Until Get_Node_Type(nodY) = "stack"
                    sStackKey = nodY.Key
                    sMachineKey = nodY.Parent.Key
                    sHistoryKey = nodY.Parent.Parent.Key
                    mnuUseGroup.Checked = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).UseGroup
                    PopupMenu mnuGroup
                Case "filter"
                    sFilterKey = NodSelected.Key
                    sGroupKey = NodSelected.Parent.Key
                    Set nodY = NodSelected.Parent
                    Do
                        Set nodY = nodY.Parent
                    Loop Until Get_Node_Type(nodY) = "stack"
                    sStackKey = nodY.Key
                    sMachineKey = nodY.Parent.Key
                    sHistoryKey = nodY.Parent.Parent.Key
                    mnuUseFilter.Checked = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).UseFilter
                    PopupMenu mnuFilter
                Case "processor"
                    sProcessorKey = NodSelected.Tag
                    sFilterKey = NodSelected.Parent.Key
                    sGroupKey = NodSelected.Parent.Parent.Key
                    Set nodY = NodSelected
                    Do
                        Set nodY = nodY.Parent
                    Loop Until Get_Node_Type(nodY) = "stack"
                    sStackKey = nodY.Key
                    sMachineKey = nodY.Parent.Key
                    sHistoryKey = nodY.Parent.Parent.Key
                    mnuUseProcessor.Checked = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).ProcessorItem(sProcessorKey).UseProcessor
                    PopupMenu mnuProcessor
                Case "trigger"
                    sTriggerKey = NodSelected.Tag
                    sFilterKey = NodSelected.Parent.Key
                    sGroupKey = NodSelected.Parent.Parent.Key
                    Set nodY = NodSelected
                    Do
                        Set nodY = nodY.Parent
                    Loop Until Get_Node_Type(nodY) = "stack"
                    sStackKey = nodY.Key
                    sMachineKey = nodY.Parent.Key
                    sHistoryKey = nodY.Parent.Parent.Key
                    mnuUseTrigger.Checked = Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).FilterItem(sFilterKey).TriggerItem(sTriggerKey).UseTrigger
                    PopupMenu mnuTrigger
                Case "selector"
                    PopupMenu mnuSelector
                Case "ai"
                    PopupMenu mnuAILifeForm
            End Select
        End If
    End If
    DoEvents

End Sub

Private Function Get_Node_Type(NodeIn As Node) As String

    Select Case Left$(NodeIn.Key, 3)
        Case "roo"
            Get_Node_Type = "root"
        Case "fil"
            Get_Node_Type = "filter"
        Case "mac"
            Get_Node_Type = "machine"
        Case "sta"
            Get_Node_Type = "stack"
        Case "gro"
            Get_Node_Type = "group"
        Case "pro"
            Get_Node_Type = "processor"
        Case "tri"
            Get_Node_Type = "trigger"
        Case "sel"
            Get_Node_Type = "selector"
        Case Else
            Select Case Left$(NodeIn.Image, 3)
                Case "his"
                    Get_Node_Type = "history"
                Case "unl"
                    Get_Node_Type = "unloadedhistory"
                Case "thi"  'thinking
                    Get_Node_Type = "ai"
            End Select
    End Select

End Function

Private Sub Add_Lottery_Item(NodSelected As Node, ExpectedNode As String)
Dim sHistoryKey As String
Dim sMachineKey As String
Dim sStackKey As String
Dim sGroupKey As String
Dim nodY As Node

    
    Do Until sHistoryKey <> ""
        
    Loop
    
    
    Set NodSelected = tvFolders.SelectedItem
    If Get_Node_Type(NodSelected) = "group" Then
        sGroupKey = NodSelected.Key
        Set nodY = NodSelected
        Do
            Set nodY = nodY.Parent
        Loop Until Get_Node_Type(nodY) = "stack"
        sStackKey = nodY.Key
        sMachineKey = nodY.Parent.Key
        sHistoryKey = nodY.Parent.Parent.Key
    Else
        Exit Sub
    End If
    Set nodY = tvFolders.Nodes.Add(NodSelected.Key, tvwChild, "filter" & Lotto.GenerateKey, "Filter", "filter")
    Set nodY.Parent = NodSelected
    Lotto.Histories(sHistoryKey).Machines(sMachineKey).Stacks(sStackKey).Groups(sGroupKey).AddFilter (nodY.Key)
    nodY.EnsureVisible


End Sub
