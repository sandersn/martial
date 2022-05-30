VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmFrequency 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Martial 0.03"
   ClientHeight    =   6525
   ClientLeft      =   1665
   ClientTop       =   1410
   ClientWidth     =   9390
   Icon            =   "Frequency.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   3720
      TabIndex        =   59
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "<< &Back"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      TabIndex        =   25
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >>"
      Height          =   375
      Left            =   6600
      TabIndex        =   26
      Top             =   6120
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6015
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   10610
      _Version        =   327681
      Style           =   1
      Tabs            =   5
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "&Input"
      TabPicture(0)   =   "Frequency.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblRequired(2)"
      Tab(0).Control(1)=   "lblRequired(0)"
      Tab(0).Control(2)=   "lblInfoWelcome"
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(4)=   "lblMartial"
      Tab(0).Control(5)=   "lblInfoOpen"
      Tab(0).Control(6)=   "lblInfoSampleLength"
      Tab(0).Control(7)=   "lblSampleSize"
      Tab(0).Control(8)=   "lblInfoRequired"
      Tab(0).Control(9)=   "lblRequired(1)"
      Tab(0).Control(10)=   "lblOpen"
      Tab(0).Control(11)=   "txtSampleLength"
      Tab(0).Control(12)=   "cmdOpen"
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Out&put"
      TabPicture(1)   =   "Frequency.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblRequired(3)"
      Tab(1).Control(1)=   "lblResults"
      Tab(1).Control(2)=   "lblTop"
      Tab(1).Control(3)=   "lblInfoOpenOutput"
      Tab(1).Control(4)=   "lblInfoOpenWriteTable"
      Tab(1).Control(5)=   "lblInfoTopResults"
      Tab(1).Control(6)=   "lblOpenOutput"
      Tab(1).Control(7)=   "lblOpenWriteTable"
      Tab(1).Control(8)=   "lblInfoBeepOnFinish"
      Tab(1).Control(9)=   "lblWriteTableRanges"
      Tab(1).Control(10)=   "txtTopResults"
      Tab(1).Control(11)=   "cmdOpenWriteTable"
      Tab(1).Control(12)=   "cmdOpenOutput"
      Tab(1).Control(13)=   "chkBeepOnFinish"
      Tab(1).Control(14)=   "txtWriteTableRanges"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "&Formatting"
      TabPicture(2)   =   "Frequency.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblInfoIgnores"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblInfoOpenExcludeTable"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lblOpenExcludeTable"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "fraComments"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "chkUseComments"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "chkIgnoreSpace"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "chkIgnoreReturn"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "chkIgnoreTab"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdOpenExcludeTable"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).ControlCount=   9
      TabCaption(3)   =   "Analyze"
      TabPicture(3)   =   "Frequency.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblMatchCount"
      Tab(3).Control(1)=   "lblMatches"
      Tab(3).Control(2)=   "lblCalc"
      Tab(3).Control(3)=   "lblAddr"
      Tab(3).Control(4)=   "linCompletion"
      Tab(3).ControlCount=   5
      TabCaption(4)   =   "&Results"
      TabPicture(4)   =   "Frequency.frx":04B2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtResults"
      Tab(4).ControlCount=   1
      Begin VB.TextBox txtWriteTableRanges 
         Height          =   285
         Left            =   -68280
         TabIndex        =   10
         Top             =   4650
         Width           =   2535
      End
      Begin VB.CheckBox chkBeepOnFinish 
         Caption         =   "Beep W&hen Finished"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -74760
         TabIndex        =   11
         Top             =   5640
         Width           =   2295
      End
      Begin VB.CommandButton cmdOpenOutput 
         Caption         =   "Op&en Output File"
         Height          =   375
         Left            =   -74760
         TabIndex        =   7
         Top             =   4560
         Width           =   2175
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open"
         Height          =   375
         Left            =   -74820
         TabIndex        =   2
         Top             =   2040
         Width           =   1275
      End
      Begin VB.TextBox txtSampleLength 
         Height          =   285
         Left            =   -74880
         TabIndex        =   4
         Text            =   "1"
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton cmdOpenWriteTable 
         Caption         =   "Open Output Table"
         Height          =   375
         Left            =   -70680
         TabIndex        =   8
         Top             =   4560
         Width           =   2175
      End
      Begin VB.CommandButton cmdOpenExcludeTable 
         Caption         =   "Open E&xclude Table"
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   3720
         Width           =   1935
      End
      Begin VB.CheckBox chkIgnoreTab 
         Caption         =   "Ignore &Tabs"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   5520
         Width           =   2775
      End
      Begin VB.TextBox txtTopResults 
         Height          =   285
         Left            =   -74760
         MaxLength       =   4
         TabIndex        =   6
         Text            =   "25"
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chkIgnoreReturn 
         Caption         =   "Ignore &Carriage Return-Line Feeds"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   5160
         Width           =   2775
      End
      Begin VB.CheckBox chkIgnoreSpace 
         Caption         =   "Ignore &Spaces"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   4800
         Width           =   2775
      End
      Begin VB.CheckBox chkUseComments 
         Caption         =   "&Use Comments"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1360
      End
      Begin VB.Frame fraComments 
         Enabled         =   0   'False
         Height          =   3015
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   9135
         Begin VB.TextBox txtBeginComment 
            Enabled         =   0   'False
            Height          =   285
            Left            =   360
            TabIndex        =   17
            Top             =   1440
            Width           =   495
         End
         Begin VB.TextBox txtEndComment 
            Enabled         =   0   'False
            Height          =   285
            Left            =   960
            TabIndex        =   18
            Top             =   1440
            Width           =   495
         End
         Begin VB.CheckBox chkReadInsideComments 
            Caption         =   "Read only inside comments"
            Enabled         =   0   'False
            Height          =   615
            Left            =   120
            TabIndex        =   19
            Top             =   1920
            Width           =   1425
         End
         Begin VB.OptionButton optCustom 
            Caption         =   "Custom:"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1200
            Width           =   1335
         End
         Begin VB.OptionButton optThingy 
            Caption         =   "<$"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   960
            Width           =   975
         End
         Begin VB.OptionButton optC 
            Caption         =   "/*"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   480
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optCPP 
            Caption         =   "//"
            Enabled         =   0   'False
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   975
         End
         Begin VB.Label lblInfoComments 
            Caption         =   $"Frequency.frx":04CE
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   1680
            TabIndex        =   57
            Top             =   120
            Width           =   7335
         End
         Begin VB.Label lblBegin 
            AutoSize        =   -1  'True
            Caption         =   "Begin"
            Enabled         =   0   'False
            Height          =   195
            Left            =   360
            TabIndex        =   34
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lblEnd 
            AutoSize        =   -1  'True
            Caption         =   "End"
            Enabled         =   0   'False
            Height          =   195
            Left            =   960
            TabIndex        =   33
            Top             =   240
            Width           =   285
         End
         Begin VB.Label lblCEnd 
            AutoSize        =   -1  'True
            Caption         =   "*/"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1080
            TabIndex        =   32
            Top             =   480
            Width           =   135
         End
         Begin VB.Label lblCPPEnd 
            AutoSize        =   -1  'True
            Caption         =   "newl"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1080
            TabIndex        =   31
            Top             =   720
            Width           =   330
         End
         Begin VB.Label lblThingyEnd 
            AutoSize        =   -1  'True
            Caption         =   ">"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1080
            TabIndex        =   30
            Top             =   960
            Width           =   90
         End
      End
      Begin VB.TextBox txtResults 
         Height          =   5565
         Left            =   -74940
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "Frequency.frx":083A
         Top             =   360
         Width           =   9285
      End
      Begin VB.Label lblWriteTableRanges 
         AutoSize        =   -1  'True
         Caption         =   "Ranges:"
         Height          =   195
         Left            =   -68280
         TabIndex        =   9
         Top             =   4440
         Width           =   600
      End
      Begin VB.Label lblInfoBeepOnFinish 
         Caption         =   "When checked, Martial beeps when the report is ready."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -72120
         TabIndex        =   58
         Top             =   5640
         Width           =   6375
      End
      Begin VB.Label lblOpenExcludeTable 
         AutoSize        =   -1  'True
         Caption         =   "--No File Specified--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   56
         Top             =   4320
         Width           =   1785
      End
      Begin VB.Label lblInfoOpenExcludeTable 
         Caption         =   $"Frequency.frx":084C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2280
         TabIndex        =   55
         Top             =   3480
         Width           =   6975
      End
      Begin VB.Label lblInfoIgnores 
         Caption         =   $"Frequency.frx":09B9
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3240
         TabIndex        =   54
         Top             =   4800
         Width           =   6015
      End
      Begin VB.Label lblOpenWriteTable 
         AutoSize        =   -1  'True
         Caption         =   "--No File Specified--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -70680
         TabIndex        =   52
         Top             =   5160
         Width           =   1785
      End
      Begin VB.Label lblOpenOutput 
         AutoSize        =   -1  'True
         Caption         =   "--No File Specified--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74760
         TabIndex        =   51
         Top             =   5160
         Width           =   1785
      End
      Begin VB.Label lblInfoTopResults 
         Caption         =   $"Frequency.frx":0A4C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   -73200
         TabIndex        =   50
         Top             =   480
         Width           =   7455
      End
      Begin VB.Label lblInfoOpenWriteTable 
         Caption         =   $"Frequency.frx":0B73
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -70680
         TabIndex        =   49
         Top             =   1680
         Width           =   4935
      End
      Begin VB.Label lblInfoOpenOutput 
         Caption         =   $"Frequency.frx":0D71
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   -74760
         TabIndex        =   48
         Top             =   1680
         Width           =   3615
      End
      Begin VB.Label lblOpen 
         AutoSize        =   -1  'True
         Caption         =   "--No File Specified--"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -74760
         TabIndex        =   47
         Top             =   3120
         Width           =   1785
      End
      Begin VB.Label lblRequired 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Index           =   1
         Left            =   -73560
         TabIndex        =   45
         Top             =   1920
         Width           =   210
      End
      Begin VB.Label lblInfoRequired 
         AutoSize        =   -1  'True
         Caption         =   "= Required to begin analyzation process."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   -73440
         TabIndex        =   44
         Top             =   5520
         Width           =   3645
      End
      Begin VB.Label lblSampleSize 
         AutoSize        =   -1  'True
         Caption         =   "Sample &Length(s):"
         Height          =   195
         Left            =   -74820
         TabIndex        =   3
         Top             =   3720
         Width           =   1275
      End
      Begin VB.Label lblInfoSampleLength 
         Caption         =   $"Frequency.frx":0EC6
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -73440
         TabIndex        =   42
         Top             =   3840
         Width           =   5295
      End
      Begin VB.Label lblInfoOpen 
         Caption         =   $"Frequency.frx":0F98
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -73320
         TabIndex        =   41
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Label lblMartial 
         Caption         =   "Martial is named after the Roman poet of the same name."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000028&
         Height          =   255
         Left            =   -74400
         TabIndex        =   40
         Top             =   1200
         Width           =   8655
      End
      Begin VB.Label Label1 
         Caption         =   "Immodicis brevis est aetas et rara senectus."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2895
         Left            =   -68040
         TabIndex        =   39
         Top             =   3000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblInfoWelcome 
         Caption         =   $"Frequency.frx":1070
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   38
         Top             =   480
         Width           =   9135
      End
      Begin VB.Label lblMatchCount 
         Caption         =   "654321"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71520
         TabIndex        =   37
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblMatches 
         Caption         =   "Number of Matches So Far:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72240
         TabIndex        =   36
         Top             =   1800
         Width           =   3570
      End
      Begin VB.Label lblTop 
         AutoSize        =   -1  'True
         Caption         =   "Only Sho&w Top"
         Height          =   195
         Left            =   -74760
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblResults 
         AutoSize        =   -1  'True
         Caption         =   "Results"
         Height          =   195
         Left            =   -73860
         TabIndex        =   35
         Top             =   1005
         Width           =   525
      End
      Begin VB.Label lblCalc 
         Caption         =   "Calculating... Current Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -72240
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   3570
      End
      Begin VB.Label lblAddr 
         AutoSize        =   -1  'True
         Caption         =   "123456"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71520
         TabIndex        =   27
         Top             =   1200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Line linCompletion 
         BorderColor     =   &H00C00000&
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   -75000
         X2              =   -65640
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label lblRequired 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Index           =   0
         Left            =   -73800
         TabIndex        =   43
         Top             =   5400
         Width           =   210
      End
      Begin VB.Label lblRequired 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Index           =   2
         Left            =   -73560
         TabIndex        =   46
         Top             =   3600
         Width           =   210
      End
      Begin VB.Label lblRequired 
         AutoSize        =   -1  'True
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Index           =   3
         Left            =   -73680
         TabIndex        =   53
         Top             =   600
         Width           =   210
      End
   End
   Begin VB.CommandButton cmdAnalyze 
      Caption         =   "&Analyze"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   6120
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "All Files|*.*|Text Files(*.txt)|*.txt"
   End
End
Attribute VB_Name = "frmFrequency"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'by Nathan Sanders, ZackMan to some
'sandersn@hotmail.com
'version 0.01(which is to say, Alpha):going to release this as sort of a test and because people have responded with interest
'to it.
'version 0.02:
'2001:
'14 Jan:
'bunches of new TODO:
'17 Jan:
'OK , looks Like we 're about to finish up! Except for necro's suggestion...Anyway, comments are working, and so is the funky
'ability to specify lengths like 1,3-5,7 etc etc. Still need to test the cases 1-3,5,7 and 1,3,5-7 though.
'and also need to see if comments are messed up with long entries where you read aaaa/* and end up having aaaa/ as a match
'by mistake even though / should have been excluded.
'Later this is sort of fixed, except that for comments longer than one character, you still get matches for all the way up to that character len-1
'eg. comment=~!@ you'll still get matches for abc~! and zabc~ but not bc~!@ (ie the full comment) I need to figure a way around this...
'Done:
'1.fix bugs (the one where not all sample strings are added correctly when reading multiple lengths)
'2.Implement comments--whenever the reader comes across the comment sign, it ignores everything until the end comment sign
    '4 styles: C, C++, Thingy, and custom. With custom I'm thinking max two chars but maybe not. Also, a blank for the second one means
    'stop at newl.
'3.Implement read ONLY inside comments. eg the reader ignores everything until it reaches the comment sign and reads until it hits the end of comment sign
'4.A built in byte counter. But I think I already have this basically..
'5 Context help. In a separate function which you pass text to and it displays it in contexty style. Somehow. I'm considering AoE2 style if Win9x's transparency
'skillz can handle it. Later: This is in and functional but I need to write help for everything now.
'6.More bugs: In ParseSampleLengths, set cStart to SOMETHING after each loop so we don't get stuff forever!
'3.read table file and exclude samples matching table entries
    'this would just require that we read everything after the '=' on each line
    'and stick an if in the 'Top X' loop to exclude matches....
'version 0.03:
'20 Jul:
'I've got a whole lot of optimisation ideas so here they are
'Later: Squashed an existing bug where the top results were not the ones screend for containing spaces/tab/crlfs
'changed report format somewhat and reenabled viewing of \0s
'got hex display working
'21 Jul:Going to implement revised optimisation. It won't be quite as dramatic as previously thought, but will be better. And comments are
'handled faster and more gracefully
'finished working bugs out of optimisation. It seems to be faster in some cases and in others not. More investigation needed
'also added elapsed time with a minimum of trouble
'23 Jul: Added binary search last night but was too tired to describe. Since I can't use pointers or Vectors I began investigating
'others methods of inserting values into an array. I came up with a Sorted Listbox, but it only has capacity for 32K matches, so now I'm
'trying a ListView which has capacity for 2G. yay!
'31 Jul: Whoooo I have been lax about documenting what I've done. That's because it's depressing when nothing
'works right, I guess. Well, I'm about to try yet ANOTHER method at using a hash that's not a ListView. The Listview idea has not gone away,
'however, because it's so cool and good at sorting and stuff and things™. But I don't think I can use it to actually hold ALL data vals
'BTW, this hash is encapsulated in a class. VB is improving!
'1 Aug: Well, the hash was SLOOW. So I've fallen back on the idea of the 3D array: 1st dimension is 1 2D array for each length. 2nd dimension
'is 256 long, 1 1D array for each possible Chr$() value (err Ascii vals 0-255). I hope this works...
'2 Aug: It works, but is deadly on memory. I'm going for a 2D array of arrays instead of a true 3D array to save memory. You can do this in
'Java, but doubts exist as for VB...(later)after searching Help, it should work!!
'ITOWKRDSSS!! Got to put it through the torture test now..... -_- but FIRST!, I'm going to add the Len(Char) to the report for happiness and stuff
'and things™
'passed torture test in 6:30!!!! This version is goood. Going to do some other stuff that won't take very long
'3 Aug: Going to start conversion to 'Wizard Version'
'4 Aug: Almost done with Wizard conversion. Left is: 1)Add new controls for new functionality 2)Add new help
'1 Sep: Back to work on Martial. Going to try to finish this up pretty soon.
'8 Oct: Doing a little work since netbeans crashed (!)(actually just threw repeated index oob exceptions)
'9 Oct: Got the add matches to table file feature *almost* working...a few bugs left in it.
'3.26: I have fixed all the bugs and squashed one in ParseWordTextStyle from last version: when you had
'1-3,4-6,7-9 etc  the 3 and 6 parsed with commas like so "3," "6," but I never noticed because Val() just looks until it can't find anymore numbers
'however, when parsing hex, it has to see the whole number so your parsing must be PERFECT. Well, now it is...
'10 Oct:Tested the write report function--it works perfectly. Also caught that I wasn't closing the file 0_o so I added that
'Now I'm going to add the multiple files ability
'10.39: That's done, so I'm going to package this up pretty soon. For now, I have to take a Data Structures test, though
'17 Oct: Having discovered binary trees, I'm going to try them in Martial to see if they're faster.

'TODO:
'17.-- Next version-- Make a .lng file that is a Sequential file which can be translated easily


'Done:
'1.Implement main optimisation::
    'I'm going to divide hufSeq (the matches found) into a 3D array where the first dimension is the length of the sample and 2nd is ascii val
    'of first character in string(assumes strlen > 0 ^_^)
'that way, even the worst case where all matches are new is only as slow as the old method. However, every match found for each length
'speeds us up because we don't have to check all the remaining matches *for that length and ascii val*. The addition of the ascii val is very
'important because it reduces most lists to be very short compared to their previous lengths. That way, they are much faster to search. However,
'the 'find top X matches' is a little slower because of that, but since it takes not very long anyway, it's no big deal
'2.Show more feedback:show the open file (pref in the title bar) and what's happening when.
    'also make a cool animated 'completion bar' which really wouldn't have to be updated very often since it's so cool
'3.Add a filetype to the Open dialog!!
'4.read table file and exclude samples matching table entries
    'this would just require that we read everything after the '=' on each line
    'and stick an if in the 'Top X' loop to exclude matches....
'5.Update help and add help (maybe) to the calculation status label
'6.Maybe add more info per match, like hex (helpful for the blocks that show up a lot) and strlen()
'7.Add ability to add matches to a table file
'8.Add ability to write LONG report to file in addition to the truncated version in the text box
'9.Work on tab order cause it seems like the added controls aren't at all right
'10.Record the elapsed time that the search took and display with the report
'11.Add a 'beep on finish' option
'12.Maybe turn some/all of the options into menus--this turned into a Wizard GUI remake :)
'14.Change code in cmdOpen so that it gives an error message for non CancelError messages
'15.Test comments to make sure they still work and test exclusion of tables
'16. On the download page, include all necessary .ocx
'18. Add ability to 'open' multiple files for batch analyzation. It won't be very hard--just
'make the lblOpen longer and make an array of strFilename. Then create another loop outside the Do which
'Fors through the filenames opening and closing them. Easy!
'19. Add a reset button since you might need to reset the files if you change your mind

'2000:
'5 Dec:
'Done:
'1. Combine Analyze/Cancel button
'4.Fix that pesky 0 terminating the string bug.
    'trap anything with 0 in it and substitute something inoccuous
'5.calc BEST (top 10?) set of substrings(of specific or variable length)
    ' for a given substring block(sample length? or text analyzed?)
    'a substring block is a certain number of substrings(of specific or variable lengths)
    'so to accomplish this all you need is a 'Show Top ___ Results' with another label "Enter the number of substrings
    'available to you here" (or just rollover help or something)
    'WAIT! You have to mult the sample len times the occurence to calc the total space savings!! So a len 10 of 15 times
    'is better than a len 2 of 70 times!
'7.Allow ranges of sample lengths
'Notes:

'TODO:

'2.xx no demand..yet xx possibly show current address button in hex (contigent on user demand)
' xx not needed xx 6.Also mebbe get rid of double counting somehow(this might be a problem with frequency analysation of spaces especially since
    'for Sample Len=2, the code counts each space in a run at least 2 long instead of the number of double-spaces in each run
    'at least 2 long...have to look into this, eh? (also might have troubles with double counting somehow at the end of runs too)
    '6b.or just get rid of subset samples: ex. if you have lens 2-3, if you count "*.*" remove all instances of "*." and ".*" because
    'that optimisation would be pointless and displaying them is just confusing to the average dumb ROM hacker.
    'Decided this wasn't needed
    
Dim intFileno As Integer       'for main file
Dim strFilename() As String   'for main file(s)
Dim strFilenameTableExclude As String   'for exclude table
Dim intFilenoReport As Integer              'for report
Dim strFilenameReport As String             'meme
Dim intFilenoTableWrite As Integer      'for write table
Dim strFilenameTableWrite As String     'mismo

Dim bCancel As Boolean      'whether to cancel mid-read
Dim bComments As Boolean    'whether to use comments or not
Dim bTableExclude As Boolean    'whether to exclude table values or not
Dim bTableWrite As Boolean  'whether to write table values to file or not
Dim bOutput As Boolean  'whether to write report to file or not
Dim bBeep As Boolean    'whether or not to beep when finished

Dim intLens() As Integer    'a global (ugh can't pass variable len array :( that keeps the sample lengths requested for analysation
Dim hufTableVals() As Freq 'a list of the table values that should be skipped
Dim maxFreq As Freq 'a global(don't want to waste mem by using recursive var)
Dim maxNode As HufNode  'and the pointer to the max so we can zero it out

Private Type Freq
    Char As String
    Frequency As Long
End Type
Private Type FreqAra    'This is a HACK to have a 2D array of arrays rather than a true 3D array which wastes too much memory
    hufAra() As Freq        'native in Java, but not in VB
End Type
'these are codes to indicate that this match should be skipped for containing an unwanted char
Private Const vbSpacePresent = -256
Private Const vbTabPresent = -128
Private Const vbCrLfPresent = -64
Private Const vbTableValPresent = -32
'parses strLen 'word style' and stores number results in a variable length int array called intLens
'returns the number of matches found
Private Function ParseTextWordStyle(strLen As String, intLens() As Integer) As Integer
'loop through each pair of commas looking for dashes. If no dashes, just grab the number between the commas.
'else separate each number from the dash.
'then fill an array with every length that needs to be analyzed.
'return the array.
'how's that for stepwise refinement? Guess I need practice...:P
Dim cStart As Integer   'keeps track of where the range starts in the string
Dim cEnd As Integer     'guess
Dim intResult As Integer    'mainly just holds results from mid$
Dim intTop As Integer, intBottom As Integer 'top and bottom of ranges--used in loops to fill intLens()
Dim strParse As String  'temp storage
Dim cLens As Integer    'number of sample lens
Dim i As Integer, j As Integer    'guess
    ReDim intLens(0 To 99)   'this number should not normally be exceeded, but no guarantees!
'    cStart = 1
Do
    cStart = cEnd + 1 'inc cStart to match the end of the previous sample
    cEnd = InStr(cStart, strLen, ",")
    If cEnd = 0 Then cEnd = Len(strLen) + 1 'if we're at the end(ie no more commas), set the end of sample to the end of string
    intResult = InStr(cStart, strLen, "-")
    If intResult = 0 Or intResult >= cEnd Then  'not a range
        strParse = Trim$(Mid$(strLen, cStart, cEnd - cStart + 1))
        If IsNumeric(strParse) Then
            intLens(cLens) = Val(strParse)
            cLens = cLens + 1
            If cLens > UBound(intLens) Then
                ReDim Preserve intLens(0 To cLens + 99)
            End If
        'Else 'not numeric
            'ignore it or maybe warn the user
        End If
    Else    'else range
        strParse = Trim$(Mid$(strLen, cStart, intResult - cStart))
        If IsNumeric(strParse) Then
            intBottom = Val(strParse)
        Else
            intBottom = 1
        End If
        strParse = Mid$(strLen, intResult + 1, cEnd - intResult - 1)
        If IsNumeric(strParse) Then
            intTop = Val(strParse)
        Else
            intTop = 2
        End If
        For i = intBottom To intTop
            intLens(cLens) = i
            cLens = cLens + 1
            If cLens > UBound(intLens) Then
                ReDim Preserve intLens(0 To cLens + 99)
            End If
        Next i
        'have to use intresult to mid$ the two numbers apart from each other, cstart with len cstart-intresult(+1?)
        'and intresult with len intresult-cend(+1?)
        'then for loop through each frequency
    End If
Loop While InStr(cStart, strLen, ",") > 0
ParseTextWordStyle = cLens

End Function
Private Function ParseSampleLengths() As Boolean
'loop through each pair of commas looking for dashes. If no dashes, just grab the number between the commas.
'else separate each number from the dash.
'then fill an array with every length that needs to be analyzed.
'return the array.
'how's that for stepwise refinement? Guess I need practice...:P
Dim strLen As String    'temp storage for txtSampleLength cause its a pain to type
Dim cLens As Integer    'number of sample lens
Dim i As Integer, j As Integer    'guess
'    cStart = 1
    strLen = txtSampleLength.Text
    cLens = ParseTextWordStyle(strLen, intLens)
    ReDim Preserve intLens(0 To cLens - 1)  'size down to make sure we're not wasting any
    If cLens > 10 Then
        MsgBox "Warning: This calculation will be very lengthy since there are over ten sample lengths to be analyzed.", vbOKOnly
    End If
    'now sort intLens
    'sort lens ascending
    For i = 0 To cLens - 1
        For j = i To cLens - 1
            If intLens(i) > intLens(j) Then
                SwapInt intLens(i), intLens(j)
            End If
        Next j
    Next i
    If (intLens(cLens - 1) > 200) Then
        If MsgBox("You have a sample length that's way too long. Are you sure you want to use it?", vbYesNo) = vbNo Then
            txtSampleLength.SelStart = 0
            txtSampleLength.SelLength = Len(txtSampleLength.Text)
            txtSampleLength.SetFocus
            ParseSampleLengths = False
            Exit Function 'may need to abort differently
        End If
    End If
    ParseSampleLengths = True
End Function
'this sub opens a table, reads all 'normal' values, stores them in an array for exclusion, and 'returns' the array
'it's not actually returned, just filled as the only way to do it is globally.
Private Function ParseTableValues() As Long
Dim strLine As String
Dim ch As String
Dim intResult As Integer
Dim intFilenoTableExclude As Integer
ReDim hufTableVals(0 To 255) As Freq
Dim lIdx As Long
    lIdx = 0
    intFilenoTableExclude = FreeFile
    Open strFilenameTableExclude For Input As #intFilenoTableExclude
    Do Until EOF(intFilenoTableExclude)
        Line Input #intFilenoTableExclude, strLine
        ch = LCase$(Left$(strLine, 1))
        'analyze--if the first character isn't hex, then ignore the line. else it's probably good so search for =
        If (ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "f") Then
            'if we find an =, then it's good! parse it this way:
            intResult = InStr(strLine, "=")
            If intResult > 0 Then
                lIdx = lIdx + 1
                If UBound(hufTableVals) < lIdx Then
                    ReDim Preserve hufTableVals(0 To UBound(hufTableVals) + 256)
                End If
                hufTableVals(lIdx).Char = Right$(strLine, Len(strLine) - intResult)
            End If
        End If

        '1]up to = is the hex (trash it cause we don't care :)
        '2]past = is the actual value (keep it and assign to an array of huffmen, but don't fill in a frequency becuase they're not used that way)
    Loop
    Close intFilenoTableExclude
    ParseTableValues = lIdx
End Function
'this sub accepts an array of frequencies, opens the table specified by the user(global), converts the array
'into a Thingy standard table format, and appends the array to the file. yay
'it also requires that the array of desired table values has been put in a String array(global) by ParseTableEntries()
'user must know to prefix their desired hex digits with &H
Private Sub WriteTableValues(hufseq() As Freq)
Dim cLens As Integer    'how many spots they want filled in the table
Dim i As Integer
Dim line As String
Dim intRanges() As Integer
'pseudocode:
'first, read in requested table values
'loop through all table values requested to be filled. If we run out of values to fill in, quit.
'for each loop write a line(in Sequential Output mode) in FF=match format
    cLens = ParseTextWordStyle(txtWriteTableRanges, intRanges)
    intFilenoTableWrite = FreeFile
    Open strFilenameTableWrite For Append As intFilenoTableWrite 'wheee! I still remember it!
    Print #intFilenoTableWrite, ""  'print a blank line first
    For i = 0 To cLens - 1 Step 1
        'if we're out of matches, quit
        If i > UBound(hufseq) Then Exit For
        'otherwise write in table format
        If hufseq(i).Frequency > -1 Then    'only write non-skipped ones
            line = Hex$(intRanges(i))
            If (Len(line) = 1) Then
                line = "0" & line
            End If
            line = line & "=" & hufseq(i).Char
            Print #intFilenoTableWrite, line
        End If
    Next i
    Close intFilenoTableWrite
End Sub
'this sub accepts a report to be written and the file to write it to.
'then it writes it--shaboom!
Private Sub WriteReport(strReport As String)
    intFilenoReport = FreeFile
    Open strFilenameReport For Output As intFilenoReport
    Print #intFilenoReport, strReport
    Close #intFilenoReport
End Sub
Private Function max(int1, int2) As Integer
    max = IIf(int1 > int2, int1, int2)
End Function
Private Sub SetEnablement(bEnabled As Boolean)
    'flip the GUI status according to bEnabled's value.
    cmdAnalyze.Caption = IIf(bEnabled, "&Analyze", "Cancel") 'cancel/analyze button
    cmdAnalyze.Cancel = Not bEnabled
    SSTab1.Enabled = bEnabled
    lblCalc.Visible = Not bEnabled 'reading visual feedback
    lblAddr.Visible = Not bEnabled
    linCompletion.Visible = Not bEnabled
    cmdOpen.Enabled = bEnabled
End Sub

Private Sub SwapInt(huf1 As Integer, huf2 As Integer)
Dim hufTemp As Integer
    hufTemp = huf2
    huf2 = huf1
    huf1 = hufTemp
End Sub

Private Sub chkBeepOnFinish_Click()
    bBeep = IIf(chkBeepOnFinish.Value = vbChecked, True, False)
End Sub

Private Sub chkUseComments_Click()
    'flip comment status
    bComments = Not bComments
    'and twiddle GUI accordingly
    fraComments.Enabled = bComments
    chkReadInsideComments.Enabled = bComments
    optC.Enabled = bComments
    optCPP.Enabled = bComments
    optThingy.Enabled = bComments
    optCustom.Enabled = bComments
    lblCEnd.Enabled = bComments
    lblCPPEnd.Enabled = bComments
    lblThingyEnd.Enabled = bComments
    txtBeginComment.Enabled = bComments
    txtEndComment.Enabled = bComments
    lblBegin.Enabled = bComments
    lblEnd.Enabled = bComments
End Sub

Private Sub cmdAnalyze_Click()
    If cmdAnalyze.Cancel = True Then    'if we're running, then cancel
        bCancel = True
    Else
        SSTab1.Tab = 3 'analyze tab
        SetEnablement False 'disable GUI
        cmdAnalyze.Caption = "Cancel"
        cmdAnalyze.Cancel = True
        AnalyzeFrequencyBinary '***
        SetEnablement True  'enable GUI
        SSTab1.TabEnabled(4) = True 'results tab
        SSTab1.Tab = 4
        txtResults.SetFocus
    End If
End Sub
Private Sub AnalyzeFrequency()
'get the totals for each sample
'and create a wonderful report to be displayed in txtResults
'**=undone
Dim hufseq() As FreqAra '3D array of frequency values. 1st dimension is 1 2D array for each length. 2nd dimension
'is 256 long, 1 1D array for each possible Chr$() value (err Ascii vals 0-255), 3rd dimension is 0 based list of all matches for that 2D position
Dim intAsc As Integer    'this holds the ascii val of the first character of ch so that the match can be further categorised

Dim pAddr As Long   'the current address of the file that we're reading
Dim i As Long, j As Long, k As Long    'counter
Dim cFiles As Long  'another counter which loops through every file we hit
Dim intResult As Integer    'holds the index of a comment in the sample, if one exists

Dim ch As String    'holds the characters to be analysed
Dim maxSample As Integer    'the longest of 1.The longest char sample 2.The begin comment tag 3.The end comment tag:
                                            ':determines how many chars are read at one time
Dim cChars As Long 'keep track of how many different chars are actually used (total value)
Dim cCharsLen() As Long 'keep track of how many different chars are actually used (an array which has one element for each length)
'this is now 2D because it's UBound(intLens) * 256 (those are the lengths of the first two dimensions of the array
Dim cBytesCounted As Long 'keep track of bytes actually analysed (ie skip the ones commented out)
Dim cTotalBytes As Long       'keep track of how many bytes are in all of the files
Dim cTableExcludes As Long  'keep track of the total number of table values in the table value which should be excluded from the count

Dim maxFreq As Freq         'the actual match of rank X
Dim maxFreqIndex As Long 'the 1stD index of the highest current match when determing Top x Results
Dim maxFreqLenIndex As Long 'the 3rdD index of the highest current match (added to account for 2D hufSeq array)
Dim maxFreqAscIndex         'the 2ndD index of the highest current match (added to account for 3D hufSeq array)
Dim topResults() As Freq    'the top X results
Dim ubTopResults As Long    'holds UBound(topResults) for easy access

Dim hexChars As String 'a composition string to display the hex value of the result chars
Dim hexTemp As String 'each hex char right after its been calced by Hex$(). Used so that I can prepend a "0" on single digit hex
Dim strMsg As String    'a composition string to display in txtResults
Dim beginTime As Date   'two variables to display the elapsed time in analysing the file
Dim dTime As Date

Dim strBeginComment As String   'these hold the comment tags ie /* and */. CurComment just holds whichever one we're looking for right now
Dim strEndComment As String
Dim strCurComment As String
Dim bIgnoring As Boolean    'whether we're just flipping through the values ignoring them or analysing them
Dim bBegin As Boolean   'whether the current comment tag to search for is Begin tag or End tag.

'heyheyhey here, the code goes

    beginTime = Now 'and we're off!
    
    'get sample lengths and alloc mem
    If ParseSampleLengths = False Then Exit Sub
    ReDim hufseq(0 To UBound(intLens), 0 To 255)   'dim the first two dimensions as a 2D array
    For i = 0 To UBound(intLens) Step 1 'and dim the last dimension as a seperate array--a true 3D array wastes too much mem
        For j = 0 To 255
            ReDim hufseq(i, j).hufAra(0 To 255) 'the 255 is fairly arbitrary
        Next j
    Next i
    ReDim cCharsLen(0 To UBound(intLens), 0 To 255)   'match this array to be parallel with the number of lengths we have
    
        'test string
    'see if we want to exclude table values and if yes, get them from the proper file
    If bTableExclude Then
        cTableExcludes = ParseTableValues
    End If
    'see if we're using comments and if yes, get which one
    If bComments = True Then
        bBegin = True   'we start off looking for the start comment value
        If chkReadInsideComments.Value = vbChecked Then bIgnoring = True    'start ignoring data until hit begin comment tag
        'fill str|Begin|End|Comment with the correct values (right now there is a 2 char limit on comments since I'm not sure how
        'they'll affect performance)
        If optC.Value = True Then
            strBeginComment = "/*"
            strEndComment = "*/"
        ElseIf optCPP.Value = True Then
            strBeginComment = "//"
            strEndComment = vbCrLf
        ElseIf optThingy.Value = True Then
            strBeginComment = "<$"
            strEndComment = ">"
        Else    'we hope custom style
            strBeginComment = txtBeginComment.Text
            If strBeginComment = "" Then
                bComments = False  'whoa! no text so forget the whole thing
            Else
                strEndComment = txtEndComment.Text
                If txtEndComment.Text = "" Then strEndComment = vbCrLf
            End If
        End If
        strCurComment = strBeginComment 'init cur comment tag to the begin tag
    End If
    
    'if we're using comments, add those to the sample length so we can see them coming ahead of time
    maxSample = max(Len(strBeginComment), Len(strEndComment))
    ch = Space$(intLens(UBound(intLens)) + maxSample) 'replace with max len after parsed in ParseSampleLength
    
    'get the totals of each sample
    'first twiddle the GUI
    cmdOpen.Enabled = False
    
    
For cFiles = 0 To UBound(strFilename) - 1 Step 1  'superimposed for loop ^_^ to read multiple files
    'open file and start at the beginning
    intFileno = FreeFile
    Open strFilename(cFiles) For Binary As #intFileno
    cTotalBytes = cTotalBytes + LOF(intFileno)
    pAddr = 1
    'and twiddle gui again
    lblCalc.Caption = "Reading... Current Address of File " & cFiles + 1 & "/" & UBound(strFilename) & ":"
    lblAddr.Caption = pAddr
    lblMatchCount.Caption = cChars
    linCompletion.X2 = 0
    
    Do
        Get #intFileno, pAddr, ch
        'see if we've hit a comment tag(begin or end)
        If bComments And Left$(ch, Len(strCurComment)) = strCurComment Then
            bIgnoring = Not bIgnoring
            bBegin = Not bBegin
            'inc past the comment tag so we don't get caught by /*/ confusing like /**/
            pAddr = pAddr + (Len(strCurComment))
            Get #intFileno, pAddr, ch
            'save the current comment tag for easy access
            strCurComment = IIf(bBegin, strBeginComment, strEndComment)
        End If
        'loop through current matches looking to see if new sample matches an old one
        If bIgnoring = False Then   'only count hits if we're actively reading
            'detect comments
            If strCurComment <> "" Then 'because InStr returns 1 for "", we have to manually set intResult to 0 ourselves.
                intResult = InStr(ch, strCurComment)    'detect comments, if any
            Else
                intResult = 0
            End If
            'save the Asc val of the first character to further categorise the match
            intAsc = Asc(Left$(ch, 1))

'quick lesson: len("aabs/*Hey, mon") = 14 but above Instr = 5 so you want to quit when intLens > 4; ie: >=5; ie: intLens(i) >= intResult
            'see if exists
            For i = 0 To UBound(intLens) Step 1 'loop through every length
               'check to see if we've run into a comment (and if there's a comment in the first place)
                If intResult > 0 And intLens(i) >= intResult Then
                    Exit For
                End If
                'loop through all matches of this length
                For j = 0 To cCharsLen(i, intAsc) - 1 Step 1
                    If hufseq(i, intAsc).hufAra(j).Char = Left$(ch, intLens(i)) Then    'if it's a match
                        'update it
                        hufseq(i, intAsc).hufAra(j).Frequency = hufseq(i, intAsc).hufAra(j).Frequency + 1
                        Exit For    'and we can quit this length's match search
                    End If
                Next j
                'add new
                If j = cCharsLen(i, intAsc) Then   'we know we looped all the way through without finding anything
                    'so add it
                    'but first see if we need to alloc more mem
                    If cCharsLen(i, intAsc) >= UBound(hufseq(i, intAsc).hufAra) Then 'this increases it only for this sample length/ascii val result list
                        ReDim Preserve hufseq(i, intAsc).hufAra(0 To UBound(hufseq(i, intAsc).hufAra) + 256) 'alloc mem 256 huffmen at once
                    End If
                    hufseq(i, intAsc).hufAra(cCharsLen(i, intAsc)).Frequency = 1 'first sighting
                    hufseq(i, intAsc).hufAra(cCharsLen(i, intAsc)).Char = Left$(ch, intLens(i))  'and snap off the actual sample
                    cCharsLen(i, intAsc) = cCharsLen(i, intAsc) + 1 'update this length's total
                    cChars = cChars + 1                 'and the total total
                End If
            Next i
            cBytesCounted = cBytesCounted + 1
        End If  'end if bIgnoring=false
        'twiddle GUI again
        If pAddr Mod 50 = 0 Then    'only update every 50 bytes cause its faster if you don't mangle the screen so much
            lblAddr.Caption = pAddr
            lblMatchCount.Caption = cChars
            If pAddr Mod 500 = 0 Then   'only update line every 500 bytes
                linCompletion.X2 = (pAddr / LOF(intFileno)) * SSTab1.Width 'run this and you will see how it works
            End If
        End If
        DoEvents
        'check for cancelation of processing
        If bCancel = True Then
            bCancel = False
            ReDim hufseq(0 To 1)    'release most of mem
            Exit Sub
        End If
        pAddr = pAddr + 1
    Loop Until EOF(intFileno)   'end read loop
    Close intFileno
Next cFiles 'end superimposed for loop to allow multiple files

    'analyse totals
    'first twiddle GUI a little:
    lblCalc.Caption = "Calculating... Current Match:"
    'make sure the top results is number ^^
    If Not IsNumeric(txtTopResults.Text) Then
        txtTopResults.Text = "25"
    End If
    ReDim topResults(0 To IIf(txtTopResults.Text < cChars, txtTopResults.Text, cChars - 1)) 'an array of the Top X results
    ubTopResults = UBound(topResults)                                                                'and its UBound, so I don't have to call UBound all the time
    'grab the top x matches
    For i = 0 To ubTopResults Step 1
        maxFreqIndex = 0    'reinit to 0 each time
        maxFreqLenIndex = 0
        maxFreqAscIndex = 0
        maxFreq.Char = "": maxFreq.Frequency = -1
        'loop through and find max
        For j = 0 To UBound(intLens) Step 1
            For intAsc = 0 To 255 Step 1
                For k = 0 To cCharsLen(j, intAsc) Step 1
                    If (maxFreq.Frequency * Len(maxFreq.Char)) _
                    < (hufseq(j, intAsc).hufAra(k).Frequency * Len(hufseq(j, intAsc).hufAra(k).Char)) Then
                        maxFreq = hufseq(j, intAsc).hufAra(k)
                        maxFreqIndex = k    'and save where we found it, too
                        maxFreqAscIndex = intAsc
                        maxFreqLenIndex = j
                    End If
                Next k
            Next intAsc
        Next j
        'put this match in the top x results
        topResults(i) = maxFreq
        'and take it out of the test
        hufseq(maxFreqLenIndex, maxFreqAscIndex).hufAra(maxFreqIndex).Frequency = -1
        lblAddr.Caption = i & " of " & txtTopResults.Text
        DoEvents
        If bCancel = True Then
            bCancel = False
            ReDim hufseq(0 To 1)    'release most of mem
            Exit Sub
        End If
    Next i
    'now flag space(32)/tab(9) and carriage return/line feed(10 and 13) if they've been checked
' GUI feedback
    lblCalc.Caption = "Formatting Top Results..."
    DoEvents
    
    If chkIgnoreSpace.Value = vbChecked Then
        For i = 0 To ubTopResults Step 1
            If InStr(topResults(i).Char, Chr$(32)) > 0 Then 'space
                topResults(i).Frequency = vbSpacePresent
            End If
        Next i
    End If
    If chkIgnoreReturn.Value = vbChecked Then
        For i = 0 To ubTopResults Step 1
            If InStr(topResults(i).Char, Chr$(10)) > 0 And InStr(topResults(i).Char, Chr$(13)) > 0 Then  'carriage return+linefeed
                topResults(i).Frequency = vbCrLfPresent
            End If
        Next i
    End If
    If chkIgnoreTab.Value = vbChecked Then
        For i = 0 To ubTopResults Step 1
            If InStr(topResults(i).Char, Chr$(9)) > 0 Then  'tab
                topResults(i).Frequency = vbTabPresent
            End If
        Next i
    End If
    If bTableExclude Then
        For i = 0 To ubTopResults Step 1
            For j = 0 To cTableExcludes Step 1
                If topResults(i).Char = hufTableVals(j).Char Then
                    topResults(i).Frequency = vbTableValPresent
                End If
            Next j
        Next i
    End If
'GUI feedback
    If bTableWrite Then
        lblCalc.Caption = "Writing Table Values..."
        WriteTableValues topResults
    End If
    lblCalc.Caption = "Generating Report..."
    DoEvents
    'print results
    For i = 0 To ubTopResults Step 1
        strMsg = strMsg & i & ":" 'first append the index of the result
        Select Case topResults(i).Frequency
        'append the reason for skippage
        Case vbSpacePresent
            strMsg = strMsg & "--skipped for containing a space--" & vbCrLf
        Case vbTabPresent
            strMsg = strMsg & "--skipped for containing a tab--" & vbCrLf
        Case vbCrLfPresent
            strMsg = strMsg & "--skipped for containing a carriage return linefeed--" & vbCrLf
        Case vbTableValPresent
            strMsg = strMsg & "--skipped for containing a table value--" & vbCrLf
        Case Else   'not skipped!
            strMsg = strMsg & " total space = " & (Len(topResults(i).Char) * topResults(i).Frequency) & "; hits = " & topResults(i).Frequency
            If InStr(topResults(i).Char, Chr$(0)) = 0 Then  'if there is no \0 you can display the actual text, otherwise only hex
                strMsg = strMsg & "; string = " & topResults(i).Char
            Else
                strMsg = strMsg & "; --string not displayable--"
            End If
            
            hexChars = ""   'reinit hexChars
            For j = 1 To Len(topResults(i).Char) Step 1
                hexTemp = Hex$(Asc(Mid$(topResults(i).Char, j, 1)))
                If Len(hexTemp) = 1 Then hexTemp = "0" & hexTemp
                hexChars = hexChars & hexTemp
            Next j
            strMsg = strMsg & "; hex = " & hexChars
            strMsg = strMsg & "; len = " & Len(topResults(i).Char) & vbCrLf
        End Select  'end select skip reason
        'here is the new and improved format:
        '3: total space = 384; hits = 128; string = foo; hex = 666F6F; len = 3
        'or
        '4: total space = 384; hits = 128; string = --not displayable--; hex = 101300; len = 3
    Next i
    strMsg = vbCrLf & vbCrLf & strMsg
    dTime = beginTime - Now 'figure the total time taken
    strMsg = vbCrLf & "Elapsed time: " & Format(dTime, "hh:nn:ss") & strMsg
    For i = UBound(intLens) To 0 Step -1
        strMsg = intLens(i) & " " & strMsg
    Next i
    strMsg = "Total characters: " & cTotalBytes & "     Characters actually counted: " & cBytesCounted & vbCrLf & "Number of different strings: " & cChars & "    Sample lengths requested:" & strMsg
    For i = UBound(strFilename) - 1 To 0 Step -1
        strMsg = "Filename:" & strFilename(i) & vbCrLf & strMsg
    Next i
    If bOutput Then 'if they requested to write a long report, do so
        WriteReport strMsg
    End If
    If Len(strMsg) > 32767 Then 'now truncate the displayed version
        strMsg = "--Report truncated because size is too great for text box. To remedy this situation," _
          & " reduce the number of results to be shown that aren't really relevant anyway.--" & vbCrLf & vbCrLf & _
        Left$(strMsg, 32000) & vbCrLf & vbCrLf & "--Report truncated because size is too great for text box. To remedy this situation," _
          & " reduce the number of results to be shown that aren't really relevant anyway.--"
    End If
    txtResults.Text = strMsg
    If bBeep Then Beep  'beep on finish
    bCancel = False 'and reset cancel flag
End Sub
Private Sub AnalyzeFrequencyBinary()
'get the totals for each sample
'and create a wonderful report to be displayed in txtResults
'**=undone
Dim hufseq() As FreqAra '3D array of frequency values. 1st dimension is 1 2D array for each length. 2nd dimension
'is 256 long, 1 1D array for each possible Chr$() value (err Ascii vals 0-255), 3rd dimension is 0 based list of all matches for that 2D position
Dim intAsc As Integer    'this holds the ascii val of the first character of ch so that the match can be further categorised
Dim top As New HufNode  'the top of the tree of values***TEMP
    top.Freq = 1
    top.Strin = " "
Dim current As HufNode  'the current one we're looking..at
Dim bContinue As Boolean    'whether we should stop.

Dim pAddr As Long   'the current address of the file that we're reading
Dim i As Long, j As Long, k As Long    'counter
Dim cFiles As Long  'another counter which loops through every file we hit
Dim intResult As Integer    'holds the index of a comment in the sample, if one exists

Dim ch As String    'holds the characters to be analysed
Dim maxSample As Integer    'the longest of 1.The longest char sample 2.The begin comment tag 3.The end comment tag:
                                            ':determines how many chars are read at one time
Dim cChars As Long 'keep track of how many different chars are actually used (total value)
Dim cCharsLen() As Long 'keep track of how many different chars are actually used (an array which has one element for each length)
'this is now 2D because it's UBound(intLens) * 256 (those are the lengths of the first two dimensions of the array
Dim cBytesCounted As Long 'keep track of bytes actually analysed (ie skip the ones commented out)
Dim cTotalBytes As Long       'keep track of how many bytes are in all of the files
Dim cTableExcludes As Long  'keep track of the total number of table values in the table value which should be excluded from the count

Dim maxFreq As Freq         'the actual match of rank X
Dim maxFreqIndex As Long 'the 1stD index of the highest current match when determing Top x Results
Dim maxFreqLenIndex As Long 'the 3rdD index of the highest current match (added to account for 2D hufSeq array)
Dim maxFreqAscIndex         'the 2ndD index of the highest current match (added to account for 3D hufSeq array)
Dim topResults() As Freq    'the top X results
Dim ubTopResults As Long    'holds UBound(topResults) for easy access

Dim hexChars As String 'a composition string to display the hex value of the result chars
Dim hexTemp As String 'each hex char right after its been calced by Hex$(). Used so that I can prepend a "0" on single digit hex
Dim strMsg As String    'a composition string to display in txtResults
Dim beginTime As Date   'two variables to display the elapsed time in analysing the file
Dim dTime As Date

Dim strBeginComment As String   'these hold the comment tags ie /* and */. CurComment just holds whichever one we're looking for right now
Dim strEndComment As String
Dim strCurComment As String
Dim bIgnoring As Boolean    'whether we're just flipping through the values ignoring them or analysing them
Dim bBegin As Boolean   'whether the current comment tag to search for is Begin tag or End tag.

'heyheyhey here, the code goes

    beginTime = Now 'and we're off!
    
    'get sample lengths and alloc mem
    If ParseSampleLengths = False Then Exit Sub
    ReDim hufseq(0 To UBound(intLens), 0 To 255)   'dim the first two dimensions as a 2D array
    For i = 0 To UBound(intLens) Step 1 'and dim the last dimension as a seperate array--a true 3D array wastes too much mem
        For j = 0 To 255
            ReDim hufseq(i, j).hufAra(0 To 255) 'the 255 is fairly arbitrary
        Next j
    Next i
    ReDim cCharsLen(0 To UBound(intLens), 0 To 255)   'match this array to be parallel with the number of lengths we have
    
        'test string
    'see if we want to exclude table values and if yes, get them from the proper file
    If bTableExclude Then
        cTableExcludes = ParseTableValues
    End If
    'see if we're using comments and if yes, get which one
    If bComments = True Then
        bBegin = True   'we start off looking for the start comment value
        If chkReadInsideComments.Value = vbChecked Then bIgnoring = True    'start ignoring data until hit begin comment tag
        'fill str|Begin|End|Comment with the correct values (right now there is a 2 char limit on comments since I'm not sure how
        'they'll affect performance)
        If optC.Value = True Then
            strBeginComment = "/*"
            strEndComment = "*/"
        ElseIf optCPP.Value = True Then
            strBeginComment = "//"
            strEndComment = vbCrLf
        ElseIf optThingy.Value = True Then
            strBeginComment = "<$"
            strEndComment = ">"
        Else    'we hope custom style
            strBeginComment = txtBeginComment.Text
            If strBeginComment = "" Then
                bComments = False  'whoa! no text so forget the whole thing
            Else
                strEndComment = txtEndComment.Text
                If txtEndComment.Text = "" Then strEndComment = vbCrLf
            End If
        End If
        strCurComment = strBeginComment 'init cur comment tag to the begin tag
    End If
    
    'if we're using comments, add those to the sample length so we can see them coming ahead of time
    maxSample = max(Len(strBeginComment), Len(strEndComment))
    ch = Space$(intLens(UBound(intLens)) + maxSample) 'replace with max len after parsed in ParseSampleLength
    
    'get the totals of each sample
    'first twiddle the GUI
    cmdOpen.Enabled = False
    
    
For cFiles = 0 To UBound(strFilename) - 1 Step 1  'superimposed for loop ^_^ to read multiple files
    'open file and start at the beginning
    intFileno = FreeFile
    Open strFilename(cFiles) For Binary As #intFileno
    cTotalBytes = cTotalBytes + LOF(intFileno)
    pAddr = 1
    'and twiddle gui again
    lblCalc.Caption = "Reading... Current Address of File " & cFiles + 1 & "/" & UBound(strFilename) & ":"
    lblAddr.Caption = pAddr
    lblMatchCount.Caption = cChars
    linCompletion.X2 = 0
    
    Do
        Get #intFileno, pAddr, ch
        'see if we've hit a comment tag(begin or end)
        If bComments And Left$(ch, Len(strCurComment)) = strCurComment Then
            bIgnoring = Not bIgnoring
            bBegin = Not bBegin
            'inc past the comment tag so we don't get caught by /*/ confusing like /**/
            pAddr = pAddr + (Len(strCurComment))
            Get #intFileno, pAddr, ch
            'save the current comment tag for easy access
            strCurComment = IIf(bBegin, strBeginComment, strEndComment)
        End If
        'loop through current matches looking to see if new sample matches an old one
        If bIgnoring = False Then   'only count hits if we're actively reading
            'detect comments
            If strCurComment <> "" Then 'because InStr returns 1 for "", we have to manually set intResult to 0 ourselves.
                intResult = InStr(ch, strCurComment)    'detect comments, if any
            Else
                intResult = 0
            End If
            'save the Asc val of the first character to further categorise the match
            intAsc = Asc(Left$(ch, 1))

'quick lesson: len("aabs/*Hey, mon") = 14 but above Instr = 5 so you want to quit when intLens > 4; ie: >=5; ie: intLens(i) >= intResult
            'see if exists
            For i = 0 To UBound(intLens) Step 1 'loop through every length
               'check to see if we've run into a comment (and if there's a comment in the first place)
                If intResult > 0 And intLens(i) >= intResult Then
                    Exit For
                End If
                'loop through all matches of this length
                Set current = top
                Dim hops As Long
                Do
                    bContinue = False
                    Dim sample As String
                    sample = Left$(ch, intLens(i))
                    If (current.Strin = sample) Then
                        current.Freq = current.Freq + 1
                    ElseIf current.Strin > sample Then
                        If current.IsLeft Then
                            Set current = current.Left
                            bContinue = True
                        Else    'add
                            Set current.Left = New HufNode
                            current.Left.Strin = sample
                            current.Left.Freq = 1
                            current.IsLeft = True
                            cChars = cChars + 1                 'update total unique chars
                        End If
                    Else    'current.strin > target
                        If current.IsRight Then
                            Set current = current.Right
                            bContinue = True
                        Else
                            Set current.Right = New HufNode
                            current.Right.Strin = sample
                            current.Right.Freq = 1
                            current.IsRight = True
                            cChars = cChars + 1                 'update total unique chars
                        End If
                    End If
                    hops = hops + 1
                Loop While (bContinue)
'***OLD CODE
'                For j = 0 To cCharsLen(i, intAsc) - 1 Step 1
'                    If hufseq(i, intAsc).hufAra(j).Char = Left$(ch, intLens(i)) Then    'if it's a match
'                        'update it
'                        hufseq(i, intAsc).hufAra(j).Frequency = hufseq(i, intAsc).hufAra(j).Frequency + 1
'                        Exit For    'and we can quit this length's match search
'                    End If
'                Next j

                'add new
'                If j = cCharsLen(i, intAsc) Then   'we know we looped all the way through without finding anything
'                    'so add it
'                    'but first see if we need to alloc more mem
'                    If cCharsLen(i, intAsc) >= UBound(hufseq(i, intAsc).hufAra) Then 'this increases it only for this sample length/ascii val result list
'                        ReDim Preserve hufseq(i, intAsc).hufAra(0 To UBound(hufseq(i, intAsc).hufAra) + 256) 'alloc mem 256 huffmen at once
'                    End If
'                    hufseq(i, intAsc).hufAra(cCharsLen(i, intAsc)).Frequency = 1 'first sighting
'                    hufseq(i, intAsc).hufAra(cCharsLen(i, intAsc)).Char = Left$(ch, intLens(i))  'and snap off the actual sample
'                    cCharsLen(i, intAsc) = cCharsLen(i, intAsc) + 1 'update this length's total

'                End If
'***END OLD
            Next i
            cBytesCounted = cBytesCounted + 1
        End If  'end if bIgnoring=false
        'twiddle GUI again
        If pAddr Mod 50 = 0 Then    'only update every 50 bytes cause its faster if you don't mangle the screen so much
            lblAddr.Caption = pAddr
            lblMatchCount.Caption = cChars
            If pAddr Mod 500 = 0 Then   'only update line every 500 bytes
                linCompletion.X2 = (pAddr / LOF(intFileno)) * SSTab1.Width 'run this and you will see how it works
            End If
        End If
        DoEvents
        'check for cancelation of processing
        If bCancel = True Then
            bCancel = False
            ReDim hufseq(0 To 1)    'release most of mem
            Exit Sub
        End If
        pAddr = pAddr + 1
    Loop Until EOF(intFileno)   'end read loop
    Close intFileno
Next cFiles 'end superimposed for loop to allow multiple files

    'analyse totals
    'first twiddle GUI a little:
    lblCalc.Caption = "Calculating... Current Match:"
    'make sure the top results is number ^^
    If Not IsNumeric(txtTopResults.Text) Then
        txtTopResults.Text = "25"
    End If
    ReDim topResults(0 To IIf(txtTopResults.Text < cChars, txtTopResults.Text, cChars - 1)) 'an array of the Top X results
    ubTopResults = UBound(topResults)                                                                'and its UBound, so I don't have to call UBound all the time
    'grab the top x matches
    For i = 0 To ubTopResults Step 1
        maxFreq.Frequency = -1
        TopSearch top
        topResults(i).Frequency = maxNode.Freq
        topResults(i).Char = maxNode.Strin
        maxNode.Freq = -1   'remove from searchability
'***OLD CODE
'        maxFreqIndex = 0    'reinit to 0 each time
'        maxFreqLenIndex = 0
'        maxFreqAscIndex = 0
'        maxFreq.Char = "": maxFreq.Frequency = -1
'        'loop through and find max
'        For j = 0 To UBound(intLens) Step 1
'            For intAsc = 0 To 255 Step 1
'                For k = 0 To cCharsLen(j, intAsc) Step 1
'                    If (maxFreq.Frequency * Len(maxFreq.Char)) _
'                    < (hufseq(j, intAsc).hufAra(k).Frequency * Len(hufseq(j, intAsc).hufAra(k).Char)) Then
'                        maxFreq = hufseq(j, intAsc).hufAra(k)
'                        maxFreqIndex = k    'and save where we found it, too
'                        maxFreqAscIndex = intAsc
'                        maxFreqLenIndex = j
'                    End If
'                Next k
'            Next intAsc
'        Next j
'        'put this match in the top x results
'        topResults(i) = maxFreq
'        'and take it out of the test
'        hufseq(maxFreqLenIndex, maxFreqAscIndex).hufAra(maxFreqIndex).Frequency = -1
        lblAddr.Caption = i & " of " & txtTopResults.Text
        DoEvents
        If bCancel = True Then
            bCancel = False
            ReDim hufseq(0 To 1)    'release most of mem
            Exit Sub
        End If
'***END OLD
    Next i
    'now flag space(32)/tab(9) and carriage return/line feed(10 and 13) if they've been checked
' GUI feedback
    lblCalc.Caption = "Formatting Top Results..."
    DoEvents
    
    If chkIgnoreSpace.Value = vbChecked Then
        For i = 0 To ubTopResults Step 1
            If InStr(topResults(i).Char, Chr$(32)) > 0 Then 'space
                topResults(i).Frequency = vbSpacePresent
            End If
        Next i
    End If
    If chkIgnoreReturn.Value = vbChecked Then
        For i = 0 To ubTopResults Step 1
            If InStr(topResults(i).Char, Chr$(10)) > 0 And InStr(topResults(i).Char, Chr$(13)) > 0 Then  'carriage return+linefeed
                topResults(i).Frequency = vbCrLfPresent
            End If
        Next i
    End If
    If chkIgnoreTab.Value = vbChecked Then
        For i = 0 To ubTopResults Step 1
            If InStr(topResults(i).Char, Chr$(9)) > 0 Then  'tab
                topResults(i).Frequency = vbTabPresent
            End If
        Next i
    End If
    If bTableExclude Then
        For i = 0 To ubTopResults Step 1
            For j = 0 To cTableExcludes Step 1
                If topResults(i).Char = hufTableVals(j).Char Then
                    topResults(i).Frequency = vbTableValPresent
                End If
            Next j
        Next i
    End If
'GUI feedback
    If bTableWrite Then
        lblCalc.Caption = "Writing Table Values..."
        WriteTableValues topResults
    End If
    lblCalc.Caption = "Generating Report..."
    DoEvents
    'print results
    For i = 0 To ubTopResults Step 1
        strMsg = strMsg & i & ":" 'first append the index of the result
        Select Case topResults(i).Frequency
        'append the reason for skippage
        Case vbSpacePresent
            strMsg = strMsg & "--skipped for containing a space--" & vbCrLf
        Case vbTabPresent
            strMsg = strMsg & "--skipped for containing a tab--" & vbCrLf
        Case vbCrLfPresent
            strMsg = strMsg & "--skipped for containing a carriage return linefeed--" & vbCrLf
        Case vbTableValPresent
            strMsg = strMsg & "--skipped for containing a table value--" & vbCrLf
        Case Else   'not skipped!
            strMsg = strMsg & " total space = " & (Len(topResults(i).Char) * topResults(i).Frequency) & "; hits = " & topResults(i).Frequency
            If InStr(topResults(i).Char, Chr$(0)) = 0 Then  'if there is no \0 you can display the actual text, otherwise only hex
                strMsg = strMsg & "; string = " & topResults(i).Char
            Else
                strMsg = strMsg & "; --string not displayable--"
            End If
            
            hexChars = ""   'reinit hexChars
            For j = 1 To Len(topResults(i).Char) Step 1
                hexTemp = Hex$(Asc(Mid$(topResults(i).Char, j, 1)))
                If Len(hexTemp) = 1 Then hexTemp = "0" & hexTemp
                hexChars = hexChars & hexTemp
            Next j
            strMsg = strMsg & "; hex = " & hexChars
            strMsg = strMsg & "; len = " & Len(topResults(i).Char) & vbCrLf
        End Select  'end select skip reason
        'here is the new and improved format:
        '3: total space = 384; hits = 128; string = foo; hex = 666F6F; len = 3
        'or
        '4: total space = 384; hits = 128; string = --not displayable--; hex = 101300; len = 3
    Next i
    strMsg = vbCrLf & vbCrLf & strMsg
    dTime = beginTime - Now 'figure the total time taken
    strMsg = vbCrLf & "Elapsed time: " & Format(dTime, "hh:nn:ss") & strMsg
    For i = UBound(intLens) To 0 Step -1
        strMsg = intLens(i) & " " & strMsg
    Next i
    cBytesCounted = cBytesCounted + (intLens(UBound(intLens)) - 1)  'doctor up cbytescounted so that it accounts for read method
    strMsg = "Total characters: " & cTotalBytes & "     Characters actually counted: " & cBytesCounted & vbCrLf & "Number of different strings: " & cChars & "    Sample lengths requested:" & strMsg
    For i = UBound(strFilename) - 1 To 0 Step -1
        strMsg = "Filename:" & strFilename(i) & vbCrLf & strMsg
    Next i
    If bOutput Then 'if they requested to write a long report, do so
        WriteReport strMsg
    End If
    If Len(strMsg) > 32767 Then 'now truncate the displayed version
        strMsg = "--Report truncated because size is too great for text box. To remedy this situation," _
          & " reduce the number of results to be shown that aren't really relevant anyway.--" & vbCrLf & vbCrLf & _
        Left$(strMsg, 32000) & vbCrLf & vbCrLf & "--Report truncated because size is too great for text box. To remedy this situation," _
          & " reduce the number of results to be shown that aren't really relevant anyway.--"
    End If
    txtResults.Text = strMsg
    If bBeep Then Beep  'beep on finish
    bCancel = False 'and reset cancel flag
    MsgBox "Average Hops:" & (hops / cChars)
End Sub
Private Sub TopSearch(node As HufNode)
    If maxFreq.Frequency < node.Freq Then
        maxFreq.Frequency = node.Freq   'save so that future compares will be accurate
        Set maxNode = node 'and return this node so that it can be 0d out if it's the max.
    End If
    If node.IsLeft Then
        TopSearch node.Left
    End If
    If node.IsRight Then
        TopSearch node.Right
    End If
End Sub
Private Sub cmdBack_Click()
    If SSTab1.Tab = 4 Then  'go ALL the way back
        SSTab1.Tab = 0
    Else
        SSTab1.Tab = SSTab1.Tab - 1
    End If
End Sub

Private Sub cmdNext_Click()
    SSTab1.Tab = SSTab1.Tab + 1
End Sub

Private Sub cmdOpen_Click()
On Error GoTo OpenErr   'just quit if they hit cancel
    CommonDialog1.Filter = "All Files|*.*|Text Files(*.txt)|*.txt"
    CommonDialog1.FilterIndex = 1
    CommonDialog1.filename = ""
    CommonDialog1.ShowOpen
    Close   'close any already open files (just in case)
    
    ReDim Preserve strFilename(0 To UBound(strFilename) + 1)    'alloc another filename first
    strFilename(UBound(strFilename) - 1) = CommonDialog1.filename
    
    cmdAnalyze.Enabled = True
    If UBound(strFilename) = 1 Then
        lblOpen.Caption = strFilename(UBound(strFilename) - 1)
    Else
        lblOpen.Caption = lblOpen.Caption & ";" & strFilename(UBound(strFilename) - 1)
    End If
    Exit Sub
OpenErr:
    If Err.Number <> cdlCancel Then
        MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Hit OK to close Martial. If you believe you received this error in error please contact the creator at sandersn@hotmail.com", vbOKOnly, "Huge Error!"
    End If
End Sub

Private Sub cmdOpenExcludeTable_Click()
On Error GoTo OpenExcludeTableErr   'just quit if they hit cancel
    bTableExclude = True

    CommonDialog1.Filter = "All Files|*.*|Table Files(*.tbl)|*.tbl"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.filename = ""
    CommonDialog1.ShowOpen  'because we're going to just read here
    strFilenameTableExclude = CommonDialog1.filename
    lblOpenExcludeTable.Caption = strFilenameTableExclude
    Exit Sub
OpenExcludeTableErr:
    If Err.Number <> cdlCancel Then
        MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Hit OK to close Martial. If you believe you received this error in error please contact the creator at sandersn@hotmail.com", vbOKOnly, "Huge Error!"
    End If
    
End Sub

Private Sub cmdOpenOutput_Click()
On Error GoTo OpenOutputErr   'just quit if they hit cancel
    bOutput = True
    CommonDialog1.Filter = "All Files|*.*|Text Files(*.txt)|*.txt|"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.filename = ""
    CommonDialog1.ShowSave  'because we're going to append here
    strFilenameReport = CommonDialog1.filename
    lblOpenOutput.Caption = strFilenameReport
    Exit Sub
OpenOutputErr:
    If Err.Number <> cdlCancel Then
        MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Hit OK to close Martial. If you believe you received this error in error please contact the creator at sandersn@hotmail.com", vbOKOnly, "Huge Error!"
    End If

End Sub

Private Sub cmdOpenWriteTable_Click()
On Error GoTo OpenWriteTableErr   'just quit if they hit cancel
    bTableWrite = True
    CommonDialog1.Filter = "All Files|*.*|Table Files(*.tbl)|*.tbl"
    CommonDialog1.FilterIndex = 2
    CommonDialog1.filename = ""
    CommonDialog1.ShowSave  'because we're going to append here
    strFilenameTableWrite = CommonDialog1.filename
    lblOpenWriteTable.Caption = strFilenameTableWrite
    Exit Sub
OpenWriteTableErr:
    If Err.Number <> cdlCancel Then
        MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & vbCrLf & "Hit OK to close Martial. If you believe you received this error in error please contact the creator at sandersn@hotmail.com", vbOKOnly, "Huge Error!"
    End If

End Sub

Private Sub cmdReset_Click()
Dim i As Integer
    'reset all global variables to their initial state
    ReDim strFilename(0 To 0) As String
    strFilenameTableExclude = ""
    strFilenameReport = ""
    strFilenameTableWrite = ""
    bComments = False
    bTableExclude = False
    bTableWrite = False
    bOutput = False
    bBeep = False
    'GUI stuffs
    cmdAnalyze.Enabled = False
    lblOpen.Caption = "--No File Specified--"
    txtSampleLength.Text = "1"
    txtTopResults.Text = "25"
    lblOpenOutput.Caption = "--No File Specified--"
    chkBeepOnFinish.Value = vbUnchecked
    lblOpenWriteTable.Caption = "--No File Specified--"
    txtWriteTableRanges.Text = ""
    chkUseComments.Value = vbUnchecked
        fraComments.Enabled = False
        chkReadInsideComments.Enabled = False
        chkReadInsideComments.Value = vbUnchecked
        optC.Value = True
        optC.Enabled = False
        optCPP.Enabled = False
        optThingy.Enabled = False
        optCustom.Enabled = False
        lblCEnd.Enabled = False
        lblCPPEnd.Enabled = False
        lblThingyEnd.Enabled = False
        txtBeginComment.Enabled = False
        txtBeginComment.Text = ""
        txtEndComment.Enabled = False
        txtEndComment.Text = ""
        lblBegin.Enabled = False
        lblEnd.Enabled = False
    lblOpenExcludeTable.Caption = "--No File Specified--"
    chkIgnoreSpace.Value = vbUnchecked
    chkIgnoreReturn.Value = vbUnchecked
    chkIgnoreTab.Value = vbUnchecked
    'and set to initial view too
    SSTab1.Tab = 0
    cmdOpen.SetFocus
End Sub

Private Sub Form_Load()
    SSTab1.TabEnabled(3) = False    'because you can't set this from the property sheets
    SSTab1.TabEnabled(4) = False
    SSTab1.Tab = 0  'just for safekeeping because else itstarts on the last tab the *programmer* looked at
    ReDim strFilename(0 To 0) As String
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Close
End Sub


Private Sub lblMartial_DblClick()
    Label1.Visible = Not Label1.Visible
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)    'this is really like a changefocus of tab event..
    Select Case SSTab1.Tab
    Case 0  'input
        cmdBack.Enabled = False
        cmdNext.Enabled = True
    Case 1  'output
        cmdBack.Enabled = True
        cmdNext.Enabled = True
    Case 2  'formatting
        cmdBack.Enabled = True
        cmdNext.Enabled = False
    Case 3  'analyse
        cmdBack.Enabled = False
        cmdNext.Enabled = False
    Case 4  'results
        cmdBack.Enabled = True
        cmdNext.Enabled = False
    End Select
End Sub

Private Sub txtSampleLength_GotFocus()
    txtSampleLength.SelStart = 0
    txtSampleLength.SelLength = Len(txtSampleLength.Text)
End Sub
Private Sub txtTopResults_GotFocus()
    txtTopResults.SelStart = 0
    txtTopResults.SelLength = Len(txtTopResults.Text)
End Sub
Private Sub txtWriteTableRanges_GotFocus()
    txtWriteTableRanges.SelStart = 0
    txtWriteTableRanges.SelLength = Len(txtWriteTableRanges.Text)
End Sub
