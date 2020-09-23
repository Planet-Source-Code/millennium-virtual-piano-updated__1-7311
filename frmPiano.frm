VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yamaha XG Virtual Piano - Copyright 2001 Ipong Corporation"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11910
   ForeColor       =   &H00000000&
   Icon            =   "frmPiano.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmPiano.frx":0442
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   11910
   Begin VB.HScrollBar Panpot 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   1440
      Max             =   127
      TabIndex        =   166
      TabStop         =   0   'False
      Top             =   6585
      Value           =   64
      Width           =   1155
   End
   Begin VB.HScrollBar Pitch 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   10620
      Max             =   127
      TabIndex        =   122
      TabStop         =   0   'False
      Top             =   5625
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar VariationChorus 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   90
      Max             =   127
      TabIndex        =   159
      TabStop         =   0   'False
      Top             =   3375
      Value           =   5
      Width           =   1215
   End
   Begin VB.HScrollBar VariationReverb 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   90
      Max             =   127
      TabIndex        =   157
      TabStop         =   0   'False
      Top             =   2895
      Value           =   5
      Width           =   1215
   End
   Begin VB.HScrollBar VariationDry 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   90
      Max             =   127
      TabIndex        =   155
      TabStop         =   0   'False
      Top             =   3855
      Value           =   5
      Width           =   1215
   End
   Begin VB.HScrollBar Decay 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   5340
      Max             =   127
      TabIndex        =   153
      TabStop         =   0   'False
      Top             =   5145
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar ModulationAMod 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   10620
      Max             =   127
      TabIndex        =   151
      TabStop         =   0   'False
      Top             =   2190
      Value           =   5
      Width           =   1215
   End
   Begin VB.HScrollBar ModulationFMod 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   10620
      Max             =   127
      TabIndex        =   149
      TabStop         =   0   'False
      Top             =   1710
      Value           =   5
      Width           =   1215
   End
   Begin VB.HScrollBar ModulationPMod 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   10620
      Max             =   127
      TabIndex        =   147
      TabStop         =   0   'False
      Top             =   1230
      Value           =   10
      Width           =   1215
   End
   Begin VB.HScrollBar ModulationAmp 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   10620
      Max             =   127
      TabIndex        =   145
      TabStop         =   0   'False
      Top             =   750
      Value           =   64
      Width           =   1215
   End
   Begin VB.OptionButton VoiceOption 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "VL"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   3
      Left            =   1920
      TabIndex        =   108
      TabStop         =   0   'False
      Top             =   5940
      Width           =   465
   End
   Begin VB.OptionButton VoiceOption 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "Drum"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   2
      Left            =   2880
      TabIndex        =   110
      TabStop         =   0   'False
      Top             =   5940
      Width           =   690
   End
   Begin VB.OptionButton VoiceOption 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "FX"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   1
      Left            =   2385
      TabIndex        =   109
      TabStop         =   0   'False
      Top             =   5940
      Width           =   510
   End
   Begin VB.OptionButton VoiceOption 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      Caption         =   "Ins"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   316
      Index           =   0
      Left            =   1410
      TabIndex        =   107
      TabStop         =   0   'False
      Top             =   5940
      Value           =   -1  'True
      Width           =   525
   End
   Begin VB.CheckBox SoftPedal 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Soft Pedal"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   117
      TabStop         =   0   'False
      Top             =   6525
      Width           =   855
   End
   Begin VB.CheckBox Sostenuto 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Sostenuto"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   116
      TabStop         =   0   'False
      Top             =   6225
      Width           =   855
   End
   Begin VB.CheckBox Sustain 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Sustain"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   5925
      Width           =   855
   End
   Begin VB.HScrollBar VelSenseOffset 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   90
      Max             =   127
      TabIndex        =   142
      TabStop         =   0   'False
      Top             =   7590
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar VelSenseDepth 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   90
      Max             =   127
      TabIndex        =   141
      TabStop         =   0   'False
      Top             =   7110
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar VLFilterEqDepth 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   1440
      Max             =   127
      TabIndex        =   139
      TabStop         =   0   'False
      Top             =   5625
      Value           =   64
      Width           =   1155
   End
   Begin VB.HScrollBar VLGrowl 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   9300
      Max             =   127
      TabIndex        =   132
      TabStop         =   0   'False
      Top             =   5625
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar VLBreath 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   7980
      Max             =   127
      TabIndex        =   131
      TabStop         =   0   'False
      Top             =   5625
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar VLScream 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   6660
      Max             =   127
      TabIndex        =   130
      TabStop         =   0   'False
      Top             =   5625
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar VLTonging 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   5340
      Max             =   127
      TabIndex        =   129
      TabStop         =   0   'False
      Top             =   5625
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar VLEmbrouchure 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   4020
      Max             =   127
      TabIndex        =   128
      TabStop         =   0   'False
      Top             =   5625
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar VLPressure 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   2700
      Max             =   127
      TabIndex        =   127
      TabStop         =   0   'False
      Top             =   5625
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar PitchRelTime 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   90
      Max             =   127
      TabIndex        =   126
      TabStop         =   0   'False
      Top             =   6585
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar PitchRelLev 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   90
      Max             =   127
      TabIndex        =   125
      TabStop         =   0   'False
      Top             =   6105
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar PitchAttTime 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   90
      Max             =   127
      TabIndex        =   124
      TabStop         =   0   'False
      Top             =   5625
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar PitchInitLev 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   90
      Max             =   127
      TabIndex        =   123
      TabStop         =   0   'False
      Top             =   5145
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar Velocity 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   1440
      Max             =   127
      TabIndex        =   121
      TabStop         =   0   'False
      Top             =   5145
      Value           =   127
      Width           =   1155
   End
   Begin VB.HScrollBar PortamentoCtrl 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   10620
      Max             =   127
      TabIndex        =   118
      TabStop         =   0   'False
      Top             =   4650
      Value           =   5
      Width           =   1215
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   720
   End
   Begin VB.ListBox chord 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      Columns         =   1
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   885
      IntegralHeight  =   0   'False
      ItemData        =   "frmPiano.frx":074C
      Left            =   4650
      List            =   "frmPiano.frx":0796
      TabIndex        =   115
      TabStop         =   0   'False
      Top             =   5910
      Width           =   7230
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   720
   End
   Begin VB.CommandButton cmdHook 
      Caption         =   "Hook"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   114
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2685
      TabIndex        =   112
      TabStop         =   0   'False
      Top             =   6585
      Width           =   450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3135
      TabIndex        =   111
      TabStop         =   0   'False
      Top             =   6585
      Width           =   435
   End
   Begin VB.HScrollBar Transpose 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   2700
      Max             =   48
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   5145
      Value           =   24
      Width           =   1215
   End
   Begin VB.CheckBox Preset 
      BackColor       =   &H00808080&
      Caption         =   "Preset Song"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   225
      Style           =   1  'Graphical
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   435
      Width           =   945
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   1800
      Top             =   720
   End
   Begin MSComctlLib.ListView ListView1 
      CausesValidation=   0   'False
      Height          =   4860
      Left            =   1410
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   45
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   8573
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   8421376
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "V"
         Text            =   "Voice"
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "B"
         Text            =   "Bank"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "P"
         Text            =   "Program"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.HScrollBar VariationTime 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   90
      Max             =   127
      TabIndex        =   101
      TabStop         =   0   'False
      Top             =   4335
      Value           =   5
      Width           =   1215
   End
   Begin VB.ComboBox VarType 
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      ItemData        =   "frmPiano.frx":0896
      Left            =   90
      List            =   "frmPiano.frx":08EC
      Style           =   2  'Dropdown List
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   2355
      Width           =   1215
   End
   Begin VB.CheckBox Variation 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Off"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   4605
      Width           =   1215
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "A#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   70
      Left            =   11175
      Style           =   1  'Graphical
      TabIndex        =   70
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "G#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   68
      Left            =   10935
      Style           =   1  'Graphical
      TabIndex        =   68
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "F#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   66
      Left            =   10695
      Style           =   1  'Graphical
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "D#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   63
      Left            =   10215
      Style           =   1  'Graphical
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "C#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   61
      Left            =   9975
      Style           =   1  'Graphical
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "A#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   58
      Left            =   9495
      Style           =   1  'Graphical
      TabIndex        =   58
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "G#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   56
      Left            =   9255
      Style           =   1  'Graphical
      TabIndex        =   56
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "F#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   54
      Left            =   9015
      Style           =   1  'Graphical
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "D#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   51
      Left            =   8535
      Style           =   1  'Graphical
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "C#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   49
      Left            =   8295
      Style           =   1  'Graphical
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "A#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   46
      Left            =   7815
      Style           =   1  'Graphical
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "G#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   44
      Left            =   7575
      Style           =   1  'Graphical
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "F#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   42
      Left            =   7335
      Style           =   1  'Graphical
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "D#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   39
      Left            =   6855
      Style           =   1  'Graphical
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "C#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   37
      Left            =   6615
      Style           =   1  'Graphical
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "A#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   34
      Left            =   6135
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "G#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   32
      Left            =   5895
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "F#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   30
      Left            =   5655
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "D#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   27
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "C#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   25
      Left            =   4935
      Style           =   1  'Graphical
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "A#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   22
      Left            =   4455
      Style           =   1  'Graphical
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "G#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   20
      Left            =   4215
      Style           =   1  'Graphical
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "F#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   18
      Left            =   3975
      Style           =   1  'Graphical
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "D#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   15
      Left            =   3495
      Style           =   1  'Graphical
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "C#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   13
      Left            =   3255
      Style           =   1  'Graphical
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "A#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   10
      Left            =   2775
      Style           =   1  'Graphical
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "G#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   8
      Left            =   2535
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "F#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   6
      Left            =   2295
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "D#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   3
      Left            =   1830
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         C"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   72
      Left            =   11535
      Style           =   1  'Graphical
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         B"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   71
      Left            =   11295
      Style           =   1  'Graphical
      TabIndex        =   71
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         A"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   69
      Left            =   11055
      Style           =   1  'Graphical
      TabIndex        =   69
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         G"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   67
      Left            =   10815
      Style           =   1  'Graphical
      TabIndex        =   67
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         F"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   65
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         E"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   64
      Left            =   10335
      Style           =   1  'Graphical
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         D"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   62
      Left            =   10095
      Style           =   1  'Graphical
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         C"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   60
      Left            =   9855
      Style           =   1  'Graphical
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         B"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   59
      Left            =   9615
      Style           =   1  'Graphical
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         A"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   57
      Left            =   9375
      Style           =   1  'Graphical
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         G"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   55
      Left            =   9135
      Style           =   1  'Graphical
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         F"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   53
      Left            =   8895
      Style           =   1  'Graphical
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         E"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   52
      Left            =   8655
      Style           =   1  'Graphical
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         D"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   50
      Left            =   8415
      Style           =   1  'Graphical
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         C"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   48
      Left            =   8175
      Style           =   1  'Graphical
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         B"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   47
      Left            =   7935
      Style           =   1  'Graphical
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         A"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   45
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         G"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   43
      Left            =   7455
      Style           =   1  'Graphical
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         F"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   41
      Left            =   7215
      Style           =   1  'Graphical
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         E"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   40
      Left            =   6975
      Style           =   1  'Graphical
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         D"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   38
      Left            =   6735
      Style           =   1  'Graphical
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         C"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   36
      Left            =   6495
      Style           =   1  'Graphical
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         B"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   35
      Left            =   6255
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         A"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   33
      Left            =   6015
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         G"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   31
      Left            =   5775
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         F"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   29
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         E"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   28
      Left            =   5295
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         D"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   26
      Left            =   5055
      Style           =   1  'Graphical
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         C"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   24
      Left            =   4815
      Style           =   1  'Graphical
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         B"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   23
      Left            =   4575
      Style           =   1  'Graphical
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         A"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   21
      Left            =   4335
      Style           =   1  'Graphical
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         G"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   19
      Left            =   4095
      Style           =   1  'Graphical
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         F"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   17
      Left            =   3855
      Style           =   1  'Graphical
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         E"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   16
      Left            =   3615
      Style           =   1  'Graphical
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         D"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   14
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         C"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   12
      Left            =   3135
      Style           =   1  'Graphical
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         B"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   11
      Left            =   2895
      Style           =   1  'Graphical
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         A"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   9
      Left            =   2655
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         G"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   7
      Left            =   2415
      Style           =   1  'Graphical
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         F"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   5
      Left            =   2175
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         E"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   4
      Left            =   1935
      Style           =   1  'Graphical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Caption         =   "C#"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   1
      Left            =   1605
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   6840
      Width           =   165
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         D"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   2
      Left            =   1695
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox touch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "         C"
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      ForeColor       =   &H0080C0FF&
      Height          =   1005
      Index           =   0
      Left            =   1455
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   6840
      Width           =   255
   End
   Begin VB.CheckBox Portamento 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Portamento"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10620
      Style           =   1  'Graphical
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   5040
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.HScrollBar PortamentoTime 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   10620
      Max             =   127
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   4170
      Value           =   5
      Width           =   1215
   End
   Begin VB.HScrollBar Modulation 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   10620
      Max             =   127
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   270
      Value           =   5
      Width           =   1215
   End
   Begin VB.HScrollBar Harmonic 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   9300
      Max             =   127
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   5145
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar Bright 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   7980
      Max             =   127
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   5145
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar Release 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   6660
      Max             =   127
      TabIndex        =   95
      TabStop         =   0   'False
      Top             =   5145
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar Attack 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   4020
      Max             =   127
      TabIndex        =   93
      TabStop         =   0   'False
      Top             =   5145
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar VibDelay 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   10620
      Max             =   127
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   3660
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar VibDepth 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   10620
      Max             =   127
      TabIndex        =   89
      TabStop         =   0   'False
      Top             =   3180
      Value           =   64
      Width           =   1215
   End
   Begin VB.HScrollBar VibRate 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   10620
      Max             =   127
      TabIndex        =   87
      TabStop         =   0   'False
      Top             =   2700
      Value           =   64
      Width           =   1215
   End
   Begin VB.ComboBox ChoType 
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      ItemData        =   "frmPiano.frx":09DE
      Left            =   90
      List            =   "frmPiano.frx":0A04
      Style           =   2  'Dropdown List
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   1725
      Width           =   1215
   End
   Begin VB.ComboBox RevType 
      BackColor       =   &H00FFFFFF&
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      ItemData        =   "frmPiano.frx":0A7D
      Left            =   90
      List            =   "frmPiano.frx":0AA3
      Style           =   2  'Dropdown List
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   885
      Width           =   1215
   End
   Begin VB.HScrollBar Reverb 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   90
      Max             =   127
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   1230
      Value           =   5
      Width           =   1215
   End
   Begin VB.HScrollBar Chorus 
      CausesValidation=   0   'False
      Height          =   195
      Left            =   90
      Max             =   127
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   2070
      Value           =   5
      Width           =   1215
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Panpot"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1440
      TabIndex        =   167
      Top             =   6375
      UseMnemonic     =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pitch Bend"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   10620
      TabIndex        =   74
      Top             =   5415
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   10560
      Top             =   5400
      Width           =   1320
   End
   Begin VB.Label Label39 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   165
      Top             =   4125
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label38 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Release Time"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   164
      Top             =   6375
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label37 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Release Level"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   163
      Top             =   5895
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label36 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Attack Time"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   162
      Top             =   5415
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label35 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Initial Level"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   161
      Top             =   4935
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label34 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send To Chrs"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   160
      Top             =   3165
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label33 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send To Rvrb"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   158
      Top             =   2685
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label32 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dry"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   156
      Top             =   3645
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Decay"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5340
      TabIndex        =   154
      Top             =   4935
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amp Depth"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   152
      Top             =   1980
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LPF Depth"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   150
      Top             =   1500
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pitch Depth"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   148
      Top             =   1020
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label27 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amp Ctrl"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   146
      Top             =   540
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vel Offset"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   144
      Top             =   7380
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vel Depth"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   143
      Top             =   6900
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label26 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VL Eq Depth"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   1440
      TabIndex        =   140
      Top             =   5415
      UseMnemonic     =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label25 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Embrouchure"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4020
      TabIndex        =   138
      Top             =   5415
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label24 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tonging"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5340
      TabIndex        =   137
      Top             =   5415
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label23 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scream"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6660
      TabIndex        =   136
      Top             =   5415
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Breath"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7980
      TabIndex        =   135
      Top             =   5415
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Growl"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9300
      TabIndex        =   134
      Top             =   5415
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pressure"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2700
      TabIndex        =   133
      Top             =   5415
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Porta.Time"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   120
      Top             =   3960
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Porta.Ctrl"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   119
      Top             =   4440
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      Enabled         =   0   'False
      Height          =   360
      Left            =   45
      Picture         =   "frmPiano.frx":0B06
      Stretch         =   -1  'True
      Top             =   60
      Width           =   1305
   End
   Begin VB.Label midiStatus 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   195
      Left            =   2685
      TabIndex        =   113
      Top             =   6375
      UseMnemonic     =   0   'False
      Width           =   870
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Transpose"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2700
      TabIndex        =   82
      Top             =   4935
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Chorus"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   78
      Top             =   1515
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reverb"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   75
      Top             =   675
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Modulation"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   84
      Top             =   60
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   1440
      TabIndex        =   73
      Top             =   4935
      UseMnemonic     =   0   'False
      Width           =   1155
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Harmonic"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9300
      TabIndex        =   98
      Top             =   4935
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Brightness"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   7980
      TabIndex        =   96
      Top             =   4935
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Release"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6660
      TabIndex        =   94
      Top             =   4935
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Attack"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4020
      TabIndex        =   92
      Top             =   4935
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vibrato Delay"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   90
      Top             =   3450
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vibrato Depth"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   88
      Top             =   2970
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Vibrato Rate"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   10620
      TabIndex        =   86
      Top             =   2490
      UseMnemonic     =   0   'False
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1680
      Left            =   30
      Top             =   645
      Width           =   1350
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1440
      Left            =   10560
      Top             =   3945
      Width           =   1320
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1920
      Left            =   30
      Top             =   4920
      Width           =   1350
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H0000C000&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   1410
      Top             =   5400
      Width           =   9135
   End
   Begin VB.Shape Shape5 
      BackColor       =   &H000040C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   1410
      Top             =   4920
      Width           =   9135
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00800080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1005
      Left            =   30
      Top             =   6840
      Width           =   1350
   End
   Begin VB.Shape Shape8 
      BackColor       =   &H00008080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2580
      Left            =   30
      Top             =   2325
      Width           =   1350
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00000080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   2400
      Left            =   10560
      Top             =   45
      Width           =   1320
   End
   Begin VB.Shape Shape10 
      BackColor       =   &H00008080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1440
      Left            =   10560
      Top             =   2475
      Width           =   1320
   End
   Begin VB.Shape Shape11 
      BackColor       =   &H00C000C0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   465
      Left            =   1410
      Top             =   6360
      Width           =   1230
   End
   Begin VB.Menu MIDI_Device 
      Caption         =   "MIDI Device"
      Begin VB.Menu Device 
         Caption         =   ""
         Index           =   0
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   7
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu Device 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   9
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Base_Note 
      Caption         =   "Base Note"
      Begin VB.Menu Base 
         Caption         =   "C1 (12)"
         Checked         =   -1  'True
         Index           =   12
      End
      Begin VB.Menu Base 
         Caption         =   "C2 (24)"
         Index           =   24
      End
      Begin VB.Menu Base 
         Caption         =   "C3 (36)"
         Index           =   36
      End
   End
   Begin VB.Menu Sel_Channel 
      Caption         =   "Channel"
      Begin VB.Menu Chn 
         Caption         =   "1"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu Chn 
         Caption         =   "2"
         Index           =   1
      End
      Begin VB.Menu Chn 
         Caption         =   "3"
         Index           =   2
      End
      Begin VB.Menu Chn 
         Caption         =   "4"
         Index           =   3
      End
      Begin VB.Menu Chn 
         Caption         =   "5"
         Index           =   4
      End
      Begin VB.Menu Chn 
         Caption         =   "6"
         Index           =   5
      End
      Begin VB.Menu Chn 
         Caption         =   "7"
         Index           =   6
      End
      Begin VB.Menu Chn 
         Caption         =   "8"
         Index           =   7
      End
      Begin VB.Menu Chn 
         Caption         =   "9"
         Index           =   8
      End
      Begin VB.Menu Chn 
         Caption         =   "10 (Drum)"
         Index           =   9
      End
      Begin VB.Menu Chn 
         Caption         =   "11"
         Index           =   10
      End
      Begin VB.Menu Chn 
         Caption         =   "12"
         Index           =   11
      End
      Begin VB.Menu Chn 
         Caption         =   "13"
         Index           =   12
      End
      Begin VB.Menu Chn 
         Caption         =   "14"
         Index           =   13
      End
      Begin VB.Menu Chn 
         Caption         =   "15"
         Index           =   14
      End
      Begin VB.Menu Chn 
         Caption         =   "16"
         Index           =   15
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu About 
         Caption         =   "About"
      End
      Begin VB.Menu Panic 
         Caption         =   "Panic"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Initialize()
    If App.PrevInstance Then End
    dumpResFile
End Sub
Private Sub Form_Load()
    ' Hooking cmdHook for mci callback
    pOldProc = SetWindowLong(cmdHook.hwnd, GWL_WNDPROC, AddressOf WindowProc)
    ' Location of midi file
    locSong = GetShortName(validPath(App.Path) & "song.mid")
    ' set Yamaha SysEx ID
    YMH_MSG = Chr(&HF0) & Chr(&H43) & Chr(&H10) & Chr(&H4C)
    ' Set default chord
    chord.Text = chord.List(0)
    ' Set default channel
    Channel = 0
    ' Set default basenote
    baseNote = 12
    ' Set default volume
    volume = 127
    ' Get and set MIDI Device
    Dim i As Long
    Dim caps As MIDIOUTCAPS
    ' Set the first device as midi mapper
    Device(0).Caption = "MIDI Mapper"
    Device(0).Visible = True
    Device(0).Enabled = True
    ' Get the rest of the midi devices
    numDevices = midiOutGetNumDevs()
    For i = 0 To (numDevices - 1)
        midiOutGetDevCaps i, caps, Len(caps)
        Device(i + 1).Caption = caps.szPname
        Device(i + 1).Visible = True
        Device(i + 1).Enabled = True
        If InStr(1, caps.szPname, "Yamaha", vbTextCompare) = 1 Then YmhDvc = i + 1
        ' maximum item for midi devices is 10
        If i = 8 Then Exit For
    Next
    device_Click (YmhDvc)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    rc = SetWindowLong(cmdHook.hwnd, GWL_WNDPROC, pOldProc)
    rc = midiOutClose(hmidi)
    rc = mciSendString("stop midi", 0&, 0, 0)
    rc = mciSendString("close midi", 0&, 0, 0)
    killResFile
End Sub
Private Sub touch_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    StopNote LastDrag, lastChordInDrag
End Sub
Private Sub touch_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Select Case State
        Case 0
            touch(Index).DragIcon = Form1.MouseIcon
            lastChordInDrag = chord.ItemData(chord.ListIndex)
            StartNote Index, lastChordInDrag
            LastDrag = Index
        Case 1
            StopNote LastDrag, lastChordInDrag
    End Select
End Sub
' Press the button and play midi note
Sub PlayMIDINote(ByVal nNote As Integer, ByVal nAdd As Integer)
    midiMsg = &H90 + ((baseNote + nNote + nAdd) * &H100) + (volume * &H10000) + Channel
    midiOutShortMsg hmidi, midiMsg
    touch(nNote + nAdd).Value = 1
End Sub
' send midi start event
Sub StartNote(Index As Integer, nChord As Integer)
    If (touch(Index).Value = 1) Then Exit Sub
    If nChord <> -1 Then
        If Index + 12 > 72 Or Index - 12 < 0 Then Exit Sub
    End If
    Select Case nChord
        Case Dominant7thMin5
            Select Case Index Mod 12
                Case 0 To 3, 11
                    ' 0;2;-6;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 2
                    PlayMIDINote Index, -6
                    PlayMIDINote Index, -2
                Case 4
                    ' 0;-10;-6;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -10
                    PlayMIDINote Index, -6
                    PlayMIDINote Index, -2
                Case 5 To 7
                    ' 0;2;6;10
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 2
                    PlayMIDINote Index, 6
                    PlayMIDINote Index, 10
                Case 8 To 10
                    ' 0;2;6;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 2
                    PlayMIDINote Index, 6
                    PlayMIDINote Index, -2
            End Select
        Case Diminish
            Select Case Index Mod 12
                Case 0, 1, 11
                    ' 0;3;-6;-3
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, -6
                    PlayMIDINote Index, -3
                Case 2 To 4
                    ' 0;-9;-6;-3
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -9
                    PlayMIDINote Index, -6
                    PlayMIDINote Index, -3
                Case 5 To 8
                    ' 0;3;6;9
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, 6
                    PlayMIDINote Index, 9
                Case 9, 10
                    ' 0;3;6;-3
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, 6
                    PlayMIDINote Index, -3
            End Select
        Case Minor7thMin5
            Select Case Index Mod 12
                Case 0, 1, 11
                    ' 0;3;-6;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, -6
                    PlayMIDINote Index, -2
                Case 2 To 4
                    ' 0;-9;-6;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -9
                    PlayMIDINote Index, -6
                    PlayMIDINote Index, -2
                Case 5 To 7
                    ' 0;3;6;10
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, 6
                    PlayMIDINote Index, 10
                Case 8 To 10
                    ' 0;3;6;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, 6
                    PlayMIDINote Index, -2
            End Select
        Case Minor6th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;3;-5;-3
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -3
                Case 2 To 4
                    ' 0;-9;-5;-3
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -9
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -3
                Case 5 To 7
                    ' 0;3;7;9
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, 9
                Case 8, 9
                    ' 0;3;7;-3
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, -3
            End Select
        Case Minor7th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;3;-5;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 2 To 4
                    ' 0;-9;-5;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -9
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 5 To 7
                    ' 0;3;7;10
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, 10
                Case 8, 9
                    ' 0;3;7;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, -2
            End Select
        Case Minor
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;3;-5
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, -5
                Case 2 To 4
                    ' 0;-5;-9
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -9
                Case 5 To 9
                    ' 0;3;7
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, 7
            End Select
        Case Major6th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;4;-5;-3
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -3
                Case 2 To 4
                    ' 0;-8;-5;-3
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -8
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -3
                Case 5 To 7
                    ' 0;4;7;9
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, 9
                Case 8, 9
                    ' 0;4;7;-3
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, -3
            End Select
        Case Dominant7th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;4;-5;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 2 To 4
                    ' 0;-8;-5;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -8
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 5 To 7
                    ' 0;4;7;10
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, 10
                Case 8, 9
                    ' 0;4;7;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, -2
            End Select
        Case Major7th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;4;-5;-1
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -1
                Case 2 To 4
                    ' 0;-8;-5;-1
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -8
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -1
                Case 5 To 7
                    ' 0;4;7;11
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, 11
                Case 8, 9
                    ' 0;4;7;-1
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, -1
            End Select
        Case Major
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;4;-5
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, -5
                Case 2 To 4
                    ' 0;-5;-8
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -8
                Case 5 To 9
                    ' 0;4;7
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 7
            End Select
        Case Augmented7th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;4;-4;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, -4
                    PlayMIDINote Index, -2
                Case 2 To 4
                    ' 0;-8;-4;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -8
                    PlayMIDINote Index, -4
                    PlayMIDINote Index, -2
                Case 5, 6
                    ' 0;4;8;10
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 8
                    PlayMIDINote Index, 10
                Case 7 To 9
                    ' 0;4;8;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 8
                    PlayMIDINote Index, -2
            End Select
        Case Augment
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;4;-4
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, -4
                Case 2 To 4
                    ' 0;-8;-4
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -8
                    PlayMIDINote Index, -4
                Case 5 To 9
                    ' 0;4;8
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 8
            End Select
        Case Dominant7thSuspended4th
            Select Case Index Mod 12
                Case 0, 10, 11
                    ' 0;5;-5;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 5
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 1 To 4
                    ' 0;-7;-5;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -7
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 5 To 7
                    ' 0;5;7;10
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 5
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, 10
                Case 8, 9
                    ' 0;5;7;-2
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 5
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, -2
            End Select
        Case Suspended4th
            Select Case Index Mod 12
                Case 0, 10, 11
                    ' 0;5;-5
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 5
                    PlayMIDINote Index, -5
                Case 1 To 4
                    ' 0;-7;-5
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, -7
                    PlayMIDINote Index, -5
                Case 5 To 9
                    ' 0;5;7
                    PlayMIDINote Index, 0
                    PlayMIDINote Index, 5
                    PlayMIDINote Index, 7
            End Select
        Case Minor7thMin9
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 1;3;-5;-2
                    PlayMIDINote Index, 1
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 2 To 4
                    ' 1;-9;-5;-2
                    PlayMIDINote Index, 1
                    PlayMIDINote Index, -9
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 5 To 7
                    ' 1;3;7;10
                    PlayMIDINote Index, 1
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, 10
                Case 8, 9
                    ' 1;3;7;-2
                    PlayMIDINote Index, 1
                    PlayMIDINote Index, 3
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, -2
            End Select
        Case MajorMin9
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 1;4;-5;-2
                    PlayMIDINote Index, 1
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 2 To 4
                    ' 1;-8;-5;-2
                    PlayMIDINote Index, 1
                    PlayMIDINote Index, -8
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 5 To 7
                    ' 1;4;7;10
                    PlayMIDINote Index, 1
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, 10
                Case 8, 9
                    ' 1;4;7;-2
                    PlayMIDINote Index, 1
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, -2
            End Select
        Case MajorMin9Suspended4th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 1;5;-5;-2
                    PlayMIDINote Index, 1
                    PlayMIDINote Index, 5
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 2 To 4
                    ' 1;-7;-5;-2
                    PlayMIDINote Index, 1
                    PlayMIDINote Index, -7
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 5 To 7
                    ' 1;5;7;10
                    PlayMIDINote Index, 1
                    PlayMIDINote Index, 5
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, 10
                Case 8, 9
                    ' 1;5;7;-2
                    PlayMIDINote Index, 1
                    PlayMIDINote Index, 5
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, -2
            End Select
        Case Major9th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 2;4;-5;-2
                    PlayMIDINote Index, 2
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 2 To 4
                    ' 2;-8;-5;-2
                    PlayMIDINote Index, 2
                    PlayMIDINote Index, -8
                    PlayMIDINote Index, -5
                    PlayMIDINote Index, -2
                Case 5 To 7
                    ' 2;4;7;10
                    PlayMIDINote Index, 2
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, 10
                Case 8, 9
                    ' 2;4;7;-2
                    PlayMIDINote Index, 2
                    PlayMIDINote Index, 4
                    PlayMIDINote Index, 7
                    PlayMIDINote Index, -2
            End Select
        Case Octave
            ' 0;12
            PlayMIDINote Index, 0
            PlayMIDINote Index, 12
        Case Else
            ' Single Note
            PlayMIDINote Index, 0
    End Select
End Sub
' Raise the button and stop midi note
Sub StopMIDINote(ByVal nNote As Integer, ByVal nAdd As Integer)
    midiMsg = &H80 + ((baseNote + nNote + nAdd) * &H100) + Channel
    midiOutShortMsg hmidi, midiMsg
    touch(nNote + nAdd).Value = 0
End Sub
' send midi stop event
Private Sub StopNote(Index As Integer, nChord As Integer)
    If nChord <> -1 Then
        If Index + 12 > 72 Or Index - 12 < 0 Then Exit Sub
    End If
    Select Case nChord
        Case Dominant7thMin5
            Select Case Index Mod 12
                Case 0 To 3, 11
                    ' 0;2;-6;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, 2
                    StopMIDINote Index, -6
                    StopMIDINote Index, -2
                Case 4
                    ' 0;-10;-6;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, -10
                    StopMIDINote Index, -6
                    StopMIDINote Index, -2
                Case 5 To 7
                    ' 0;2;6;10
                    StopMIDINote Index, 0
                    StopMIDINote Index, 2
                    StopMIDINote Index, 6
                    StopMIDINote Index, 10
                Case 8 To 10
                    ' 0;2;6;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, 2
                    StopMIDINote Index, 6
                    StopMIDINote Index, -2
            End Select
        Case Diminish
            Select Case Index Mod 12
                Case 0, 1, 11
                    ' 0;3;-6;-3
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, -6
                    StopMIDINote Index, -3
                Case 2 To 4
                    ' 0;-9;-6;-3
                    StopMIDINote Index, 0
                    StopMIDINote Index, -9
                    StopMIDINote Index, -6
                    StopMIDINote Index, -3
                Case 5 To 8
                    ' 0;3;6;9
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, 6
                    StopMIDINote Index, 9
                Case 5 To 8
                    ' 0;3;6;9
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, 6
                    StopMIDINote Index, 9
                Case 9, 10
                    ' 0;3;6;-3
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, 6
                    StopMIDINote Index, -3
            End Select
        Case Minor7thMin5
            Select Case Index Mod 12
                Case 0, 1, 11
                    ' 0;3;-6;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, -6
                    StopMIDINote Index, -2
                Case 2 To 4
                    ' 0;-9;-6;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, -9
                    StopMIDINote Index, -6
                    StopMIDINote Index, -2
                Case 5 To 7
                    ' 0;3;6;10
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, 6
                    StopMIDINote Index, 10
                Case 8 To 10
                    ' 0;3;6;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, 6
                    StopMIDINote Index, -2
            End Select
        Case Minor6th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;3;-5;-3
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, -5
                    StopMIDINote Index, -3
                Case 2 To 4
                    ' 0;-9;-5;-3
                    StopMIDINote Index, 0
                    StopMIDINote Index, -9
                    StopMIDINote Index, -5
                    StopMIDINote Index, -3
                Case 5 To 7
                    ' 0;3;7;9
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, 7
                    StopMIDINote Index, 9
                Case 8, 9
                    ' 0;3;7;-3
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, 7
                    StopMIDINote Index, -3
            End Select
        Case Minor7th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;3;-5;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 2 To 4
                    ' 0;-9;-5;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, -9
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 5 To 7
                    ' 0;3;7;10
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, 7
                    StopMIDINote Index, 10
                Case 8, 9
                    ' 0;3;7;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, 7
                    StopMIDINote Index, -2
            End Select
        Case Minor
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;3;-5
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, -5
                Case 2 To 4
                    ' 0;-5;-9
                    StopMIDINote Index, 0
                    StopMIDINote Index, -5
                    StopMIDINote Index, -9
                Case 5 To 9
                    ' 0;3;7
                    StopMIDINote Index, 0
                    StopMIDINote Index, 3
                    StopMIDINote Index, 7
            End Select
        Case Major6th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;4;-5;-3
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, -5
                    StopMIDINote Index, -3
                Case 2 To 4
                    ' 0;-8;-5;-3
                    StopMIDINote Index, 0
                    StopMIDINote Index, -8
                    StopMIDINote Index, -5
                    StopMIDINote Index, -3
                Case 5 To 7
                    ' 0;4;7;9
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, 7
                    StopMIDINote Index, 9
                Case 8, 9
                    ' 0;4;7;-3
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, 7
                    StopMIDINote Index, -3
            End Select
        Case Dominant7th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;4;-5;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 2 To 4
                    ' 0;-8;-5;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, -8
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 5 To 7
                    ' 0;4;7;10
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, 7
                    StopMIDINote Index, 10
                Case 8, 9
                    ' 0;4;7;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, 7
                    StopMIDINote Index, -2
            End Select
        Case Major7th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;4;-5;-1
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, -5
                    StopMIDINote Index, -1
                Case 2 To 4
                    ' 0;-8;-5;-1
                    StopMIDINote Index, 0
                    StopMIDINote Index, -8
                    StopMIDINote Index, -5
                    StopMIDINote Index, -1
                Case 5 To 7
                    ' 0;4;7;11
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, 7
                    StopMIDINote Index, 11
                Case 8, 9
                    ' 0;4;7;-1
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, 7
                    StopMIDINote Index, -1
            End Select
        Case Major
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;4;-5
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, -5
                Case 2 To 4
                    ' 0;-5;-8
                    StopMIDINote Index, 0
                    StopMIDINote Index, -5
                    StopMIDINote Index, -8
                Case 5 To 9
                    ' 0;4;7
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, 7
            End Select
        Case Augmented7th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;4;-4;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, -4
                    StopMIDINote Index, -2
                Case 2 To 4
                    ' 0;-8;-4;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, -8
                    StopMIDINote Index, -4
                    StopMIDINote Index, -2
                Case 5, 6
                    ' 0;4;8;10
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, 8
                    StopMIDINote Index, 10
                Case 7 To 9
                    ' 0;4;8;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, 8
                    StopMIDINote Index, -2
            End Select
        Case Augment
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 0;4;-4
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, -4
                Case 2 To 4
                    ' 0;-8;-4
                    StopMIDINote Index, 0
                    StopMIDINote Index, -8
                    StopMIDINote Index, -4
                Case 5 To 9
                    ' 0;4;8
                    StopMIDINote Index, 0
                    StopMIDINote Index, 4
                    StopMIDINote Index, 8
            End Select
        Case Dominant7thSuspended4th
            Select Case Index Mod 12
                Case 0, 10, 11
                    ' 0;5;-5;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, 5
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 1 To 4
                    ' 0;-7;-5;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, -7
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 5 To 7
                    ' 0;5;7;10
                    StopMIDINote Index, 0
                    StopMIDINote Index, 5
                    StopMIDINote Index, 7
                    StopMIDINote Index, 10
                Case 8, 9
                    ' 0;5;7;-2
                    StopMIDINote Index, 0
                    StopMIDINote Index, 5
                    StopMIDINote Index, 7
                    StopMIDINote Index, -2
            End Select
        Case Suspended4th
            Select Case Index Mod 12
                Case 0, 10, 11
                    ' 0;5;-5
                    StopMIDINote Index, 0
                    StopMIDINote Index, 5
                    StopMIDINote Index, -5
                Case 1 To 4
                    ' 0;-7;-5
                    StopMIDINote Index, 0
                    StopMIDINote Index, -7
                    StopMIDINote Index, -5
                Case 5 To 9
                    ' 0;5;7
                    StopMIDINote Index, 0
                    StopMIDINote Index, 5
                    StopMIDINote Index, 7
            End Select
        Case Minor7thMin9
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 1;3;-5;-2
                    StopMIDINote Index, 1
                    StopMIDINote Index, 3
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 2 To 4
                    ' 1;-9;-5;-2
                    StopMIDINote Index, 1
                    StopMIDINote Index, -9
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 5 To 7
                    ' 1;3;7;10
                    StopMIDINote Index, 1
                    StopMIDINote Index, 3
                    StopMIDINote Index, 7
                    StopMIDINote Index, 10
                Case 8, 9
                    ' 1;3;7;-2
                    StopMIDINote Index, 1
                    StopMIDINote Index, 3
                    StopMIDINote Index, 7
                    StopMIDINote Index, -2
            End Select
        Case MajorMin9
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 1;4;-5;-2
                    StopMIDINote Index, 1
                    StopMIDINote Index, 4
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 2 To 4
                    ' 1;-8;-5;-2
                    StopMIDINote Index, 1
                    StopMIDINote Index, -8
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 5 To 7
                    ' 1;4;7;10
                    StopMIDINote Index, 1
                    StopMIDINote Index, 4
                    StopMIDINote Index, 7
                    StopMIDINote Index, 10
                Case 8, 9
                    ' 1;4;7;-2
                    StopMIDINote Index, 1
                    StopMIDINote Index, 4
                    StopMIDINote Index, 7
                    StopMIDINote Index, -2
            End Select
        Case MajorMin9Suspended4th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 1;5;-5;-2
                    StopMIDINote Index, 1
                    StopMIDINote Index, 5
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 2 To 4
                    ' 1;-7;-5;-2
                    StopMIDINote Index, 1
                    StopMIDINote Index, -7
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 5 To 7
                    ' 1;5;7;10
                    StopMIDINote Index, 1
                    StopMIDINote Index, 5
                    StopMIDINote Index, 7
                    StopMIDINote Index, 10
                Case 8, 9
                    ' 1;5;7;-2
                    StopMIDINote Index, 1
                    StopMIDINote Index, 5
                    StopMIDINote Index, 7
                    StopMIDINote Index, -2
            End Select
        Case Major9th
            Select Case Index Mod 12
                Case 0, 1, 10, 11
                    ' 2;4;-5;-2
                    StopMIDINote Index, 2
                    StopMIDINote Index, 4
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 2 To 4
                    ' 2;-8;-5;-2
                    StopMIDINote Index, 2
                    StopMIDINote Index, -8
                    StopMIDINote Index, -5
                    StopMIDINote Index, -2
                Case 5 To 7
                    ' 2;4;7;10
                    StopMIDINote Index, 2
                    StopMIDINote Index, 4
                    StopMIDINote Index, 7
                    StopMIDINote Index, 10
                Case 8, 9
                    ' 2;4;7;-2
                    StopMIDINote Index, 2
                    StopMIDINote Index, 4
                    StopMIDINote Index, 7
                    StopMIDINote Index, -2
            End Select
        Case Octave
            ' 0;12
            StopMIDINote Index, 0
            StopMIDINote Index, 12
        Case Else
            ' Single Note
            StopMIDINote Index, 0
    End Select
End Sub
Private Sub Command1_Click()
    rc = mciSendString("stop midi", 0&, 0, 0)
    rc = mciSendString("close midi", 0&, 0, 0)
    rc = mciSendString("open " & locSong & " type sequencer alias midi", 0&, 0, 0)
    If rc <> 0 Then GoTo errHandler
    Sleep 50
    If YmhDvc - 1 <> curDevice Then
        rc = mciSendString("set midi port " & YmhDvc - 1, 0&, 0, 0)
        If rc <> 0 Then GoTo errHandler
    End If
    rc = mciSendString("set midi time format SMPTE 30", 0&, 0, 0)
    If rc <> 0 Then GoTo errHandler
    rc = mciSendString("play midi notify", 0&, 0, cmdHook.hwnd)
    If rc = 0 Then Timer4.Enabled = True
errExit:
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
    Exit Sub
errHandler:
    rc = mciSendString("stop midi", 0&, 0, 0)
    rc = mciSendString("close midi", 0&, 0, 0)
    midiStatus = ""
    Resume errExit
End Sub
Private Sub Command2_Click()
    rc = mciSendString("stop midi", 0&, 0, 0)
    rc = mciSendString("close midi", 0&, 0, 0)
    Timer4.Enabled = False
    midiStatus = ""
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
End Sub
Private Sub Timer4_Timer()
    rc = mciSendString("status midi position", RetSts, 128, 0)
    midiStatus = RetSts
End Sub
Private Sub Label1_Click()
    VibRate.Value = 64
End Sub
Private Sub Label2_Click()
    VibDepth.Value = 64
End Sub
Private Sub Label3_Click()
    VibDelay.Value = 64
End Sub
Private Sub Label4_Click()
    Attack.Value = 64
End Sub
Private Sub Label5_Click()
    Release.Value = 64
End Sub
Private Sub Label6_Click()
    Bright.Value = 64
End Sub
Private Sub Label7_Click()
    Harmonic.Value = 64
End Sub
Private Sub Label8_Click()
    Velocity.Value = 100
End Sub
Private Sub Label9_Click()
    Modulation.Value = 0
End Sub
Private Sub Label10_Click()
    Pitch.Value = 64
End Sub
Private Sub Label11_Click()
    Reverb.Value = 40
End Sub
Private Sub Label12_Click()
    Chorus.Value = 0
End Sub
Private Sub Label14_Click()
    Transpose.Value = 24
End Sub
Private Sub Label15_Click()
    PortamentoCtrl.Value = 0
End Sub
Private Sub Label16_Click()
    PortamentoTime.Value = 0
End Sub
Private Sub Label17_Click()
    Panpot.Value = 64
End Sub
Private Sub Label18_Click()
    If VLPressure.Enabled Then VLPressure.Value = 64
End Sub
Private Sub Label19_Click()
    VelSenseDepth.Value = 64
End Sub
Private Sub Label20_Click()
    VelSenseOffset.Value = 64
End Sub
Private Sub Label21_Click()
    If VLGrowl.Enabled Then VLGrowl.Value = 64
End Sub
Private Sub Label22_Click()
    If VLBreath.Enabled Then VLBreath.Value = 64
End Sub
Private Sub Label23_Click()
    If VLScream.Enabled Then VLScream.Value = 64
End Sub
Private Sub Label24_Click()
    If VLTonging.Enabled Then VLTonging.Value = 64
End Sub
Private Sub Label25_Click()
    If VLEmbrouchure.Enabled Then VLEmbrouchure.Value = 64
End Sub
Private Sub Label26_Click()
    If VLFilterEqDepth.Enabled Then VLFilterEqDepth.Value = 64
End Sub
Private Sub Label27_Click()
    ModulationAmp.Value = 64
End Sub
Private Sub Label28_Click()
    ModulationPMod.Value = 10
End Sub
Private Sub Label29_Click()
    ModulationFMod.Value = 0
End Sub
Private Sub Label30_Click()
    ModulationAMod.Value = 0
End Sub
Private Sub Label31_Click()
    Decay.Value = 64
End Sub
Private Sub Label32_Click()
    VariationDry.Value = 127
End Sub
Private Sub Label33_Click()
    VariationReverb.Value = 0
End Sub
Private Sub Label34_Click()
    VariationChorus.Value = 0
End Sub
Private Sub Label35_Click()
    PitchInitLev.Value = 64
End Sub
Private Sub Label36_Click()
    PitchAttTime.Value = 64
End Sub
Private Sub Label37_Click()
    PitchRelLev.Value = 64
End Sub
Private Sub Label38_Click()
    PitchRelTime.Value = 64
End Sub
Private Sub Label39_Click()
    VariationTime.Value = 0
End Sub
Private Sub Panpot_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&HE) & Chr(Val("&H" & Hex(Panpot.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    If Panpot.Value = 0 Then
        Label17 = "Panpot = Rnd"
    Else
        Label17 = "Panpot = " & Panpot.Value - 64
    End If
End Sub
Private Sub Modulation_Change()
    SetController 1, Modulation.Value, Channel
    Label9 = "Modulation = " & Modulation.Value
End Sub
Private Sub ModulationAMod_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&H22) & Chr(Val("&H" & Hex(ModulationAMod.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label30 = "Amp Depth = " & ModulationAMod.Value
End Sub
Private Sub ModulationAmp_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&H1F) & Chr(Val("&H" & Hex(ModulationAmp.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label27 = "Amp Ctrl = " & ModulationAmp.Value - 64
End Sub
Private Sub ModulationFMod_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&H21) & Chr(Val("&H" & Hex(ModulationFMod.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label29 = "LPF Depth = " & ModulationFMod.Value
End Sub
Private Sub ModulationPMod_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&H20) & Chr(Val("&H" & Hex(ModulationPMod.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label28 = "Pitch Depth = " & ModulationPMod.Value
End Sub
Private Sub PitchAttTime_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&H6A) & Chr(Val("&H" & Hex(PitchAttTime.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label36 = "Attack Time = " & PitchAttTime.Value - 64
End Sub
Private Sub PitchInitLev_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&H69) & Chr(Val("&H" & Hex(PitchInitLev.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label35 = "Initial Level = " & PitchInitLev.Value - 64
End Sub
Private Sub PitchRelLev_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&H6B) & Chr(Val("&H" & Hex(PitchRelLev.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label37 = "Release Level = " & PitchRelLev.Value - 64
End Sub
Private Sub PitchRelTime_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&H6C) & Chr(Val("&H" & Hex(PitchRelTime.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label38 = "Release Time = " & PitchRelTime.Value - 64
End Sub
Private Sub Portamento_Click()
    If Portamento.Value = 0 Then
        Portamento.Caption = "Portamento Off"
        PortamentoTime.Value = 0
        PortamentoCtrl.Value = 0
        SetController 65, 0, Channel
    Else
        Portamento.Caption = "Portamento On"
        SetController 65, 127, Channel
    End If
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
End Sub
Private Sub PortamentoTime_Change()
    SetController 5, PortamentoTime.Value, Channel
    Label16 = "Porta.Time = " & PortamentoTime.Value
End Sub
Private Sub PortamentoCtrl_Change()
    SetController 84, PortamentoCtrl.Value, Channel
    Label15 = "Porta.Ctrl = " & PortamentoCtrl.Value
End Sub
Private Sub SoftPedal_Click()
    If SoftPedal.Value = 0 Then
        SetController 67, 0, Channel
    Else
        SetController 67, 127, Channel
    End If
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
End Sub
Private Sub Sostenuto_Click()
    If Sostenuto.Value = 0 Then
        SetController 66, 0, Channel
    Else
        SetController 66, 127, Channel
    End If
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
End Sub
Private Sub Sustain_Click()
    If Sustain.Value = 0 Then
        SetController 64, 0, Channel
    Else
        SetController 64, 127, Channel
    End If
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
End Sub
Private Sub Transpose_Change()
    sendmsg = YMH_MSG & Chr(&H0) & Chr(&H0) & Chr(&H6) & Chr(Transpose.Value + 40) & Chr(&HF7)
    SetSysEx sendmsg
    Label14 = "Transpose = " & Transpose.Value - 24
End Sub
Private Sub VariationChorus_Change()
    sendmsg = YMH_MSG & Chr(&H2) & Chr(&H1) & Chr(&H59) & Chr(Val("&H" & Hex(VariationChorus.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label34 = "Send To Chrs = " & VariationChorus.Value
End Sub
Private Sub VariationDry_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&H11) & Chr(Val("&H" & Hex(VariationDry.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label32 = "Dry = " & VariationDry.Value
End Sub
Private Sub VariationReverb_Change()
    sendmsg = YMH_MSG & Chr(&H2) & Chr(&H1) & Chr(&H58) & Chr(Val("&H" & Hex(VariationReverb.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label33 = "Send To Rvrb = " & VariationReverb.Value
End Sub
Private Sub VelSenseDepth_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&HC) & Chr(Val("&H" & Hex(VelSenseDepth.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label19 = "Vel Depth = " & VelSenseDepth.Value - 64
End Sub
Private Sub VelSenseOffset_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&HD) & Chr(Val("&H" & Hex(VelSenseOffset.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label20 = "Vel Offset = " & VelSenseOffset.Value - 64
End Sub
Private Sub VibRate_Change()
    SetController 99, 1, Channel
    SetController 98, 8, Channel
    SetController 6, VibRate.Value, Channel
    SetController 101, 127, Channel
    SetController 100, 127, Channel
    Label1 = "Vibrato Rate = " & VibRate.Value - 64
End Sub
Private Sub VibDepth_Change()
    SetController 99, 1, Channel
    SetController 98, 9, Channel
    SetController 6, VibDepth.Value, Channel
    SetController 101, 127, Channel
    SetController 100, 127, Channel
    Label2 = "Vibrato Depth = " & VibDepth.Value - 64
End Sub
Private Sub VibDelay_Change()
    SetController 99, 1, Channel
    SetController 98, 10, Channel
    SetController 6, VibDelay.Value, Channel
    SetController 101, 127, Channel
    SetController 100, 127, Channel
    Label3 = "Vibrato Delay = " & VibDelay.Value - 64
End Sub
Private Sub Attack_Change()
    SetController 73, Attack.Value, Channel
    Label4 = "Attack = " & Attack.Value - 64
End Sub
Private Sub Decay_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(Val("&H" & Hex(Channel))) & Chr(&H1B) & Chr(Val("&H" & Hex(Decay.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label31 = "Decay = " & Decay.Value - 64
End Sub
Private Sub Release_Change()
    SetController 72, Release.Value, Channel
    Label5 = "Release = " & Release.Value - 64
End Sub
Private Sub Bright_Change()
    SetController 74, Bright.Value, Channel
    Label6 = "Bright = " & Bright.Value - 64
End Sub
Private Sub Harmonic_Change()
    SetController 71, Harmonic.Value, Channel
    Label7 = "Harmonic = " & Harmonic.Value - 64
End Sub
Private Sub VLBreath_Change()
    sendmsg = YMH_MSG & Chr(&H9) & Chr(&H0) & Chr(&HC) & Chr(Val("&H" & Hex(VLBreath.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label22 = "Breath = " & VLBreath.Value - 64
End Sub
Private Sub VLEmbrouchure_Change()
    sendmsg = YMH_MSG & Chr(&H9) & Chr(&H0) & Chr(&H6) & Chr(Val("&H" & Hex(VLEmbrouchure.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label25 = "Embrouchure = " & VLEmbrouchure.Value - 64
End Sub
Private Sub VLFilterEqDepth_Change()
    sendmsg = YMH_MSG & Chr(&H8) & Chr(&H0) & Chr(&H71) & Chr(Val("&H" & Hex(VLFilterEqDepth.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label26 = "VL Eq Depth = " & VLFilterEqDepth.Value - 64
End Sub
Private Sub VLGrowl_Change()
    sendmsg = YMH_MSG & Chr(&H9) & Chr(&H0) & Chr(&HE) & Chr(Val("&H" & Hex(VLGrowl.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label21 = "Growl = " & VLGrowl.Value - 64
End Sub
Private Sub VLPressure_Change()
    sendmsg = YMH_MSG & Chr(&H9) & Chr(&H0) & Chr(&H4) & Chr(Val("&H" & Hex(VLPressure.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label18 = "Pressure = " & VLPressure.Value - 64
End Sub
Private Sub VLScream_Change()
    sendmsg = YMH_MSG & Chr(&H9) & Chr(&H0) & Chr(&HA) & Chr(Val("&H" & Hex(VLScream.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label23 = "Scream = " & VLScream.Value - 64
End Sub
Private Sub VLTonging_Change()
    sendmsg = YMH_MSG & Chr(&H9) & Chr(&H0) & Chr(&H8) & Chr(Val("&H" & Hex(VLTonging.Value))) & Chr(&HF7)
    SetSysEx sendmsg
    Label24 = "Tonging = " & VLTonging.Value - 64
End Sub
Private Sub Velocity_Change()
    SetController 7, Velocity.Value, Channel
    Label8 = "Volume = " & Velocity.Value
End Sub
Private Sub Pitch_Change()
    If Timer1.Enabled And pitchPos <> 64 Then
        Label10 = "Pitch Bend = " & Pitch.Value - 64
        Exit Sub
    End If
    If Pitch.Value = 64 Then
        Timer1.Enabled = False
        Label10 = "Pitch Bend"
    Else
        pitchPos = Pitch.Value
        SetController 101, 0, Channel
        SetController 100, 0, Channel
        SetController 6, 6, Channel
        SetController 101, 127, Channel
        SetController 100, 127, Channel
        Timer1.Enabled = True
    End If
End Sub
Private Sub Timer1_Timer()
    Select Case pitchPos
        Case Is > 64
            pitchPos = pitchPos - 1
        Case Is < 64
            pitchPos = pitchPos + 1
    End Select
    SetPitchBend 0, pitchPos, Channel
    Pitch.Value = pitchPos
    DoEvents
End Sub
Private Sub Reverb_Change()
    SetController 91, Reverb.Value, Channel
    Label11 = "Reverb = " & Reverb.Value
End Sub
Private Sub RevType_Click()
    sendmsg = YMH_MSG & Chr(&H2) & Chr(&H1) & Chr(&H0)
    Select Case RevType.ItemData(RevType.ListIndex)
        Case 0 ' Hall1
            sendmsg = sendmsg + Chr(&H1) & Chr(&H0)
        Case 1 ' Hall2
            sendmsg = sendmsg + Chr(&H1) & Chr(&H1)
        Case 2 ' Room1
            sendmsg = sendmsg + Chr(&H2) & Chr(&H0)
        Case 3 ' Room2
            sendmsg = sendmsg + Chr(&H2) & Chr(&H1)
        Case 4 ' Room3
            sendmsg = sendmsg + Chr(&H2) & Chr(&H2)
        Case 5 ' Stage1
            sendmsg = sendmsg + Chr(&H3) & Chr(&H0)
        Case 6 ' Stage2
            sendmsg = sendmsg + Chr(&H3) & Chr(&H1)
        Case 7 ' Plate
            sendmsg = sendmsg + Chr(&H4) & Chr(&H0)
        Case 8 ' White Room
            sendmsg = sendmsg + Chr(&H10) & Chr(&H0)
        Case 9 ' Tunnel
            sendmsg = sendmsg + Chr(&H11) & Chr(&H0)
        Case 10 ' Basement
            sendmsg = sendmsg + Chr(&H13) & Chr(&H0)
    End Select
    'always end with &HF7
    sendmsg = sendmsg + Chr(&HF7)
    SetSysEx sendmsg
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
End Sub
Private Sub Chorus_Change()
    SetController 93, Chorus.Value, Channel
    Label12 = "Chorus = " & Chorus.Value
End Sub
Private Sub ChoType_Click()
    sendmsg = YMH_MSG & Chr(&H2) & Chr(&H1) & Chr(&H20)
    Select Case ChoType.ItemData(ChoType.ListIndex)
        Case 0 ' Chorus1
            sendmsg = sendmsg + Chr(&H41) & Chr(&H0)
        Case 1 ' Chorus2
            sendmsg = sendmsg + Chr(&H41) & Chr(&H1)
        Case 2 ' Chorus3
            sendmsg = sendmsg + Chr(&H41) & Chr(&H2)
        Case 3 ' Chorus4
            sendmsg = sendmsg + Chr(&H41) & Chr(&H8)
        Case 4 ' Celeste1
            sendmsg = sendmsg + Chr(&H42) & Chr(&H0)
        Case 5 ' Celeste2
            sendmsg = sendmsg + Chr(&H42) & Chr(&H1)
        Case 6 ' Celeste3
            sendmsg = sendmsg + Chr(&H42) & Chr(&H2)
        Case 7 ' Celeste4
            sendmsg = sendmsg + Chr(&H42) & Chr(&H8)
        Case 8 ' Flanger1
            sendmsg = sendmsg + Chr(&H43) & Chr(&H0)
        Case 9 ' Flanger2
            sendmsg = sendmsg + Chr(&H43) & Chr(&H1)
        Case 10 ' Flanger3
            sendmsg = sendmsg + Chr(&H43) & Chr(&H8)
    End Select
    'always end with &HF7
    sendmsg = sendmsg + Chr(&HF7)
    SetSysEx sendmsg
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
End Sub
Private Sub VariationTime_Change()
    SetController 94, VariationTime.Value, Channel
    Label39 = "Level = " & VariationTime.Value
End Sub
Private Sub VarType_Click()
    sendmsg = YMH_MSG & Chr(&H2) & Chr(&H1) & Chr(&H40)
    Select Case VarType.ItemData(VarType.ListIndex)
        Case 0 ' Delay LCR
            sendmsg = sendmsg + Chr(&H5) & Chr(&H0)
        Case 1 ' Delay LR
            sendmsg = sendmsg + Chr(&H6) & Chr(&H0)
        Case 2 ' Echo
            sendmsg = sendmsg + Chr(&H7) & Chr(&H0)
        Case 3 ' Cross Delay
            sendmsg = sendmsg + Chr(&H8) & Chr(&H0)
        Case 4 ' ER1
            sendmsg = sendmsg + Chr(&H9) & Chr(&H0)
        Case 5 ' ER2
            sendmsg = sendmsg + Chr(&H9) & Chr(&H1)
        Case 6 ' Gate Reverb
            sendmsg = sendmsg + Chr(&HA) & Chr(&H0)
        Case 7 ' Reverse Gate
            sendmsg = sendmsg + Chr(&HB) & Chr(&H0)
        Case 8 ' Karaoke1
            sendmsg = sendmsg + Chr(&H14) & Chr(&H0)
        Case 9 ' Karaoke2
            sendmsg = sendmsg + Chr(&H14) & Chr(&H1)
        Case 10 ' Karaoke3
            sendmsg = sendmsg + Chr(&H14) & Chr(&H2)
        Case 11 ' Symphonic
            sendmsg = sendmsg + Chr(&H44) & Chr(&H0)
        Case 12 ' Rotary Speaker
            sendmsg = sendmsg + Chr(&H45) & Chr(&H0)
        Case 13 ' Tremolo
            sendmsg = sendmsg + Chr(&H46) & Chr(&H0)
        Case 14 ' Auto Pan
            sendmsg = sendmsg + Chr(&H47) & Chr(&H0)
        Case 15 ' Phaser1
            sendmsg = sendmsg + Chr(&H48) & Chr(&H0)
        Case 16 ' Phaser2
            sendmsg = sendmsg + Chr(&H48) & Chr(&H8)
        Case 17 ' Distortion
            sendmsg = sendmsg + Chr(&H49) & Chr(&H0)
        Case 18 ' Overdrive
            sendmsg = sendmsg + Chr(&H4A) & Chr(&H0)
        Case 19 ' Amp Simulator
            sendmsg = sendmsg + Chr(&H4B) & Chr(&H0)
        Case 20 ' 3 Band Equalizer
            sendmsg = sendmsg + Chr(&H4C) & Chr(&H0)
        Case 21 ' 2 Band Equalizer
            sendmsg = sendmsg + Chr(&H4D) & Chr(&H0)
        Case 22 ' Auto Wah
            sendmsg = sendmsg + Chr(&H4E) & Chr(&H0)
    End Select
    'always end with &HF7
    sendmsg = sendmsg + Chr(&HF7)
    SetSysEx sendmsg
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
End Sub
Private Sub Variation_Click()
    If Variation.Value = 0 Then
        ' set connection to insertion
        sendmsg = YMH_MSG & Chr(&H2) & Chr(&H1) & Chr(&H5A) & Chr(&H0) & Chr(&HF7)
        SetSysEx sendmsg
        Variation.Caption = "Off"
    Else
        ' set connection to system
        sendmsg = YMH_MSG & Chr(&H2) & Chr(&H1) & Chr(&H5A) & Chr(&H1) & Chr(&HF7)
        SetSysEx sendmsg
        Variation.Caption = "On"
    End If
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
End Sub
Private Sub VoiceOption_Click(Index As Integer)
    If Index = 3 And Channel <> 0 Then
        VoiceOption.Item(0).Value = True
        MsgBox "Sorry, VL Plug-in is only for channel 1."
        Index = 0
    End If
    VoiceOption.Item(Index).Value = True
    Dim sVoice As String, nBank As Integer, nProgram As Integer
    Dim itmX As ListItem
    ListView1.ListItems.Clear
    Select Case Index
        Case 0
            SetController 0, 0, Channel
            Open "Ins.txt" For Input As #1
        Case 1
            SetController 0, 64, Channel
            Open "SFX.txt" For Input As #1
        Case 2
            SetController 0, 127, Channel
            Open "Drum.txt" For Input As #1
        Case 3
            SetController 0, 81, Channel
            Open "VL.txt" For Input As #1
    End Select
    Do While Not EOF(1)
        Input #1, sVoice, nBank, nProgram
        Set itmX = ListView1.ListItems.Add(, , sVoice)
        itmX.SubItems(1) = nBank
        itmX.SubItems(2) = nProgram
    Loop
    Close #1
    Set itmX = Nothing
    ' Set to first item of the list
    SetController 32, ListView1.SelectedItem.SubItems(1), Channel
    SetInstrument ListView1.SelectedItem.SubItems(2), Channel
    If Index = 3 Then
        'VL Pressure assign to controller 21
        SetSysEx YMH_MSG & Chr(&H9) & Chr(&H0) & Chr(&H3) & Chr(&H15) & Chr(&HF7)
        'VL Embrouchure assign to controller 22
        SetSysEx YMH_MSG & Chr(&H9) & Chr(&H0) & Chr(&H5) & Chr(&H16) & Chr(&HF7)
        'VL Tonging assign to controller 23
        SetSysEx YMH_MSG & Chr(&H9) & Chr(&H0) & Chr(&H7) & Chr(&H17) & Chr(&HF7)
        'VL Scream assign to controller 24
        SetSysEx YMH_MSG & Chr(&H9) & Chr(&H0) & Chr(&H9) & Chr(&H18) & Chr(&HF7)
        'VL Breath assign to controller 25
        SetSysEx YMH_MSG & Chr(&H9) & Chr(&H0) & Chr(&HB) & Chr(&H19) & Chr(&HF7)
        'VL Growl assign to controller 26
        SetSysEx YMH_MSG & Chr(&H9) & Chr(&H0) & Chr(&HD) & Chr(&H1A) & Chr(&HF7)
        VLPressure.Enabled = True
        VLEmbrouchure.Enabled = True
        VLTonging.Enabled = True
        VLScream.Enabled = True
        VLBreath.Enabled = True
        VLGrowl.Enabled = True
        VLFilterEqDepth.Enabled = True
        VLPressure.Value = 64
        VLEmbrouchure.Value = 64
        VLTonging.Value = 64
        VLScream.Value = 64
        VLBreath.Value = 64
        VLGrowl.Value = 64
        VLFilterEqDepth.Value = 64
        chord.Text = chord.List(0)
        chord.Enabled = False
    Else
        VLPressure.Enabled = False
        VLEmbrouchure.Enabled = False
        VLTonging.Enabled = False
        VLScream.Enabled = False
        VLBreath.Enabled = False
        VLGrowl.Enabled = False
        VLFilterEqDepth.Enabled = False
        Label18 = "Pressure"
        Label21 = "Growl"
        Label22 = "Breath"
        Label23 = "Scream"
        Label24 = "Tonging"
        Label25 = "Embrouchure"
        Label26 = "VL Eq Depth"
        If Preset.Value = 0 Then chord.Enabled = True
    End If
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
End Sub
Private Sub ListView1_Click()
    SetController 32, ListView1.SelectedItem.SubItems(1), Channel
    SetInstrument ListView1.SelectedItem.SubItems(2), Channel
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
End Sub
Private Sub Preset_Click()
    If Preset.Value = 0 Then
        StopNote lastNote, -1
        Timer2.Enabled = False
        MIDI_Device.Enabled = True
        Base_Note.Enabled = True
        Sel_Channel.Enabled = True
        If VoiceOption.Item(3).Value = False Then
            chord.Enabled = True
            MySetFocus chord.hwnd
        Else
            MySetFocus Form1.hwnd
        End If
    Else
        MIDI_Device.Enabled = False
        chord.Text = chord.List(0)
        chord.Enabled = False
        Base_Note.Enabled = False
        Sel_Channel.Enabled = False
        autoPlay = 0
        Timer2.Enabled = True
        MySetFocus Form1.hwnd
    End If
End Sub
Private Sub Timer2_Timer()
    StopNote lastNote, -1
    Select Case autoPlay
        Case 0
            StartNote 51, -1
            lastNote = 51
            Timer2.Interval = 1400
        Case 1
            StartNote 50, -1
            lastNote = 50
            Timer2.Interval = 350
        Case 2
            StartNote 51, -1
            lastNote = 51
            Timer2.Interval = 350
        Case 3
            StartNote 53, -1
            lastNote = 53
            Timer2.Interval = 700
        Case 4
            StartNote 48, -1
            lastNote = 48
            Timer2.Interval = 1400
        Case 5
            StartNote 46, -1
            lastNote = 46
            Timer2.Interval = 1400
        Case 6
            StartNote 44, -1
            lastNote = 44
            Timer2.Interval = 1400
        Case 7
            StartNote 43, -1
            lastNote = 43
            Timer2.Interval = 350
        Case 8
            StartNote 44, -1
            lastNote = 44
            Timer2.Interval = 350
        Case 9
            StartNote 46, -1
            lastNote = 46
            Timer2.Interval = 700
        Case 10
            StartNote 41, -1
            lastNote = 41
            Timer2.Interval = 2800
        Case 11
            StartNote 43, -1
            lastNote = 43
            Timer2.Interval = 1400
        Case 12
            StartNote 45, -1
            lastNote = 45
            Timer2.Interval = 350
        Case 13
            StartNote 46, -1
            lastNote = 46
            Timer2.Interval = 700
        Case 14
            StartNote 48, -1
            lastNote = 48
            Timer2.Interval = 350
        Case 15
            StartNote 53, -1
            lastNote = 53
            Timer2.Interval = 1400
        Case 16
            StartNote 46, -1
            lastNote = 46
            Timer2.Interval = 1400
        Case 17
            StartNote 51, -1
            lastNote = 51
            Timer2.Interval = 1350
        Case 18
            ' Take a breath
            Timer2.Interval = 50
        Case 19
            StartNote 51, -1
            lastNote = 51
            Timer2.Interval = 350
        Case 20
            StartNote 50, -1
            lastNote = 50
            Timer2.Interval = 700
        Case 21
            StartNote 48, -1
            lastNote = 48
            Timer2.Interval = 350
        Case 22
            StartNote 46, -1
            lastNote = 46
            Timer2.Interval = 2800
        Case 23
            StartNote 51, -1
            lastNote = 51
            Timer2.Interval = 1400
        Case 24
            StartNote 50, -1
            lastNote = 50
            Timer2.Interval = 350
        Case 25
            StartNote 51, -1
            lastNote = 51
            Timer2.Interval = 350
        Case 26
            StartNote 53, -1
            lastNote = 53
            Timer2.Interval = 700
        Case 27
            StartNote 48, -1
            lastNote = 48
            Timer2.Interval = 1400
        Case 28
            StartNote 46, -1
            lastNote = 46
            Timer2.Interval = 1400
        Case 29
            StartNote 44, -1
            lastNote = 44
            Timer2.Interval = 1400
        Case 30
            StartNote 43, -1
            lastNote = 43
            Timer2.Interval = 350
        Case 31
            StartNote 44, -1
            lastNote = 44
            Timer2.Interval = 350
        Case 32
            StartNote 46, -1
            lastNote = 46
            Timer2.Interval = 700
        Case 33
            StartNote 41, -1
            lastNote = 41
            Timer2.Interval = 2800
        Case 34
            StartNote 43, -1
            lastNote = 43
            Timer2.Interval = 1400
        Case 35
            StartNote 45, -1
            lastNote = 45
            Timer2.Interval = 350
        Case 36
            StartNote 46, -1
            lastNote = 46
            Timer2.Interval = 700
        Case 37
            StartNote 48, -1
            lastNote = 48
            Timer2.Interval = 350
        Case 38
            StartNote 53, -1
            lastNote = 53
            Timer2.Interval = 1400
        Case 39
            StartNote 46, -1
            lastNote = 46
            Timer2.Interval = 1400
        Case 40
            StartNote 56, -1
            lastNote = 56
            Timer2.Interval = 1350
        Case 41
            ' Take a breath
            Timer2.Interval = 50
        Case 42
            StartNote 56, -1
            lastNote = 56
            Timer2.Interval = 350
        Case 43
            StartNote 55, -1
            lastNote = 55
            Timer2.Interval = 700
        Case 44
            StartNote 53, -1
            lastNote = 53
            Timer2.Interval = 350
        Case 45
            StartNote 55, -1
            lastNote = 55
            Timer2.Interval = 2800
        Case 46
            StartNote 48, -1
            lastNote = 48
            Timer2.Interval = 1400
        Case 47
            StartNote 50, -1
            lastNote = 50
            Timer2.Interval = 350
        Case 48
            StartNote 51, -1
            lastNote = 51
            Timer2.Interval = 700
        Case 49
            StartNote 53, -1
            lastNote = 53
            Timer2.Interval = 350
        Case 50
            StartNote 46, -1
            lastNote = 46
            Timer2.Interval = 1400
        Case 51
            StartNote 51, -1
            lastNote = 51
            Timer2.Interval = 1350
        Case 52
            ' Take a breath
            Timer2.Interval = 50
        Case 53
            StartNote 51, -1
            lastNote = 51
            Timer2.Interval = 1400
        Case 54
            StartNote 56, -1
            lastNote = 56
            Timer2.Interval = 350
        Case 55
            StartNote 55, -1
            lastNote = 55
            Timer2.Interval = 700
        Case 56
            StartNote 53, -1
            lastNote = 53
            Timer2.Interval = 350
        Case 57
            StartNote 51, -1
            lastNote = 51
            Timer2.Interval = 2100
        Case 58
            autoPlay = -1
    End Select
    autoPlay = autoPlay + 1
End Sub
' <<<<< MENUBAR SECTION >>>>>
Private Sub About_Click()
    MsgBox "XG is a trademark from Yamaha Corporation", , "About Yamaha XG Virtual Piano"
End Sub
' Open the midi device selected in the menu. The menu index equals the
' midi device number + 1.
Private Sub device_Click(Index As Integer)
    Device(curDevice + 1).Checked = False
    Device(Index).Checked = True
    curDevice = Index - 1
    rc = midiOutClose(hmidi)
    rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
    If (rc <> 0) Then
        getMIDIoutErrText "opening MIDI device", rc
        Exit Sub
    End If
    Panic_Click
End Sub
' Base note for our piano
Private Sub Base_Click(Index As Integer)
    Base(baseNote).Checked = False
    baseNote = Index
    Base(Index).Checked = True
End Sub
' Output channel for our piano
Private Sub Chn_Click(Index As Integer)
    If VoiceOption.Item(3).Value = True Then
        If Index <> 0 Then
            MsgBox "Sorry, VL Plug-in is only for channel 1."
            Exit Sub
        End If
    End If
    Chn(Channel).Checked = False
    Channel = Index
    Chn(Index).Checked = True
    Panic_Click
End Sub
' Reset stuck notes and controllers
Private Sub Panic_Click()
    rc = midiOutReset(hmidi)
    If rc <> 0 Then
        getMIDIoutErrText "resetting midi device", rc
        Exit Sub
    End If
    MousePointer = 11
    Sleep 50
    
    'XG System on
    sendmsg = YMH_MSG & Chr(&H0) & Chr(&H0) & Chr(&H7E) & Chr(&H0) & Chr(&HF7)
    SetSysEx sendmsg
    
    'All parameter reset
    'sendmsg = YMH_MSG & Chr(&H0) & Chr(&H0) & Chr(&H7F) & Chr(&H0) & Chr(&HF7)
    'SetSysEx sendmsg
    
    Base_Click (12)
    Pitch.Value = 64
    
    'SysEx message
    RevType.Text = RevType.List(0)
    ChoType.Text = ChoType.List(0)
    Variation.Value = 0
    VarType.Text = VarType.List(0)
    VariationDry.Value = 127
    VariationReverb.Value = 0
    VariationChorus.Value = 0
    PitchAttTime.Value = 64
    PitchInitLev.Value = 64
    PitchRelLev.Value = 64
    PitchRelTime.Value = 64
    ModulationAmp.Value = 64
    ModulationPMod.Value = 10
    ModulationFMod.Value = 0
    ModulationAMod.Value = 0
    Decay.Value = 64
    Transpose.Value = 24
    VelSenseDepth.Value = 64
    VelSenseOffset.Value = 64
    Panpot.Value = 64
    
    'Controller Message
    Modulation.Value = 0
    Velocity.Value = 100
    Reverb.Value = 40
    Chorus.Value = 0
    VariationTime.Value = 0
    Portamento.Value = 0
    Sustain.Value = 0
    Sostenuto.Value = 0
    SoftPedal.Value = 0
    Attack.Value = 64
    Release.Value = 64
    Bright.Value = 64
    Harmonic.Value = 64
    VibRate.Value = 64
    VibDepth.Value = 64
    VibDelay.Value = 64
    If Channel = 9 Then
        VoiceOption_Click (2)
        VoiceOption(0).Enabled = False
        VoiceOption(1).Enabled = False
        VoiceOption(3).Enabled = False
    Else
        VoiceOption(0).Enabled = True
        VoiceOption(1).Enabled = True
        VoiceOption(3).Enabled = True
        VoiceOption.Item(0).Value = True
        VoiceOption_Click (0)
    End If
    MousePointer = 0
    If chord.Enabled Then
        MySetFocus chord.hwnd
    Else
        MySetFocus Form1.hwnd
    End If
End Sub
