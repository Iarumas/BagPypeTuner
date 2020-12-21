VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3E2E5A60-5763-11D5-A29B-F938E4C62136}#10.5#0"; "levelm.ocx"
Object = "{86135EDC-6265-45AA-8A47-6C463280490B}#1.0#0"; "AudioControls2.ocx"
Begin VB.Form Form_Main 
   AutoRedraw      =   -1  'True
   ClientHeight    =   11400
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   17550
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   24
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "PipeTuner.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   760
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1170
   Begin VB.PictureBox picFrameNotes 
      BackColor       =   &H8000000A&
      Height          =   9135
      Left            =   2580
      ScaleHeight     =   605
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   36
      Top             =   360
      Width           =   540
      Begin VB.PictureBox picDrones 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   64
         Top             =   6840
         Width           =   480
      End
      Begin VB.PictureBox picDrones 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   63
         Top             =   7440
         Width           =   480
      End
      Begin VB.PictureBox picDrones 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   62
         Top             =   8040
         Width           =   480
      End
      Begin VB.PictureBox picScaleNote 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   11
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   47
         Top             =   0
         Width           =   480
      End
      Begin VB.PictureBox picScaleNote 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   10
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   46
         Top             =   600
         Width           =   480
      End
      Begin VB.PictureBox picScaleNote 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   9
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   45
         Top             =   1200
         Width           =   480
      End
      Begin VB.PictureBox picScaleNote 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   8
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   44
         Top             =   1800
         Width           =   480
      End
      Begin VB.PictureBox picScaleNote 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   7
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   43
         Top             =   2400
         Width           =   480
      End
      Begin VB.PictureBox picScaleNote 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   6
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   42
         Top             =   3000
         Width           =   480
      End
      Begin VB.PictureBox picScaleNote 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   5
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   41
         Top             =   3600
         Width           =   480
      End
      Begin VB.PictureBox picScaleNote 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   4
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   40
         Top             =   4200
         Width           =   480
      End
      Begin VB.PictureBox picScaleNote 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   3
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   39
         Top             =   4800
         Width           =   480
      End
      Begin VB.PictureBox picScaleNote 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   2
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   38
         Top             =   5400
         Width           =   480
      End
      Begin VB.PictureBox picScaleNote 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BorderStyle     =   0  'Kein
         ForeColor       =   &H80000008&
         Height          =   480
         Index           =   1
         Left            =   0
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   37
         Top             =   6000
         Width           =   480
      End
   End
   Begin VB.HScrollBar hsb_ScrollTime 
      Height          =   180
      LargeChange     =   20
      Left            =   3120
      SmallChange     =   5
      TabIndex        =   30
      Top             =   9480
      Width           =   12000
   End
   Begin AUDIOCONTROLS2Lib.Axis AxisSeconds 
      Height          =   345
      Left            =   3120
      TabIndex        =   29
      Top             =   0
      Width           =   12060
      _Version        =   65536
      _ExtentX        =   21272
      _ExtentY        =   609
      _StockProps     =   13
      BackColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      EndValue        =   15
      TransparentAxis =   -1  'True
      OffsetBegin     =   5
      OffsetEnd       =   62
   End
   Begin VB.Timer tmr_CountDown 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2040
      Top             =   960
   End
   Begin VB.CommandButton cmd_Start_Stop 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   23
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Frame frame_Audio 
      Caption         =   "Audio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   12
      Top             =   4920
      Width           =   2415
      Begin VB.CommandButton cmd_Set_Volume 
         Caption         =   "Set &Volume"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         TabIndex        =   33
         Top             =   1800
         Width           =   1245
      End
      Begin VB.ComboBox cbo_Audio_SampleRate 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "PipeTuner.frx":0442
         Left            =   120
         List            =   "PipeTuner.frx":046E
         Style           =   2  'Dropdown-Liste
         TabIndex        =   15
         Top             =   360
         Width           =   1200
      End
      Begin VB.ComboBox cbo_Audio_Bits 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "PipeTuner.frx":049E
         Left            =   480
         List            =   "PipeTuner.frx":04A6
         Style           =   2  'Dropdown-Liste
         TabIndex        =   14
         Top             =   840
         Width           =   840
      End
      Begin VB.ComboBox cbo_Audio_Channels 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "PipeTuner.frx":04AE
         Left            =   480
         List            =   "PipeTuner.frx":04B8
         Style           =   2  'Dropdown-Liste
         TabIndex        =   13
         Top             =   1320
         Width           =   840
      End
      Begin levelm.LevelMeter obj_Level_Meter 
         Height          =   1935
         Left            =   2040
         Top             =   240
         Width           =   285
         _ExtentX        =   503
         _ExtentY        =   3413
         Level           =   0
         Horizontal      =   0   'False
         Reverse         =   0   'False
         PeakDelay       =   500
         Gradient        =   -1  'True
         Solid           =   -1  'True
      End
      Begin ComctlLib.Slider sld_Audio_Volume 
         Height          =   1785
         Left            =   1440
         TabIndex        =   31
         Top             =   480
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   3149
         _Version        =   327682
         MousePointer    =   7
         OLEDropMode     =   1
         Orientation     =   1
         LargeChange     =   10
         SmallChange     =   5
         Max             =   100
         SelectRange     =   -1  'True
         SelStart        =   50
         TickFrequency   =   10
         Value           =   50
      End
      Begin VB.Label lbl_Audio_Volume 
         Alignment       =   1  'Rechts
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1440
         TabIndex        =   32
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.PictureBox picNote 
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11520
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   6
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Timer tmr_FRAME 
      Interval        =   1000
      Left            =   2040
      Top             =   480
   End
   Begin VB.Frame frame_Ref_Freq 
      Caption         =   "Reference Frequency"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   255
      Width           =   2415
      Begin VB.TextBox txt_Ref_A 
         Alignment       =   2  'Zentriert
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   640
         Locked          =   -1  'True
         TabIndex        =   28
         Text            =   "480.0"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txt_Ref_Bb 
         Alignment       =   2  'Zentriert
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   640
         Locked          =   -1  'True
         MousePointer    =   4  'Symbol
         TabIndex        =   27
         Text            =   "453.1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.PictureBox pic_Invisible 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2160
         ScaleHeight     =   915
         ScaleWidth      =   15
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.CommandButton cmd_Ref_Freq_Up 
         Caption         =   "F8 >>"
         DragIcon        =   "PipeTuner.frx":04C2
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1760
         TabIndex        =   11
         Top             =   1320
         Width           =   600
      End
      Begin VB.CommandButton cmd_Ref_Freq_Up 
         Caption         =   "F7 >"
         DragIcon        =   "PipeTuner.frx":07CC
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   10
         Top             =   1320
         Width           =   560
      End
      Begin VB.CommandButton cmd_Ref_Freq_Dwn 
         Caption         =   "< F6"
         DragIcon        =   "PipeTuner.frx":0AD6
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   640
         TabIndex        =   9
         Top             =   1320
         Width           =   560
      End
      Begin VB.CommandButton cmd_Ref_Freq_Dwn 
         Caption         =   "<< F5"
         DragIcon        =   "PipeTuner.frx":0DE0
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   40
         TabIndex        =   8
         Top             =   1320
         Width           =   600
      End
      Begin VB.CommandButton cmd_Ref_Freq_Dwn 
         Caption         =   "<<< F4"
         DragIcon        =   "PipeTuner.frx":10EA
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   40
         TabIndex        =   4
         Top             =   1680
         Width           =   1160
      End
      Begin VB.CommandButton cmd_Ref_Freq_Up 
         Caption         =   "F9 >>>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   3
         Top             =   1680
         Width           =   1160
      End
      Begin VB.Label lbl_Ref_Bb 
         Alignment       =   2  'Zentriert
         BackStyle       =   0  'Transparent
         Caption         =   "b"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lbl_Ref_A 
         Alignment       =   2  'Zentriert
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lbl_Ref_Bb 
         Alignment       =   2  'Zentriert
         BackColor       =   &H80000018&
         BackStyle       =   0  'Transparent
         Caption         =   "B"
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
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame frame_Start_Stop 
      Caption         =   "Start/Stop"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   16
      Top             =   3000
      Width           =   2415
      Begin VB.TextBox txt_File_Length 
         Alignment       =   1  'Rechts
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "3600.1"
         Top             =   640
         Width           =   615
      End
      Begin VB.TextBox txt_Timer 
         Alignment       =   1  'Rechts
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   25
         Text            =   "3600.00"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox txt_Length 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   20
         Text            =   "3600"
         Top             =   640
         Width           =   495
      End
      Begin VB.TextBox txt_Start 
         Alignment       =   1  'Rechts
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   840
         TabIndex        =   19
         Text            =   "5"
         Top             =   280
         Width           =   495
      End
      Begin VB.Label lbl_Max 
         Caption         =   "max."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   35
         Top             =   360
         Width           =   375
      End
      Begin VB.Label lbl_Length_Sec 
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         Top             =   720
         Width           =   255
      End
      Begin VB.Label lbl_Start_Sec 
         Caption         =   "s"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   360
         Width           =   255
      End
      Begin VB.Label lbl_Length 
         Caption         =   "Length"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lbl_Start 
         Caption         =   "Start in"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.PictureBox picDisplay 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9105
      Left            =   3120
      ScaleHeight     =   603
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   796
      TabIndex        =   24
      Top             =   360
      Width           =   12000
      Begin VB.Label labDrones 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DroneName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   3
         Left            =   0
         TabIndex        =   61
         Top             =   8520
         Width           =   1500
      End
      Begin VB.Label labDrones 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DroneName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   2
         Left            =   0
         TabIndex        =   60
         Top             =   8160
         Width           =   1500
      End
      Begin VB.Label labDrones 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "DroneName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   1
         Left            =   0
         TabIndex        =   59
         Top             =   7800
         Width           =   1500
      End
      Begin VB.Label labScaleNote 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NoteName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   10
         Left            =   0
         TabIndex        =   58
         Top             =   6600
         Width           =   1500
      End
      Begin VB.Label labScaleNote 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NoteName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   9
         Left            =   0
         TabIndex        =   57
         Top             =   6120
         Width           =   1500
      End
      Begin VB.Label labScaleNote 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NoteName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   8
         Left            =   0
         TabIndex        =   56
         Top             =   5640
         Width           =   1500
      End
      Begin VB.Label labScaleNote 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NoteName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   7
         Left            =   0
         TabIndex        =   55
         Top             =   5160
         Width           =   1500
      End
      Begin VB.Label labScaleNote 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NoteName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   6
         Left            =   0
         TabIndex        =   54
         Top             =   4680
         Width           =   1500
      End
      Begin VB.Label labScaleNote 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NoteName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   5
         Left            =   0
         TabIndex        =   53
         Top             =   4200
         Width           =   1500
      End
      Begin VB.Label labScaleNote 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NoteName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   4
         Left            =   0
         TabIndex        =   52
         Top             =   3720
         Width           =   1500
      End
      Begin VB.Label labScaleNote 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NoteName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   3
         Left            =   0
         TabIndex        =   51
         Top             =   3240
         Width           =   1500
      End
      Begin VB.Label labScaleNote 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NoteName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   2
         Left            =   0
         TabIndex        =   50
         Top             =   2760
         Width           =   1500
      End
      Begin VB.Label labScaleNote 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NoteName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   1
         Left            =   0
         TabIndex        =   49
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label labScaleNote 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NoteName"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   480
         Index           =   11
         Left            =   0
         TabIndex        =   48
         Top             =   1800
         Width           =   1500
      End
   End
   Begin VB.Label lab_NoteName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   14880
      TabIndex        =   5
      Top             =   1440
      Width           =   255
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open File"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save File"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save File As"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuTunerMode 
      Caption         =   "&Mode"
      Begin VB.Menu mnuModeBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTunerModeLive 
         Caption         =   "Tuner Mode: &Live"
      End
      Begin VB.Menu mnuTunerModeRecord 
         Caption         =   "Tuner Mode: &Record"
      End
      Begin VB.Menu mnuTunerModeWAV 
         Caption         =   "Tuner Mode &WAV File "
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuModeBar2 
         Caption         =   "-"
      End
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "&Configuration"
      Begin VB.Menu mnuConfigBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRefFreqConstant 
         Caption         =   "Reference Frequency &Constant"
         Checked         =   -1  'True
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuRefFreqTenor 
         Caption         =   "from &Tenor Drone"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuRefFreqBass 
         Caption         =   "from &Bass Drone"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuConfigBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAudioDeviceSelect 
         Caption         =   "Audio &Device Select"
         Begin VB.Menu mnuAudioDeviceBar1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAudioDevice 
            Caption         =   "Device Dummy"
            Index           =   0
         End
         Begin VB.Menu menuAudioDeviceBar2 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuAudioLineSelect 
         Caption         =   "Audio &Input Select"
         Begin VB.Menu mnuAudioLineBar1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAudioLine 
            Caption         =   "Line Dummy"
            Index           =   0
         End
         Begin VB.Menu mnuAudioLineBar2 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuConfigBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChanterScale 
         Caption         =   "Configure Scale Chanter / Drones"
      End
      Begin VB.Menu mnuConfigBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFFT 
         Caption         =   "FFT &Parameter"
         Begin VB.Menu mnuFFTBar1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFFTSampleLength 
            Caption         =   "Sample &Data Length"
            Begin VB.Menu mnuFFTSampleLengthN 
               Caption         =   "    128"
               Index           =   0
            End
            Begin VB.Menu mnuFFTSampleLengthN 
               Caption         =   "    256"
               Index           =   1
            End
            Begin VB.Menu mnuFFTSampleLengthN 
               Caption         =   "    512"
               Index           =   2
            End
            Begin VB.Menu mnuFFTSampleLengthN 
               Caption         =   "  1024"
               Index           =   3
            End
            Begin VB.Menu mnuFFTSampleLengthN 
               Caption         =   "  2048"
               Index           =   4
            End
            Begin VB.Menu mnuFFTSampleLengthN 
               Caption         =   "  4096"
               Index           =   5
            End
            Begin VB.Menu mnuFFTSampleLengthN 
               Caption         =   "  8192"
               Index           =   6
            End
            Begin VB.Menu mnuFFTSampleLengthN 
               Caption         =   "16384"
               Index           =   7
            End
            Begin VB.Menu mnuFFTSampleLengthN 
               Caption         =   "32768"
               Index           =   8
            End
            Begin VB.Menu mnuFFTSampleLengthN 
               Caption         =   "65536"
               Index           =   9
            End
         End
         Begin VB.Menu mnuFFTSampleInterval 
            Caption         =   "Sample &Interval"
            Begin VB.Menu mnuFFTSampleIntervalN 
               Caption         =   "    128"
               Index           =   0
            End
            Begin VB.Menu mnuFFTSampleIntervalN 
               Caption         =   "    256"
               Index           =   1
            End
            Begin VB.Menu mnuFFTSampleIntervalN 
               Caption         =   "    512"
               Index           =   2
            End
            Begin VB.Menu mnuFFTSampleIntervalN 
               Caption         =   "  1024"
               Index           =   3
            End
            Begin VB.Menu mnuFFTSampleIntervalN 
               Caption         =   "  2048"
               Index           =   4
            End
            Begin VB.Menu mnuFFTSampleIntervalN 
               Caption         =   "  4096"
               Index           =   5
            End
            Begin VB.Menu mnuFFTSampleIntervalN 
               Caption         =   "  8192"
               Index           =   6
            End
            Begin VB.Menu mnuFFTSampleIntervalN 
               Caption         =   "16384"
               Index           =   7
            End
            Begin VB.Menu mnuFFTSampleIntervalN 
               Caption         =   "32768"
               Index           =   8
            End
            Begin VB.Menu mnuFFTSampleIntervalN 
               Caption         =   "65536"
               Index           =   9
            End
         End
         Begin VB.Menu mnuFFTBar2 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuConfigBar5 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents DirectSoundRecord As frmDX_Record
Attribute DirectSoundRecord.VB_VarHelpID = -1
Private WithEvents AudioBuffer As cls_AudioBuffer
Attribute AudioBuffer.VB_VarHelpID = -1

Private clsMix  As clsWinMixer
Private mlngAudioDeviceMode As Long
Private mlngAudioDestination As Long
Private mlngAudioDeviceIndex As Long
Private mlngAudioLineIndex As Long

'Private RecordVolume   As cls_RecordMixer
Private AudioLevelOkFIFO As cls_FIFO_int

Private mbytAudioSample() As Byte
Private mintAudioData() As Integer
Private mdblSpectrum() As Double
Private mstrHeader As String

Private mintFFTExponentMin As Integer
Private mintFFTExponentMax As Integer

Private mdblAllFrequencies() As Double

Private TunerMode As TunerType
Private TunerState As TunerStatus



Private Sub Form_Load()

    Dim i As Integer
    Dim intPicWidth As Integer
    Dim intPicHeight As Integer
    Dim SizeX As Integer
    Dim SizeY As Integer
    
    SizeX = Screen.Width / Screen.TwipsPerPixelX
    'SizeX = 800
    SizeY = (Screen.Height - GetTaskbarHeight) / Screen.TwipsPerPixelY
    'SizeY = 600
    
    Me.Move 0, 0, SizeX * Screen.TwipsPerPixelX, SizeY * Screen.TwipsPerPixelY

    intPicWidth = SizeX - 220
    intPicHeight = SizeY - 100
    
    picDisplay.Width = intPicWidth
    picDisplay.Height = intPicHeight
    picFrameNotes.Height = intPicHeight
    
    hsb_ScrollTime.Width = intPicWidth
    hsb_ScrollTime.Top = picDisplay.Top + intPicHeight + 4
    
    AxisSeconds.Width = intPicWidth + 4
    
'    ReSample.test
    
'    ' add minimize/maximize buttons in frame
'    FormAddMinMaxButtons Me.hwnd, True, True
   
    Set DirectSoundRecord = New frmDX_Record
    Set AudioBuffer = New cls_AudioBuffer

    'Set RecordVolume = New cls_RecordMixer
    Set clsMix = New clsWinMixer
    
    Set AudioLevelOkFIFO = New cls_FIFO_int
    
    ' if NUMLOCK was active -> deactivate
    If (GetKeyState(vbKeyNumlock) = 1) Then
        keybd_event VK_NUMLOCK, 1, 0, 0
        keybd_event VK_NUMLOCK, 1, KEYEVENTF_KEYUP, 0
    End If
    
    ' display refresh interval
    gsng_RefreshInterval = 0.2          ' 0.25s
    
    ' show audio devices
'    For i = 0 To RecordVolume.DeviceCount - 1
'        If i <> 0 Then Load mnuAudioDevice(i)
'        mnuAudioDevice(i).Caption = RecordVolume.DeviceName(i)
'    Next

    Call ConfigAudioDeviceMenu
    Call mnuAudioDevice_Click(0)
    
    ' select 1.device
'    mnuAudioDevice_Click (0)
    ' select 1.device or mircophone if present
'    mnuAudioLine_Click (0)
'    For i = 0 To RecordVolume.MixerLineCount - 1
'        If RecordVolume.MixerLineType(i) = MIXERLINE_MICROPHONE Then mnuAudioLine_Click (i)
'    Next i
    

    ' set audio channel options
    cbo_Audio_Channels.Clear
    cbo_Audio_Channels.AddItem "1ch"
    cbo_Audio_Channels.AddItem "2ch"
    
    ' clear audio bits
    cbo_Audio_Bits.Clear
            
    ' set wav input parameters
    WavInFile.SampleRate = 22050
    WavInFile.BitsPerSample = 8
    WavInFile.Channels = 2
    WavInFile.SampleLength = 4096
    WavInFile.SampleInterval = 1024
    WavInFile.SectionTimeStart = 0             ' Start after 0s for WAV
    WavInFile.SectionTimeLength = 40            ' Length 40 for WAV
    
    ' set wav output parameters
    WavOutFile.SampleRate = 44100
    WavOutFile.BitsPerSample = 16
    WavOutFile.Channels = 1
    WavOutFile.SampleLength = 8192
    WavOutFile.SampleInterval = 2048
    WavOutFile.SectionTimeStart = 0             ' Start time 0s for Record
    WavOutFile.SectionTimeLength = 30           ' Length 30 for Record
    
    ' wav file for output (write)
    WavOutFile.ReadWrite = True

    ' notebook1/2 or PC at home
    WavInFile.FileName = "D:\Profiles\ABGRAUST\My Documents\Pipes\FFT\Samples 44kHz\boum mono.wav"
    WavOutFile.FileName = "D:\Profiles\ABGRAUST\My Documents\Pipes\FFT\Samples 44kHz\test.wav"
    If WavInFile.Exists Then GoTo labelFileSelected
        
    WavInFile.FileName = "C:\D-Laufwerk\Pipes\FFT\Samples 44kHz\boum mono.wav"
    WavOutFile.FileName = "C:\D-Laufwerk\Pipes\FFT\Samples 44kHz\test.wav"
    If WavInFile.Exists Then GoTo labelFileSelected

    WavInFile.FileName = "D:\Pipes\FFT\Samples 44kHz\boum mono.wav"
    WavOutFile.FileName = "D:\Pipes\FFT\Samples 44kHz\test.wav"
    If WavInFile.Exists Then GoTo labelFileSelected

labelFileSelected:
    
    ' Select Tuner Mode
    TunerMode.WAV = True
    Set WavFile = WavInFile                 ' get file info for default wavinfile
    
    mnuTunerModeWAV_Click
    
    'common wav parameters
    mintFFTExponentMin = 7              ' 2^7 = 128
    mintFFTExponentMax = 16             ' 2^16 = 65536
    mnuFFTSampleLengthN_Click (6)       ' select 2^(7+6)=2^13 = 8192
    mnuFFTSampleIntervalN_Click (4)     ' select 2^(7+4)=2^11 = 2048
    
    Call WAV_File_Info
    
    'mnuTunerModeLive_Click             ' Live
    'mnuTunerModeRecord_Click           ' Record
    mnuTunerModeWAV_Click              ' WAV
 
    If TunerMode.WAV Then txt_Start.Text = WavFile.SectionTimeStart
    If TunerMode.Record Then txt_Start.Text = TunerState.CountDown
    If TunerMode.Live Then txt_Start.Text = 0
    txt_Length.Text = WavFile.SectionTimeLength
    
    ' level meter
    dB_Meter.SetLevelMeterObject obj_Level_Meter
    dB_Meter.Set_dB_Limit -10                   ' set display limit for meter to -10 dB
    
    gdblReferenceFrequency = 478                ' Reference Frequency
    Call NoteDefinitions.NoteInit
    Call NoteDefinitions.StandardSetting
    Call FrequencyDetection.Set_RefFreq(gdblReferenceFrequency)
    Call FrequencyDetection.Init
    
    Call Settings
    Call Display.SetDefaultValues
    
    ' adjust reference frequency display
    Call Ref_Freq_Update
    
    'picDisplay.Visible = False
    
    cmd_Start_Stop.TabIndex = 1

End Sub

Public Sub Settings()
    
    Dim i As Integer
        
    NoteBuffer.NoteBufferInit (picDisplay.ScaleWidth)
    
    Call SetNoteLabels
    
    Display.SetCentAtHalfScale 20               ' 20 cent at half distance between lines
    Display.SetPicture picDisplay, Form_Main    ' define picture and form for display
    'Display.SetDefaultValues                    ' set display default vlaues
    Display.Draw
    Display.SetAxis AxisSeconds                 ' define axis obj
    
    ' scroll bar set to 0
    hsb_ScrollTime.Value = 0
    hsb_ScrollTime.Max = 0
    
        
    'For i = LBound(gudtNotes) To UBound(gudtNotes)
    '    Debug.Print gudtNotes(i).Name, gudtNotes(i).CentSelected, _
    '    gudtNotes(i).Numerator, gudtNotes(i).Denominator, gudtNotes(i).Ratio, _
    '    gudtNotes(i).AbsoluteCent, gudtNotes(i).RelativeCent
    'Next i
    'Debug.Print
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    DirectSoundRecord.SoundStop
    Unload DirectSoundRecord
    Set DirectSoundRecord = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'Set RecordVolume = Nothing
    Set clsMix = Nothing

End Sub

'Private Sub Form_Resize()
    
    'Me.WindowState = vbMaximized
    
'End Sub

Private Sub mnuTunerModeLive_Click()
    SetTabStrip (1)
    mnuTunerModeLive.Checked = True
    mnuTunerModeRecord.Checked = False
    mnuTunerModeWAV.Checked = False
End Sub
Private Sub mnuTunerModeRecord_Click()
    SetTabStrip (2)
    mnuTunerModeLive.Checked = False
    mnuTunerModeRecord.Checked = True
    mnuTunerModeWAV.Checked = False
End Sub
Private Sub mnuTunerModeWAV_Click()
    SetTabStrip (3)
    mnuTunerModeLive.Checked = False
    mnuTunerModeRecord.Checked = False
    mnuTunerModeWAV.Checked = True
End Sub

Private Sub mnuFFTSampleLengthN_Click(index As Integer)
    
    Dim i As Integer
    
    For i = 0 To mintFFTExponentMax - mintFFTExponentMin
        mnuFFTSampleLengthN(i).Checked = (i = index)
        mnuFFTSampleIntervalN(i).Visible = (i < index)
    Next i
    WavFile.SampleLength = 2 ^ (mintFFTExponentMin + index)
    
End Sub
Private Sub mnuFFTSampleIntervalN_Click(index As Integer)
    
    Dim i As Integer

    For i = 0 To mintFFTExponentMax - mintFFTExponentMin
        mnuFFTSampleIntervalN(i).Checked = (i = index)
        mnuFFTSampleLengthN(i).Visible = (i > index)
    Next i
    WavFile.SampleInterval = 2 ^ (mintFFTExponentMin + index)

End Sub

Private Sub mnuRefFreqConstant_Click()
    mnuRefFreqConstant.Checked = True
    mnuRefFreqTenor.Checked = False
    mnuRefFreqBass.Checked = False
End Sub
Private Sub mnuRefFreqTenor_Click()
    mnuRefFreqConstant.Checked = False
    mnuRefFreqTenor.Checked = True
    mnuRefFreqBass.Checked = False
End Sub
Private Sub mnuRefFreqBass_Click()
    mnuRefFreqConstant.Checked = False
    mnuRefFreqTenor.Checked = False
    mnuRefFreqBass.Checked = True
End Sub

Private Sub ConfigAudioDeviceMenu()

    Dim i       As Long

    mlngAudioDeviceMode = MIXER_OPENBY_WAVEIN_ID
    'mlngAudioDeviceMode = MIXER_OPENBY_WAVEOUT_ID
    'mlngAudioDeviceMode = MIXER_OPENBY_MIDIIN_ID
    'mlngAudioDeviceMode = MIXER_OPENBY_MIDIOUT_ID
    'mlngAudioDeviceMode = MIXER_OPENBY_MIXER_ID

    clsMix.DeviceClose

    For i = 0 To clsMix.DeviceCount(mlngAudioDeviceMode) - 1
        If i <> 0 Then Load mnuAudioDevice(i)
        mnuAudioDevice(i).Caption = clsMix.DeviceName(i, mlngAudioDeviceMode)
    Next

End Sub

Private Sub ConfigAudioRecordMenu()
    
    Dim i   As Long
    Dim j   As Long

    ' we want to select the microphone as the
    ' recording source for the WaveIn
    
    ' go through every destination
    For i = 0 To clsMix.DestinationCount - 1
        ' if the destination's type is WaveIn (recording)...
        If clsMix.DestinationType(i) = MIXERLINE_DST_WAVEIN Then
'        If clsMix.DestinationType(i) = MIXERLINE_DST_SPEAKERS Then
            
            ' unmnute all destinations
            clsMix.DestinationMute(i) = False
            mlngAudioDestination = i
            Call UpdateAudioRecordMenu

            ' ... search for a connected microphone source
            For j = 0 To clsMix.SourceCount(i) - 1
                
                ' deselect all lines
                mnuAudioLine(j).Checked = False
                ' mute all lines
                clsMix.SourceMute(i, j) = False
                
                
                If clsMix.SourceType(i, j) = MIXERLINE_SRC_MICROPHONE Then
                    ' found it / select it / unmute it
                     mnuAudioLine_Click (j)

                    
                    'clsMix.SourceSelected(i, j) = True
                    'clsMix.SourceMute(i, j) = False
                    'mnuAudioLine(j).Checked = True
                    'mlngAudioLineIndex = j
                    ''Exit For
                End If
                
            Next

        End If
    Next

End Sub

Private Sub UpdateAudioRecordMenu()
    
    Dim i As Long
    
    ' make all audio lines visible and deselect them
    For i = mnuAudioLine.LBound To mnuAudioLine.UBound
        mnuAudioLine(i).Visible = True
        mnuAudioLine(i).Checked = False
    Next i
   
    ' add mnu items for audio line in and deselect them
    For i = mnuAudioLine.UBound + 1 To clsMix.SourceCount(mlngAudioDestination) - 1
        Load mnuAudioLine(i)
        mnuAudioLine(i).Visible = True
        mnuAudioLine(i).Checked = False
    Next i
    
    ' make mnu items invisible that are not available
    For i = clsMix.SourceCount(mlngAudioDestination) To mnuAudioLine.UBound
        mnuAudioLine(i).Visible = False
    Next i
    
    ' name mnu items
    For i = 0 To clsMix.SourceCount(mlngAudioDestination) - 1
        mnuAudioLine(i).Caption = clsMix.SourceName(mlngAudioDestination, i)
    Next i


End Sub

Private Sub mnuAudioDevice_Click(index As Integer)
    
    Dim i  As Long
    
    clsMix.DeviceClose
    
    mlngAudioDeviceIndex = index

    If Not clsMix.DeviceOpen(mlngAudioDeviceIndex, mlngAudioDeviceMode) Then
        MsgBox "Couldn't open the device!", vbExclamation
        Exit Sub
    End If
    
    ' select audio device
    For i = mnuAudioDevice.LBound To mnuAudioDevice.UBound
        mnuAudioDevice(i).Checked = (i = mlngAudioDeviceIndex)
    Next i
    
    
    'Debug.Print clsMix.DeviceName(mlngAudioDeviceIndex, mlngAudioDeviceMode)
    'MsgBox (clsMix.DeviceName(mlngAudioDeviceIndex, mlngAudioDeviceMode))
    
    Call ConfigAudioRecordMenu
    
End Sub

Private Sub mnuAudioLine_Click(index As Integer)

    Dim i As Integer
    Dim lngVolume As Long
    
    mlngAudioLineIndex = index

    ' select audio line
    clsMix.SourceSelected(mlngAudioDestination, mlngAudioLineIndex) = True
    
    ' name mnu items and select/unmute one / mute all others
    For i = mnuAudioLine.LBound To mnuAudioLine.UBound
        mnuAudioLine(i).Checked = (i = mlngAudioLineIndex)
        clsMix.SourceMute(mlngAudioDestination, mlngAudioLineIndex) = Not (i = mlngAudioLineIndex)
    Next i
    
    lngVolume = 100 - sld_Audio_Volume.Value
    ' set the volume of the source to 50% for every channel (-1)
    clsMix.SourceVolume(mlngAudioDestination, mlngAudioLineIndex, -1) = lngVolume

    'sld_Audio_Volume.Value = 65535 - RecordVolume.MixerLineVolume
'    RecordVolume.MixerLineVolume = 65535 / 100 * (100 - sld_Audio_Volume.Value)

    'MsgBox (clsMix.SourceName(mlngAudioDestination, mlngAudioLineIndex) & "/" & _
            clsMix.SourceType(mlngAudioDestination, mlngAudioLineIndex))
    
End Sub


Private Sub OmnuAudioDevice_Click(index As Integer)
    
    Dim i As Integer
    
    If Not RecordVolume.SelectDevice(index) Then
        MsgBox "Couldn't select device!", vbExclamation
        Exit Sub
    End If
    
    ' select audio device
    For i = mnuAudioDevice.LBound To mnuAudioDevice.UBound
        mnuAudioDevice(i).Checked = (i = index)
    Next i
    
    ' make all audio lines visibale and deselect them
    For i = mnuAudioLine.LBound To mnuAudioLine.UBound
        mnuAudioLine(i).Visible = True
        mnuAudioLine(i).Checked = False
    Next i
   
    ' add mnu items for audio line in and deselect them
    For i = mnuAudioLine.UBound + 1 To RecordVolume.MixerLineCount - 1
        Load mnuAudioLine(i)
        mnuAudioLine(i).Visible = True
        mnuAudioLine(i).Checked = False
    Next i
    
    ' make mnu items invisible that are not available
    For i = RecordVolume.MixerLineCount To mnuAudioLine.UBound
        mnuAudioLine(i).Visible = False
    Next i
    
    ' name mnu items and select one
    For i = 0 To RecordVolume.MixerLineCount - 1
        mnuAudioLine(i).Caption = RecordVolume.MixerLineName(i)
        mnuAudioLine(i).Checked = (i = index)
    Next i
    
    ' choose 1st item
    mnuAudioLine_Click (0)
    ' choose microphone
    For i = 0 To RecordVolume.MixerLineCount - 1
        If RecordVolume.MixerLineType(i) = MIXERLINE_MICROPHONE Then mnuAudioLine_Click (i)
    Next i
    
End Sub

Private Sub OmnuAudioLine_Click(index As Integer)

    Dim i As Integer

    ' select audio line
    If Not RecordVolume.SelectMixerLine(index) Then
        MsgBox "Couldn't select mixer line!", vbExclamation
        Exit Sub
    End If
    
    'RecordVolume.SelectMixerLine (Index)
    Debug.Print RecordVolume.SelectedMixerLine

    ' MixerLineType can be used to automaticaly find and set
    ' the line you want to record from, e.g. microphone.
    ' MixerLine also accepts a line id as a parameter,
    ' pass -1 and the currently selected line is returned.
    
    'Debug.Print "Line Type: ";
    Select Case RecordVolume.MixerLineType(index)
    'Select Case RecordVolume.SelectedMixerLine
    
        Case MIXERLINE_ANALOG:      'Debug.Print "Analog"
        Case MIXERLINE_AUXILIARY:   'Debug.Print "Auxiliary"
        Case MIXERLINE_COMPACTDISC: 'Debug.Print "Compact Disc"
        Case MIXERLINE_DIGITAL:     'Debug.Print "Digital"
        Case MIXERLINE_LINE:        'Debug.Print "Line-In"
        Case MIXERLINE_MICROPHONE:  'Debug.Print "Microphone"
        Case MIXERLINE_PCSPEAKER:   'Debug.Print "PC Speaker"
        Case MIXERLINE_SYNTHESIZER: 'Debug.Print "Synthesizer"
        Case MIXERLINE_TELEPHONE:   'Debug.Print "Telephone"
        Case MIXERLINE_UNDEFINED:   'Debug.Print "Undefined"
        Case MIXERLINE_WAVEOUT:     'Debug.Print "WaveOut"
        Case Else:                  'Debug.Print "Unknown"
    End Select
    

    'sld_Audio_Volume.Value = 65535 - RecordVolume.MixerLineVolume
    RecordVolume.MixerLineVolume = 65535 - sld_Audio_Volume.Value
    
    For i = mnuAudioLine.LBound To mnuAudioLine.UBound
        mnuAudioLine(i).Checked = (i = index)
    Next i
    
End Sub

Private Sub mnuFileOpen_Click()
    
    Dim strExt  As String
    ' file extension
    strExt = "wav"

    WavFile.FileName = frmDialog.File_Open(strExt)
    mstrHeader = String(24, " ") + "File Source: " + WavFile.FileName
    Call WAV_File_Info
    
    If WavFile.FileName = "" Then
        cmd_Start_Stop.Enabled = False
    Else
        cmd_Start_Stop.Enabled = True
    End If

End Sub

Private Sub mnuFileSaveAs_Click()
    
    Dim strExt  As String
    ' file extension
    strExt = "wav"

    WavFile.FileDestination = frmDialog.File_Save(strExt)
    mstrHeader = String(24, " ") + "File Destination:  " + WavFile.FileDestination

    'save/copy file
    WavFile.Move WavFile.FileDestination

End Sub

Private Sub mnuFileSave_Click()

    If WavFile.FileDestination = "" Then
        mnuFileSaveAs_Click
    Else
        WavFile.Move WavFile.FileDestination
    End If

End Sub

Private Sub mnuChanterScale_Click()

    frmConfig.Show
    
End Sub

Private Sub SetTabStrip(ByVal nTab As Integer)
    
    Dim i As Integer
    
    If TunerMode.WAV Then
        Set WavInFile = WavFile             ' store .wav settings in WavInFile obj
    Else
        Set WavOutFile = WavFile            ' store live/reord settings in WavOutFile obj
    End If
        
    Select Case nTab                        ' set live/record/wav mode
        Case 1:
            TunerMode.Live = True
            TunerMode.Record = False
            TunerMode.WAV = False
            cmd_Start_Stop.Caption = "&Start Live"
        Case 2:
            TunerMode.Live = False
            TunerMode.Record = True
            TunerMode.WAV = False
            cmd_Start_Stop.Caption = "&Start Record"
        Case 3:
            TunerMode.Live = False
            TunerMode.Record = False
            TunerMode.WAV = True
            cmd_Start_Stop.Caption = "&Start WAV"
    End Select
        
    If TunerMode.WAV Then
        Set WavFile = WavInFile             ' get .wav settings from Wavfile obj
        Call WAV_File_Info                  ' get .wav file information
        If WavFile.FileName = "" Then
            cmd_Start_Stop.Enabled = False
        Else
            cmd_Start_Stop.Enabled = True
        End If
    Else
        Set WavFile = WavOutFile            ' get live/record settings from WavFile obj
        cmd_Start_Stop.Enabled = True
    End If
             
    Call Update_Settings                    ' update the audio related settings

    ' Fokus auf das erste Element setzen
    On Local Error Resume Next

    On Local Error GoTo 0
     
End Sub

Private Sub SetNoteLabels()
    
    Dim i  As Integer
    Dim sngScaleFactor As Single
    
    sngScaleFactor = picFrameNotes.ScaleHeight / (UBound(gudtNotes) + UBound(gudtDrones) + 2)
    
    picFrameNotes.BackColor = &HFFFFFF
    For i = 1 To UBound(gudtNotes)
        labScaleNote(i).Caption = gudtNotes(i).Name
        labScaleNote(i).Width = 35
        labScaleNote(i).Height = 20
        labScaleNote(i).FontSize = 8
        labScaleNote(i).ForeColor = &HFFFFFF
        labScaleNote(i).Top = picFrameNotes.ScaleHeight - picScaleNote(1).Height / 2 - (i + 1 + UBound(gudtDrones)) * sngScaleFactor
        labScaleNote(i).Left = picDisplay.ScaleWidth - labScaleNote(1).Width - 16
        Set picScaleNote(i).Picture = gudtNotes(i).Pic
        picScaleNote(i).Top = labScaleNote(i).Top + 3
        picScaleNote(i).Visible = True
    Next i
    For i = UBound(gudtNotes) + 1 To UBound(gudtNoteDefaults)
        Set picScaleNote(i).Picture = Nothing
        picScaleNote(i).Top = 0
        picScaleNote(i).Visible = False
        labScaleNote(i).Caption = ""
        labScaleNote(i).Top = 0
    Next i
    
    For i = 1 To UBound(gudtDrones)
        labDrones(i).Caption = gudtDrones(i).Name
        labDrones(i).Width = 35
        labDrones(i).Height = 20
        labDrones(i).FontSize = 8
        labDrones(i).ForeColor = &HFFFFFF
        labDrones(i).Top = picFrameNotes.ScaleHeight - picDrones(1).Height / 2 - i * sngScaleFactor
        labDrones(i).Left = picDisplay.ScaleWidth - labDrones(1).Width - 16
        Set picDrones(i).Picture = gudtDrones(i).Pic
        picDrones(i).Top = labDrones(i).Top + 3
        picDrones(i).Visible = True
    Next i
    For i = UBound(gudtDrones) + 1 To UBound(gudtDroneDefaults)
        labDrones(i).Caption = ""
        labDrones(i).Top = 0
        Set picDrones(i).Picture = Nothing
        picDrones(i).Top = 0
        picDrones(i).Visible = False
    Next i
    
End Sub
Private Sub Update_Settings()

    Dim strBitsPerSample As String
    Dim strSampleRate As String
    Dim strChannels As String

    ' set header for file name
    If TunerMode.WAV Then mstrHeader = String(24, " ") + "File Source: " + WavFile.FileName
    If TunerMode.Record Then mstrHeader = String(24, " ") + "File Destination:  " + WavFile.FileDestination
    If TunerMode.Live Then mstrHeader = ""
    
    'MsgBox (WavInFile.SectionTimeStart & " / " & WavOutFile.SectionTimeStart)
    'txt_Start_Change
    If TunerMode.WAV Then txt_Start.Text = WavFile.SectionTimeStart
    If TunerMode.Record Then txt_Start.Text = TunerState.CountDown
    If TunerMode.Live Then txt_Start.Text = 0
    txt_Length.Text = WavFile.SectionTimeLength
    'MsgBox (WavInFile.SectionTimeStart & " / " & WavOutFile.SectionTimeStart)
    
    ' set timer text
    If TunerMode.WAV Then txt_Timer.Text = Format(WavFile.SectionTimeStart, "###0.00")
    If TunerMode.Record Then txt_Timer.Text = Format(TunerState.CountDown, "###0.00")
    If TunerMode.Live Then txt_Timer.Text = Format(0, "###0.00")
    
    frame_Start_Stop.Visible = Not TunerMode.Live       ' no Start Stop for live / only record and .wav file
    
    ' no save file buttons for .wav files (must not be overwritten)
    mnuFileSave.Enabled = TunerMode.Record
    mnuFileSaveAs.Enabled = TunerMode.Record
    mnuFileOpen.Enabled = TunerMode.WAV
    mnuFile.Enabled = Not TunerMode.Live
    
    ' select audio device only for live and record (not .wav)
    mnuAudioDeviceSelect.Enabled = Not TunerMode.WAV
    mnuAudioLineSelect.Enabled = Not TunerMode.WAV
    
    sld_Audio_Volume.Visible = Not TunerMode.WAV
    lbl_Audio_Volume.Visible = Not TunerMode.WAV
    cmd_Set_Volume.Visible = Not TunerMode.WAV

    
    txt_Timer.Visible = Not TunerMode.Live              ' timer not visible for live
    hsb_ScrollTime.Visible = Not TunerMode.Live
    
    If TunerMode.WAV Then
        
        cbo_Audio_SampleRate.BackColor = &H80000018
        cbo_Audio_Bits.BackColor = &H80000018
        cbo_Audio_Channels.BackColor = &H80000018

        cbo_Audio_SampleRate.Clear
        cbo_Audio_SampleRate.AddItem "   4000 Hz"                    ' 4kHz supported in .wav file
        cbo_Audio_SampleRate.AddItem "   8000 Hz"                    ' 8kHz supported in .wav file
        cbo_Audio_SampleRate.AddItem "  11025 Hz"                    ' 11.025kHz supported in .wav file
        cbo_Audio_SampleRate.AddItem "  16000 Hz"                    ' 16kHz supported in .wav file
        cbo_Audio_SampleRate.AddItem "  22050 Hz"                    ' 22.05kHz supported in.wav file
        cbo_Audio_SampleRate.AddItem "  24000 Hz"                    ' 24kHz supported in .wav file
        cbo_Audio_SampleRate.AddItem "  32000 Hz"                    ' 32kHz supported in .wav file
        cbo_Audio_SampleRate.AddItem "  44100 Hz"                    ' 44.1kHz supported in .wav file
        cbo_Audio_SampleRate.AddItem "  48000 Hz"                    ' 48kHz supported in .wav file
        cbo_Audio_SampleRate.AddItem "  88200 Hz"                    ' 88.2kHz supported in .wav file
        cbo_Audio_SampleRate.AddItem "  96000 Hz"                    ' 96kHz supported in .wav file
        cbo_Audio_SampleRate.AddItem " 192000 Hz"                    ' 192kHz supported in .wav file

        cbo_Audio_Bits.Clear
        cbo_Audio_Bits.AddItem "   8b"                        ' 8 bit supported in .wav file
        cbo_Audio_Bits.AddItem "  12b"                       ' 12 bit supported in .wav file
        cbo_Audio_Bits.AddItem "  16b"                       ' 16 bit supported in .wav file
        cbo_Audio_Bits.AddItem "  20b"                       ' 20 bit supported in .wav file
        cbo_Audio_Bits.AddItem "  24b"                       ' 24 bit supported in .wav file
        cbo_Audio_Bits.AddItem "  32b"                       ' 32 bit supported in .wav file
        cbo_Audio_Bits.AddItem "  32f"                      ' 32 bit float supported in .wav file
        
        cbo_Audio_Channels.Clear
        cbo_Audio_Channels.AddItem "  1ch"                    ' mono supported in .wav file
        cbo_Audio_Channels.AddItem "  2ch"                    ' stereo supported in .wav file
        cbo_Audio_Channels.AddItem "  4ch"                    ' 4 channel supported in .wav file
        cbo_Audio_Channels.AddItem "  6ch"                    ' 5.1 channel supported in .wav file
        cbo_Audio_Channels.AddItem "  8ch"                    ' 7.1 channel supported in .wav file
        
        cbo_Audio_SampleRate.Enabled = False            ' read only in .wav file
        cbo_Audio_Bits.Enabled = False                  ' read only in .wav file
        cbo_Audio_Channels.Enabled = False              ' read only in .wav file

    Else
        
        cbo_Audio_SampleRate.BackColor = &H80000014
        cbo_Audio_Bits.BackColor = &H80000014
        cbo_Audio_Channels.BackColor = &H80000014

        cbo_Audio_SampleRate.Clear
        cbo_Audio_SampleRate.AddItem "   4000 Hz"                    ' 4kHz supported for live/record
        cbo_Audio_SampleRate.AddItem "   8000 Hz"                    ' 8kHz supported for live/record"
        cbo_Audio_SampleRate.AddItem "  11025 Hz"                    ' 11.025kHz supported for live/record
        cbo_Audio_SampleRate.AddItem "  16000 Hz"                    ' 16kHz supported for live/record
        cbo_Audio_SampleRate.AddItem "  22050 Hz"                    ' 22.05kHz supported for live/record
        cbo_Audio_SampleRate.AddItem "  24000 Hz"                    ' 24kHz supported for live/record
        cbo_Audio_SampleRate.AddItem "  32000 Hz"                    ' 32kHz supported for live/record
        cbo_Audio_SampleRate.AddItem "  44100 Hz"                    ' 44.1kHz supported for live/record
        cbo_Audio_SampleRate.AddItem "  48000 Hz"                    ' 48kHz supported for live/record
        cbo_Audio_SampleRate.AddItem "  88200 Hz"                    ' 88.2kHz supported for live/record
        cbo_Audio_SampleRate.AddItem "  96000 Hz"                    ' 96kHz supported for live/record
        'cbo_Audio_SampleRate.AddItem " 192000Hz"                    ' 192kHz supported for live/record
        
        cbo_Audio_Bits.Clear
        cbo_Audio_Bits.AddItem "   8b"                        ' 8 bit supported for live/record
        cbo_Audio_Bits.AddItem "  16b"                        ' 16 bit supported for live/record
        cbo_Audio_Bits.AddItem "  24b"                        ' 24 bit supported for live/record
        cbo_Audio_Bits.AddItem "  32b"                        ' 32 bit supported for live/record
        'cbo_Audio_Bits.AddItem "  32f"                        ' 32 bit float supported for live/record
        
        cbo_Audio_Channels.Clear
        cbo_Audio_Channels.AddItem "  1ch"                    ' mono supported for live/record
        cbo_Audio_Channels.AddItem "  2ch"                    ' stereo supported for live/record
        
        cbo_Audio_SampleRate.Enabled = True             ' combo box enabled for live/record
        cbo_Audio_Bits.Enabled = True                   ' combo box enabled for live/record
        cbo_Audio_Channels.Enabled = True               ' combo box enabled for live/record
        
    End If
    
    strBitsPerSample = CStr(WavFile.BitsPerSample)
    strSampleRate = CStr(WavFile.SampleRate)
    strChannels = CStr(WavFile.Channels)
    
    ' 4 digits bits per sample
    Do While Len(strBitsPerSample) < 4
        strBitsPerSample = " " & strBitsPerSample
    Loop
    
    ' 7 digits for sample rate
    Do While Len(strSampleRate) < 7
        strSampleRate = " " & strSampleRate
    Loop
    
    ' 3 digits for channels
    Do While Len(strChannels) < 3
        strChannels = " " & strChannels
    Loop
    
    ' add "b" or "f" for bit or float as unit
    Select Case WavFile.FormatTag
        Case 1: strBitsPerSample = strBitsPerSample & "b"
        Case 3: strBitsPerSample = strBitsPerSample & "f"
    End Select
    
    ' add units
    strChannels = strChannels & "ch"
    strSampleRate = strSampleRate & " Hz"
    
    'check if the audio parameters match with the options in the combo boxes
    Call CheckComboBoxForValue(cbo_Audio_SampleRate, strSampleRate)
    Call CheckComboBoxForValue(cbo_Audio_Bits, strBitsPerSample)
    Call CheckComboBoxForValue(cbo_Audio_Channels, strChannels)
    
End Sub

Private Sub CheckComboBoxForValue(ByRef myCombobox As ComboBox, ByVal Value As String)
    'look after Value in myComboBox
    
    Dim i As Long
    
    For i = 0 To myCombobox.ListCount - 1
        If myCombobox.List(i) = Value Then
            myCombobox.ListIndex = i
            Exit Sub
        End If
    Next i
    
    ' if no selectable in combo box
    MsgBox (myCombobox.Name & " / " & "Value Not Supported")
    
End Sub

Sub form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' hot keys
    
    Dim wait_key_time As Single             ' min. allowed time delay between to key events
    Dim new_key_time As Single              ' time for current key event
    Dim delta_key_time As Single            ' time between latest and previous key event
    Dim last_key_time As Single             ' time for previsous key event (that was not ignored)
    Dim slide_move As Integer

    wait_key_time = 0.2
    new_key_time = Timer                            ' set time for current key event
    delta_key_time = new_key_time - last_key_time

    If delta_key_time > wait_key_time Then          ' if delay is shorter than wait time the key events are ignored
    
        last_key_time = Timer                       ' set time for latest key event
     
        Select Case KeyCode
            Case vbKeyEscape: End
            Case vbKeySpace: cmd_Start_Stop_Click
            Case vbKeyF1:
                ' dropdown box only enabled if not in wav mode and buttons are enabled
                If (TunerMode.Live Or TunerMode.Record) Then
                    cbo_Audio_SampleRate.SetFocus
                    Call DropDown(cbo_Audio_SampleRate)
                End If
            Case vbKeyF2:
                ' dropdown box only enabled if not in wav mode and buttons are enabled
                If (TunerMode.Live Or TunerMode.Record) Then
                    cbo_Audio_Bits.SetFocus
                    Call DropDown(cbo_Audio_Bits)
                End If
            Case vbKeyF3:
                ' dropdown box only enabled if not in wav mode and buttons are enabled
                If (TunerMode.Live Or TunerMode.Record) Then
                    cbo_Audio_Channels.SetFocus
                    Call DropDown(cbo_Audio_Channels)
                End If
            Case vbKeyF10: cmd_Start_Stop_Click
            Case vbKeyS: cmd_Start_Stop_Click
            Case vbKeyW: mnuTunerModeWAV_Click      ' tuner mode wav
            Case vbKeyR: mnuTunerModeRecord_Click   ' tuner mode record
            Case vbKeyL: mnuTunerModeLive_Click     ' tuner mode live
            'Case vbKeyI: mnuAudioLineSelect
            'Case vbKeyD: mnuAudioDeviceSelect
            Case vbKeyC: mnuRefFreqConstant_Click
            Case vbKeyT: mnuRefFreqTenor_Click
            Case vbKeyB: mnuRefFreqBass_Click
            Case vbKeyP: 'cmd_Configuration_Scale_Click
            Case vbKeyA: 'cmd_Analyser_Click
            Case vbKeyV: cmd_Set_Volume_Click
            Case vbKeyF:
                    If TunerMode.WAV Then mnuFileOpen_Click
                    If TunerMode.Record Then mnuFileSave_Click
            
            ' adjust reference frequency with hotkeys
            Case vbKeyF4: Call Ref_Freq_Up(2, 0)        ' - 10Hz
            Case vbKeyF5: Call Ref_Freq_Up(1, 0)        ' - 1Hz
            Case vbKeyF6: Call Ref_Freq_Up(0, 0)        ' - 0.1Hz
            Case vbKeyF7: Call Ref_Freq_Up(0, 1)        ' + 0.1Hz
            Case vbKeyF8: Call Ref_Freq_Up(1, 1)        ' + 1Hz
            Case vbKeyF9: Call Ref_Freq_Up(2, 1)        ' + 10Hz
            
            Case Shift = 2 And vbKeyLeft:   Call Ref_Freq_Up(2, 0)              ' - 10Hz
            Case Shift = 1 And vbKeyLeft:   Call Ref_Freq_Up(1, 0)              ' - 1Hz
            'Case Shift = 0 And vbKeyLeft:   Call Ref_Freq_Up(0, 0)              ' - 0.1Hz
            'Case Shift = 0 And vbKeyRight:  Call Ref_Freq_Up(0, 1)              ' + 0.1Hz
            Case Shift = 1 And vbKeyRight:  Call Ref_Freq_Up(1, 1)              ' + 1Hz
            Case Shift = 2 And vbKeyRight:  Call Ref_Freq_Up(2, 1)              ' + 10Hz
            
            Case Shift = 2 And vbKeySubtract:   Call Ref_Freq_Up(2, 0)           ' - 10Hz
            Case Shift = 1 And vbKeySubtract:   Call Ref_Freq_Up(1, 0)           ' - 1Hz
            Case Shift = 0 And vbKeySubtract:   Call Ref_Freq_Up(0, 0)           ' - 0.1Hz
            Case Shift = 0 And vbKeyAdd:        Call Ref_Freq_Up(0, 1)           ' + 0.1Hz
            Case Shift = 1 And vbKeyAdd:        Call Ref_Freq_Up(1, 1)           ' + 1Hz
            Case Shift = 2 And vbKeyAdd:        Call Ref_Freq_Up(2, 1)           ' + 10Hz

        End Select
   
    End If

    DoEvents

End Sub


Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' show spectrum of sample when slected via middle mouse button
    
    ' only works when tuner is not runnig or tuner mode is record/wav
        If TunerState.Running Then Exit Sub
        If TunerMode.Live Then Exit Sub
        
        Dim sngZoom_X As Single
        Dim index As Long

        If Button = vbMiddleButton And Shift = 0 Then
            ' watch zoom factor of display pic
            sngZoom_X = Display.Zoom
            index = CLng(X / sngZoom_X)
            'MsgBox (Index)
            Form_Monitor.glngIndex = index
            Form_Monitor.ShowSpectrum
            Form_Monitor.Show
        End If
        
End Sub

Sub txt_Ref_Bb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'adjust reference frequency with shift key and left/right mousebuttons
    If Button = vbLeftButton Then Call Ref_Freq_Up(Shift, False)      ' 0: 0.1Hz / 1: 1 Hz / 2: 10Hz
    If Button = vbRightButton Then
        ' trick for right click / would do somthing else otherwise
        Call MouseNoRightClick(pic_Invisible)
        Call Ref_Freq_Up(Shift, True)
    End If
End Sub
Sub txt_Ref_A_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'adjust reference frequency with shift key and left/right mousebuttons
    If Button = vbLeftButton Then Call Ref_Freq_Up(Shift, False)      ' 0: 0.1Hz / 1: 1 Hz / 2: 10Hz
    If Button = vbRightButton Then
        ' trick for right click / would do somthing else otherwise
        Call MouseNoRightClick(pic_Invisible)
        Call Ref_Freq_Up(Shift, True)
    End If
End Sub

Private Sub cmd_Ref_Freq_Up_Click(index As Integer)
    ' increase reference frequency
    Call Ref_Freq_Up(index, 1)
End Sub
Private Sub cmd_Ref_Freq_Dwn_Click(index As Integer)
    ' decrease reference frequency
    Call Ref_Freq_Up(index, 0)
End Sub

Private Sub Ref_Freq_Up(index As Integer, bln_Up As Boolean)
    'adjust reference frequency (index = 0: 0.1Hz / 1: 1Hz / 2: 10Hz)
    'increase if bln_Up else decrease
    If bln_Up Then gdblReferenceFrequency = gdblReferenceFrequency + 10 ^ (index - 1)
    If Not bln_Up Then gdblReferenceFrequency = gdblReferenceFrequency - 10 ^ (index - 1)
    Call Ref_Freq_Update

End Sub

Private Sub Ref_Freq_Update()
    'gdblReferenceFrequency is global variable

    Dim dblRef_A As Double
    Dim dblRef_Bb   As Double
    Dim dblSemiToneFactor As Double
    
    Dim dblFreqs() As Double

    ' factor for one semitone in chromatic scale
    dblSemiToneFactor = 2 ^ (1 / 12)
    
    ' update frequency for A and Bb
    dblRef_A = gdblReferenceFrequency
    dblRef_Bb = dblRef_A / dblSemiToneFactor
    
    txt_Ref_A.Text = Format(dblRef_A, "000.0")
    txt_Ref_Bb.Text = Format(dblRef_Bb, "000.0")
    
    ' update cent buffer
    NoteBuffer.UpdateCent gdblBufferFrequency.FIFO_Buffer, gdblBufferFrequency.FIFO_Position

    ' update display
    Display.Draw
   
End Sub

Private Sub opt_Ref_Source_Const_Click()
    DoEvents
End Sub

Private Sub opt_Ref_Source_Tenor_Click()
    DoEvents
End Sub

Private Sub opt_Ref_Source_Bass_Click()
    DoEvents
End Sub

Private Sub hsb_ScrollTime_Change()
    ' scroll display window
    hsb_ScrollTime_Scroll
    
End Sub
Private Sub hsb_ScrollTime_Scroll()
    ' scroll display window
    
    Dim i As Long
    
    If TunerState.Running Then Exit Sub
    'If gdblBufferFrequency.FIFO_StartPosition <= 0 Then Exit Sub

    ' frequencies array (buffersize x (chanter plus drones))
    Dim dblFrequencies() As Double
    ReDim dblFrequencies(0 To gdblBufferFrequency.FIFO_Elements - 1, 0 To UBound(gudtDrones))
    
    ' copy frequencies from the array that contains all frequencies
    For i = 0 To UBound(gudtDrones)
        CopyMemory dblFrequencies(0, i), mdblAllFrequencies(hsb_ScrollTime.Value, i), Len(dblFrequencies(0, 0)) * gintBufferNote.FIFO_Elements
    Next i
    
    ' set buffers
    gdblBufferFrequency.FIFO_Clear
    gdblBufferFrequency.Set_FIFO_Buffer dblFrequencies
    gdblBufferFrequency.FIFO_StartPosition = hsb_ScrollTime.Value
    ' update cent buffer
    NoteBuffer.UpdateCent gdblBufferFrequency.FIFO_Buffer, gdblBufferFrequency.FIFO_Elements
    ' update display
    Display.Draw
    ' update axis
    Display.SetAxis AxisSeconds
   
End Sub

Private Sub cbo_Audio_SampleRate_Click()
    ' set audio sample rate
    
    Dim strText As String
    Dim strFFTSamples As String
    Dim strFFTTime As String
    Dim i As Integer
        
    strText = cbo_Audio_SampleRate.List(cbo_Audio_SampleRate.ListIndex)
    WavFile.SampleRate = CLng(Left$(strText, Len(strText) - 2))

    ' update FFT interval and length with new sample rate
    'txt_Time_FFT_Interval.Text = Format(WavFile.SampleInterval / WavFile.SampleRate * 1000, "##0.0")
    'txt_Time_FFT_Length.Text = Format(WavFile.SampleLength / WavFile.SampleRate * 1000, "##0.0")
    
    For i = mintFFTExponentMin To mintFFTExponentMax
        strFFTSamples = CStr(2 ^ i)
        strFFTTime = CStr(Format(2 ^ i / WavFile.SampleRate * 1000, "#####"))
        strFFTSamples = String(2 * (5 - Len(strFFTSamples)), " ") + strFFTSamples
        strFFTTime = String(2 * (5 - Len(strFFTTime)), " ") + strFFTTime
        mnuFFTSampleLengthN(i - mintFFTExponentMin).Caption = strFFTSamples & "     = " & strFFTTime & " ms"
        mnuFFTSampleIntervalN(i - mintFFTExponentMin).Caption = strFFTSamples & "     = " & strFFTTime & " ms"
    Next i
    
End Sub

Private Sub cbo_Audio_Bits_Click()
    ' set bits per sample
    
    Dim strText As String
    
    strText = cbo_Audio_Bits.List(cbo_Audio_Bits.ListIndex)
    WavFile.BitsPerSample = CInt(Left$(strText, Len(strText) - 1))
    
    Select Case Right$(strText, 1)
        Case "b": WavFile.FormatTag = 1     ' for PCM
        Case "f": WavFile.FormatTag = 3     ' for IEEE float format
    End Select
    
End Sub

Private Sub cbo_Audio_Channels_Click()
    ' set number of channels
    
    Dim strText As String
    
    strText = cbo_Audio_Channels.List(cbo_Audio_Channels.ListIndex)
    WavFile.Channels = CInt(Left$(strText, Len(strText) - 2))
    
End Sub

Private Sub WAV_File_Info()
                            
    If WavFile.Exists = False Then
        'MsgBox ("File does not exist !!!")
        Exit Sub
    End If
                
    WavFile.OpenFile
    WavFile.GetWavFileInfo
    WavFile.CloseFile
    txt_File_Length.Text = Format(WavFile.DataTime, "###0.0")
               
    Call Update_Settings
 
End Sub

Private Sub Enable_Buttons(blnEnable As Boolean)
    
    'enable/disable all the cmd buttons, frames, comboboxes, texts, ...
        
    txt_Start.Enabled = blnEnable
    txt_Length.Enabled = blnEnable

    mnuAudioDeviceSelect = blnEnable
    mnuAudioLineSelect = blnEnable

    'frame_Audio.Enabled = blnEnable
    cbo_Audio_SampleRate.Enabled = blnEnable And Not TunerMode.WAV
    cbo_Audio_Bits.Enabled = blnEnable And Not TunerMode.WAV
    cbo_Audio_Channels.Enabled = blnEnable And Not TunerMode.WAV
    
    mnuFile.Enabled = blnEnable
    mnuFileOpen.Enabled = TunerMode.WAV
    mnuFileSave.Enabled = TunerMode.Record
    mnuFileSaveAs.Enabled = TunerMode.Record


End Sub

Sub tmr_Frame_timer()
    'show time in frame

    Form_Main.Caption = Time$ + mstrHeader
    
End Sub

Sub tmr_CountDown_timer()

    txt_Timer.ForeColor = RGB(255, 0, 0)                ' color: red
    TunerState.CountDown = TunerState.CountDown - 1     ' count down
       
    ' stop timer when count down reaches 0
    If TunerState.CountDown <= 0 Then
        TunerState.CountDown = 0
        tmr_CountDown.Enabled = False
        txt_Timer.ForeColor = RGB(0, 0, 0)              ' color black
        Call cmd_Start_Stop_Click                       ' press Start/Stop
    End If
        
    txt_Timer.Text = Format(TunerState.CountDown, "###0.00")

End Sub

Private Sub txt_Start_Change()

    If txt_Start.Text = "" Then Exit Sub
    
    ' set start value
    If TunerMode.Live Then WavFile.SectionTimeStart = 0
    If TunerMode.WAV Then WavFile.SectionTimeStart = CSng(txt_Start.Text)
    If WavFile.SectionTimeStart < 0 Then WavFile.SectionTimeStart = 0
    If TunerMode.Record Then TunerState.CountDown = CLng(txt_Start.Text)
    If TunerState.CountDown < 0 Then TunerState.CountDown = 0
    
   
    'when in TuneMode = Record then CountDown ( in red )
    If TunerMode.Record And TunerState.CountDown > 0 Then           ' or if Index = 0  / this is TunerMode.Record
        txt_Timer.ForeColor = RGB(255, 0, 0)
        'txt_Timer.Text = Format(TunerState.CountDown, "###0")
    Else
        'TunerState.CountDown = 0
        txt_Timer.ForeColor = RGB(0, 0, 0)
        'txt_Timer.Text = Format(TunerState.CountDown, "###0.00")
    End If
    
    If TunerMode.Live Then txt_Timer.Text = Format(0, "###0.00")
    If TunerMode.Record Then txt_Timer.Text = Format(TunerState.CountDown, "###0.00")
    If TunerMode.WAV Then txt_Timer.Text = Format(WavFile.SectionTimeStart, "###0.00")
        
End Sub

Private Sub txt_Length_Change()
    
    'Set start value
    If txt_Length.Text <> "" Then
        If CSng(txt_Length.Text) > 0 Then WavFile.SectionTimeLength = CSng(txt_Length.Text)
    End If
    
End Sub

Private Sub cmd_Start_Stop_Click()

    Dim strLeft As String
    strLeft = Left$(cmd_Start_Stop.Caption, 5)
    
    'Select Case cmd_Start_Stop.Caption
    Select Case strLeft
    
        Case "&Star"
                
            If TunerMode.Record And TunerState.CountDown >= 1 Then
                cmd_Start_Stop.Caption = "CountDown"
                TunerState.Running = False
                tmr_CountDown.Enabled = True
            Else
                cmd_Start_Stop.Caption = "&Stop"
                TunerState.Running = True
            End If
            
            Call AudioInStart
            
        Case "&Stop"
            
            BlockInput True
            cmd_Start_Stop.Enabled = False
            Call AudioInStop
            If TunerMode.Live Then cmd_Start_Stop.Caption = "&Start Live"
            If TunerMode.Record Then cmd_Start_Stop.Caption = "&Start Record"
            If TunerMode.WAV Then cmd_Start_Stop.Caption = "&Start WAV"
            cmd_Start_Stop.Enabled = True
            BlockInput False
            Call txt_Start_Change
            
            
        Case "Count"
            
            If TunerState.CountDown = 0 Then
                cmd_Start_Stop.Caption = "&Stop"
                tmr_CountDown.Enabled = False
                TunerState.Running = True
                TunerState.Stop = False
            Else
                If TunerMode.Live Then cmd_Start_Stop.Caption = "&Start Live"
                If TunerMode.Record Then cmd_Start_Stop.Caption = "&Start Record"
                If TunerMode.WAV Then cmd_Start_Stop.Caption = "&Start WAV"
                tmr_CountDown.Enabled = False
                Call AudioInStop
            End If
    
    End Select
    
End Sub

Private Sub sld_Audio_Volume_Click()

    sld_Audio_Volume_Scroll
    
End Sub

Private Sub sld_Audio_Volume_Scroll()

    Dim i As Long
    Dim j As Long
        
    'RecordVolume.MixerLineVolume = 65535 / 100 * (100 - sld_Audio_Volume.Value)
    For i = 0 To clsMix.DestinationCount - 1
        clsMix.DestinationVolume(i, -1) = 100 - sld_Audio_Volume.Value
        For j = 0 To clsMix.SourceCount(i) - 1
            clsMix.SourceVolume(i, j, -1) = 100 - sld_Audio_Volume.Value
            Debug.Print clsMix.DestinationType(i), clsMix.DestinationName(i), clsMix.SourceName(i, j)
        Next j
    Next i
    
    lbl_Audio_Volume.Caption = 100 - Fix(sld_Audio_Volume.Value / sld_Audio_Volume.Max * 100) & "%"
    
End Sub

Private Sub cmd_Set_Volume_Click()
    
    Dim Vol_Ratio As Single     ' current value (0-1)
    Dim new_Vol As Single       ' new value (0-1)
    
    ' adjust/set volume to 50% when tuner is running or counting down
    If TunerState.Running Or cmd_Start_Stop.Caption = "CountDown" Then
        Vol_Ratio = 10 ^ (dB_Meter.dB_Value / 10)                       ' current ratio (Vol/Max): 0-1
        new_Vol = 0.5 / Vol_Ratio * (65535 - sld_Audio_Volume.Value)    ' new silder value (0-65535) to set volume to 50%
        sld_Audio_Volume.Value = IIf(65535 - new_Vol < sld_Audio_Volume.Min, sld_Audio_Volume.Min, new_Vol)
        Call sld_Audio_Volume_Click
    End If
    
End Sub

Private Sub DirectSoundRecord_GotWaveData(bytBuffer() As Byte)
'Private Sub AudioTest(bytBuffer() As Byte)
    
    If WavFile.DataTime + WavFile.ReadTime >= WavFile.SectionTimeStop Then
        Call cmd_Start_Stop_Click
        Exit Sub
    End If
    
    'CopyMemory mbytAudioSample(0), bytBuffer(0), WavFile.BlockAlign * WavFile.ReadLength
    mbytAudioSample = bytBuffer
    WavFile.WavData = mbytAudioSample
    WavFile.ConvertWavData

'    Debug.Print mbytAudioSample(0), mbytAudioSample(1), mbytAudioSample(2), mbytAudioSample(3),
'    Debug.Print
    
    'Call AudioInStop

    If TunerMode.Record And TunerState.Running Then         ' Write to File if Recording
        Call WavFile.WriteWavData
    End If
        
    AudioBuffer.InputBuffer = WavFile.ReadData              ' Write into Audio Buffer
    AudioBuffer.BufferWrite                                 ' Write into Audio Buffer
    
End Sub

Private Sub AudioInStart()
    
    ' setup tuner
    Call InitializeTuner
    
    If TunerMode.Live Or TunerMode.Record Then
        
        Dim ErrReturn As String
        ReDim mbytAudioSample(0 To WavFile.ReadByteLength - 1)
        
        With DirectSoundRecord
        
            ' initialize DirectSound with exactly the same sound format as you expect to write in the file
            ErrReturn = .Initialize(WavFile.FormatTag, _
                                    WavFile.SampleRate, _
                                    WavFile.BitsPerSample, _
                                    WavFile.Channels, _
                                    WavFile.ReadByteLength)
                                
            If Len(ErrReturn) = 0 Then
                ' if there was no error
                ' start recording
                .SoundPlay

            Else
                MsgBox ErrReturn, vbExclamation, "DirectSound Error"
            End If
            
        End With
    
    End If
    
    If TunerMode.WAV Then
        ' prepare file
        WavFile.OpenFile
        WavFile.SampleStart = WavFile.SectionTimeStart * WavFile.SampleRate
        
        ' read as long as current sample position < stop position of samples in file
        Do While TunerState.Running And _
            (WavFile.SampleStart + WavFile.ReadLength) < WavFile.SectionTimeStop * WavFile.SampleRate _
            And (WavFile.SampleStart + WavFile.ReadLength) < WavFile.DataSamples

                WavFile.ReadWavData
                AudioBuffer.InputBuffer = WavFile.ReadData              ' Write into Audio Buffer
                AudioBuffer.BufferWrite                                 ' Write into Audio Buffer
                DoEvents                        '( might be stopped by pressing Start/Stop button)
        Loop
        
        ' if tuner is still running stop it
        If TunerState.Running Then cmd_Start_Stop_Click
        ' set level meter to 0
        obj_Level_Meter.Level = 0
        
    End If
       
End Sub

Private Sub AudioInStop()
    
    ' set tuner status
    TunerState.Stop = True
    TunerState.Running = False
    
    If TunerMode.Live Or TunerMode.Record Then
    
        With DirectSoundRecord
            ' stop recording
            .SoundStop
        
            ' un-initialize DirectSound
            .UninitializeSound
        End With
    
    End If
    
    If TunerMode.WAV Or TunerMode.Record Then Call WavFile.CloseFile
    
    ' set buttons
    Call Enable_Buttons(TunerState.Stop)
    ' set level meter to 0
    obj_Level_Meter.Level = 0
    
    'Debug.Print RecordVolume.SelectedDevice, RecordVolume.SelectedMixerLine, RecordVolume.MixerLineVolume
    
End Sub

Private Sub InitializeTuner()
    
    Dim lngNumberOfElements As Long
    
    ' set buttons
    TunerState.Stop = False
    Call Enable_Buttons(TunerState.Stop)

    ' set length for Input : must be multiple of FFT_Interval (sample interval)
    WavFile.ReadLength = WavFile.SampleInterval * _
                Round(gsng_RefreshInterval * WavFile.SampleRate / WavFile.SampleInterval, 0)
    If WavFile.ReadLength < WavFile.SampleInterval Then WavFile.ReadLength = WavFile.SampleInterval


    With AudioBuffer
        
        .Channels = WavFile.Channels
        ' set interval for FFT
        .BufferStepLength = WavFile.SampleInterval
        ' set length for FFT
        .OutputBufferLength = WavFile.SampleLength
        ' set length for Input
        .InputBufferLength = WavFile.ReadLength
        
        .BufferClear            ' clear audio buffer

    End With
    
    With WavFile
        ' wait until audio input good for a period of time: 4s
        AudioLevelOkFIFO.FIFO_Elements = 4 * Round(.SampleRate / .SampleInterval)
        ' stop after audio inout not good for another period of time: 1s
        AudioLevelOkFIFO.FIFO_MaxNonValidElements = 1 * Round(.SampleRate / .SampleInterval)
        ' clear buffer
        AudioLevelOkFIFO.FIFO_Clear
    End With
    
    
    If TunerMode.Record Then
        ' prepare new file and header
        WavFile.Delete
        WavFile.OpenFile
        WavFile.WriteHeader
    End If
           
    If TunerMode.Record Or TunerMode.WAV Then
    ' resize allfrequencies array to save values / also set scrollbar
        lngNumberOfElements = CLng(WavFile.SectionTimeLength * WavFile.SampleRate / WavFile.SampleInterval)
        ReDim mdblAllFrequencies(0 To lngNumberOfElements - 1, 0 To UBound(gudtDrones))
        hsb_ScrollTime.Value = 0
        hsb_ScrollTime.Max = lngNumberOfElements
        
    End If
    
    If TunerMode.WAV Then
        ' compensate buffer start position for wav file
        WavFile.StartTime = AudioBuffer.BufferStartPosition / WavFile.SampleRate
    Else
        ' start = 0 for live and record
        gdblBufferFrequency.FIFO_StartPosition = 0
    End If
    
    FrequencyDetection.Set_RefFreq (gdblReferenceFrequency)
    FrequencyDetection.Init
    Display.Init
    Display.SetAxis AxisSeconds

End Sub

Private Sub AudioBuffer_BufferReady()

    Dim i As Long
    Dim lngFFTCounter As Long
    Dim dblFrequencies() As Double
    
    ' Write Audio Level Meter
    dB_Meter.Update_dB_Level WavFile.dB_Value
    
    ' read as long as there is datas in the audio read buffer
    Do While AudioBuffer.BufferRead <> 0

        mintAudioData = AudioBuffer.OutputBuffer
        
'        Debug.Print mintAudioData(0, 0)

        
        If dB_Meter.Level_Good Then
            AudioLevelOkFIFO.FIFO_Fill (1)                   ' Set to 1 for volume ok
        Else
            AudioLevelOkFIFO.FIFO_Fill (0)                   ' Set to 0 for volume to not ok
        End If
        
        ' start live analysis if audio level was ok for some time and stop if it is not for a short time
        ' parameters set in AudioInStart
        If TunerState.Running And Not (TunerMode.Live And AudioLevelOkFIFO.FIFO_Result = 0) Then
            FrequencyDetection.Set_RefFreq (gdblReferenceFrequency)
            NoteBuffer.BufferPipes (FrequencyDetection.MeasureFrequencies(mintAudioData))
        End If
        
        If (TunerMode.WAV Or TunerMode.Record) And TunerState.Running Then
            'copy frequencies from frequencies buffer
            dblFrequencies = gdblBufferFrequency.FIFO_Input
            ' current FFT position
            lngFFTCounter = gdblBufferFrequency.FIFO_StartPosition + gdblBufferFrequency.FIFO_Position - 1
            ' copy current frequencies into array with all frequencies
            For i = 0 To UBound(gudtDrones)
                mdblAllFrequencies(lngFFTCounter, i) = dblFrequencies(i)
            Next i
        End If
        
    Loop
    
    If Not TunerMode.Live Then
        ' scroll to current buffer start position
        hsb_ScrollTime.Max = gdblBufferFrequency.FIFO_StartPosition
        hsb_ScrollTime.Value = gdblBufferFrequency.FIFO_StartPosition
    End If
    If TunerMode.Record And TunerState.Running Then
        ' text = current time of recording
        txt_Timer.Text = Format(WavFile.DataTime, "###0.00")
        ' stop if current time is longer then stop time (= start time + length)
        If WavFile.DataTime >= WavFile.SectionTimeStop Then cmd_Start_Stop_Click
    End If
    If TunerMode.WAV Then
        ' text = current time of wav file
        txt_Timer.Text = Format(WavFile.SampleStart / WavFile.SampleRate, "###0.00")
    End If
    
    ' refresh display
    Display.Draw
    ' update axis
    Display.SetAxis AxisSeconds

End Sub
