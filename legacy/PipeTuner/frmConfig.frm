VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   Caption         =   "No Drone"
   ClientHeight    =   7440
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   496
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   696
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   855
      Left            =   8520
      TabIndex        =   226
      Top             =   4920
      Width           =   1815
   End
   Begin VB.PictureBox picColorDrone 
      Height          =   495
      Index           =   0
      Left            =   1320
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   225
      Top             =   8400
      Width           =   495
   End
   Begin VB.CheckBox chkDrone 
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   223
      Top             =   8400
      Width           =   255
   End
   Begin VB.Frame labDrones 
      Caption         =   "Drones"
      Height          =   2655
      Left            =   8520
      TabIndex        =   216
      Top             =   720
      Width           =   1815
      Begin VB.PictureBox picColorDrone 
         Height          =   495
         Index           =   2
         Left            =   1080
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   222
         Top             =   1080
         Width           =   495
      End
      Begin VB.CheckBox chkDrone 
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   220
         Top             =   1080
         Width           =   255
      End
      Begin VB.PictureBox picColorDrone 
         Height          =   495
         Index           =   1
         Left            =   1080
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   219
         Top             =   240
         Width           =   495
      End
      Begin VB.CheckBox chkDrone 
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   217
         Top             =   240
         Width           =   255
      End
      Begin VB.Label labDrone 
         Caption         =   "No Drone"
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
         Index           =   2
         Left            =   600
         TabIndex        =   221
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label labDrone 
         Caption         =   "No Drone"
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
         Index           =   1
         Left            =   600
         TabIndex        =   218
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   11
      Left            =   0
      TabIndex        =   188
      Top             =   720
      Width           =   8175
      Begin VB.PictureBox picSymbolNote 
         Height          =   540
         Index           =   11
         Left            =   600
         Picture         =   "frmConfig.frx":0000
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   204
         Top             =   56
         Width           =   540
      End
      Begin VB.Frame frameCentNote 
         Height          =   615
         Index           =   11
         Left            =   4920
         TabIndex        =   200
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   11
            Left            =   960
            TabIndex        =   202
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   201
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   11
            Left            =   720
            TabIndex        =   203
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1000
            Min             =   -1000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   11
         Left            =   4560
         TabIndex        =   199
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   11
         Left            =   2160
         TabIndex        =   198
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Height          =   615
         Index           =   11
         Left            =   2520
         TabIndex        =   193
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   11
            Left            =   1080
            TabIndex        =   195
            Text            =   "1"
            Top             =   240
            Width           =   165
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   11
            Left            =   600
            TabIndex        =   194
            Text            =   "1"
            Top             =   240
            Width           =   165
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   11
            Left            =   1440
            TabIndex        =   196
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(11)"
            BuddyDispid     =   196622
            BuddyIndex      =   11
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorNote 
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   197
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(11)"
            BuddyDispid     =   196623
            BuddyIndex      =   11
            OrigLeft        =   120
            OrigTop         =   240
            OrigRight       =   375
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioNote 
            Index           =   11
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzNote 
         Height          =   615
         Index           =   11
         Left            =   7200
         TabIndex        =   191
         Top             =   0
         Width           =   975
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Index           =   11
            Left            =   120
            TabIndex        =   192
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   11
         Left            =   1320
         ScaleHeight     =   32
         ScaleMode       =   0  'Benutzerdefiniert
         ScaleWidth      =   32
         TabIndex        =   190
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   11
         Left            =   240
         TabIndex        =   189
         Top             =   80
         Width           =   255
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Ratio"
      Height          =   495
      Left            =   2520
      TabIndex        =   182
      Top             =   0
      Width           =   1815
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "compared to reference"
         Height          =   255
         Left            =   120
         TabIndex        =   183
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frequency"
      Height          =   495
      Left            =   7200
      TabIndex        =   181
      Top             =   0
      Width           =   975
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "in Hz"
         Height          =   255
         Left            =   360
         TabIndex        =   184
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Cent"
      Height          =   495
      Left            =   4920
      TabIndex        =   177
      Top             =   0
      Width           =   1815
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "absolute"
         Height          =   255
         Left            =   960
         TabIndex        =   179
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "relative*"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   178
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   10
      Left            =   0
      TabIndex        =   157
      Top             =   1320
      Width           =   8175
      Begin VB.PictureBox picSymbolNote 
         Height          =   540
         Index           =   10
         Left            =   600
         Picture         =   "frmConfig.frx":00CA
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   205
         Top             =   56
         Width           =   540
      End
      Begin VB.Frame frameCentNote 
         Height          =   615
         Index           =   10
         Left            =   4920
         TabIndex        =   169
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   10
            Left            =   960
            TabIndex        =   171
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   170
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   10
            Left            =   720
            TabIndex        =   172
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1000
            Min             =   -1000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   10
         Left            =   4560
         TabIndex        =   168
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   10
         Left            =   2160
         TabIndex        =   167
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Height          =   615
         Index           =   10
         Left            =   2520
         TabIndex        =   162
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   10
            Left            =   1080
            TabIndex        =   164
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   10
            Left            =   600
            TabIndex        =   163
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   10
            Left            =   1440
            TabIndex        =   165
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(10)"
            BuddyDispid     =   196622
            BuddyIndex      =   10
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorNote 
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   166
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(10)"
            BuddyDispid     =   196623
            BuddyIndex      =   10
            OrigLeft        =   120
            OrigTop         =   240
            OrigRight       =   375
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioNote 
            Index           =   10
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzNote 
         Height          =   615
         Index           =   10
         Left            =   7200
         TabIndex        =   160
         Top             =   0
         Width           =   975
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   161
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   10
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   159
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   10
         Left            =   240
         TabIndex        =   158
         Top             =   80
         Width           =   255
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   9
      Left            =   0
      TabIndex        =   141
      Top             =   1920
      Width           =   8175
      Begin VB.PictureBox picSymbolNote 
         Height          =   540
         Index           =   9
         Left            =   600
         Picture         =   "frmConfig.frx":0194
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   206
         Top             =   56
         Width           =   540
      End
      Begin VB.Frame frameCentNote 
         Height          =   615
         Index           =   9
         Left            =   4920
         TabIndex        =   153
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   9
            Left            =   960
            TabIndex        =   155
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   154
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   9
            Left            =   720
            TabIndex        =   156
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1000
            Min             =   -1000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   9
         Left            =   4560
         TabIndex        =   152
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   9
         Left            =   2160
         TabIndex        =   151
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Height          =   615
         Index           =   9
         Left            =   2520
         TabIndex        =   146
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   9
            Left            =   1080
            TabIndex        =   148
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   9
            Left            =   600
            TabIndex        =   147
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   9
            Left            =   1440
            TabIndex        =   149
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(9)"
            BuddyDispid     =   196622
            BuddyIndex      =   9
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorNote 
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   150
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(9)"
            BuddyDispid     =   196623
            BuddyIndex      =   9
            OrigLeft        =   120
            OrigTop         =   240
            OrigRight       =   375
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioNote 
            Index           =   9
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzNote 
         Height          =   615
         Index           =   9
         Left            =   7200
         TabIndex        =   144
         Top             =   0
         Width           =   975
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   145
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   9
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   143
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   9
         Left            =   240
         TabIndex        =   142
         Top             =   80
         Width           =   255
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   8
      Left            =   0
      TabIndex        =   125
      Top             =   2520
      Width           =   8175
      Begin VB.PictureBox picSymbolNote 
         Height          =   540
         Index           =   8
         Left            =   600
         Picture         =   "frmConfig.frx":025E
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   207
         Top             =   56
         Width           =   540
      End
      Begin VB.Frame frameCentNote 
         Height          =   615
         Index           =   8
         Left            =   4920
         TabIndex        =   137
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   8
            Left            =   960
            TabIndex        =   139
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   138
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   8
            Left            =   720
            TabIndex        =   140
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1000
            Min             =   -1000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   8
         Left            =   4560
         TabIndex        =   136
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   8
         Left            =   2160
         TabIndex        =   135
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Height          =   615
         Index           =   8
         Left            =   2520
         TabIndex        =   130
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   8
            Left            =   1080
            TabIndex        =   132
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   8
            Left            =   600
            TabIndex        =   131
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   8
            Left            =   1440
            TabIndex        =   133
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(8)"
            BuddyDispid     =   196622
            BuddyIndex      =   8
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorNote 
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   134
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(8)"
            BuddyDispid     =   196623
            BuddyIndex      =   8
            OrigLeft        =   120
            OrigTop         =   240
            OrigRight       =   375
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioNote 
            Index           =   8
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzNote 
         Height          =   615
         Index           =   8
         Left            =   7200
         TabIndex        =   128
         Top             =   0
         Width           =   975
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   129
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   8
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   127
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   8
         Left            =   240
         TabIndex        =   126
         Top             =   120
         Width           =   255
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   7
      Left            =   0
      TabIndex        =   109
      Top             =   3120
      Width           =   8175
      Begin VB.PictureBox picSymbolNote 
         Height          =   540
         Index           =   7
         Left            =   600
         Picture         =   "frmConfig.frx":0328
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   215
         Top             =   56
         Width           =   540
      End
      Begin VB.Frame frameCentNote 
         Height          =   615
         Index           =   7
         Left            =   4920
         TabIndex        =   121
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   7
            Left            =   960
            TabIndex        =   123
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   122
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   7
            Left            =   720
            TabIndex        =   124
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1000
            Min             =   -1000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   7
         Left            =   4560
         TabIndex        =   120
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   7
         Left            =   2160
         TabIndex        =   119
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Height          =   615
         Index           =   7
         Left            =   2520
         TabIndex        =   114
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   116
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   7
            Left            =   600
            TabIndex        =   115
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   7
            Left            =   1440
            TabIndex        =   117
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(7)"
            BuddyDispid     =   196622
            BuddyIndex      =   7
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorNote 
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   118
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(7)"
            BuddyDispid     =   196623
            BuddyIndex      =   7
            OrigLeft        =   120
            OrigTop         =   240
            OrigRight       =   375
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioNote 
            Index           =   7
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzNote 
         Height          =   615
         Index           =   7
         Left            =   7200
         TabIndex        =   112
         Top             =   0
         Width           =   975
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   113
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   7
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   111
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   7
         Left            =   240
         TabIndex        =   110
         Top             =   80
         Width           =   255
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   6
      Left            =   0
      TabIndex        =   93
      Top             =   3720
      Width           =   8175
      Begin VB.PictureBox picSymbolNote 
         Height          =   540
         Index           =   6
         Left            =   600
         Picture         =   "frmConfig.frx":03F2
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   208
         Top             =   56
         Width           =   540
      End
      Begin VB.Frame frameCentNote 
         Height          =   615
         Index           =   6
         Left            =   4920
         TabIndex        =   105
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   6
            Left            =   960
            TabIndex        =   107
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   106
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   6
            Left            =   720
            TabIndex        =   108
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1000
            Min             =   -1000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   6
         Left            =   4560
         TabIndex        =   104
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   6
         Left            =   2160
         TabIndex        =   103
         Top             =   120
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Height          =   615
         Index           =   6
         Left            =   2520
         TabIndex        =   98
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   6
            Left            =   1080
            TabIndex        =   100
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   6
            Left            =   600
            TabIndex        =   99
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   6
            Left            =   1440
            TabIndex        =   101
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(6)"
            BuddyDispid     =   196622
            BuddyIndex      =   6
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorNote 
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   102
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(6)"
            BuddyDispid     =   196623
            BuddyIndex      =   6
            OrigLeft        =   120
            OrigTop         =   240
            OrigRight       =   375
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioNote 
            Index           =   6
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzNote 
         Height          =   615
         Index           =   6
         Left            =   7200
         TabIndex        =   96
         Top             =   0
         Width           =   975
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   97
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   6
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   95
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   6
         Left            =   240
         TabIndex        =   94
         Top             =   80
         Width           =   255
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   5
      Left            =   0
      TabIndex        =   77
      Top             =   4320
      Width           =   8175
      Begin VB.PictureBox picSymbolNote 
         Height          =   540
         Index           =   5
         Left            =   600
         Picture         =   "frmConfig.frx":04BC
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   209
         Top             =   56
         Width           =   540
      End
      Begin VB.Frame frameCentNote 
         Height          =   615
         Index           =   5
         Left            =   4920
         TabIndex        =   89
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   5
            Left            =   960
            TabIndex        =   91
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   90
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   5
            Left            =   720
            TabIndex        =   92
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1000
            Min             =   -1000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   5
         Left            =   4560
         TabIndex        =   88
         Top             =   120
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   5
         Left            =   2160
         TabIndex        =   87
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Height          =   615
         Index           =   5
         Left            =   2520
         TabIndex        =   82
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   84
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   5
            Left            =   600
            TabIndex        =   83
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   5
            Left            =   1440
            TabIndex        =   85
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(5)"
            BuddyDispid     =   196622
            BuddyIndex      =   5
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorNote 
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   86
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(5)"
            BuddyDispid     =   196623
            BuddyIndex      =   5
            OrigLeft        =   120
            OrigTop         =   240
            OrigRight       =   375
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioNote 
            Index           =   5
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzNote 
         Height          =   615
         Index           =   5
         Left            =   7200
         TabIndex        =   80
         Top             =   0
         Width           =   975
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   81
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   5
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   79
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   5
         Left            =   240
         TabIndex        =   78
         Top             =   80
         Width           =   255
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   4
      Left            =   0
      TabIndex        =   65
      Top             =   4920
      Width           =   8175
      Begin VB.PictureBox picSymbolNote 
         Height          =   540
         Index           =   4
         Left            =   600
         Picture         =   "frmConfig.frx":0586
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   210
         Top             =   56
         Width           =   540
      End
      Begin VB.Frame frameCentNote 
         Height          =   615
         Index           =   4
         Left            =   4920
         TabIndex        =   173
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   175
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   4
            Left            =   960
            TabIndex        =   174
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   4
            Left            =   720
            TabIndex        =   176
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1000
            Min             =   -1000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   4
         Left            =   4560
         TabIndex        =   76
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   4
         Left            =   2160
         TabIndex        =   75
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Height          =   615
         Index           =   4
         Left            =   2520
         TabIndex        =   70
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   72
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   4
            Left            =   600
            TabIndex        =   71
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   4
            Left            =   1440
            TabIndex        =   73
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(4)"
            BuddyDispid     =   196622
            BuddyIndex      =   4
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorNote 
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   74
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(4)"
            BuddyDispid     =   196623
            BuddyIndex      =   4
            OrigLeft        =   120
            OrigTop         =   240
            OrigRight       =   375
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioNote 
            Index           =   4
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzNote 
         Height          =   615
         Index           =   4
         Left            =   7200
         TabIndex        =   68
         Top             =   0
         Width           =   975
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   69
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   4
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   67
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   4
         Left            =   240
         TabIndex        =   66
         Top             =   80
         Width           =   255
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   3
      Left            =   0
      TabIndex        =   49
      Top             =   5520
      Width           =   8175
      Begin VB.PictureBox picSymbolNote 
         Height          =   540
         Index           =   3
         Left            =   600
         Picture         =   "frmConfig.frx":0650
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   211
         Top             =   56
         Width           =   540
      End
      Begin VB.Frame frameCentNote 
         Height          =   615
         Index           =   3
         Left            =   4920
         TabIndex        =   61
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   3
            Left            =   960
            TabIndex        =   63
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   62
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   3
            Left            =   720
            TabIndex        =   64
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1000
            Min             =   -1000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   3
         Left            =   4560
         TabIndex        =   60
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   3
         Left            =   2160
         TabIndex        =   59
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Height          =   615
         Index           =   3
         Left            =   2520
         TabIndex        =   54
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   56
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   3
            Left            =   600
            TabIndex        =   55
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   3
            Left            =   1440
            TabIndex        =   57
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(3)"
            BuddyDispid     =   196622
            BuddyIndex      =   3
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorNote 
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(3)"
            BuddyDispid     =   196623
            BuddyIndex      =   3
            OrigLeft        =   120
            OrigTop         =   240
            OrigRight       =   375
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioNote 
            Index           =   3
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzNote 
         Height          =   615
         Index           =   3
         Left            =   7200
         TabIndex        =   52
         Top             =   0
         Width           =   975
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   53
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   3
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   51
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   3
         Left            =   240
         TabIndex        =   50
         Top             =   80
         Width           =   255
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   2
      Left            =   0
      TabIndex        =   33
      Top             =   6120
      Width           =   8175
      Begin VB.PictureBox picSymbolNote 
         Height          =   540
         Index           =   2
         Left            =   600
         Picture         =   "frmConfig.frx":071A
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   212
         Top             =   56
         Width           =   540
      End
      Begin VB.Frame frameCentNote 
         Height          =   615
         Index           =   2
         Left            =   4920
         TabIndex        =   45
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   960
            TabIndex        =   47
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   46
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   48
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1000
            Min             =   -1000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   2
         Left            =   4560
         TabIndex        =   44
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   2
         Left            =   2160
         TabIndex        =   43
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Height          =   615
         Index           =   2
         Left            =   2520
         TabIndex        =   38
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   40
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   630
            TabIndex        =   39
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   2
            Left            =   1440
            TabIndex        =   41
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(2)"
            BuddyDispid     =   196622
            BuddyIndex      =   2
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorNote 
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(2)"
            BuddyDispid     =   196623
            BuddyIndex      =   2
            OrigLeft        =   120
            OrigTop         =   240
            OrigRight       =   375
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioNote 
            Index           =   2
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzNote 
         Height          =   615
         Index           =   2
         Left            =   7200
         TabIndex        =   36
         Top             =   0
         Width           =   975
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   2
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   35
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   2
         Left            =   240
         TabIndex        =   34
         Top             =   80
         Width           =   255
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   1
      Left            =   0
      TabIndex        =   17
      Top             =   6720
      Width           =   8175
      Begin VB.PictureBox picSymbolNote 
         Height          =   540
         Index           =   1
         Left            =   600
         Picture         =   "frmConfig.frx":07E4
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   213
         Top             =   56
         Width           =   540
      End
      Begin VB.Frame frameCentNote 
         Height          =   615
         Index           =   1
         Left            =   4920
         TabIndex        =   29
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   960
            TabIndex        =   31
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   32
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1000
            Min             =   -1000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   1
         Left            =   4560
         TabIndex        =   28
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   1
         Left            =   2160
         TabIndex        =   27
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Height          =   615
         Index           =   1
         Left            =   2520
         TabIndex        =   22
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   24
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   630
            TabIndex        =   23
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   1
            Left            =   1440
            TabIndex        =   25
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(1)"
            BuddyDispid     =   196622
            BuddyIndex      =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorNote 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(1)"
            BuddyDispid     =   196623
            BuddyIndex      =   1
            OrigLeft        =   120
            OrigTop         =   240
            OrigRight       =   375
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioNote 
            Index           =   1
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzNote 
         Height          =   615
         Index           =   1
         Left            =   7200
         TabIndex        =   20
         Top             =   0
         Width           =   975
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   1
         Left            =   1320
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   19
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   18
         Top             =   80
         Width           =   255
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   7560
      Width           =   8175
      Begin VB.PictureBox picSymbolNote 
         Height          =   540
         Index           =   0
         Left            =   600
         Picture         =   "frmConfig.frx":08AE
         ScaleHeight     =   32
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   214
         Top             =   56
         Width           =   540
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   80
         Width           =   255
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   0
         Left            =   1320
         ScaleHeight     =   32
         ScaleMode       =   0  'Benutzerdefiniert
         ScaleWidth      =   32
         TabIndex        =   15
         Top             =   80
         Width           =   495
      End
      Begin VB.Frame frameHzNote 
         Height          =   615
         Index           =   0
         Left            =   7200
         TabIndex        =   13
         Top             =   0
         Width           =   975
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Enabled         =   0   'False
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame frameRatioNote 
         Height          =   615
         Index           =   0
         Left            =   2520
         TabIndex        =   8
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   600
            TabIndex        =   10
            Text            =   "1"
            Top             =   240
            Width           =   420
         End
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   9
            Text            =   "1"
            Top             =   240
            Width           =   420
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   11
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(0)"
            BuddyDispid     =   196622
            BuddyIndex      =   0
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorNote 
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(0)"
            BuddyDispid     =   196623
            BuddyIndex      =   0
            OrigLeft        =   120
            OrigTop         =   240
            OrigRight       =   375
            OrigBottom      =   525
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioNote 
            Index           =   0
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   0
         Left            =   2160
         TabIndex        =   7
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   0
         Left            =   4560
         TabIndex        =   6
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameCentNote 
         Height          =   615
         Index           =   0
         Left            =   4920
         TabIndex        =   2
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   960
            TabIndex        =   3
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   5
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   1000
            Min             =   -1000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   855
      Left            =   8520
      TabIndex        =   0
      Top             =   3720
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog objStdDialog 
      Left            =   1920
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line2 
      Index           =   1
      X1              =   696
      X2              =   0
      Y1              =   40
      Y2              =   40
   End
   Begin VB.Label labDrone 
      Caption         =   "No Drone"
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
      Index           =   0
      Left            =   720
      TabIndex        =   224
      Top             =   8520
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Note"
      Height          =   255
      Left            =   720
      TabIndex        =   187
      Top             =   120
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "Active"
      Height          =   255
      Left            =   120
      TabIndex        =   186
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   " Color"
      Height          =   375
      Left            =   1320
      TabIndex        =   185
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "* relative to chromatic scale"
      Height          =   495
      Index           =   2
      Left            =   8520
      TabIndex        =   180
      Top             =   120
      Width           =   1335
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   560
      X2              =   560
      Y1              =   -48
      Y2              =   496
   End
   Begin VB.Menu menuTemperament 
      Caption         =   "&Temperament"
      Index           =   1
      Begin VB.Menu menuSetTemperamentAsDefault 
         Caption         =   "Set Current Temperament as &Default"
         Index           =   1
      End
      Begin VB.Menu menuFileBar1 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu menuSaveTemperament 
         Caption         =   "&Save Current Temperament"
         Index           =   1
      End
      Begin VB.Menu menuLoadTemperament 
         Caption         =   "&Load Temperament"
         Index           =   1
      End
   End
   Begin VB.Menu menuColor 
      Caption         =   "&Color Settings"
      Index           =   1
      Begin VB.Menu menuSetColorsAsDefault 
         Caption         =   "Set Current Colors as &Default"
         Index           =   1
      End
      Begin VB.Menu menuFileBar2 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu menuSaveColors 
         Caption         =   "&Save Current Colors"
         Index           =   1
      End
      Begin VB.Menu menuLoadColors 
         Caption         =   "&Load Colors"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mudtNoteDefaults() As NoteAttributes
Private mudtDroneDefaults() As NoteAttributes

Private Sub Form_Load()

    Dim i As Integer

    frmConfig.Caption = "Scale Configuration: Chanter / Drones"
    
    mudtNoteDefaults = gudtNoteDefaults
    mudtDroneDefaults = gudtDroneDefaults
    
    'For i = LBound(gudtNoteDefaults) To UBound(gudtNoteDefaults)
    '    Debug.Print gudtNoteDefaults(i).Name, gudtNoteDefaults(i).CentSelected, _
    '    gudtNoteDefaults(i).Numerator, gudtNoteDefaults(i).Denominator, gudtNoteDefaults(i).Ratio, _
    '    gudtNoteDefaults(i).AbsoluteCent, gudtNoteDefaults(i).RelativeCent
    'Next i
    'Debug.Print
    
    For i = LBound(gudtNoteDefaults) To UBound(gudtNoteDefaults)
        
        'mudtNoteDefaults(i) = gudtNoteDefaults(i)
        
        UpDownRatioNumeratorNote(i).Wrap = False
        UpDownRatioDenominatorNote(i).Wrap = False
        'UpDownRatioNumeratorNote(i).BuddyControl = txtRatioNumeratorNote(i)
        'UpDownRatioDenominatorNote(i).BuddyControl = txtRatioDenominatorNote(i)
        'UpDownRatioNumeratorNote(i).BuddyProperty = "Text"
        'UpDownRatioDenominatorNote(i).BuddyProperty = "Text"
        'UpDownRatioNumeratorNote(i).SyncBuddy = True
        'UpDownRatioDenominatorNote(i).SyncBuddy = True
        UpDownRatioNumeratorNote(i).Min = 0
        UpDownRatioDenominatorNote(i).Min = 1
        UpDownRatioNumeratorNote(i).Max = 99
        UpDownRatioDenominatorNote(i).Max = 99
        
        'UpDownRatioNumeratorNote(i).Left = 120
        'UpDownRatioNumeratorNote(i).Top = 240
        'UpDownRatioNumeratorNote(i).Width = 255
        'UpDownRatioNumeratorNote(i).Height = 285
        'UpDownRatioDenominatorNote(i).Left = 1440
        'UpDownRatioDenominatorNote(i).Top = 240
        'UpDownRatioDenominatorNote(i).Width = 255
        'UpDownRatioDenominatorNote(i).Height = 285
        txtRatioNumeratorNote(i).Left = 360
        txtRatioNumeratorNote(i).Top = 240
        txtRatioNumeratorNote(i).Width = 360
        txtRatioNumeratorNote(i).Height = 285
        txtRatioDenominatorNote(i).Left = 1080
        txtRatioDenominatorNote(i).Top = 240
        txtRatioDenominatorNote(i).Width = 360
        txtRatioDenominatorNote(i).Height = 285
        
    Next i
    
    Call GetDefaults
    
End Sub

Private Sub GetDefaults()

    Dim i As Integer
    
    For i = LBound(gudtNoteDefaults) To UBound(gudtNoteDefaults)
        picColorNote(i).BackColor = gudtNoteDefaults(i).Color
        picSymbolNote(i).Picture = gudtNoteDefaults(i).Pic
        UpDownRatioNumeratorNote(i).Value = gudtNoteDefaults(i).Numerator
        UpDownRatioDenominatorNote(i).Value = gudtNoteDefaults(i).Denominator
        
        chkNote(i).Value = 1
        optRatioNote(i).Value = Not gudtNoteDefaults(i).CentSelected
        optCentNote(i).Value = gudtNoteDefaults(i).CentSelected
        
    Next i
    
    chkNote(4).Value = 0
    chkNote(8).Value = 0
    
     
End Sub

Private Sub picColorNote_Click(Index As Integer)
     On Error Resume Next
     objStdDialog.Color = picColorNote(Index).BackColor
     objStdDialog.CancelError = True
     objStdDialog.flags = cdlCCFullOpen + cdlCCRGBInit
     objStdDialog.ShowColor
     If Err = 0 Then picColorNote(Index).BackColor = objStdDialog.Color
     MsgBox (Hex(picColorNote(Index).BackColor))
End Sub
Private Sub picColorDrone_Click(Index As Integer)
     On Error Resume Next
     objStdDialog.Color = picColorDrone(Index).BackColor
     objStdDialog.CancelError = True
     objStdDialog.flags = cdlCCFullOpen + cdlCCRGBInit
     objStdDialog.ShowColor
     If Err = 0 Then picColorDrone(Index).BackColor = objStdDialog.Color
     MsgBox (Hex(picColorDrone(Index).BackColor))
End Sub

Private Sub updownCentNote_Change(Index As Integer)
    If optCentNote(Index).Value = False Then Exit Sub
    UpDownCentNote(Index).Increment = 10
    mudtNoteDefaults(Index).RelativeCent = UpDownCentNote(Index).Value / 10
    mudtNoteDefaults(Index).AbsoluteCent = mudtNoteDefaults(Index).RelativeCent + mudtNoteDefaults(Index).ChromaticCent
    mudtNoteDefaults(Index).Ratio = 2 ^ (mudtNoteDefaults(Index).AbsoluteCent / 1200)
    txtRelCentNote(Index).Text = Format(mudtNoteDefaults(Index).RelativeCent, "##0.0")
    txtAbsCentNote(Index).Text = Format(mudtNoteDefaults(Index).AbsoluteCent, "####0.0")
    txtHzNote(Index).Text = Format(mudtNoteDefaults(Index).Ratio * gdblReferenceFrequency, "###0.0")
End Sub
Private Sub UpDownRatioNumeratorNote_Change(Index As Integer)
    If optRatioNote(Index).Value = False Then Exit Sub
    txtRatioNumeratorNote(Index).Enabled = True
    txtRatioDenominatorNote(Index).Enabled = True
    Call optRatioNote_Click(Index)
End Sub
Private Sub UpDownRatioDenominatorNote_Change(Index As Integer)
    If optRatioNote(Index).Value = False Then Exit Sub
    txtRatioNumeratorNote(Index).Enabled = True
    txtRatioDenominatorNote(Index).Enabled = True
    Call optRatioNote_Click(Index)
End Sub
Private Sub optRatioNote_Click(Index As Integer)
    
    'optRatioNote(Index).Value = True
    UpDownCentNote(Index).Enabled = False
    txtAbsCentNote(Index).Enabled = False
    txtRelCentNote(Index).Enabled = False
    txtRatioNumeratorNote(Index).Enabled = True
    txtRatioDenominatorNote(Index).Enabled = True
    
    mudtNoteDefaults(Index).CentSelected = False
    mudtNoteDefaults(Index).Numerator = UpDownRatioNumeratorNote(Index).Value
    mudtNoteDefaults(Index).Denominator = UpDownRatioDenominatorNote(Index).Value
    mudtNoteDefaults(Index).Ratio = mudtNoteDefaults(Index).Numerator / mudtNoteDefaults(Index).Denominator
    mudtNoteDefaults(Index).AbsoluteCent = ConvertFrequencyInAbsoluteCent(1, mudtNoteDefaults(Index).Ratio)
    mudtNoteDefaults(Index).RelativeCent = mudtNoteDefaults(Index).AbsoluteCent - _
                                            mudtNoteDefaults(Index).ChromaticCent
                    
    txtRelCentNote(Index).Text = Format(mudtNoteDefaults(Index).RelativeCent, "##0.0")
    txtAbsCentNote(Index).Text = Format(mudtNoteDefaults(Index).AbsoluteCent, "####0.0")
    txtHzNote(Index).Text = Format(mudtNoteDefaults(Index).Ratio * gdblReferenceFrequency, "###0.0")

End Sub
Private Sub optCentNote_Click(Index As Integer)
    'optCentNote(Index).Value = True
    mudtNoteDefaults(Index).CentSelected = True
    txtRatioNumeratorNote(Index).Enabled = False
    txtRatioDenominatorNote(Index).Enabled = False
    txtAbsCentNote(Index).Enabled = True
    txtRelCentNote(Index).Enabled = True
    UpDownCentNote(Index).Enabled = True
    UpDownCentNote(Index).Value = 10 * Round(mudtNoteDefaults(Index).RelativeCent)
End Sub

Private Sub chkNote_Click(Index As Integer)
    mudtNoteDefaults(Index).Selected = chkNote(Index).Value
End Sub

Private Sub cmdApply_Click()
    
    Dim i As Integer
    
    For i = LBound(gudtNoteDefaults) To UBound(gudtNoteDefaults)
        gudtNoteDefaults(i) = mudtNoteDefaults(i)
    Next i
    
    Me.Hide
    
    Call NoteDefinitions.MapNotes(gudtNotes, gudtNoteDefaults)
    Call Form_Main.Settings

    
    'For i = LBound(gudtNotes) To UBound(gudtNotes)
    '    Debug.Print gudtNotes(i).Name, gudtNotes(i).CentSelected, _
    '    gudtNotes(i).Numerator, gudtNotes(i).Denominator, gudtNotes(i).Ratio, _
    '    gudtNotes(i).AbsoluteCent, gudtNotes(i).RelativeCent
    'Next i
    'Debug.Print

End Sub

Private Sub cmdCancel_Click()
    
    Call GetDefaults
    Me.Hide
    
End Sub

