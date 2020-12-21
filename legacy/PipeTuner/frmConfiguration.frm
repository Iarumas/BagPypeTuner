VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfiguration 
   Caption         =   "Form1"
   ClientHeight    =   9825
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame frameDrone 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   2
      Left            =   0
      TabIndex        =   250
      Top             =   8520
      Width           =   10335
      Begin VB.OptionButton optCentDrone 
         Height          =   495
         Index           =   2
         Left            =   5520
         TabIndex        =   267
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioDrone 
         Height          =   495
         Index           =   2
         Left            =   3000
         TabIndex        =   266
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioDrone 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   2
         Left            =   3360
         TabIndex        =   261
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   263
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   615
            TabIndex        =   262
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorDrone 
            Height          =   285
            Index           =   2
            Left            =   1440
            TabIndex        =   264
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtRatioDenominatorDrone(2)"
            BuddyDispid     =   196613
            BuddyIndex      =   2
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   9
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorDrone 
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   265
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorDrone(2)"
            BuddyDispid     =   196614
            BuddyIndex      =   2
            OrigLeft        =   480
            OrigTop         =   240
            OrigRight       =   735
            OrigBottom      =   525
            Max             =   9
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioDrone 
            Index           =   2
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzDrone 
         Caption         =   "Hz"
         Height          =   615
         Index           =   2
         Left            =   9360
         TabIndex        =   259
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   260
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picColorDrone 
         Height          =   495
         Index           =   2
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   258
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkDrone 
         Height          =   495
         Index           =   2
         Left            =   1560
         TabIndex        =   257
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameCentDrone 
         Caption         =   "Cent"
         Height          =   615
         Index           =   2
         Left            =   5880
         TabIndex        =   251
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtRelCentDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   253
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtAbsCentDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   252
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin MSComCtl2.UpDown UpDownCentDrone 
            Height          =   285
            Index           =   2
            Left            =   1320
            TabIndex        =   254
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
            Max             =   12000
            Min             =   -12000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin VB.Label labRelCentDrone 
            Caption         =   "relative"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   256
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labAbsCentDrone 
            Caption         =   "absolute"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   255
            Top             =   280
            Width           =   615
         End
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
         Left            =   240
         TabIndex        =   268
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameDrone 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   1
      Left            =   0
      TabIndex        =   231
      Top             =   7920
      Width           =   10335
      Begin VB.OptionButton optCentDrone 
         Height          =   495
         Index           =   1
         Left            =   5520
         TabIndex        =   248
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioDrone 
         Height          =   495
         Index           =   1
         Left            =   3000
         TabIndex        =   247
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioDrone 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   1
         Left            =   3360
         TabIndex        =   242
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   244
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   615
            TabIndex        =   243
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorDrone 
            Height          =   285
            Index           =   1
            Left            =   1440
            TabIndex        =   245
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtRatioDenominatorDrone(1)"
            BuddyDispid     =   196613
            BuddyIndex      =   1
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   9
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorDrone 
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   246
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorDrone(1)"
            BuddyDispid     =   196614
            BuddyIndex      =   1
            OrigLeft        =   480
            OrigTop         =   240
            OrigRight       =   735
            OrigBottom      =   525
            Max             =   9
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioDrone 
            Index           =   1
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.Frame frameHzDrone 
         Caption         =   "Hz"
         Height          =   615
         Index           =   1
         Left            =   9360
         TabIndex        =   240
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   241
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picColorDrone 
         Height          =   495
         Index           =   1
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   239
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkDrone 
         Height          =   495
         Index           =   1
         Left            =   1560
         TabIndex        =   238
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameCentDrone 
         Caption         =   "Cent"
         Height          =   615
         Index           =   1
         Left            =   5880
         TabIndex        =   232
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtRelCentDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   234
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtAbsCentDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   1560
            TabIndex        =   233
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin MSComCtl2.UpDown UpDownCentDrone 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   235
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
            Max             =   12000
            Min             =   -12000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin VB.Label labRelCentDrone 
            Caption         =   "relative"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   237
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labAbsCentDrone 
            Caption         =   "absolute"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   236
            Top             =   280
            Width           =   615
         End
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
         Left            =   240
         TabIndex        =   249
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   10
      Left            =   0
      TabIndex        =   212
      Top             =   6480
      Width           =   10335
      Begin VB.Frame frameCentNote 
         Caption         =   "Cent"
         Height          =   615
         Index           =   10
         Left            =   5880
         TabIndex        =   224
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   10
            Left            =   1560
            TabIndex        =   226
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   10
            Left            =   720
            TabIndex        =   225
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   10
            Left            =   1320
            TabIndex        =   227
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
         Begin VB.Label labAbsCentNote 
            Caption         =   "absolute"
            Height          =   255
            Index           =   10
            Left            =   2400
            TabIndex        =   229
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labRelCentNote 
            Caption         =   "relative"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   228
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   10
         Left            =   5520
         TabIndex        =   223
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   10
         Left            =   3000
         TabIndex        =   222
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   10
         Left            =   3360
         TabIndex        =   217
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   10
            Left            =   1080
            TabIndex        =   219
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   10
            Left            =   600
            TabIndex        =   218
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   10
            Left            =   1440
            TabIndex        =   220
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(10)"
            BuddyDispid     =   196635
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
            TabIndex        =   221
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(10)"
            BuddyDispid     =   196636
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
         Caption         =   "Hz"
         Height          =   615
         Index           =   10
         Left            =   9360
         TabIndex        =   215
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   216
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   10
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   214
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   10
         Left            =   1560
         TabIndex        =   213
         Top             =   80
         Width           =   255
      End
      Begin VB.Label labNote 
         Caption         =   "No Note"
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
         Index           =   10
         Left            =   240
         TabIndex        =   230
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   9
      Left            =   0
      TabIndex        =   193
      Top             =   5880
      Width           =   10335
      Begin VB.Frame frameCentNote 
         Caption         =   "Cent"
         Height          =   615
         Index           =   9
         Left            =   5880
         TabIndex        =   205
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   9
            Left            =   1560
            TabIndex        =   207
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   9
            Left            =   720
            TabIndex        =   206
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   9
            Left            =   1320
            TabIndex        =   208
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
         Begin VB.Label labAbsCentNote 
            Caption         =   "absolute"
            Height          =   255
            Index           =   9
            Left            =   2400
            TabIndex        =   210
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labRelCentNote 
            Caption         =   "relative"
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   209
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   9
         Left            =   5520
         TabIndex        =   204
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   9
         Left            =   3000
         TabIndex        =   203
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   9
         Left            =   3360
         TabIndex        =   198
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   9
            Left            =   1080
            TabIndex        =   200
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   9
            Left            =   600
            TabIndex        =   199
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   9
            Left            =   1440
            TabIndex        =   201
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(9)"
            BuddyDispid     =   196635
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
            TabIndex        =   202
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(9)"
            BuddyDispid     =   196636
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
         Caption         =   "Hz"
         Height          =   615
         Index           =   9
         Left            =   9360
         TabIndex        =   196
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   9
            Left            =   120
            TabIndex        =   197
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   9
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   195
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   9
         Left            =   1560
         TabIndex        =   194
         Top             =   80
         Width           =   255
      End
      Begin VB.Label labNote 
         Caption         =   "No Note"
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
         Index           =   9
         Left            =   240
         TabIndex        =   211
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   8
      Left            =   0
      TabIndex        =   174
      Top             =   5280
      Width           =   10335
      Begin VB.Frame frameCentNote 
         Caption         =   "Cent"
         Height          =   615
         Index           =   8
         Left            =   5880
         TabIndex        =   186
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   8
            Left            =   1560
            TabIndex        =   188
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   8
            Left            =   720
            TabIndex        =   187
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   8
            Left            =   1320
            TabIndex        =   189
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
         Begin VB.Label labAbsCentNote 
            Caption         =   "absolute"
            Height          =   255
            Index           =   8
            Left            =   2400
            TabIndex        =   191
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labRelCentNote 
            Caption         =   "relative"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   190
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   8
         Left            =   5520
         TabIndex        =   185
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   8
         Left            =   3000
         TabIndex        =   184
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   8
         Left            =   3360
         TabIndex        =   179
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   8
            Left            =   1080
            TabIndex        =   181
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   8
            Left            =   600
            TabIndex        =   180
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   8
            Left            =   1440
            TabIndex        =   182
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(8)"
            BuddyDispid     =   196635
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
            TabIndex        =   183
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(8)"
            BuddyDispid     =   196636
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
         Caption         =   "Hz"
         Height          =   615
         Index           =   8
         Left            =   9360
         TabIndex        =   177
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   178
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   8
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   176
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   8
         Left            =   1560
         TabIndex        =   175
         Top             =   80
         Width           =   255
      End
      Begin VB.Label labNote 
         Caption         =   "No Note"
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
         Index           =   8
         Left            =   240
         TabIndex        =   192
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   7
      Left            =   0
      TabIndex        =   155
      Top             =   4680
      Width           =   10335
      Begin VB.Frame frameCentNote 
         Caption         =   "Cent"
         Height          =   615
         Index           =   7
         Left            =   5880
         TabIndex        =   167
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   7
            Left            =   1560
            TabIndex        =   169
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   7
            Left            =   720
            TabIndex        =   168
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   7
            Left            =   1320
            TabIndex        =   170
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
         Begin VB.Label labAbsCentNote 
            Caption         =   "absolute"
            Height          =   255
            Index           =   7
            Left            =   2400
            TabIndex        =   172
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labRelCentNote 
            Caption         =   "relative"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   171
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   7
         Left            =   5520
         TabIndex        =   166
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   7
         Left            =   3000
         TabIndex        =   165
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   7
         Left            =   3360
         TabIndex        =   160
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   162
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   7
            Left            =   600
            TabIndex        =   161
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   7
            Left            =   1440
            TabIndex        =   163
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(7)"
            BuddyDispid     =   196635
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
            TabIndex        =   164
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(7)"
            BuddyDispid     =   196636
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
         Caption         =   "Hz"
         Height          =   615
         Index           =   7
         Left            =   9360
         TabIndex        =   158
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   7
            Left            =   120
            TabIndex        =   159
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   7
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   157
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   7
         Left            =   1560
         TabIndex        =   156
         Top             =   80
         Width           =   255
      End
      Begin VB.Label labNote 
         Caption         =   "No Note"
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
         Index           =   7
         Left            =   240
         TabIndex        =   173
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   6
      Left            =   0
      TabIndex        =   136
      Top             =   4080
      Width           =   10335
      Begin VB.Frame frameCentNote 
         Caption         =   "Cent"
         Height          =   615
         Index           =   6
         Left            =   5880
         TabIndex        =   148
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   6
            Left            =   1560
            TabIndex        =   150
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   6
            Left            =   720
            TabIndex        =   149
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   6
            Left            =   1320
            TabIndex        =   151
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
         Begin VB.Label labAbsCentNote 
            Caption         =   "absolute"
            Height          =   255
            Index           =   6
            Left            =   2400
            TabIndex        =   153
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labRelCentNote 
            Caption         =   "relative"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   152
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   6
         Left            =   5520
         TabIndex        =   147
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   6
         Left            =   3000
         TabIndex        =   146
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   6
         Left            =   3360
         TabIndex        =   141
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   6
            Left            =   1080
            TabIndex        =   143
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   6
            Left            =   600
            TabIndex        =   142
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   6
            Left            =   1440
            TabIndex        =   144
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(6)"
            BuddyDispid     =   196635
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
            TabIndex        =   145
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(6)"
            BuddyDispid     =   196636
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
         Caption         =   "Hz"
         Height          =   615
         Index           =   6
         Left            =   9360
         TabIndex        =   139
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   6
            Left            =   120
            TabIndex        =   140
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   6
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   138
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   6
         Left            =   1560
         TabIndex        =   137
         Top             =   80
         Width           =   255
      End
      Begin VB.Label labNote 
         Caption         =   "No Note"
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
         Index           =   6
         Left            =   240
         TabIndex        =   154
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   5
      Left            =   0
      TabIndex        =   117
      Top             =   3480
      Width           =   10335
      Begin VB.Frame frameCentNote 
         Caption         =   "Cent"
         Height          =   615
         Index           =   5
         Left            =   5880
         TabIndex        =   129
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   5
            Left            =   1560
            TabIndex        =   131
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   5
            Left            =   720
            TabIndex        =   130
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   5
            Left            =   1320
            TabIndex        =   132
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
         Begin VB.Label labAbsCentNote 
            Caption         =   "absolute"
            Height          =   255
            Index           =   5
            Left            =   2400
            TabIndex        =   134
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labRelCentNote 
            Caption         =   "relative"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   133
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   5
         Left            =   5520
         TabIndex        =   128
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   5
         Left            =   3000
         TabIndex        =   127
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   5
         Left            =   3360
         TabIndex        =   122
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   124
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   5
            Left            =   600
            TabIndex        =   123
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   5
            Left            =   1440
            TabIndex        =   125
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(5)"
            BuddyDispid     =   196635
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
            TabIndex        =   126
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(5)"
            BuddyDispid     =   196636
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
         Caption         =   "Hz"
         Height          =   615
         Index           =   5
         Left            =   9360
         TabIndex        =   120
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   121
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   5
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   119
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   5
         Left            =   1560
         TabIndex        =   118
         Top             =   80
         Width           =   255
      End
      Begin VB.Label labNote 
         Caption         =   "No Note"
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
         Index           =   5
         Left            =   240
         TabIndex        =   135
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   4
      Left            =   0
      TabIndex        =   98
      Top             =   2880
      Width           =   10335
      Begin VB.Frame frameCentNote 
         Caption         =   "Cent"
         Height          =   615
         Index           =   4
         Left            =   5880
         TabIndex        =   110
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   4
            Left            =   1560
            TabIndex        =   112
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   4
            Left            =   720
            TabIndex        =   111
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   4
            Left            =   1320
            TabIndex        =   113
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
         Begin VB.Label labAbsCentNote 
            Caption         =   "absolute"
            Height          =   255
            Index           =   4
            Left            =   2400
            TabIndex        =   115
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labRelCentNote 
            Caption         =   "relative"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   114
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   4
         Left            =   5520
         TabIndex        =   109
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   4
         Left            =   3000
         TabIndex        =   108
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   4
         Left            =   3360
         TabIndex        =   103
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   105
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   4
            Left            =   600
            TabIndex        =   104
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   4
            Left            =   1440
            TabIndex        =   106
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(4)"
            BuddyDispid     =   196635
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
            TabIndex        =   107
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(4)"
            BuddyDispid     =   196636
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
         Caption         =   "Hz"
         Height          =   615
         Index           =   4
         Left            =   9360
         TabIndex        =   101
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   4
            Left            =   120
            TabIndex        =   102
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   4
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   100
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   4
         Left            =   1560
         TabIndex        =   99
         Top             =   80
         Width           =   255
      End
      Begin VB.Label labNote 
         Caption         =   "No Note"
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
         Index           =   4
         Left            =   240
         TabIndex        =   116
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   3
      Left            =   0
      TabIndex        =   79
      Top             =   2280
      Width           =   10335
      Begin VB.Frame frameCentNote 
         Caption         =   "Cent"
         Height          =   615
         Index           =   3
         Left            =   5880
         TabIndex        =   91
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   3
            Left            =   1560
            TabIndex        =   93
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   3
            Left            =   720
            TabIndex        =   92
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   3
            Left            =   1320
            TabIndex        =   94
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
         Begin VB.Label labAbsCentNote 
            Caption         =   "absolute"
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   96
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labRelCentNote 
            Caption         =   "relative"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   95
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   3
         Left            =   5520
         TabIndex        =   90
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   3
         Left            =   3000
         TabIndex        =   89
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   3
         Left            =   3360
         TabIndex        =   84
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   86
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   3
            Left            =   600
            TabIndex        =   85
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   3
            Left            =   1440
            TabIndex        =   87
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(3)"
            BuddyDispid     =   196635
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
            TabIndex        =   88
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(3)"
            BuddyDispid     =   196636
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
         Caption         =   "Hz"
         Height          =   615
         Index           =   3
         Left            =   9360
         TabIndex        =   82
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   3
            Left            =   120
            TabIndex        =   83
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   3
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   81
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   3
         Left            =   1560
         TabIndex        =   80
         Top             =   80
         Width           =   255
      End
      Begin VB.Label labNote 
         Caption         =   "No Note"
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
         Index           =   3
         Left            =   240
         TabIndex        =   97
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   2
      Left            =   0
      TabIndex        =   60
      Top             =   1680
      Width           =   10335
      Begin VB.Frame frameCentNote 
         Caption         =   "Cent"
         Height          =   615
         Index           =   2
         Left            =   5880
         TabIndex        =   72
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   74
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   720
            TabIndex        =   73
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   2
            Left            =   1320
            TabIndex        =   75
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
         Begin VB.Label labAbsCentNote 
            Caption         =   "absolute"
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   77
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labRelCentNote 
            Caption         =   "relative"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   76
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   2
         Left            =   5520
         TabIndex        =   71
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   2
         Left            =   3000
         TabIndex        =   70
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   2
         Left            =   3360
         TabIndex        =   65
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   67
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   630
            TabIndex        =   66
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   2
            Left            =   1440
            TabIndex        =   68
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(2)"
            BuddyDispid     =   196635
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
            TabIndex        =   69
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(2)"
            BuddyDispid     =   196636
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
         Caption         =   "Hz"
         Height          =   615
         Index           =   2
         Left            =   9360
         TabIndex        =   63
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   64
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   2
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   62
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   2
         Left            =   1560
         TabIndex        =   61
         Top             =   80
         Width           =   255
      End
      Begin VB.Label labNote 
         Caption         =   "No Note"
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
         Left            =   240
         TabIndex        =   78
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   1
      Left            =   0
      TabIndex        =   41
      Top             =   1080
      Width           =   10335
      Begin VB.Frame frameCentNote 
         Caption         =   "Cent"
         Height          =   615
         Index           =   1
         Left            =   5880
         TabIndex        =   53
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   1560
            TabIndex        =   55
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   720
            TabIndex        =   54
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   56
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
         Begin VB.Label labAbsCentNote 
            Caption         =   "absolute"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   58
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labRelCentNote 
            Caption         =   "relative"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   57
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   1
         Left            =   5520
         TabIndex        =   52
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optRatioNote 
         Height          =   495
         Index           =   1
         Left            =   3000
         TabIndex        =   51
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameRatioNote 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   1
         Left            =   3360
         TabIndex        =   46
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   48
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   630
            TabIndex        =   47
            Text            =   "1"
            Top             =   240
            Width           =   150
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   1
            Left            =   1440
            TabIndex        =   49
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(1)"
            BuddyDispid     =   196635
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
            TabIndex        =   50
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(1)"
            BuddyDispid     =   196636
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
         Caption         =   "Hz"
         Height          =   615
         Index           =   1
         Left            =   9360
         TabIndex        =   44
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   1
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   43
         Top             =   80
         Width           =   495
      End
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   1
         Left            =   1560
         TabIndex        =   42
         Top             =   80
         Width           =   255
      End
      Begin VB.Label labNote 
         Caption         =   "No Note"
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
         Left            =   240
         TabIndex        =   59
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameNote 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   22
      Top             =   480
      Width           =   10335
      Begin VB.CheckBox chkNote 
         Height          =   495
         Index           =   0
         Left            =   1560
         TabIndex        =   39
         Top             =   80
         Width           =   255
      End
      Begin VB.PictureBox picColorNote 
         Height          =   495
         Index           =   0
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   38
         Top             =   80
         Width           =   495
      End
      Begin VB.Frame frameHzNote 
         Caption         =   "Hz"
         Height          =   615
         Index           =   0
         Left            =   9360
         TabIndex        =   36
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame frameRatioNote 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   0
         Left            =   3360
         TabIndex        =   31
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioNumeratorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   600
            TabIndex        =   33
            Text            =   "1"
            Top             =   240
            Width           =   165
         End
         Begin VB.TextBox txtRatioDenominatorNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   32
            Text            =   "1"
            Top             =   240
            Width           =   165
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorNote 
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   34
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtRatioDenominatorNote(0)"
            BuddyDispid     =   196635
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
            TabIndex        =   35
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorNote(0)"
            BuddyDispid     =   196636
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
         Left            =   3000
         TabIndex        =   30
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optCentNote 
         Height          =   495
         Index           =   0
         Left            =   5520
         TabIndex        =   29
         Top             =   80
         Width           =   255
      End
      Begin VB.Frame frameCentNote 
         Caption         =   "Cent"
         Height          =   615
         Index           =   0
         Left            =   5880
         TabIndex        =   23
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtRelCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   25
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin VB.TextBox txtAbsCentNote 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   24
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin MSComCtl2.UpDown UpDownCentNote 
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   26
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
         Begin VB.Label labRelCentNote 
            Caption         =   "relative"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labAbsCentNote 
            Caption         =   "absolute"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   27
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.Label labNote 
         Caption         =   "No Note"
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
         Left            =   240
         TabIndex        =   40
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.Frame frameDrone 
      BorderStyle     =   0  'Kein
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   7320
      Width           =   10335
      Begin VB.Frame frameCentDrone 
         Caption         =   "Cent"
         Height          =   615
         Index           =   0
         Left            =   5880
         TabIndex        =   16
         Top             =   0
         Width           =   3135
         Begin VB.TextBox txtAbsCentDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   18
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtRelCentDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   720
            TabIndex        =   17
            Text            =   "xxx.x"
            Top             =   240
            Width           =   615
         End
         Begin MSComCtl2.UpDown UpDownCentDrone 
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   19
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
            Max             =   12000
            Min             =   -12000
            Wrap            =   -1  'True
            Enabled         =   -1  'True
         End
         Begin VB.Label labAbsCentDrone 
            Caption         =   "absolute"
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   21
            Top             =   280
            Width           =   615
         End
         Begin VB.Label labRelCentDrone 
            Caption         =   "relative"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   20
            Top             =   280
            Width           =   615
         End
      End
      Begin VB.CheckBox chkDrone 
         Height          =   495
         Index           =   0
         Left            =   1560
         TabIndex        =   14
         Top             =   80
         Width           =   255
      End
      Begin VB.PictureBox picColorDrone 
         Height          =   495
         Index           =   0
         Left            =   2160
         ScaleHeight     =   435
         ScaleWidth      =   435
         TabIndex        =   13
         Top             =   80
         Width           =   495
      End
      Begin VB.Frame frameHzDrone 
         Caption         =   "Hz"
         Height          =   615
         Index           =   0
         Left            =   9360
         TabIndex        =   11
         Top             =   0
         Width           =   855
         Begin VB.TextBox txtHzDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   12
            Text            =   "xxxx.x"
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame frameRatioDrone 
         Caption         =   "Ratio"
         Height          =   615
         Index           =   0
         Left            =   3360
         TabIndex        =   6
         Top             =   0
         Width           =   1815
         Begin VB.TextBox txtRatioNumeratorDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   600
            TabIndex        =   8
            Text            =   "1"
            Top             =   240
            Width           =   165
         End
         Begin VB.TextBox txtRatioDenominatorDrone 
            Alignment       =   1  'Rechts
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   7
            Text            =   "1"
            Top             =   240
            Width           =   165
         End
         Begin MSComCtl2.UpDown UpDownRatioDenominatorDrone 
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   9
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtRatioDenominatorDrone(0)"
            BuddyDispid     =   196613
            BuddyIndex      =   0
            OrigLeft        =   1440
            OrigTop         =   240
            OrigRight       =   1695
            OrigBottom      =   525
            Max             =   9
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDownRatioNumeratorDrone 
            Height          =   285
            Index           =   0
            Left            =   119
            TabIndex        =   10
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   503
            _Version        =   393216
            Value           =   1
            Alignment       =   0
            BuddyControl    =   "txtRatioNumeratorDrone(0)"
            BuddyDispid     =   196614
            BuddyIndex      =   0
            OrigLeft        =   480
            OrigTop         =   240
            OrigRight       =   735
            OrigBottom      =   525
            Max             =   9
            Min             =   1
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Line lineRatioDrone 
            Index           =   0
            X1              =   840
            X2              =   960
            Y1              =   480
            Y2              =   240
         End
      End
      Begin VB.OptionButton optRatioDrone 
         Height          =   495
         Index           =   0
         Left            =   3000
         TabIndex        =   5
         Top             =   80
         Width           =   255
      End
      Begin VB.OptionButton optCentDrone 
         Height          =   495
         Index           =   0
         Left            =   5520
         TabIndex        =   4
         Top             =   80
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
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   160
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   9360
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   9360
      Width           =   2775
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   9360
      Width           =   2775
   End
   Begin MSComDlg.CommonDialog objStdDialog 
      Left            =   9960
      Top             =   9360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000011&
      X1              =   0
      X2              =   10920
      Y1              =   9240
      Y2              =   9240
   End
   Begin VB.Line Line2 
      Index           =   0
      X1              =   10440
      X2              =   10440
      Y1              =   120
      Y2              =   9960
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
Attribute VB_Name = "frmConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private i As Integer

Private Type Temperament
    blnCheck As Boolean
    intNumerator As Integer
    intDenominator As Integer
    dblRatio As Double
    dblAbsoluteCents As Double
    dblRelativeCents As Double
End Type

Private mReadWriteColors(0 To 10) As Long

Private mudtNotes(0 To 10) As Temperament
Private mudtDrones(0 To 2) As Temperament

Private mintAllNotes(0 To 10) As Integer
Private mintAllDrones(0 To 2) As Integer

Private mudtAllNotes(0 To 10) As NoteAttributes
Private mudtAllDrones(0 To 2) As NoteAttributes

Private Sub Form_load()

    frmConfiguration.Caption = "Scale Configuration: Chanter / Drones"

    mintAllNotes(0) = 11
    mintAllNotes(1) = 10
    mintAllNotes(2) = 9
    mintAllNotes(3) = 8
    mintAllNotes(4) = 7
    mintAllNotes(5) = 6
    mintAllNotes(6) = 5
    mintAllNotes(7) = 4
    mintAllNotes(8) = 3
    mintAllNotes(9) = 2
    mintAllNotes(10) = 1
    
    mintAllDrones(0) = 3
    mintAllDrones(1) = 2
    mintAllDrones(2) = 1
    
    
    For i = LBound(mintAllNotes) To UBound(mintAllNotes)
    'For i = LBound(mudtAllNotes) To UBound(mudtAllNotes)
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
        
        picColorNote(i).BackColor = udtNoteDefaults(mintAllNotes(i)).Color
        labNote(i).Caption = udtNoteDefaults(mintAllNotes(i)).Name
        optRatioNote(i).Value = True
    Next i
    
    For i = LBound(mintAllDrones) To UBound(mintAllDrones)
    'For i = LBound(mudtAllDrones) To UBound(mudtAllDrones)
        UpDownRatioNumeratorDrone(i).Wrap = False
        UpDownRatioDenominatorDrone(i).Wrap = False
        'UpDownRatioNumeratorDrone(i).BuddyControl = txtRatioNumeratorDrone(i)
        'UpDownRatioDenominatorDrone(i).BuddyControl = txtRatioDenominatorDrone(i)
        'UpDownRatioNumeratorDrone(i).BuddyProperty = "Text"
        'UpDownRatioDenominatorDrone(i).BuddyProperty = "Text"
        'UpDownRatioNumeratorDrone(i).SyncBuddy = True
        'UpDownRatioDenominatorDrone(i).SyncBuddy = True
        UpDownRatioNumeratorDrone(i).Min = 0
        UpDownRatioDenominatorDrone(i).Min = 1
        UpDownRatioNumeratorDrone(i).Max = 99
        UpDownRatioDenominatorDrone(i).Max = 99
        
        'UpDownRatioNumeratorDrone(i).Left = 120
        'UpDownRatioNumeratorDrone(i).Top = 240
        'UpDownRatioNumeratorDrone(i).Width = 255
        'UpDownRatioNumeratorDrone(i).Height = 285
        'UpDownRatioDenominatorDrone(i).Left = 1440
        'UpDownRatioDenominatorDrone(i).Top = 240
        'UpDownRatioDenominatorDrone(i).Width = 255
        'UpDownRatioDenominatorDrone(i).Height = 285
        txtRatioNumeratorDrone(i).Left = 360
        txtRatioNumeratorDrone(i).Top = 240
        txtRatioNumeratorDrone(i).Width = 360
        txtRatioNumeratorDrone(i).Height = 285
        txtRatioDenominatorDrone(i).Left = 1080
        txtRatioDenominatorDrone(i).Top = 240
        txtRatioDenominatorDrone(i).Width = 360
        txtRatioDenominatorDrone(i).Height = 285
        
        picColorDrone(i).BackColor = udtDroneDefaults(mintAllDrones(i)).Color
        labDrone(i).Caption = udtDroneDefaults(mintAllDrones(i)).Name
        optRatioDrone(i).Value = True
    Next i
    
    UpDownRatioNumeratorNote(0).Value = 2
    UpDownRatioDenominatorNote(0).Value = 1
    UpDownRatioNumeratorNote(1).Value = 7
    UpDownRatioDenominatorNote(1).Value = 4
    UpDownRatioNumeratorNote(2).Value = 5
    UpDownRatioDenominatorNote(2).Value = 3
    UpDownRatioNumeratorNote(3).Value = 8
    UpDownRatioDenominatorNote(3).Value = 5
    UpDownRatioNumeratorNote(4).Value = 3
    UpDownRatioDenominatorNote(4).Value = 2
    UpDownRatioNumeratorNote(5).Value = 4
    UpDownRatioDenominatorNote(5).Value = 3
    UpDownRatioNumeratorNote(6).Value = 5
    UpDownRatioDenominatorNote(6).Value = 4
    UpDownRatioNumeratorNote(7).Value = 6
    UpDownRatioDenominatorNote(7).Value = 5
    UpDownRatioNumeratorNote(8).Value = 9
    UpDownRatioDenominatorNote(8).Value = 8
    UpDownRatioNumeratorNote(9).Value = 1
    UpDownRatioDenominatorNote(9).Value = 1
    UpDownRatioNumeratorNote(10).Value = 7
    UpDownRatioDenominatorNote(10).Value = 8
    
    UpDownRatioNumeratorDrone(0).Value = 1
    UpDownRatioDenominatorDrone(0).Value = 2
    UpDownRatioNumeratorDrone(1).Value = 3
    UpDownRatioDenominatorDrone(1).Value = 8
    UpDownRatioNumeratorDrone(2).Value = 1
    UpDownRatioDenominatorDrone(2).Value = 4
   
    For i = LBound(mintAllNotes) To UBound(mintAllNotes)
        chkNote(i).Value = 1
    Next i
    chkNote(3).Value = 0
    chkNote(7).Value = 0
    
    chkDrone(0).Value = 1
    chkDrone(2).Value = 1
    
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
    mudtAllNotes(Index).RelativeCents = UpDownCentNote(Index).Value / 10
    mudtAllNotes(Index).AbsoluteCents = mudtAllNotes(Index).RelativeCents + udtNoteChromatic(mintAllNotes(Index)).AbsoluteCents
    mudtAllNotes(Index).Ratio = 2 ^ (mudtAllNotes(Index).AbsoluteCents / 1200)
    txtRelCentNote(Index).Text = Format(mudtAllNotes(Index).RelativeCents, "##0.0")
    txtAbsCentNote(Index).Text = Format(mudtAllNotes(Index).AbsoluteCents, "####0.0")
    txtHzNote(Index).Text = Format(mudtAllNotes(Index).Ratio * gdblReferenceFrequency, "###0.0")
End Sub
Private Sub updownCentDrone_Change(Index As Integer)
    If optCentDrone(Index).Value = False Then Exit Sub
    mudtAllDrones(Index).RelativeCents = UpDownCentDrone(Index).Value / 10
    mudtAllDrones(Index).AbsoluteCents = mudtAllDrones(Index).RelativeCents + udtDroneDefaults(mintAllDrones(Index)).AbsoluteCents
    mudtAllDrones(Index).Ratio = 2 ^ (mudtAllDrones(Index).AbsoluteCents / 1200)
    txtRelCentDrone(Index).Text = Format(mudtAllDrones(Index).RelativeCents, "##0.0")
    txtAbsCentDrone(Index).Text = Format(mudtAllDrones(Index).AbsoluteCents, "####0.0")
    txtHzDrone(Index).Text = Format(mudtAllDrones(Index).Ratio * gdblReferenceFrequency, "###0.0")
End Sub
Private Sub UpDownRatioNumeratorNote_Change(Index As Integer)
    If optRatioNote(Index).Value = False Then Exit Sub
    txtRatioNumeratorNote(Index).Enabled = True
    txtRatioDenominatorNote(Index).Enabled = True
    Call optRatioNote_Click(Index)
End Sub
Private Sub UpDownRatioNumeratorDrone_Change(Index As Integer)
    If optRatioNote(Index).Value = False Then Exit Sub
    txtRatioNumeratorDrone(Index).Enabled = True
    txtRatioDenominatorDrone(Index).Enabled = True
    Call optRatioDrone_Click(Index)
End Sub
Private Sub UpDownRatioDenominatorNote_Change(Index As Integer)
    If optRatioNote(Index).Value = False Then Exit Sub
    txtRatioNumeratorNote(Index).Enabled = True
    txtRatioDenominatorNote(Index).Enabled = True
    Call optRatioNote_Click(Index)
End Sub
Private Sub UpDownRatioDenominatorDrone_Change(Index As Integer)
    If optRatioNote(Index).Value = False Then Exit Sub
    txtRatioNumeratorDrone(Index).Enabled = True
    txtRatioDenominatorDrone(Index).Enabled = True
    Call optRatioDrone_Click(Index)
End Sub
Private Sub optRatioNote_Click(Index As Integer)
    
    'optRatioNote(Index).Value = True
    UpDownCentNote(Index).Enabled = False
    txtAbsCentNote(Index).Enabled = False
    txtRelCentNote(Index).Enabled = False
    txtRatioNumeratorNote(Index).Enabled = True
    txtRatioDenominatorNote(Index).Enabled = True
    
    mudtAllNotes(Index).Numerator = UpDownRatioNumeratorNote(Index).Value
    mudtAllNotes(Index).Denominator = UpDownRatioDenominatorNote(Index).Value
    mudtAllNotes(Index).Ratio = mudtAllNotes(Index).Numerator / mudtAllNotes(Index).Denominator
    mudtAllNotes(Index).AbsoluteCents = ConvertFrequencyInCent(1, mudtAllNotes(Index).Ratio)
    mudtAllNotes(Index).RelativeCents = mudtAllNotes(Index).AbsoluteCents - _
                       udtNoteChromatic(mintAllNotes(Index)).AbsoluteCents
                    
    txtRelCentNote(Index).Text = Format(mudtAllNotes(Index).RelativeCents, "##0.0")
    txtAbsCentNote(Index).Text = Format(mudtAllNotes(Index).AbsoluteCents, "####0.0")
    txtHzNote(Index).Text = Format(mudtAllNotes(Index).Ratio * gdblReferenceFrequency, "###0.0")

End Sub
Private Sub optRatioDrone_Click(Index As Integer)
    
    'optRatioDrone(Index).Value = True
    UpDownCentDrone(Index).Enabled = False
    txtAbsCentDrone(Index).Enabled = False
    txtRelCentDrone(Index).Enabled = False
    txtRatioNumeratorDrone(Index).Enabled = True
    txtRatioDenominatorDrone(Index).Enabled = True
    
    mudtAllDrones(Index).Numerator = UpDownRatioNumeratorDrone(Index).Value
    mudtAllDrones(Index).Denominator = UpDownRatioDenominatorDrone(Index).Value
    mudtAllDrones(Index).Ratio = mudtAllDrones(Index).Numerator / mudtAllDrones(Index).Denominator
    mudtAllDrones(Index).AbsoluteCents = ConvertFrequencyInCent(1, mudtAllDrones(Index).Ratio)
    mudtAllDrones(Index).RelativeCents = mudtAllDrones(Index).AbsoluteCents - _
                        udtDroneDefaults(mintAllDrones(Index)).AbsoluteCents
                           
    txtRelCentDrone(Index).Text = Format(mudtAllDrones(Index).RelativeCents, "##0.0")
    txtAbsCentDrone(Index).Text = Format(mudtAllDrones(Index).AbsoluteCents, "####0.0")
    txtHzDrone(Index).Text = Format(mudtAllDrones(Index).Ratio * gdblReferenceFrequency, "###0.0")

End Sub
Private Sub optCentNote_Click(Index As Integer)
    'optCentNote(Index).Value = True
    txtRatioNumeratorNote(Index).Enabled = False
    txtRatioDenominatorNote(Index).Enabled = False
    txtAbsCentNote(Index).Enabled = True
    txtRelCentNote(Index).Enabled = True
    UpDownCentNote(Index).Enabled = True
    UpDownCentNote(Index).Value = 10 * Round(mudtAllNotes(Index).RelativeCents)
End Sub
Private Sub optCentDrone_Click(Index As Integer)
    'optCentDrone(Index).Value = True
    txtRatioNumeratorDrone(Index).Enabled = False
    txtRatioDenominatorDrone(Index).Enabled = False
    'txtAbsCentDrone(Index).Enabled = True
    txtRelCentDrone(Index).Enabled = True
    UpDownCentDrone(Index).Enabled = True
    UpDownCentDrone(Index).Value = 10 * (Round(mudtAllDrones(Index).RelativeCents))
End Sub
Private Sub cmdApply_Click()
    
    Dim i As Integer
    Dim Index As Integer
    Dim NoteIndex() As Integer
    Dim intCounter As Integer
    
    ReDim NoteIndex(0 To UBound(mintAllNotes))
    
    intCounter = 0
    
    For Index = LBound(mintAllNotes) To UBound(mintAllNotes)
        udtNoteDefaults(mintAllNotes(Index)).Numerator = mudtAllNotes(Index).Numerator
        udtNoteDefaults(mintAllNotes(Index)).Denominator = mudtAllNotes(Index).Denominator
        udtNoteDefaults(mintAllNotes(Index)).Ratio = mudtAllNotes(Index).Ratio
        udtNoteDefaults(mintAllNotes(Index)).AbsoluteCents = mudtAllNotes(Index).AbsoluteCents
        udtNoteDefaults(mintAllNotes(Index)).RelativeCents = mudtAllNotes(Index).RelativeCents
        udtNoteDefaults(mintAllNotes(Index)).Color = picColorNote(Index).BackColor
        udtNoteDefaults(mintAllNotes(Index)).Index = mintAllNotes(Index)
        If chkNote(Index).Value = 1 Then
            intCounter = intCounter + 1
            NoteIndex(intCounter) = mintAllNotes(Index)
            Debug.Print intCounter, NoteIndex(intCounter)
        End If
        'Debug.Print intCounter, NoteIndex(intCounter)
    Next Index
    Debug.Print "-"
    
    ReDim Preserve NoteIndex(0 To intCounter)
   
    
    ReDim udtNote(LBound(NoteIndex) To UBound(NoteIndex))
    ReDim gobjChanterBufferCent(LBound(NoteIndex) To UBound(NoteIndex))
    
    For Index = LBound(NoteIndex) + 1 To UBound(NoteIndex)
        udtNote(UBound(NoteIndex) - Index + 1).Name = udtNoteDefaults(NoteIndex(Index)).Name
        udtNote(UBound(NoteIndex) - Index + 1).Numerator = udtNoteDefaults(NoteIndex(Index)).Numerator
        udtNote(UBound(NoteIndex) - Index + 1).Denominator = udtNoteDefaults(NoteIndex(Index)).Denominator
        udtNote(UBound(NoteIndex) - Index + 1).Ratio = udtNoteDefaults(NoteIndex(Index)).Ratio
        udtNote(UBound(NoteIndex) - Index + 1).AbsoluteCents = udtNoteDefaults(NoteIndex(Index)).AbsoluteCents
        udtNote(UBound(NoteIndex) - Index + 1).RelativeCents = udtNoteDefaults(NoteIndex(Index)).RelativeCents
        udtNote(UBound(NoteIndex) - Index + 1).Color = udtNoteDefaults(NoteIndex(Index)).Color
        udtNote(Index).Index = Index
        Debug.Print Index, NoteIndex(Index)
    Next Index
    Debug.Print "--"
    
    For Index = LBound(NoteIndex) To UBound(NoteIndex)
        Debug.Print udtNote(Index).Index, udtNote(Index).Name, udtNote(Index).Ratio, _
                    udtNote(Index).RelativeCents, udtNote(Index).AbsoluteCents, udtNote(Index).Color
    Next Index
    Debug.Print "---"
    
    For i = 1 To UBound(NoteIndex)
        Notes.Color(i) = udtNote(i).Color
    Next i
    
    
    udtNote(1).Tolerance(-1) = -100
    For i = LBound(udtNote) + 1 To UBound(udtNote) - 1
        udtNote(i + 1).Tolerance(-1) = (udtNote(i).AbsoluteCents - udtNote(i + 1).AbsoluteCents) / 2
        udtNote(i).Tolerance(1) = (udtNote(i + 1).AbsoluteCents - udtNote(i).AbsoluteCents) / 2
    Next i
    udtNote(UBound(udtNote)).Tolerance(1) = 100

End Sub

Private Sub menuSaveColors_Click(Index As Integer)

    For i = LBound(mReadWriteColors) To UBound(mReadWriteColors)
        mReadWriteColors(i) = picColorNote(i).BackColor
    Next i
    
    MsgBox (picColorNote(10).BackColor)

End Sub
