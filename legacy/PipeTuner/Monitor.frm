VERSION 5.00
Begin VB.Form Form_Monitor 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   Caption         =   "Bitmaps im »Kingsize«-Format"
   ClientHeight    =   7125
   ClientLeft      =   1245
   ClientTop       =   1890
   ClientWidth     =   14835
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Quelle
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'ZReihenfolge
   ScaleHeight     =   475
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   989
   Begin VB.HScrollBar HLauf 
      Height          =   255
      LargeChange     =   200
      Left            =   0
      Max             =   1000
      SmallChange     =   20
      TabIndex        =   1
      Top             =   6720
      Width           =   10215
   End
   Begin VB.VScrollBar VLauf 
      Height          =   6735
      LargeChange     =   200
      Left            =   10200
      Max             =   1000
      SmallChange     =   20
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Fenster 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   0
      ScaleHeight     =   447
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   679
      TabIndex        =   2
      Top             =   0
      Width           =   10215
      Begin VB.PictureBox Monitor 
         Appearance      =   0  '2D
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   6615
         Left            =   0
         ScaleHeight     =   439
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   671
         TabIndex        =   3
         Top             =   0
         Width           =   10095
      End
   End
   Begin VB.Menu MenuMain 
      Caption         =   "Grafik-Größe"
      Begin VB.Menu MenuGröße 
         Caption         =   "&300*300 Punkte"
         Index           =   1
      End
      Begin VB.Menu MenuGröße 
         Caption         =   "&600*600"
         Index           =   2
      End
      Begin VB.Menu MenuGröße 
         Caption         =   "&900*900"
         Index           =   3
      End
      Begin VB.Menu MenuGröße 
         Caption         =   "&1200*1200"
         Checked         =   -1  'True
         Index           =   4
      End
      Begin VB.Menu MenuGröße 
         Caption         =   "&1500*1500"
         Index           =   5
      End
      Begin VB.Menu MenuGröße 
         Caption         =   "&1800*1800"
         Index           =   6
      End
      Begin VB.Menu MenuGröße 
         Caption         =   "&2100*2100"
         Index           =   7
      End
      Begin VB.Menu MenuGröße 
         Caption         =   "&2400*2400"
         Index           =   8
      End
      Begin VB.Menu MenüDummy 
         Caption         =   "-"
      End
      Begin VB.Menu MenüEnde 
         Caption         =   "&Ende"
      End
   End
End
Attribute VB_Name = "Form_Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

DefLng A-Z

Public glngIndex As Long

Private min_x As Long
Private min_y As Long
Private max_x As Long
Private max_y As Long

Private max_Freq As Long
Private maxcounter As Integer

Private MonTop As Long
Private MonLeft As Long
Private MonSizeX As Long
Private MonSizeY As Long

Private mdblPowerSpectrum() As Double
Private mdblPhaseSpectrum() As Double

Private mdblMaximumSpectrum As Double

Private Const mdblTwoPi As Double = 6.28318530717959


Private Sub Form_Load()
                
    min_x = 256
    max_x = 8192
    min_y = 128
    max_y = 2048
    max_Freq = 5000
    
    mdblMaximumSpectrum = 512
    
    'Form_Monitor.Height = 256
    'Form_Monitor.Width = 512
    'Fenster.Height = 512
    'Fenster.Width = 768
    MonTop = 10
    MonLeft = 10
    MonSizeX = 1260
    MonSizeY = 900
    
    Monitor.Top = MonTop
    Monitor.Left = MonLeft
    Monitor.Width = MonSizeX
    Monitor.Height = MonSizeY
    'Zeichne_Grafik

    'Call ShowSpectrum
        
End Sub
' Veränderung der Fenstergröße: testen, ob Bildlaufleisten
' überhaupt noch erforderlich sind; Größe des Steuerelemente
' an Fesntergröße anpassen
Private Sub Form_Resize()
  If ScaleWidth >= MonSizeX Then
    HLauf.Visible = False
  Else
    HLauf.Visible = True
  End If
  If ScaleHeight >= MonSizeY Then
    VLauf.Visible = False
  Else
    VLauf.Visible = True
  End If
  B = ScaleWidth: h = ScaleHeight
  If VLauf.Visible Then B = B - VLauf.Width
  If HLauf.Visible Then h = h - HLauf.Height
  Fenster.Width = B: Fenster.Height = h
  HLauf.Move 0, h, B, HLauf.Height
  VLauf.Move B, 0, VLauf.Width, h
  Ausschnitt_Wählen  'sichtbaren Bildausschnitt verändern
End Sub
' Programmende
Private Sub MenüEnde_Click()
  End
End Sub
' Veränderung der Bildlaufleisten
Private Sub HLauf_Change()
  Ausschnitt_Wählen
End Sub
Private Sub HLauf_Scroll()
  Ausschnitt_Wählen
End Sub
Private Sub VLauf_Change()
  Ausschnitt_Wählen
End Sub
Private Sub VLauf_Scroll()
  Ausschnitt_Wählen
End Sub
' Veränderung der Grafikgröße (mit 'intelligenter'
' Fehlerroutine)
Private Sub MenuGröße_Click(index As Integer)
Noch_Ein_Versuch:
  For i = 1 To 8: MenuGröße(i).Checked = 0: Next
  MenuGröße(index).Checked = -1
  MonSizeX = index * 300: MonSizeY = MonSizeX
  Monitor.Cls
  On Error GoTo Zu_Groß
  Monitor.Width = MonSizeX: Monitor.Height = MonSizeY
  Form_Resize
  Zeichne_Grafik
  Exit Sub
Zu_Groß:
  If index = 1 Then
    MsgBox "Viel zu wenig Speicher, Ende!"
    End
  Else
    MsgBox "Zu wenig Speicher, die Grafikgröße wird reduziert!"
    index = index - 1
    Resume Noch_Ein_Versuch
  End If
End Sub
' Grafik zeichnen (wegen AutoRedraw nur einmal bzw. nach Veränderung
' der Bildgröße notwendig)
Private Sub Zeichne_Grafik()
  Monitor.Line (0, 0)-(MonSizeX, MonSizeY)
  Monitor.Line (0, MonSizeY)-(MonSizeX, 0)
  For i = 10 To MonSizeX / 2.1 Step MonSizeX / 30
    Monitor.Circle (MonSizeX / 2, MonSizeY / 2), i
  Next i
End Sub
' Left- und Top-Koordinate des Bildfelds je nach Einstellung der
' Bildlaufleisten einstellen
Private Sub Ausschnitt_Wählen()
  If Fenster.Width > MonSizeX Then
    Monitor.Left = 0
  Else
    Monitor.Left = -(MonSizeX - Fenster.Width) * HLauf.Value / 1000
  End If
  If Fenster.Height > MonSizeY Then
    Monitor.Top = 0
  Else
    Monitor.Top = -(MonSizeY - Fenster.Height) * VLauf.Value / 1000
  End If
End Sub

Sub Monitor_KeyDown(KeyCode As Integer, Shift As Integer)

Dim wait_key_time As Single
Dim new_key_time As Single
Dim delta_key_time As Single
Dim last_key_time As Single
Dim slide_move As Integer

wait_key_time = 0.1
new_key_time = Timer
delta_key_time = new_key_time - last_key_time

If delta_key_time > wait_key_time Then
    last_key_time = Timer
     
    Select Case KeyCode
        Case vbKeyEscape: End
        Case vbKeyAdd:
            glngIndex = glngIndex + 1
            Call ShowSpectrum
            Form_Monitor.Show
        Case vbKeySubtract:
            glngIndex = glngIndex - 1
            Call ShowSpectrum
            Form_Monitor.Show
    End Select
        
End If

DoEvents

End Sub

Sub Monitor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        
        Dim max_x As Long
        Dim max_y As Long
        Dim zoom_x As Integer
        
        min_x = 256
        max_x = 8192
        min_y = 128
        max_y = 1024
        
        If Button = vbLeftButton And Shift = 0 Then
            If MonSizeX >= 2 * min_x Then MonSizeX = MonSizeX / 2
            Monitor.Width = MonSizeX
        End If
        If Button = vbRightButton And Shift = 0 Then
            If MonSizeX <= max_x / 2 Then MonSizeX = MonSizeX * 2
            Monitor.Width = MonSizeX
        End If
        If Button = vbLeftButton And Shift = 1 Then
            'If MonSizeY >= 2 * min_y Then MonSizeY = MonSizeY / 2
            'Monitor.Height = MonSizeY
            mdblMaximumSpectrum = mdblMaximumSpectrum * 2
        End If
        If Button = vbRightButton And Shift = 1 Then
            'If MonSizeY <= max_y / 2 Then MonSizeY = MonSizeY * 2
            'Monitor.Height = MonSizeY
            mdblMaximumSpectrum = mdblMaximumSpectrum / 2
        End If
        
        Call ShowSpectrum
        
End Sub

Public Sub ShowSpectrum()

    Dim i As Long
    Dim K As Long
    
    Dim istart As Long
    Dim istop As Long
    Dim delta_freq As Double
    
    Dim intNote As Integer
    Dim strNote As String
    Dim dblCent As Double
    Dim dblFreq As Double
    
    
    If Not GetData Then Exit Sub
 
    delta_freq = WavFile.SampleRate / WavFile.SampleLength
    
    
    Monitor.Cls
    Monitor.DrawWidth = 1
    Monitor.ScaleLeft = 0
    Monitor.ScaleWidth = max_Freq
    Monitor.ScaleHeight = 120
        
    istart = Int(Monitor.ScaleLeft / delta_freq)
    istop = Int(Monitor.ScaleWidth / delta_freq) - 1
    
        
    For i = istart + 1 To istop
            Monitor.Line ((i - 1) * delta_freq, Monitor.ScaleHeight / 2 - CSng(mdblPowerSpectrum(i - 1) / mdblMaximumSpectrum * 50))- _
                                 (i * delta_freq, Monitor.ScaleHeight / 2 - CSng(mdblPowerSpectrum(i) / mdblMaximumSpectrum * 50)), RGB(0, 255, 0)
            Monitor.Line ((i - 1) * delta_freq, Monitor.ScaleHeight * 3 / 4 - mdblPhaseSpectrum(i - 1))- _
                                 (i * delta_freq, Monitor.ScaleHeight * 3 / 4 - mdblPhaseSpectrum(i)), RGB(0, 255, 0)
    Next i
        
    Monitor.DrawWidth = 3
    For i = istart To istop
            Monitor.PSet (i * delta_freq, Monitor.ScaleHeight / 2 - mdblPowerSpectrum(i) / mdblMaximumSpectrum * 50), RGB(0, 255, 0)
            Monitor.PSet (i * delta_freq, Monitor.ScaleHeight * 3 / 4 - mdblPhaseSpectrum(i)), RGB(0, 255, 0)
    Next i
    
    Monitor.DrawWidth = 1
    maxcounter = Int(Monitor.ScaleWidth / gdblReferenceFrequency)
    
    'For k = 1 To Int(maxcounter / 2)
    For K = 1 To 1
    For i = 1 To 9
        Monitor.Line (gudtNotes(i).Ratio * K * gdblReferenceFrequency, 0)- _
                     (gudtNotes(i).Ratio * K * gdblReferenceFrequency, Monitor.ScaleHeight / 2), _
                      RGB(255, 255, 0)
    Next i
    Next K
    

    For i = 1 To maxcounter * 4
        Monitor.Line (i * gdblReferenceFrequency / 4, 0)- _
                     (i * gdblReferenceFrequency / 4, Monitor.ScaleHeight / 2), _
                      RGB(0, 0, 255)
    Next i
    
    For i = 1 To maxcounter * 2
        Monitor.Line (i * gdblReferenceFrequency / 2, 0)- _
                     (i * gdblReferenceFrequency / 2, Monitor.ScaleHeight / 2), _
                      RGB(255, 0, 0)
    Next i
    For i = 1 To maxcounter
        Monitor.Line (i * gdblReferenceFrequency, 0)- _
                     (i * gdblReferenceFrequency, Monitor.ScaleHeight / 2), _
                      RGB(255, 255, 255)
    Next i
                     
    'Form_Main.Show
    'Form_Main.Show
    'Form_Main.Refresh
    
    Form_Resize

End Sub

Public Function GetData() As Boolean

    Dim intData() As Integer
    Dim intNoteIndex As Integer
    Dim dblFreq As Double
    Dim dblRelCent As Double
    Dim lngReadLength As Long
    Dim sngStartTime As Single
    
    GetData = False

    ' start position = index on display + start position of buffer and compensate buffer
    WavFile.SampleStart = WavFile.SampleInterval * (glngIndex + gdblBufferFrequency.FIFO_StartPosition)
    ' add start time * sample rate
    WavFile.SampleStart = WavFile.SampleStart + (WavFile.StartTime + WavFile.SectionTimeStart) * WavFile.SampleRate
    ' store start time
    sngStartTime = WavFile.SampleStart / WavFile.SampleRate
    ' store read length
    lngReadLength = WavFile.ReadLength
    ' read length is only sample length (previos read length was size of buffer)
    WavFile.ReadLength = WavFile.SampleLength
    
    If WavFile.SampleStart + WavFile.ReadLength >= WavFile.DataSamples Then Exit Function
    
    ' read wav data
    WavFile.OpenFile
    WavFile.ReadWavData
    WavFile.CloseFile
    
    ' data read
    intData = WavFile.ReadData
    
    ' calculate spectrum and phase
    FFT.SetWindowType "Gauss"
    mdblPowerSpectrum = FFT.PowerSpectrum(intData, 0)
    mdblPhaseSpectrum = FFT.PhaseSpectrum
    FrequencyDetection.Set_Spectrum mdblPowerSpectrum
    FrequencyDetection.Set_RefFreq (gdblReferenceFrequency)
    dblFreq = FrequencyDetection.MeasureChanterFrequency
    NoteBuffer.DetectNote (dblFreq)
    intNoteIndex = NoteBuffer.ChanterNoteIndex
    dblRelCent = NoteBuffer.RelativeCent
    Form_Monitor.Caption = "Time: " & Format(sngStartTime, "0000.00") & " / " & _
                            "Note Name: " & gudtNotes(intNoteIndex).Name & " / " & _
                            "Cent: " & Format(dblRelCent, "000.00") & " / " & _
                            "Freq.: " & Format(dblFreq, "000.00")
                                                   
    ' load read length again
    WavFile.ReadLength = lngReadLength
    
    GetData = True
                            
End Function

