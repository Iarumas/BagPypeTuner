Attribute VB_Name = "NoteDefinitions"
Option Explicit

Public Type NoteAttributes
    Selected As Boolean
    Index As Integer
    Name As String
    Pic As Picture
    Color As Long
    Ratio As Double
    Numerator As Integer
    Denominator As Integer
    CentSelected As Boolean
    ChromaticCent As Double
    RelativeCent As Double
    AbsoluteCent As Double
    Tolerance(-1 To 1) As Double
End Type

Public gudtNoteDefaults(0 To 11) As NoteAttributes
Public gudtDroneDefaults(0 To 3) As NoteAttributes

Public gudtNotes() As NoteAttributes
Public gudtDrones() As NoteAttributes

Public Sub NoteInit()
    
    Call NoteDefaults
    Call DroneDefaults
    
End Sub


Public Sub StandardSetting()
    
    Dim i As Integer
    
    Call NoteDefaultColors
    Call DroneDefaultColors
    
    Call HarmonicChanterScale
    'Call ChromaticChanterScale
    'Call HarmonicChanterScaleMid7th
    'Call HarmonicChanterScaleHigh7th
    
    Call DroneScale
    'Call ChromaticDroneScale
    
    Call MapNotes(gudtNotes, gudtNoteDefaults)
    Call MapNotes(gudtDrones, gudtDroneDefaults)
    
    'For i = LBound(gudtNoteDefaults) To UBound(gudtNoteDefaults)
    '    Debug.Print gudtNoteDefaults(i).Name, gudtNoteDefaults(i).CentSelected, _
    '    gudtNoteDefaults(i).Numerator, gudtNoteDefaults(i).Denominator, gudtNoteDefaults(i).Ratio, _
    '    gudtNoteDefaults(i).AbsoluteCent, gudtNoteDefaults(i).RelativeCent
    'Next i
    'Debug.Print
    
End Sub

Public Sub MapNotes(ByRef udtNotes() As NoteAttributes, ByRef udtNoteDefaults() As NoteAttributes)
    
    Dim i As Integer
    Dim intCurrentIndex As Integer
    Dim NoteIndex() As Integer
    
    For i = LBound(udtNoteDefaults) To UBound(udtNoteDefaults)
        If udtNoteDefaults(i).Selected Then
            ReDim Preserve NoteIndex(0 To intCurrentIndex)
            NoteIndex(intCurrentIndex) = i
            intCurrentIndex = intCurrentIndex + 1
        End If
    Next i
    
    ReDim udtNotes(LBound(NoteIndex) To UBound(NoteIndex))

    For i = LBound(NoteIndex) To UBound(NoteIndex)
    
        udtNotes(i).Name = udtNoteDefaults(NoteIndex(i)).Name
        udtNotes(i).Index = i
        Set udtNotes(i).Pic = udtNoteDefaults(NoteIndex(i)).Pic
        udtNotes(i).Color = udtNoteDefaults(NoteIndex(i)).Color
        
        udtNotes(i).Numerator = udtNoteDefaults(NoteIndex(i)).Numerator
        udtNotes(i).Denominator = udtNoteDefaults(NoteIndex(i)).Denominator
        
        udtNotes(i).CentSelected = udtNoteDefaults(NoteIndex(i)).CentSelected
        udtNotes(i).ChromaticCent = udtNoteDefaults(NoteIndex(i)).ChromaticCent
        
        If udtNotes(i).CentSelected Then
            udtNotes(i).AbsoluteCent = udtNoteDefaults(NoteIndex(i)).AbsoluteCent
            udtNotes(i).Ratio = 2 ^ (udtNotes(i).AbsoluteCent / 1200)
        Else
            udtNotes(i).Ratio = udtNotes(i).Numerator / udtNotes(i).Denominator
            udtNotes(i).AbsoluteCent = ConvertFrequencyInAbsoluteCent(1, udtNotes(i).Ratio)
        End If
        
        udtNotes(i).RelativeCent = udtNotes(i).AbsoluteCent - udtNotes(i).ChromaticCent

    Next i
    
    udtNotes(1).Tolerance(-1) = -100
    
    For i = LBound(udtNotes) + 1 To UBound(udtNotes) - 1
        
        udtNotes(i + 1).Tolerance(-1) = (udtNotes(i).AbsoluteCent - udtNotes(i + 1).AbsoluteCent) / 2
        udtNotes(i).Tolerance(1) = (udtNotes(i + 1).AbsoluteCent - udtNotes(i).AbsoluteCent) / 2
    
    Next i
    
    udtNotes(UBound(udtNotes)).Tolerance(1) = 100
    
    
    'For i = LBound(udtNotes) To UBound(udtNotes)
    '    Debug.Print udtNotes(i).Name, udtNotes(i).CentSelected, udtNotes(i).Ratio, udtNotes(i).AbsoluteCent, udtNotes(i).RelativeCent
    'Next i
    'Debug.Print
        
End Sub


Public Sub NoteDefaults()
    
    Dim i As Integer
    
    gudtNoteDefaults(0).Name = "No Note Detected"
    gudtNoteDefaults(0).Selected = True
    gudtNoteDefaults(0).ChromaticCent = 0
    gudtNoteDefaults(1).Name = "LG"
    gudtNoteDefaults(1).Selected = True
    gudtNoteDefaults(1).ChromaticCent = -200
    gudtNoteDefaults(2).Name = "LA"
    gudtNoteDefaults(2).Selected = True
    gudtNoteDefaults(2).ChromaticCent = 0
    gudtNoteDefaults(3).Name = "B"
    gudtNoteDefaults(3).Selected = True
    gudtNoteDefaults(3).ChromaticCent = 200
    gudtNoteDefaults(4).Name = "C"
    gudtNoteDefaults(4).Selected = False
    gudtNoteDefaults(4).ChromaticCent = 300
    gudtNoteDefaults(5).Name = "C#"
    gudtNoteDefaults(5).Selected = True
    gudtNoteDefaults(5).ChromaticCent = 400
    gudtNoteDefaults(6).Name = "D"
    gudtNoteDefaults(6).Selected = True
    gudtNoteDefaults(6).ChromaticCent = 500
    gudtNoteDefaults(7).Name = "E"
    gudtNoteDefaults(7).Selected = True
    gudtNoteDefaults(7).ChromaticCent = 700
    gudtNoteDefaults(8).Name = "F"
    gudtNoteDefaults(8).Selected = False
    gudtNoteDefaults(8).ChromaticCent = 800
    gudtNoteDefaults(9).Name = "F#"
    gudtNoteDefaults(9).Selected = True
    gudtNoteDefaults(9).ChromaticCent = 900
    gudtNoteDefaults(10).Name = "HG"
    gudtNoteDefaults(10).Selected = True
    gudtNoteDefaults(10).ChromaticCent = 1000
    gudtNoteDefaults(11).Name = "HA"
    gudtNoteDefaults(11).Selected = True
    gudtNoteDefaults(11).ChromaticCent = 1200

    For i = LBound(gudtNoteDefaults) To UBound(gudtNoteDefaults)
        gudtNoteDefaults(i).Index = i
        'Debug.Print gudtNoteDefaults(i).Name
    Next i

    For i = LBound(gudtNoteDefaults) To UBound(gudtNoteDefaults)
        Set gudtNoteDefaults(i).Pic = LoadResPicture(101 + i, vbResBitmap)
    Next i
    
End Sub

Public Sub DroneDefaults()
    
    Dim i As Integer
    
    gudtDroneDefaults(0).Name = "No Drone Detected"
    gudtDroneDefaults(0).Selected = True
    gudtDroneDefaults(0).ChromaticCent = 0
    gudtDroneDefaults(1).Name = "Bass"
    gudtDroneDefaults(1).Selected = True
    gudtDroneDefaults(1).ChromaticCent = -2400
    gudtDroneDefaults(2).Name = "Tenor1"
    gudtDroneDefaults(2).Selected = True
    gudtDroneDefaults(2).ChromaticCent = -1200
    gudtDroneDefaults(3).Name = "Tenor2"
    gudtDroneDefaults(3).Selected = False
    gudtDroneDefaults(3).ChromaticCent = -1200
    
    For i = LBound(gudtDroneDefaults) To UBound(gudtDroneDefaults)
        gudtDroneDefaults(i).Index = i
        'Debug.Print i, gudtDroneDefaults(i).Ratio, gudtDroneDefaults(i).Name
    Next i
    
    For i = LBound(gudtDroneDefaults) To UBound(gudtDroneDefaults)
        Set gudtDroneDefaults(i).Pic = LoadResPicture(113 + i, vbResBitmap)
    Next i
      
End Sub

Public Sub NoteDefaultColors()
    
    gudtNoteDefaults(0).Color = CLng(&H0)            'No Note
    gudtNoteDefaults(1).Color = RGB(128, 128, 255)   'LG
    gudtNoteDefaults(2).Color = CLng(&HFF0000)       'LA
    gudtNoteDefaults(3).Color = CLng(&HFF8000)       'B
    gudtNoteDefaults(4).Color = CLng(&HFFFF00)       'C natural
    gudtNoteDefaults(5).Color = CLng(&HFFFF80)       'C#
    gudtNoteDefaults(6).Color = RGB(0, 255, 0)       'D
    gudtNoteDefaults(7).Color = CLng(&H80FF80)       'E
    gudtNoteDefaults(8).Color = CLng(&H80FFFF)       'F natural
    gudtNoteDefaults(9).Color = RGB(255, 255, 0)     'F#
    gudtNoteDefaults(10).Color = RGB(255, 128, 0)    'HG
    gudtNoteDefaults(11).Color = CLng(&HFF)          'HA

End Sub

Public Sub DroneDefaultColors()
    
    gudtDroneDefaults(0).Color = CLng(&H0)
    gudtDroneDefaults(1).Color = CLng(&H800080)      'Bass
    gudtDroneDefaults(2).Color = CLng(&HFF80FF)      'Tenor1
    gudtDroneDefaults(3).Color = CLng(&HFFC0FF)      'Tenor2

End Sub

Public Sub DroneScale()

    Dim i As Integer
        
    'No Drone
    gudtDroneDefaults(0).Numerator = 0
    gudtDroneDefaults(0).Denominator = 1
    'Bass
    gudtDroneDefaults(1).Numerator = 1
    gudtDroneDefaults(1).Denominator = 4
    'Tenor1
    gudtDroneDefaults(2).Numerator = 1
    gudtDroneDefaults(2).Denominator = 2
    'Tenor2
    gudtDroneDefaults(3).Numerator = 1
    gudtDroneDefaults(3).Denominator = 2
    
    For i = LBound(gudtDroneDefaults) To UBound(gudtDroneDefaults)
        gudtDroneDefaults(i).CentSelected = False
        gudtDroneDefaults(i).Ratio = gudtDroneDefaults(i).Numerator / gudtDroneDefaults(i).Denominator
        gudtDroneDefaults(i).AbsoluteCent = ConvertFrequencyInAbsoluteCent(1, gudtDroneDefaults(i).Ratio)
        gudtDroneDefaults(i).RelativeCent = 0
    Next i
    
End Sub

Public Sub ChromaticDroneScale()
    
    Dim i As Integer
    
    Call DroneScale
    
    'No Drone
    gudtDroneDefaults(0).AbsoluteCent = 0
    'Bass
    gudtDroneDefaults(1).AbsoluteCent = -2400
    'Tenor1
    gudtDroneDefaults(2).AbsoluteCent = -1200
    'Tenor2
    gudtDroneDefaults(3).AbsoluteCent = -1200

    
    For i = LBound(gudtDroneDefaults) To UBound(gudtDroneDefaults)
        gudtDroneDefaults(i).CentSelected = True
        gudtDroneDefaults(i).RelativeCent = 0
        gudtDroneDefaults(i).Ratio = 2 ^ (gudtDroneDefaults(i).AbsoluteCent / 1200)   ' ratio from chromatic scale
    Next i

End Sub

Public Sub ChromaticChanterScale()
    
    Dim i As Integer
    
    Call HarmonicChanterScale
    
    'No Note
    gudtNoteDefaults(0).AbsoluteCent = 0           'No Note
    gudtNoteDefaults(1).AbsoluteCent = -200        'LG
    gudtNoteDefaults(2).AbsoluteCent = 0           'LA
    gudtNoteDefaults(3).AbsoluteCent = 200         'B
    gudtNoteDefaults(4).AbsoluteCent = 300         'C natural
    gudtNoteDefaults(5).AbsoluteCent = 400         'C sharp
    gudtNoteDefaults(6).AbsoluteCent = 500         'D
    gudtNoteDefaults(7).AbsoluteCent = 700         'E
    gudtNoteDefaults(8).AbsoluteCent = 800         'F natural
    gudtNoteDefaults(9).AbsoluteCent = 900         'F sharp
    gudtNoteDefaults(10).AbsoluteCent = 1000       'HG
    gudtNoteDefaults(11).AbsoluteCent = 1200       'HA
    
    For i = LBound(gudtNoteDefaults) To UBound(gudtNoteDefaults)
        gudtNoteDefaults(i).CentSelected = True
        gudtNoteDefaults(i).RelativeCent = 0
        gudtNoteDefaults(i).Ratio = 2 ^ (gudtNoteDefaults(i).AbsoluteCent / 1200) ' ratio from chromatic scale
    Next i
    
End Sub

Public Sub HarmonicChanterScale()
    
    Dim i As Integer
    
    ' No Note
    gudtNoteDefaults(0).Numerator = 0
    gudtNoteDefaults(0).Denominator = 1
    ' LG
    gudtNoteDefaults(1).Numerator = 7
    gudtNoteDefaults(1).Denominator = 8
    ' LA
    gudtNoteDefaults(2).Numerator = 1
    gudtNoteDefaults(2).Denominator = 1
    ' B
    gudtNoteDefaults(3).Numerator = 9
    gudtNoteDefaults(3).Denominator = 8
    ' C natural
    gudtNoteDefaults(4).Numerator = 6
    gudtNoteDefaults(4).Denominator = 5
    ' C sharp
    gudtNoteDefaults(5).Numerator = 5
    gudtNoteDefaults(5).Denominator = 4
    ' D
    gudtNoteDefaults(6).Numerator = 4
    gudtNoteDefaults(6).Denominator = 3
    ' E
    gudtNoteDefaults(7).Numerator = 3
    gudtNoteDefaults(7).Denominator = 2
    ' F natural
    gudtNoteDefaults(8).Numerator = 8
    gudtNoteDefaults(8).Denominator = 5
    ' F sharp
    gudtNoteDefaults(9).Numerator = 5
    gudtNoteDefaults(9).Denominator = 3
    ' HG
    gudtNoteDefaults(10).Numerator = 7
    gudtNoteDefaults(10).Denominator = 4
    ' HA
    gudtNoteDefaults(11).Numerator = 2
    gudtNoteDefaults(11).Denominator = 1

    'gudtNoteDefaults(i).Ratio = 0
    For i = LBound(gudtNoteDefaults) To UBound(gudtNoteDefaults)
        gudtNoteDefaults(i).CentSelected = False
        gudtNoteDefaults(i).Ratio = gudtNoteDefaults(i).Numerator / gudtNoteDefaults(i).Denominator
        gudtNoteDefaults(i).AbsoluteCent = ConvertFrequencyInAbsoluteCent(1, gudtNoteDefaults(i).Ratio)
        gudtNoteDefaults(i).RelativeCent = gudtNoteDefaults(i).AbsoluteCent - gudtNoteDefaults(i).ChromaticCent
    Next i
    
End Sub

Public Sub HarmonicChanterScaleMid7th()
    
    Call HarmonicChanterScale
    
    ' LG
    gudtNoteDefaults(1).Numerator = 8
    gudtNoteDefaults(1).Denominator = 9
    ' HG
    gudtNoteDefaults(10).Numerator = 16
    gudtNoteDefaults(10).Denominator = 9

    gudtNoteDefaults(1).Ratio = gudtNoteDefaults(1).Numerator / gudtNoteDefaults(1).Denominator
    gudtNoteDefaults(10).Ratio = gudtNoteDefaults(10).Numerator / gudtNoteDefaults(10).Denominator
    gudtNoteDefaults(1).AbsoluteCent = ConvertFrequencyInAbsoluteCent(1, gudtNoteDefaults(1).Ratio)
    gudtNoteDefaults(1).RelativeCent = gudtNoteDefaults(1).AbsoluteCent - gudtNoteDefaults(1).ChromaticCent
    gudtNoteDefaults(10).AbsoluteCent = ConvertFrequencyInAbsoluteCent(1, gudtNoteDefaults(10).Ratio)
    gudtNoteDefaults(10).RelativeCent = gudtNoteDefaults(10).AbsoluteCent - gudtNoteDefaults(10).ChromaticCent
 
End Sub

Public Sub HarmonicChanterScaleHigh7th()

    Call HarmonicChanterScale
    
    ' LG
    gudtNoteDefaults(1).Numerator = 9
    gudtNoteDefaults(1).Denominator = 10
    ' HG
    gudtNoteDefaults(10).Numerator = 9
    gudtNoteDefaults(10).Denominator = 5

    gudtNoteDefaults(1).Ratio = gudtNoteDefaults(1).Numerator / gudtNoteDefaults(1).Denominator
    gudtNoteDefaults(10).Ratio = gudtNoteDefaults(10).Numerator / gudtNoteDefaults(10).Denominator
    gudtNoteDefaults(1).AbsoluteCent = ConvertFrequencyInAbsoluteCent(1, gudtNoteDefaults(1).Ratio)
    gudtNoteDefaults(1).RelativeCent = gudtNoteDefaults(1).AbsoluteCent - gudtNoteDefaults(1).ChromaticCent
    gudtNoteDefaults(10).AbsoluteCent = ConvertFrequencyInAbsoluteCent(1, gudtNoteDefaults(10).Ratio)
    gudtNoteDefaults(10).RelativeCent = gudtNoteDefaults(10).AbsoluteCent - gudtNoteDefaults(10).ChromaticCent
    
End Sub


Public Function ConvertFrequencyInAbsoluteCent(ByVal dblRef As Double, ByVal dblFreq As Double) As Double
' Converts Frequencies in Cent
    
    If dblFreq <= 0 Or dblRef <= 0 Then
        ConvertFrequencyInAbsoluteCent = 0
    Else
        ConvertFrequencyInAbsoluteCent = 1200 / Log(2) * Log(dblFreq / dblRef)
    End If

End Function

