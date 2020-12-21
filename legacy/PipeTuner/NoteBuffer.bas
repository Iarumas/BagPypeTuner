Attribute VB_Name = "NoteBuffer"
Option Explicit

' global Variables
' gudtNotes
' gudtDrones
' gdblReferenceFrequency
' gintBufferNote
' gdblBufferCent
' gdblBufferFrequency

' short buffer for cent / frequency
Private mdblShortBufferChanterNoteCent As New cls_Multi_FIFO_dbl
Private mdblShortBufferFrequency As New cls_Multi_FIFO_dbl

' index of chanter note:
' e.g.: 0: none / 1: LG / 2: LA / 3: B / 4: C / 5: D / 6: E / 7: F / 8: HG / 9: HA
Private mintChanterNoteIndex As Integer
' index of drone:
' e.g. 0: none / 1: bass / 2: tenor
Private mintDroneIndex As Integer

Private mdblRelativeCent As Double
Private mdblAbsoluteCent As Double

Private mlngShortBufferElements As Long
Private mlngShortBufferMaxNonValidElements As Long


Public Sub NoteBufferInit(ByVal lngLongBufferElements As Long)
    
    ' global values needed:
    ' gudtNotes
    ' gudtDrones
    
    mlngShortBufferElements = 3             ' 3 element buffer
    mlngShortBufferMaxNonValidElements = 1  ' only 1 element allowed to be 0
     
    mdblShortBufferChanterNoteCent.FIFO_Dimensions = UBound(gudtNotes) + 1
    mdblShortBufferChanterNoteCent.FIFO_Elements = mlngShortBufferElements
    mdblShortBufferChanterNoteCent.FIFO_MaxNonValidElements = mlngShortBufferMaxNonValidElements
    mdblShortBufferChanterNoteCent.FIFO_Clear
       
    ' dim 0: chanter / dim 1-n drones
    mdblShortBufferFrequency.FIFO_Dimensions = UBound(gudtDrones) + 1
    mdblShortBufferFrequency.FIFO_Elements = mlngShortBufferElements
    mdblShortBufferFrequency.FIFO_MaxNonValidElements = mlngShortBufferMaxNonValidElements
    mdblShortBufferFrequency.FIFO_Clear
    
    ' global buffers for note, cent and frequency
    gintBufferNote.FIFO_Elements = lngLongBufferElements
    gintBufferNote.FIFO_MaxNonValidElements = lngLongBufferElements - 1
    gintBufferNote.FIFO_Clear
    
    ' dim 0: chanter / dim 1-n drones
    gdblBufferCent.FIFO_Dimensions = UBound(gudtDrones) + 1
    gdblBufferCent.FIFO_Elements = lngLongBufferElements
    gdblBufferCent.FIFO_MaxNonValidElements = lngLongBufferElements - 1
    gdblBufferCent.FIFO_Clear

    ' dim 0: chanter / dim 1-n drones
    gdblBufferFrequency.FIFO_Dimensions = UBound(gudtDrones) + 1
    gdblBufferFrequency.FIFO_Elements = lngLongBufferElements
    gdblBufferFrequency.FIFO_MaxNonValidElements = lngLongBufferElements - 1
    gdblBufferFrequency.FIFO_Clear
        
End Sub

Public Function ShortBufferElements() As Long
    ShortBufferElements = mlngShortBufferElements
End Function
Public Function ChanterNoteIndex() As Integer
    ChanterNoteIndex = mintChanterNoteIndex
End Function
Public Function AbsoluteCent() As Double
    AbsoluteCent = mdblAbsoluteCent
End Function
Public Function RelativeCent() As Double
    RelativeCent = mdblRelativeCent
End Function

Public Sub UpdateCent(ByRef dblFreqArray, ByVal lngPosition As Long)
    ' dimension 1 of dblFreqArray: number of elements
    ' dimension 2 of dblFreqArray: frequencies = chanter plus number of drones
    ' lng position = current position
    
    ' global variables
    ' gdblReferenceFrequency
    ' gudtDrones
    ' gintBufferNote
    ' gdblBufferCent
    
    ' modular variables
    ' mdblRelativeCent
    ' mintChanterNoteIndex
    
    Dim i As Long
    Dim j As Long
    Dim intNote() As Integer
    Dim dblCent() As Double
    
    ReDim intNote(0 To UBound(dblFreqArray, 1))         ' number of elements
    ReDim dblCent(0 To UBound(dblFreqArray, 1), 0 To UBound(dblFreqArray, 2))       ' number of elements x frequencies
   
    For i = 0 To lngPosition - 1
        
        ' chanter
        DetectNote (dblFreqArray(i, 0))
        dblCent(i, 0) = mdblRelativeCent
        intNote(i) = mintChanterNoteIndex
        
        ' drones
        For j = 1 To UBound(dblFreqArray, 2)
            dblCent(i, j) = ConvertFrequencyInCent(gudtDrones(j).Ratio * gdblReferenceFrequency, dblFreqArray(i, j))
        Next j
        
    Next i
    
    ' store in buffer
    gintBufferNote.FIFO_Buffer = intNote
    gdblBufferCent.Set_FIFO_Buffer dblCent
    
    ' set buffer position
    gintBufferNote.FIFO_Position = lngPosition
    gdblBufferCent.FIFO_Position = lngPosition

End Sub

Public Sub BufferPipes(ByRef dblFrequency)
    ' buffers the identified notes in FIFO
    ' dblFrequency is array
    ' 0: chanter / 1-n: drones
    
    ' global  variables
    ' gudtNotes / gudtDrones / gdblReferenceFrequency
    ' gintBufferNote
    ' gdblBufferCent
    ' gdblBufferFrequency
    
    ' modular varaibels
    ' mdblRelativeCent
    ' mintChanterNoteIndex
    ' mdblShortBufferChanterNoteCent
    ' mdblShortBufferFrequency

    Dim j As Long
    Dim dblChanterNoteCent() As Double
    Dim dblCent() As Double
    Dim dblFreq() As Double
    
    ReDim dblChanterNoteCent(0 To UBound(gudtNotes))
    ReDim dblCent(0 To UBound(dblFrequency))
    ReDim dblFreq(0 To UBound(dblFrequency))
    
    ' detect note and store in chanter note buffer
    DetectNote (dblFrequency(0))
    dblChanterNoteCent(mintChanterNoteIndex) = mdblRelativeCent
    
    ' buffer chanter notes
    mdblShortBufferChanterNoteCent.FIFO_Fill (dblChanterNoteCent)
    ' result from chanter note buffer
    dblChanterNoteCent = mdblShortBufferChanterNoteCent.FIFO_Result
    
    ' restore chanter note buffer if one detection fails: note is the same as previous one
    If mintChanterNoteIndex = 0 And dblChanterNoteCent(0) <> 0 Then
        mintChanterNoteIndex = gintBufferNote.FIFO_Input
    End If
    
    ' if result from chanter note buffer = 0 the set frequency to 0 and index to 0
    If dblChanterNoteCent(mintChanterNoteIndex) = 0 Then
        mintChanterNoteIndex = 0
        dblFrequency(0) = 0
    End If
    
    ' buffer frequency
    mdblShortBufferFrequency.FIFO_Fill (dblFrequency)
    dblFreq = mdblShortBufferFrequency.FIFO_Result
    
    ' calculate cents and frequency for chanter
    dblFreq(0) = gudtNotes(mintChanterNoteIndex).Ratio * gdblReferenceFrequency * (2 ^ (dblChanterNoteCent(mintChanterNoteIndex) / 1200))
    dblCent(0) = dblChanterNoteCent(mintChanterNoteIndex)
    
    ' calculate cents and frequency for drones
    For j = 1 To UBound(dblFrequency)
        dblCent(j) = ConvertFrequencyInCent(gudtDrones(j).Ratio * gdblReferenceFrequency, dblFreq(j))
    Next j
    
    ' store in buffer
    gintBufferNote.FIFO_Fill (mintChanterNoteIndex)
    gdblBufferCent.FIFO_Fill (dblCent)
    gdblBufferFrequency.FIFO_Fill (dblFreq)

End Sub

Public Function DetectNote(dblChanterFrequency As Double) As Boolean
' caluclates index of note and cent

' global variable:
' gdblReferenceFrequency / gudtNotes

' modular variable
' mdblAbsoluteCent / mdblRelativeCent / mintNoteIndex

    DetectNote = False

    mdblAbsoluteCent = ConvertFrequencyInCent(gdblReferenceFrequency, dblChanterFrequency)
    
    If mdblAbsoluteCent <> 0 Then
        For mintChanterNoteIndex = LBound(gudtNotes) + 1 To UBound(gudtNotes)
            If mdblAbsoluteCent < gudtNotes(mintChanterNoteIndex).AbsoluteCent + gudtNotes(mintChanterNoteIndex).Tolerance(1) And _
               mdblAbsoluteCent > gudtNotes(mintChanterNoteIndex).AbsoluteCent + gudtNotes(mintChanterNoteIndex).Tolerance(-1) Then
                    
                    mdblRelativeCent = mdblAbsoluteCent - gudtNotes(mintChanterNoteIndex).AbsoluteCent
                    Exit Function
            End If
        Next mintChanterNoteIndex
        
    Else
        ' no note found if absolute cent is 0
        mintChanterNoteIndex = 0
        DetectNote = False
        Exit Function
    End If
    
    ' no note found
    If mintChanterNoteIndex > UBound(gudtNotes) Then
        mintChanterNoteIndex = 0
        Exit Function
    End If
    
    DetectNote = True
    
End Function

Private Function ConvertFrequencyInCent(ByVal dblRef As Double, ByVal dblFreq As Double) As Double
' Converts Frequencies in Cents
    
    If dblFreq <= 0 Or dblRef <= 0 Then
        ConvertFrequencyInCent = 0
    Else
        ConvertFrequencyInCent = 1200 / Log(2) * Log(dblFreq / dblRef)
    End If

End Function
