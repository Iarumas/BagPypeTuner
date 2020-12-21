VERSION 5.00
Begin VB.Form frmDX_Record 
   BorderStyle     =   0  'Kein
   Caption         =   "Form2"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3105
   LinkTopic       =   "Form2"
   ScaleHeight     =   1905
   ScaleWidth      =   3105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
End
Attribute VB_Name = "frmDX_Record"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Made by Michael Ciurescu
'
' DirectSound Tutorial:
' http://www.vbforums.com/showthread.php?t=388562

Option Explicit

' Direct Sound objects
Private DX As New DirectX8
Private SEnum As DirectSoundEnum8
Private DISCap As DirectSoundCapture8

' buffer, and buffer description
Private Buff As DirectSoundCaptureBuffer8
Private BuffDesc As DSCBUFFERDESC

' For the events
Private EventsNotify() As DSBPOSITIONNOTIFY
Private EndEvent As Long, MidEvent As Long, StartEvent As Long

' to know the buffer size
Private BuffLen As Long, HalfBuffLen As Long

Implements DirectXEvent8

Public Event GotWaveData(Buffer() As Byte)

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)

    Dim WaveBuffer() As Byte
    
    ' make sure that Buff object is actually initialized to a buffer instance
    If Not (Buff Is Nothing) Then
        ReDim WaveBuffer(HalfBuffLen - 1)
    
        Select Case eventid
        Case StartEvent
            ' we got the event that the write cursor is at the beginning of the buffer
            ' therefore read from the middle of the buffer to the end
            Buff.ReadBuffer HalfBuffLen, HalfBuffLen, WaveBuffer(0), DSCBLOCK_DEFAULT
        Case MidEvent
            ' we got an event that the write cursor is at the middle of the buffer
            ' threfore read from the beginning of the buffer to the middle
            Buff.ReadBuffer 0, HalfBuffLen, WaveBuffer(0), DSCBLOCK_DEFAULT
        Case EndEvent
            ' not used right now
        End Select
    
        If eventid = StartEvent Or eventid = MidEvent Then
            RaiseEvent GotWaveData(WaveBuffer)
        End If
    End If
End Sub

Public Function Initialize(Optional ByVal Wave_Format As Integer = WAVE_FORMAT_PCM, _
                            Optional ByVal SamplesPerSec As Long = 44100, _
                            Optional ByVal BitsPerSample As Integer = 16, _
                            Optional ByVal Channels As Integer = 2, _
                            Optional ByVal HalfBufferLen As Long = 0, _
                            Optional ByVal GUID As String = "") As String
    
    ' if there is any error go to ReturnError
    On Error GoTo ReturnError
    
    Set SEnum = DX.GetDSCaptureEnum ' get the device enumeration object
    
    ' if GUID is empty, then assign the first sound device
    If Len(GUID) = 0 Then GUID = SEnum.GetGuid(1)
    
    ' choose the sound device, and create the Direct Sound object
    Set DISCap = DX.DirectSoundCaptureCreate(GUID)
    
    ' set the format to use for recording
    With BuffDesc.fxFormat
        .nFormatTag = Wave_Format
        .nChannels = Channels

        .nBitsPerSample = BitsPerSample
        .lSamplesPerSec = SamplesPerSec
        
        .nBlockAlign = Int((.nBitsPerSample + 7) / 8) * .nChannels
        '.nBlockAlign = (.nBitsPerSample * .nChannels) / 8
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
        
        If HalfBufferLen <= 0 Then
            ' make half of the buffer to be 100 ms
            'HalfBuffLen = .lAvgBytesPerSec / 10
        Else
            ' using a "custom" size buffer
            HalfBuffLen = HalfBufferLen
        End If
        
        ' make sure the buffer is aligned
        HalfBuffLen = HalfBuffLen - (HalfBuffLen Mod .nBlockAlign)
    End With
    
    ' calculate the total size of the buffer
    BuffLen = HalfBuffLen * 2
    
    BuffDesc.lBufferBytes = BuffLen
    BuffDesc.lFlags = DSCBCAPS_DEFAULT
    
    ' create the buffer object
    Set Buff = DISCap.CreateCaptureBuffer(BuffDesc)
    
    ' Create 3 event notifications
    ReDim EventsNotify(0 To 2) As DSBPOSITIONNOTIFY
    
    ' create event to signal that DirectSound write cursor
    ' is at the beginning of the buffer
    StartEvent = DX.CreateEvent(Me)
    EventsNotify(0).hEventNotify = StartEvent
    EventsNotify(0).lOffset = 1
    
    ' create event to signal that DirectSound write cursor
    ' is at half of the buffer
    MidEvent = DX.CreateEvent(Me)
    EventsNotify(1).hEventNotify = MidEvent
    EventsNotify(1).lOffset = HalfBuffLen
    
    ' create the event to signal the sound has stopped
    EndEvent = DX.CreateEvent(Me)
    EventsNotify(2).hEventNotify = EndEvent
    EventsNotify(2).lOffset = DSBPN_OFFSETSTOP
    
    ' Assign the notification points to the buffer
    Buff.SetNotificationPositions 3, EventsNotify()
    
    Initialize = ""
    Exit Function
ReturnError:
    ' return error number, description and source
    Initialize = "Error: " & Err.Number & vbNewLine & _
        "Desription: " & Err.Description & vbNewLine & _
        "Source: " & Err.Source
    
    Err.Clear
    UninitializeSound
    Exit Function
End Function

Public Sub UninitializeSound()
    On Error Resume Next
    If UBound(EventsNotify) > 0 Then
        If Err.Number = 0 Then
            ' distroy all events
            DX.DestroyEvent EventsNotify(0).hEventNotify
            DX.DestroyEvent EventsNotify(1).hEventNotify
            DX.DestroyEvent EventsNotify(2).hEventNotify
            
            Erase EventsNotify
        End If
    End If
    
    Set Buff = Nothing
    Set DISCap = Nothing
    Set SEnum = Nothing
End Sub

Public Function SoundPlay() As Boolean
    On Error GoTo ReturnError
    
    If Not Buff Is Nothing Then Buff.Start DSCBSTART_LOOPING
    
    SoundPlay = True
    Exit Function
ReturnError:
    SoundPlay = False
    Err.Clear
End Function

Public Function SoundStop() As Boolean
    On Error GoTo ReturnError
    
    If Not Buff Is Nothing Then Buff.Stop
    
    SoundStop = True
    Exit Function
ReturnError:
    SoundStop = False
    Err.Clear
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UninitializeSound
End Sub
