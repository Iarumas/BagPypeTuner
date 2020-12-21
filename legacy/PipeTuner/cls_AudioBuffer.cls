VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_AudioBuffer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' FIFO buffer for arrays (only integer type)
    
Private Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (destination As Any, Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (destination As Any, ByVal Length As Long)

' define type of buffer : integer, long, single, double
' here: integer
Private mInputBuffer() As Integer
Private mOutputBuffer() As Integer
Private mBuffer() As Integer
' this is the only place the variables have to be set to int,lng or sng
' there are also some properties that have to be adjusted

Private mintChannels As Integer             'number of audio channels

' input buffer = sample length from audio source
' output buffer = sample length for analysis (FFT)
' step length = length between to FFTs: e.g. FFT sample length 8192 but do a FFT every 2048 sampels = step length
' buffer length = length of input buffer + outbut buffer - step length
Private mlngInputBufferLength As Long       'length of input buffer
Private mlngInputBufferSize As Long         'size of input buffer
Private mlngOutputBufferLength As Long      'length of output buffer
Private mlngOutputBufferSize As Long        'size of output buffer
Private mlngBufferLength As Long            'length of main buffer
Private mlngBufferSize As Long              'size of main buffer
Private mlngBufferStepLength As Long        'length of steps for buffer
Private mlngBufferStepSize As Long          'size of steps for buffer
Private mlngBufferReadPosition As Long      'position where to read data from main buffer
Private mlngBufferWritePosition As Long     'position where to write data to the main buffer
Private mlngBufferStartPosition As Long     'start position of main buffer

Public Event BufferReady()                  'event when buffer is full and needs to be read

' define type of buffer : integer, long, single, double
Public Property Let InputBuffer(ByRef Value() As Integer)
    mInputBuffer = Value
End Property
' define type of buffer : integer, long, single, double
Public Property Get InputBuffer() As Integer()
    InputBuffer = mInputBuffer
End Property
' define type of buffer : integer, long, single, double
Public Property Let OutputBuffer(ByRef Value() As Integer)
    mOutputBuffer = Value
End Property
' define type of buffer : integer, long, single, double
Public Property Get OutputBuffer() As Integer()
    OutputBuffer = mOutputBuffer
End Property

Public Property Let Channels(ByVal Value As Integer)
    mintChannels = Value
End Property
Public Property Get Channels() As Integer
    Channels = mintChannels
End Property

Public Property Let InputBufferLength(ByVal Value As Long)
    mlngInputBufferLength = Value
End Property
Public Property Get InputBufferLength() As Long
    InputBufferLength = mlngInputBufferLength
End Property
Public Property Let OutputBufferLength(ByVal Value As Long)
    mlngOutputBufferLength = Value
End Property
Public Property Get OutputBufferLength() As Long
    OutputBufferLength = mlngOutputBufferLength
End Property

Public Property Let BufferStepLength(ByVal Value As Long)
    mlngBufferStepLength = Value
End Property
Public Property Get BufferStepLength() As Long
    BufferStepLength = mlngBufferStepLength
End Property

Public Property Get BufferStartPosition() As Long
    BufferStartPosition = mlngBufferStartPosition
End Property

Public Sub BufferClear()
    
    ' set buffer length
    mlngBufferLength = mlngInputBufferLength + mlngOutputBufferLength - mlngBufferStepLength
    
    ' define size of buffers
    ReDim mInputBuffer(0 To mintChannels - 1, 0 To mlngInputBufferLength - 1)
    ReDim mOutputBuffer(0 To mintChannels - 1, 0 To mlngOutputBufferLength - 1)
    ReDim mBuffer(0 To mintChannels - 1, 0 To mlngBufferLength - 1)
    
    ' size in bytes
    mlngInputBufferSize = mintChannels * mlngInputBufferLength * Len(mInputBuffer(0, 0))
    mlngOutputBufferSize = mintChannels * mlngOutputBufferLength * Len(mOutputBuffer(0, 0))
    mlngBufferSize = mintChannels * mlngBufferLength * Len(mBuffer(0, 0))
    mlngBufferStepSize = mintChannels * mlngBufferStepLength * Len(mBuffer(0, 0))
     
    ' clear buffers
    Call ZeroMemory(mInputBuffer(0, 0), mlngInputBufferSize)
    Call ZeroMemory(mOutputBuffer(0, 0), mlngOutputBufferSize)
    Call ZeroMemory(mBuffer(0, 0), mlngBufferSize)
    
    ' start for 1. read
    mlngBufferReadPosition = 0
    ' start position for write:
    mlngBufferWritePosition = (mlngBufferLength Mod mlngInputBufferLength)
    If mlngBufferWritePosition <> 0 Then mlngBufferWritePosition = mlngBufferWritePosition - mlngInputBufferLength
    mlngBufferStartPosition = 0 - mlngBufferWritePosition
    ' start and write position will be 0 or negative for start
    
End Sub

Public Function BufferWrite() As Boolean
    
    Dim lngBytesToCopy As Long
    
    ' exit if buffer is full (or not enough space for another input buffer)
    If mlngBufferWritePosition > mlngBufferLength - mlngInputBufferLength Then
        Debug.Print " Buffer Full !"
        BufferWrite = False
        Exit Function
    End If
    
    If mlngBufferWritePosition >= 0 Then
        'move/push arrays: Destination, Source, Number of Bytes
        CopyMemory mBuffer(0, mlngBufferWritePosition), mInputBuffer(0, 0), mlngInputBufferSize
    Else
        ' if wp < 0 then only copy elements with positions >= 0
        lngBytesToCopy = mintChannels * Len(mBuffer(0, 0)) * (mlngInputBufferLength + mlngBufferWritePosition)
        CopyMemory mBuffer(0, 0), mInputBuffer(0, 0 - mlngBufferWritePosition), lngBytesToCopy
    End If
        
    ' update position of next input to buffer
    mlngBufferWritePosition = mlngBufferWritePosition + mlngInputBufferLength
    
    BufferWrite = True
    
    ' if buffer full raise event
    If mlngBufferWritePosition = mlngBufferLength Then
        RaiseEvent BufferReady
    End If
    
End Function

Public Function BufferRead() As Integer
    'returns the current read number (1, 2, ...) = number of steps (in units of step length)

    ' if not enough elements written to fill output buffer then exit
    If mlngBufferWritePosition < mlngOutputBufferLength Then
        BufferRead = 0
        Exit Function
    End If
    
    Dim lngBytesToMove As Long
    
    ' bytes to move = channels * size of 1 element * (buffer length - step length)
    lngBytesToMove = mintChannels * Len(mBuffer(0, 0)) * (mlngBufferLength - mlngBufferStepLength)
    
    ' Copy buffer into output buffer Destination, Source, Number of Bytes
    CopyMemory mOutputBuffer(0, 0), mBuffer(0, mlngBufferReadPosition), mlngOutputBufferSize
    
    ' update current buffer position = old position + step length
    mlngBufferReadPosition = mlngBufferReadPosition + mlngBufferStepLength
    
    ' is the current read number: 1,2,3,...
    ' not starting from 0, because read position will always be > step length when BufferRead function is called
    BufferRead = mlngBufferReadPosition / mlngBufferStepLength
    
    ' when buffer is read n times, so that 1 input buffer size was read then the buffer will be moved
    If BufferRead = mlngInputBufferLength / mlngBufferStepLength Then
        Call BufferMove
    End If
    
End Function

Private Sub BufferMove()
    
    If mlngBufferSize > mlngInputBufferSize Then
        ' move rest of buffer to the start of buffer
        MoveMemory mBuffer(0, 0), mBuffer(0, mlngInputBufferLength), mlngBufferSize - mlngInputBufferSize
    End If
    ' fill rest of buffer with zero
    'ZeroMemory mBuffer(0, mlngBufferLength - mlngInputBufferLength), mlngInputBufferSize

    ' update current buffer position
    mlngBufferReadPosition = 0
    mlngBufferWritePosition = mlngBufferLength - mlngInputBufferLength

End Sub
