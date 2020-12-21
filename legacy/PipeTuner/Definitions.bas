Attribute VB_Name = "Definitions"
Option Explicit

Public Const SPI_GETWORKAREA = 48

Public gblnDebug As Boolean
Public gdblReferenceFrequency  As Double
Public gsng_RefreshInterval As Single

Public gdblBufferFrequency As New cls_Multi_FIFO_dbl
Public gdblBufferCent As New cls_Multi_FIFO_dbl
Public gintBufferNote As New cls_FIFO_int

Public Type TunerType
    Live As Boolean
    Record As Boolean
    WAV As Boolean
End Type

Public Type TunerStatus
    CountDown As Long
    Running As Boolean
    Stop As Boolean
End Type

Public Type WAV_Type
    FormatInteger As Byte
    FormatLong As Byte
    FormatSingle As Byte
End Type

Public WavFile As New cls_WavFile
Public WavInFile As New cls_WavFile
Public WavOutFile As New cls_WavFile

Public Type RECT
  Left As Long
  Top As Integer
  Right As Long
  Bottom As Long
End Type

