VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_WavFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' file states
Const mstrFileStatusOpen = "Open"
Const mstrFileStatusClosed = "Closed"
Const mstrFileStatusReading = "Reading"
Const mstrFileStatusWriting = "Writing"

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)

Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" _
    (ByVal lpFileName As String) As Long

Private Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" _
    (ByVal lpExistingFileName As String, _
     ByVal lpNewFileName As String, _
     ByVal bFailIfExists As Long) As Long
     
Private Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" _
    (ByVal lpExistingFileName As String, _
     ByVal lpNewFileName As String) As Long

Private Type WavFileHeader
    lngRiff As Long             ' "RIFF"
    lngFileSize As Long         ' file size in bytes
    lngWave As Long             ' "WAVE"
    lngFormat As Long           ' "fmt "
    lngFormatLength As Long     ' length of format info in wav file
End Type

Private Type WavFileFormat
    intFormatTag As Integer     ' 1 for PCM
    intChannels As Integer      ' channels
    lngSamplesPerSec As Long    ' eg. 44100
    lngBytesPerSec As Long      ' bytes per sec
    intBlockAlign As Integer    ' block:  bytes per sample x channels
    intBitsPersample As Integer ' bits per sample per channeö
End Type

Private Type WavFileChunk
    lngType As Long             ' "DATA"
    lngLength As Long           ' length of data in bytes
End Type

Private mudtFileHeader As WavFileHeader
Private mudtFileFormat As WavFileFormat
Private mudtFileChunk As WavFileChunk

Private mintFileNumber As Integer           ' filenumber for file access
Private mstrFileName As String              ' path for file to read
Private mstrFileDestination As String       ' path for file to store

Private mstrFileStatus As String    ' file status
Private mblnFileOpen As Boolean     ' file is open
Private mblnFileBusy As Boolean     ' file is opened by another sub/function

Private mlngHeaderSize As Long      ' wav file header size
Private mblnReadWrite As Boolean    ' wav file read or read/write
Private mblnNewFile As Boolean      ' wav file does not exist / new file created
Private mblnExists As Boolean       ' wav file exists

Private mlngSampleLength As Long    'FFT sample length:   only needed to store information for FFT
Private mlngSampleInterval As Long  'FFT sample interval: only needed to store information for FFT

Private mlngSampleStart As Long     'start position
Private mlngReadLength As Long      'length of read sample
Private mlngStartAddress As Long    'start byte
Private mlngReadByteLength As Long  'length of read sample in bytes

Private mintBytesPerChannel As Integer      ' bytes per channel per sample
Private mbytWavData() As Byte               ' byte array for wav. data
Private mintWavData() As Integer            ' integer array for 16b wav. data
Private mlngWavData() As Long               ' long array for 32b wav. data
Private mstrWAVDataType As String           ' "Integer", "Long" or "Single"

Private mlngReadData() As Long              ' long array for 32b read wav data
Private mintReadData() As Integer           ' integer array for 16b read wav data
Private msngReadData() As Single            ' single array for 32b float read wav data

Private msngTimerStart As Single            ' variable for start of timer
Private msngTimerDelay As Single            ' variable for delay of timer (too wait for data)
Private mblnTimeOut As Boolean              ' variable for time out

Private msngSectionTimeStart As Single      ' section/sample start time
Private msngSectionTimeStop As Single       ' section/sample stop time
Private msngSectionTimeLength As Single     ' section/sample time length

Private msngStartTime As Single             ' store start time of audio buffer


Public Property Get Exists() As Boolean
    mblnExists = FileExists(mstrFileName)
    Exists = mblnExists
End Property
Public Property Let ReadWrite(ByVal Value As Boolean)
    mblnReadWrite = Value
End Property
Public Property Get ReadWrite() As Boolean
    ReadWrite = mblnReadWrite
End Property
Property Get HeaderSize() As Long
    HeaderSize = mlngHeaderSize
End Property
Public Property Let FormatTag(ByVal Value As Integer)
    mudtFileFormat.intFormatTag = Value
End Property
Public Property Get FormatTag() As Integer
    FormatTag = mudtFileFormat.intFormatTag
End Property
Public Property Get DataType() As String
    DataType = mstrWAVDataType
End Property
Public Property Let Channels(ByVal Value As Integer)
    With mudtFileFormat
        .intChannels = Value
        ' when channels are modified then block align, bytes per sample and read byte length have to be adjusted
        .intBlockAlign = .intChannels * mintBytesPerChannel
        .lngBytesPerSec = .lngSamplesPerSec * .intBlockAlign
        mlngReadByteLength = mlngReadLength * .intBlockAlign
    End With
    
    ' adjust arrays
    ReDim mbytWavData(0 To mudtFileFormat.intBlockAlign * mlngReadLength - 1)
    ReDim mintReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    
    If mstrWAVDataType = "Integer" Then ReDim mintReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    If mstrWAVDataType = "Long" Then ReDim mlngReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    If mstrWAVDataType = "Single" Then ReDim msngReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    
End Property
Public Property Get Channels() As Integer
    Channels = mudtFileFormat.intChannels
End Property
Public Property Let BitsPerSample(ByVal Value As Integer)
    With mudtFileFormat
        .intBitsPersample = Value
        ' when bitspersample are modified then block align, bytes per sample and read byte length have to be adjusted
        mintBytesPerChannel = Int((.intBitsPersample + 7) / 8)  ' 16b -> 2B, 17b - 3B, 24b -> 3B, 25b -> 4B
        mlngReadByteLength = mlngReadLength * .intBlockAlign
        .intBlockAlign = .intChannels * mintBytesPerChannel
        .lngBytesPerSec = .lngSamplesPerSec * .intBlockAlign
    End With
    
    ' adjust arrays
    ReDim mbytWavData(0 To mudtFileFormat.intBlockAlign * mlngReadLength - 1)
    ReDim mintReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    
    If mudtFileFormat.intFormatTag = 1 And mudtFileFormat.intBitsPersample > 16 Then
    ' use long for PCM and bits > 16 bit
        mstrWAVDataType = "Long"
        ReDim mlngReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    End If

    If mudtFileFormat.intFormatTag = 1 And mudtFileFormat.intBitsPersample <= 16 Then
    ' use integer for PCM and bits <= 16 bit
        mstrWAVDataType = "Integer"
        ReDim mintReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    End If

    If mudtFileFormat.intFormatTag = 3 Then
    ' IEEE float format: 32bit float -> use single
        mstrWAVDataType = "Single"
        ReDim msngReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    End If
    
    
    If mudtFileFormat.intFormatTag <> 1 And mudtFileFormat.intFormatTag <> 3 Then
        ' only PCM and IEEE float format is supported
        mstrWAVDataType = ""
        MsgBox ("File Format Not Supported")
    End If

End Property
Public Property Get BitsPerSample() As Integer
    BitsPerSample = mudtFileFormat.intBitsPersample
End Property
Public Property Get BlockAlign() As Integer
    BlockAlign = mudtFileFormat.intBlockAlign
End Property
Public Property Let SampleRate(ByVal Value As Long)
    With mudtFileFormat
        ' when sample rate is modified then also bytes per sec is modified
        .lngSamplesPerSec = Value
        .lngBytesPerSec = .lngSamplesPerSec * .intBlockAlign
    End With
End Property
Public Property Get SampleRate() As Long
    SampleRate = mudtFileFormat.lngSamplesPerSec
End Property

Public Property Get DataBytes() As Long
    ' wav file data byte size
    DataBytes = mudtFileChunk.lngLength
End Property
Public Property Get DataSamples() As Long
    ' wav file data points
    DataSamples = mudtFileChunk.lngLength / mudtFileFormat.intBlockAlign
End Property
Public Property Get DataTime() As Single
    ' wav file time length
    DataTime = mudtFileChunk.lngLength / mudtFileFormat.intBlockAlign / mudtFileFormat.lngSamplesPerSec
End Property

Public Property Let FileName(ByVal Value As String)
    mstrFileName = Value
    mblnExists = FileExists(mstrFileName)
End Property
Public Property Get FileName() As String
    FileName = mstrFileName
End Property

Public Property Let FileDestination(ByVal Value As String)
    mstrFileDestination = Value
End Property
Public Property Get FileDestination() As String
    FileDestination = mstrFileDestination
End Property

Public Property Get FileStatus() As String
    FileStatus = mstrFileStatus
End Property
Public Property Get FileOpen() As Boolean
    FileOpen = mblnFileOpen
End Property
Public Property Get FileBusy() As Boolean
    FileBusy = mblnFileBusy
End Property
Public Property Get Filenumber() As Integer
    Filenumber = mintFileNumber
End Property

Public Property Let WavData(ByRef Value() As Byte)
    mbytWavData = Value
End Property
Public Property Get WavData() As Byte()
    WavData = mbytWavData
End Property

Public Property Let IntegerData(ByRef Value() As Integer)
    ' copy integer data array into byt array
    CopyMemory mbytWavData(0), Value(0, 0), mlngReadByteLength
End Property
Public Property Let LongData(ByRef Value() As Long)
    ' copy integer data array into byt array
    CopyMemory mbytWavData(0), Value(0, 0), mlngReadByteLength
End Property

Public Property Get ReadDataLong() As Long()
    ' read long data array
    ReadDataLong = mlngReadData
End Property
Public Property Get ReadDataInteger() As Integer()
    ' read integer data array
    ReadDataInteger = mintReadData
End Property
Public Property Get ReadDataSingle() As Single()
    ' read singel data array
    ReadDataSingle = msngReadData
End Property
Public Property Get ReadData() As Integer()
    ' read data array: here integer
    ReadData = mintReadData
End Property

Public Property Let SampleStart(ByVal Value As Long)
    mlngSampleStart = Value
End Property
Public Property Get SampleStart() As Long
    SampleStart = mlngSampleStart
End Property

Public Property Let ReadLength(ByVal Value As Long)
    ' number of samples
    With mudtFileFormat
        mlngReadLength = Value
        ' when read length is modified then the read length in bytes has to be adjusted
        mlngReadByteLength = mlngReadLength * .intBlockAlign
    End With
    
    ' adjust arrays
    ReDim mbytWavData(0 To mudtFileFormat.intBlockAlign * mlngReadLength - 1)
    ReDim mintReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    
    If mstrWAVDataType = "Integer" Then ReDim mintReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    If mstrWAVDataType = "Long" Then ReDim mlngReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    If mstrWAVDataType = "Single" Then ReDim msngReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)

End Property
Public Property Get ReadLength() As Long
    ReadLength = mlngReadLength
End Property
Public Property Get ReadTime() As Single
    ' time length of read data
    ReadTime = mlngReadLength / mudtFileFormat.lngSamplesPerSec
End Property
Public Property Get ReadByteLength() As Long
    ReadByteLength = mlngReadByteLength
End Property

Public Property Let SampleLength(ByVal Value As Long)
    mlngSampleLength = Value
End Property
Public Property Get SampleLength() As Long
    SampleLength = mlngSampleLength
End Property
Public Property Let SampleInterval(ByVal Value As Long)
    mlngSampleInterval = Value
End Property
Public Property Get SampleInterval() As Long
    SampleInterval = mlngSampleInterval
End Property

Public Property Let SectionTimeStart(ByVal Value As Single)
    msngSectionTimeStart = Value
    msngSectionTimeStop = msngSectionTimeStart + msngSectionTimeLength
End Property
Public Property Get SectionTimeStart() As Single
    SectionTimeStart = msngSectionTimeStart
End Property
Public Property Let SectionTimeLength(ByVal Value As Single)
    msngSectionTimeLength = Value
    msngSectionTimeStop = msngSectionTimeStart + msngSectionTimeLength
End Property
Public Property Get SectionTimeLength() As Single
    SectionTimeLength = msngSectionTimeLength
End Property
Public Property Get SectionTimeStop() As Single
    msngSectionTimeStop = msngSectionTimeStart + msngSectionTimeLength
    SectionTimeStop = msngSectionTimeStop
End Property
Public Property Let StartTime(ByVal Value As Single)
    msngStartTime = Value
End Property
Public Property Get StartTime() As Single
    StartTime = msngStartTime
End Property

Private Sub Class_Initialize()
    Call Init
End Sub

Public Sub Init()
     
    mintFileNumber = 0
    
    mstrFileName = ""
    mstrFileStatus = mstrFileStatusClosed
    
    mlngHeaderSize = 44     ' normal PCM header size
    
    mblnExists = False
    mblnReadWrite = False
    mblnNewFile = True
    mblnFileOpen = False
    mblnFileBusy = False
    mblnTimeOut = False
    
    mlngSampleStart = 0
    mlngReadLength = 8192
    
    With mudtFileFormat
        .intFormatTag = 1
        .lngSamplesPerSec = 44100
        .intBitsPersample = 16
        mintBytesPerChannel = Int((.intBitsPersample + 7) / 8)
        .intChannels = 1
        .intBlockAlign = .intChannels * mintBytesPerChannel
        .lngBytesPerSec = .intBlockAlign * .lngSamplesPerSec
        mlngReadByteLength = .intBlockAlign * mlngReadLength
    End With
    
    mstrWAVDataType = "Integer"
    
    msngSectionTimeStart = 0
    msngSectionTimeLength = 20
    msngSectionTimeStop = msngSectionTimeStart + msngSectionTimeLength
    msngStartTime = 0

End Sub

Public Sub OpenFile()
    
    'check if file exists / if not there is a new file
    mblnNewFile = Not FileExists(mstrFileName)
    If Not mblnExists Then mblnNewFile = True Else mblnNewFile = False
    
    mintFileNumber = FreeFile
    
    'open file for read/write or read only
    If mblnReadWrite Then
        Open mstrFileName For Binary Access Read Write As #mintFileNumber
    Else
        Open mstrFileName For Binary Access Read As #mintFileNumber
    End If
    
    ' now file exists definitly
    mblnExists = True
    'file is open
    mstrFileStatus = mstrFileStatusOpen
    mblnFileOpen = True

    'file header is written already but no data
    If mblnNewFile Then Call WriteHeader
 
End Sub

Public Sub CloseFile()
    
    'check if file exists
    mblnExists = FileExists(mstrFileName)
    If Not mblnExists Then
        MsgBox ("File Does Not Exist")
        Exit Sub
    End If
    
    'check if file is open
    If Not mblnFileOpen Then
        MsgBox ("File Already Closed")
        Exit Sub
    End If
    
    'close file
    Close #mintFileNumber
    'file is closed
    mstrFileStatus = mstrFileStatusClosed
    mblnFileOpen = False

End Sub

Public Sub GetWavFileInfo()
    
    'read the header info of wav file
    
    Dim lngHeader2DATA As Long
          
    'check if file exists
    mblnExists = FileExists(mstrFileName)
    If Not mblnExists Then
        MsgBox ("File Does Not Exist")
        Exit Sub
    End If
    
    'check if file is open
    If Not mblnFileOpen Then
        MsgBox ("File Not Open")
        Exit Sub
    End If
    
    'check if file is already accessed / if yes wait
    Do While mblnFileBusy
        Debug.Print "Sub GetWavFileInfo", "File busy", mstrFileStatus
        DoEvents
    Loop
    
    'now file is accessed for reading
    mblnFileBusy = True
    mstrFileStatus = mstrFileStatusReading
       
    'read header
    Get #mintFileNumber, 1, mudtFileHeader
    Get #mintFileNumber, , mudtFileFormat
    
    'header length accoring to header information
    lngHeader2DATA = Len(mudtFileHeader) + mudtFileHeader.lngFormatLength
    Get #mintFileNumber, lngHeader2DATA + 1, mudtFileChunk
    
    'check real header length / until "DATA"
    Do While mudtFileChunk.lngType <> &H61746164    ' "DATA"
        lngHeader2DATA = lngHeader2DATA + 2
        Get #mintFileNumber, lngHeader2DATA + 1, mudtFileChunk
        'exit if header size is longer the 256 bytes
        If lngHeader2DATA > 256 Then Exit Sub
    Loop
    
    'real header size
    mlngHeaderSize = lngHeader2DATA + 8
    
    With mudtFileFormat
        mlngReadByteLength = mlngReadLength * .intBlockAlign
        mintBytesPerChannel = .intBlockAlign / .intChannels
    End With
    
    ' only PCM (1) and IEEE float format (3) is supported
    If mudtFileFormat.intFormatTag <> 1 And mudtFileFormat.intFormatTag <> 3 Then
        MsgBox ("File Format Not Supported")
        Exit Sub
    End If
    
    ' adjust arrays
    ReDim mbytWavData(0 To mudtFileFormat.intBlockAlign * mlngReadLength - 1)
    ReDim mintReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    
    If mudtFileFormat.intFormatTag = 1 And mudtFileFormat.intBitsPersample > 16 Then
    ' use long for PCM and bits > 16 bit
        mstrWAVDataType = "Long"
        ReDim mlngReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    End If

    If mudtFileFormat.intFormatTag = 1 And mudtFileFormat.intBitsPersample <= 16 Then
    ' use integer for PCM and bits<= 16 bit
        mstrWAVDataType = "Integer"
        ReDim mintReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    End If

    If mudtFileFormat.intFormatTag = 3 Then
    ' IEEE float format: 32bit float -> use single
        mstrWAVDataType = "Single"
        ReDim msngReadData(0 To mudtFileFormat.intChannels - 1, 0 To mlngReadLength - 1)
    End If
    
    'release file again for other accesses
    mblnFileBusy = False
    mstrFileStatus = mstrFileStatusOpen
        
End Sub

Public Sub WriteHeader()
        
    'check if file exists
    mblnExists = FileExists(mstrFileName)
    If Not mblnExists Then
        MsgBox ("File Does Not Exist")
        Exit Sub
    End If
    
    'check if file is open
    If Not mblnFileOpen Then
        MsgBox ("File Not Open")
        Exit Sub
    End If
    
    'check if file read/write
    If Not mblnReadWrite Then
        MsgBox ("File is Read Only")
        Exit Sub
    End If
    
    ''check if file is already accessed / if yes wait
    Do While mblnFileBusy
        Debug.Print "Sub WriteHeader", "File busy", mstrFileStatus
        DoEvents
    Loop
    
    'now file is accessed/busy
    mblnFileBusy = True
    mstrFileStatus = mstrFileStatusWriting
   
    With mudtFileFormat
        '.intFormatTag = 1   'PCM
        '.intChannels = .intChannels
        '.lngSamplesPerSec = .lngSamplesPerSec
        '.intBitsPersample = intBitsPersample
        .intBlockAlign = .intChannels * Int((.intBitsPersample + 7) / 8)
        .lngBytesPerSec = .intBlockAlign * .lngSamplesPerSec
    End With
    
    With mudtFileHeader
        .lngRiff = &H46464952 ' "RIFF"
        .lngFileSize = 0
        .lngWave = &H45564157 ' "WAVE"
        .lngFormat = &H20746D66 ' "fmt "
        .lngFormatLength = Len(mudtFileFormat)
    End With
    
    With mudtFileChunk
        .lngType = &H61746164 ' "data"
        .lngLength = 0
    End With
    
    'header size
    mlngHeaderSize = Len(mudtFileHeader) + Len(mudtFileFormat) + Len(mudtFileChunk)
    mudtFileHeader.lngFileSize = mlngHeaderSize
    
    'write to file
    Put #mintFileNumber, 1, mudtFileHeader
    Put #mintFileNumber, , mudtFileFormat
    Put #mintFileNumber, , mudtFileChunk
    
    Put #mintFileNumber, 5, mudtFileHeader.lngFileSize
    Put #mintFileNumber, mlngHeaderSize - 4 + 1, mudtFileChunk.lngLength
    
    'release file again for other accesses
    mblnFileBusy = False
    mstrFileStatus = mstrFileStatusOpen
    
End Sub

Public Sub WriteWavData()
    
    'check if file exists
    If Not mblnExists Then
        MsgBox ("File Does Not Exist")
        Exit Sub
    End If
    
    'check if file is open
    If Not mblnFileOpen Then
        MsgBox ("File Not Open")
        Exit Sub
    End If
    
    'check if file is read only
    If Not mblnReadWrite Then
        MsgBox ("File is Read Only")
        Exit Sub
    End If
    
    ''check if file is already accessed / if yes wait
    Do While mblnFileBusy
        Debug.Print "Sub WriteWavData", "File busy", mstrFileStatus
        DoEvents
    Loop
    
    'now the file is accessed
    mblnFileBusy = True
    mstrFileStatus = mstrFileStatusWriting
          
    'read what's already in the wav file
    Get #mintFileNumber, 1, mudtFileHeader
    Get #mintFileNumber, , mudtFileFormat
    Get #mintFileNumber, , mudtFileChunk
       
    'header size
    mlngHeaderSize = Len(mudtFileHeader) + Len(mudtFileFormat) + Len(mudtFileChunk)
    'length if data (chunk) in bytes
    mudtFileChunk.lngLength = mudtFileChunk.lngLength + mlngReadByteLength

    Put #mintFileNumber, mudtFileHeader.lngFileSize + 1, mbytWavData      'attach to old file size
    
    mudtFileHeader.lngFileSize = mudtFileChunk.lngLength + mlngHeaderSize  'new file size
    
    Put #mintFileNumber, 5, mudtFileHeader.lngFileSize                     'modify header
    Put #mintFileNumber, mlngHeaderSize - 4 + 1, mudtFileChunk.lngLength   'modify chunk size
    
    'file free again for other accesses
    mblnFileBusy = False
    mstrFileStatus = mstrFileStatusOpen
    
End Sub

Public Sub Delete(Optional strFileName As String = "")
    
    Dim retval As Long
    Dim blnExists As Boolean
    
    If strFileName = "" Then strFileName = mstrFileName
    
    'check if file exists
    blnExists = FileExists(strFileName)
    If Not blnExists Then
        'MsgBox ("File Does Not Exist")
        Exit Sub
    End If

    retval = DeleteFile(strFileName)
 
    If retval = 1 Then
      ' MsgBox "File deleted", vbInformation + _
        vbOKOnly, "Successfully deleted"
    Else
      MsgBox "File could not be deleted", vbInformation + _
        vbOKOnly, "File Not Deleted"
    End If

End Sub

Public Sub Copy(ByVal strNewFilename As String)
    
    Dim retval As Long
    
    'check if file exists
    mblnExists = FileExists(mstrFileName)
    If Not mblnExists Then
        MsgBox ("File Does Not Exist C")
        Exit Sub
    End If
    
    'check if file is not open or accessed
    If mblnFileOpen Or mblnFileBusy Then
        MsgBox ("File is still open")
        Exit Sub
    End If
    
    'copy file
    retval = CopyFile(mstrFileName, strNewFilename, 0)
 
    If retval = 0 Then MsgBox "File already exists", vbInformation + vbOKOnly, "File Not Copied"

End Sub

Public Sub Move(ByVal strNewFilename As String)
    
    Dim retval As Long
    
    'check if file exists
    mblnExists = FileExists(mstrFileName)
    If Not mblnExists Then
        MsgBox ("File Does Not Exist M")
        Exit Sub
    End If
    
    'check if file is not open or accessed
    If mblnFileOpen Or mblnFileBusy Then
        MsgBox ("File is still open")
        Exit Sub
    End If
    
    'copy file
    retval = CopyFile(mstrFileName, strNewFilename, 0)
 
    If retval = 0 Then
        Delete (strNewFilename)
        retval = CopyFile(mstrFileName, strNewFilename, 0)
    End If
    
    If retval = 0 Then MsgBox "File Not Moved", vbInformation + vbOKOnly

End Sub

Public Function ReadWavData() As Boolean
    
    ReadWavData = False
    
        'check if file exists
    If Not mblnExists Then
        MsgBox ("File Does Not Exist")
        Exit Function
    End If
    
    'check if file is open
    If Not mblnFileOpen Then
        MsgBox ("File Not Open")
        Exit Function
    End If
    
    mlngStartAddress = mlngSampleStart * mudtFileFormat.intBlockAlign + mlngHeaderSize + 1

    'start timer
    msngTimerStart = Timer
    
    Do
        'check if file is already accessed / if yes wait (DoEvents)
        Do While mblnFileBusy
            Debug.Print "Sub ReadWavData", "File busy", mstrFileStatus
            DoEvents
        Loop
        
        'if there is more data go on / if no more data wait/doevents until time out
        If mlngStartAddress - 1 + mlngReadByteLength <= mudtFileChunk.lngLength + mlngHeaderSize Then Exit Do
        
        DoEvents
        
        'check time
        msngTimerDelay = Timer - msngTimerStart
        'time out limit is time of 4x samples
        mblnTimeOut = msngTimerDelay > 4 * mlngReadLength / mudtFileFormat.lngSamplesPerSec
    
    Loop Until mblnTimeOut
    
    'if time out the exit
    If mblnTimeOut Then
        MsgBox ("No More Data")
        Exit Function
    End If
    
    'file busy reading
    mblnFileBusy = True
    mstrFileStatus = mstrFileStatusReading

    Get #mintFileNumber, mlngStartAddress, mbytWavData
    
    ' convert byte data to int,lng or sng
    Call ConvertWavData
    
    'release file again for other accesses
    mblnFileBusy = False
    mstrFileStatus = mstrFileStatusOpen
    
    'start of next sample = old start plus length
    mlngSampleStart = mlngSampleStart + mlngReadLength

    ReadWavData = True
    
End Function

Public Sub ConvertWavData()

    Select Case mintBytesPerChannel

        ' 25-32bit(4Bytes)
        Case 4:
            If mstrWAVDataType = "Long" Then Call ConvertWavData_32b_To_Integer
            If mstrWAVDataType = "Single" Then Call ConvertWavData_32f_To_Integer
            'If mstrWAVDataType = "Long" Then Call ConvertWavData_32b_To_Long
            'If mstrWAVDataType = "Single" Then Call ConvertWavData_32f_To_Single
        ' 17-24bit(3Bytes)
        Case 3:
            Call ConvertWavData_24b_To_Integer
            mstrWAVDataType = "Integer"
            'Call ConvertWavData_24b_To_Long
        '9-16bit(2Bytes)
        Case 2: Call ConvertWavData_16b_To_Integer
        ' 8 bit -> 1 Byte / Byte is unsigned
        Case 1: Call ConvertWavData_8b_To_Integer
            
    End Select

End Sub

Private Sub ConvertWavData_8b_To_Long()

    Dim lngNumberOfSample As Long
    Dim intChannelX As Integer
    Dim lngByteIndex As Long
    
    For lngNumberOfSample = 0 To mlngReadLength - 1
        For intChannelX = 0 To mudtFileFormat.intChannels - 1

            'lngByteIndex = mudtFileFormat.intBlockAlign * lngNumberOfSample + mintBytesPerChannel * intChannelX
            lngByteIndex = mudtFileFormat.intBlockAlign * lngNumberOfSample + intChannelX
            ' normalize -2G to +2G
            mlngReadData(intChannelX, lngNumberOfSample) = (CLng(mbytWavData(lngByteIndex)) - 128) * 16777216
            'CopyMemory ByVal VarPtr(mlngReadData(intChannelX, lngNumberOfSample)) + 3, mbytWavData(lngByteIndex), 1
            'mlngReadData(intChannelX, lngNumberOfSample) = mlngReadData(intChannelX, lngNumberOfSample) - 2147483648#
        
        Next intChannelX
    Next lngNumberOfSample

End Sub
Private Sub ConvertWavData_8b_To_Integer()

    Dim lngNumberOfSample As Long
    Dim intChannelX As Integer
    Dim lngByteIndex As Long
    
    For lngNumberOfSample = 0 To mlngReadLength - 1
        For intChannelX = 0 To mudtFileFormat.intChannels - 1
            
            'lngByteIndex = mudtFileFormat.intBlockAlign * lngNumberOfSample + mintBytesPerChannel * intChannelX
            lngByteIndex = mudtFileFormat.intBlockAlign * lngNumberOfSample + intChannelX
            ' normalize -32768 to +32767
            mintReadData(intChannelX, lngNumberOfSample) = (CInt(mbytWavData(lngByteIndex)) - 128) * 256
        
        Next intChannelX
    Next lngNumberOfSample

End Sub

Private Sub ConvertWavData_16b_to_Long()

    Dim lngNumberOfSample As Long
    Dim intChannelX As Integer
    Dim lngByteIndex As Long

    For lngNumberOfSample = 0 To mlngReadLength - 1
        For intChannelX = 0 To mudtFileFormat.intChannels - 1
            
            lngByteIndex = mudtFileFormat.intBlockAlign * lngNumberOfSample + 2 * intChannelX
            ' e.g. 2Bytes per Ch. in to Long: Byte0->LongByte2 / Byte1->LongByte3
            CopyMemory ByVal VarPtr(mlngReadData(intChannelX, lngNumberOfSample)) + 2, mbytWavData(lngByteIndex), 2

        Next intChannelX
    Next lngNumberOfSample

End Sub
Private Sub ConvertWavData_16b_To_Integer()

    CopyMemory mintReadData(0, 0), mbytWavData(0), mlngReadByteLength

End Sub

Private Sub ConvertWavData_24b_To_Long()

    Dim lngNumberOfSample As Long
    Dim intChannelX As Integer
    Dim lngByteIndex As Long

    For lngNumberOfSample = 0 To mlngReadLength - 1
        For intChannelX = 0 To mudtFileFormat.intChannels - 1
        
            lngByteIndex = mudtFileFormat.intBlockAlign * lngNumberOfSample + 3 * intChannelX
            ' e.g. 3Bytes per Ch. in to Long: Byte0->LongByte1 / Byte1->Longbyte2 / Byte2->LongByte3
            CopyMemory ByVal VarPtr(mlngReadData(intChannelX, lngNumberOfSample)) + 1, mbytWavData(lngByteIndex), 3
        
        Next intChannelX
    Next lngNumberOfSample

End Sub
Private Sub ConvertWavData_24b_To_Integer()

    Dim lngNumberOfSample As Long
    Dim intChannelX As Integer
    Dim lngByteIndex As Long

    For lngNumberOfSample = 0 To mlngReadLength - 1
        For intChannelX = 0 To mudtFileFormat.intChannels - 1
            
            lngByteIndex = mudtFileFormat.intBlockAlign * lngNumberOfSample + 3 * intChannelX + 1
            ' only the 2 higher bytes are copied: Byte1->IntegerByte0 / Byte2->IntegerByte1
            CopyMemory mintReadData(intChannelX, lngNumberOfSample), mbytWavData(lngByteIndex), 2

        Next intChannelX
    Next lngNumberOfSample

End Sub

Private Sub ConvertWavData_32b_To_Long()

    CopyMemory mlngReadData(0, 0), mbytWavData(0), mlngReadByteLength

End Sub
Private Sub ConvertWavData_32b_To_Integer()

    Dim lngNumberOfSample As Long
    Dim intChannelX As Integer
    Dim lngByteIndex As Long

    For lngNumberOfSample = 0 To mlngReadLength - 1
        For intChannelX = 0 To mudtFileFormat.intChannels - 1
            
            lngByteIndex = mudtFileFormat.intBlockAlign * lngNumberOfSample + 4 * intChannelX + 2
            ' only the 2 higher bytes are copied: Byte2->IntegerByte0 / Byte3->IntegerByte1
            CopyMemory mintReadData(intChannelX, lngNumberOfSample), mbytWavData(lngByteIndex), 2

        Next intChannelX
    Next lngNumberOfSample

End Sub
Private Sub ConvertWavData_32f_To_Single()
    
    CopyMemory msngReadData(0, 0), mbytWavData(0), mlngReadByteLength

End Sub

Private Sub ConvertWavData_32f_To_Long()
    
    Dim lngNumberOfSample As Long
    Dim intChannelX As Integer
    
    CopyMemory msngReadData(0, 0), mbytWavData(0), mlngReadByteLength
    
    For lngNumberOfSample = 0 To mlngReadLength - 1
        For intChannelX = 0 To mudtFileFormat.intChannels - 1
                
            mlngReadData(intChannelX, lngNumberOfSample) = CLng(2 ^ 31 * msngReadData(intChannelX, lngNumberOfSample))
                
        Next intChannelX
    Next lngNumberOfSample

End Sub
Private Sub ConvertWavData_32f_To_Integer()

    Dim lngNumberOfSample As Long
    Dim intChannelX As Integer
    Dim lngByteIndex As Long
    
    CopyMemory msngReadData(0, 0), mbytWavData(0), mlngReadByteLength

    For lngNumberOfSample = 0 To mlngReadLength - 1
        For intChannelX = 0 To mudtFileFormat.intChannels - 1
                
            mintReadData(intChannelX, lngNumberOfSample) = CInt(2 ^ 15 * msngReadData(intChannelX, lngNumberOfSample))
                
        Next intChannelX
    Next lngNumberOfSample

End Sub

Public Function dB_Value() As Single
    
    Dim i As Long
    
    dB_Value = 0
    
    Select Case mstrWAVDataType
    
        Case "Integer"
            Dim intValue As Integer
            Dim intMaxValue As Integer
            For i = 0 To mlngReadLength - 1
                intValue = mintReadData(0, i)
                intMaxValue = IIf(intValue > intMaxValue, intValue, intMaxValue)
            Next i
            If intMaxValue > 0 Then dB_Value = 10 * Log(intMaxValue / 2 ^ 15) / Log(10)
    
        Case "Long"
            Dim lngValue As Single
            Dim lngMaxValue As Single
            For i = 0 To mlngReadLength - 1
                lngValue = mlngReadData(0, i)
                lngMaxValue = IIf(lngValue > lngMaxValue, lngValue, lngMaxValue)
            Next i
            If lngMaxValue > 0 Then dB_Value = 10 * Log(lngMaxValue / 2 ^ 31) / Log(10)
    
        Case "Single"
            Dim sngValue As Single
            Dim sngMaxValue As Single
            For i = 0 To mlngReadLength - 1
                sngValue = msngReadData(0, i)
                sngMaxValue = IIf(sngValue > sngMaxValue, sngValue, sngMaxValue)
            Next i
            If sngMaxValue > 0 Then dB_Value = 10 * Log(sngMaxValue) / Log(10)
    
    End Select
    
End Function


