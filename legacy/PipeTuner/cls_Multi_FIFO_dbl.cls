VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_Multi_FIFO_dbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)
    
' define type of buffer : integer, long, single, double
' here: double
Private mBuffer() As Double
Private mInput() As Double
Private mOutput() As Double
' these are the only variables that have to be set to int, lng or sng
' there are also properties, subs and functions to be adjusted

Private mdblResult() As Double
Private mdblTotal() As Double
Private mdblValidAverage() As Double
Private mlngValidElements() As Long
Private mlngElements As Long
Private mintDimensions As Integer
Private mlngMaxNonValidElements As Long
Private mlngPosition As Long
Private mlngStartPosition As Long


Public Property Let FIFO_Elements(ByVal Value As Long)
    'Number of Elements / FIFO Length
    mlngElements = Value
    If Value <= mlngMaxNonValidElements Then
        mlngMaxNonValidElements = mlngElements - 1
        'at least one element must be valid -> non valid elements = elements-1
    End If
    If Value < 0 Then
        mlngElements = 1
        MsgBox ("Minimum is 1")
    End If
    ReDim mBuffer(0 To mlngElements - 1, mintDimensions - 1)
End Property
Public Property Get FIFO_Elements() As Long
    'number of elements
    FIFO_Elements = mlngElements
End Property
Public Property Let FIFO_Dimensions(ByVal Value As Integer)
    'Number of Elements / FIFO Length
    mintDimensions = Value
    
    ReDim mBuffer(0 To mlngElements - 1, 0 To mintDimensions - 1)
    ReDim mInput(0 To mintDimensions - 1)
    ReDim mOutput(0 To mintDimensions - 1)
    
    ReDim mdblResult(0 To mintDimensions - 1)
    ReDim mdblTotal(0 To mintDimensions - 1)
    ReDim mdblValidAverage(0 To mintDimensions - 1)
    ReDim mlngValidElements(0 To mintDimensions - 1)
    
End Property
Public Property Get FIFO_Dimensions() As Integer
    'number of elements
    FIFO_Dimensions = mintDimensions
End Property
Property Let FIFO_MaxNonValidElements(ByVal Value As Long)
    'maximum of non valid elements
    mlngMaxNonValidElements = Value
    If mlngMaxNonValidElements >= mlngElements Then
        'at least one elements must be valid -> MaxNonValidElements <= Elements -1
        mlngMaxNonValidElements = mlngElements - 1
        MsgBox ("Maximum is Number of Elements - 1")
    End If
    If mlngMaxNonValidElements < 0 Then
        'must be >= 0
        mlngMaxNonValidElements = 0
        MsgBox ("Minimum is 0")
    End If
End Property
Public Property Get FIFO_MaxNonValidElements() As Long
    'maximum of non valid elements
    FIFO_MaxNonValidElements = mlngMaxNonValidElements
End Property
Public Property Let FIFO_Position(ByVal Value As Long)
    'postion of next input to FIFO
    mlngPosition = Value
End Property
Public Property Get FIFO_Position() As Long
    'postion of next input to FIFO
    FIFO_Position = mlngPosition
End Property
Public Property Let FIFO_StartPosition(ByRef Value As Long)
    ' start postion of FIFO
    mlngStartPosition = Value
End Property
Public Property Get FIFO_StartPosition() As Long
    ' start postion of FIFO
    FIFO_StartPosition = mlngStartPosition
End Property

Public Property Get FIFO_ValidAverage() As Double()
    'average of valid elements
    FIFO_ValidAverage = mdblValidAverage
End Property
Public Property Get FIFO_Total() As Double()
    ' sum of all elements
    FIFO_Total = mdblTotal
End Property
Public Property Get FIFO_ValidElements() As Long()
    ' number of valid elements
    FIFO_ValidElements = mlngValidElements
End Property
' define type of buffer : integer, long, single, double
Public Property Get FIFO_Input() As Double()
    'current input to FIFO
    FIFO_Input = mInput
End Property
Public Property Get FIFO_Result() As Double()
    'result is the valid average if non valid elements < MaxNonValidElements
    FIFO_Result = mdblResult
End Property

' define type of buffer : integer, long, single, double
Public Sub Set_FIFO_Buffer(ByRef Value() As Double, Optional ByVal intDim)
    'FIFO content
    Dim intDimension As Integer
    
    mlngElements = UBound(Value, 1) + 1
    If mlngMaxNonValidElements >= mlngElements Then
        'at least one elements must be valid -> MaxNonValidElements <= Elements -1
        mlngMaxNonValidElements = mlngElements - 1
        MsgBox ("Maximum is Number of Elements - 1")
    End If
    
    If IsMissing(intDim) Then
        mintDimensions = UBound(Value, 2) + 1
        FIFO_Clear
        mBuffer = Value
        ' caluclate the buffer values: total, valid elements, ....
        For intDimension = 0 To mintDimensions - 1
            Calculate_Total (intDimension)
            Calculate_ValidAverage (intDimension)
            Calculate_Result (intDimension)
        Next intDimension
    Else
        intDimension = intDim
        CopyMemory mBuffer(0, intDimension), Value(0), Len(Value(0)) * mlngElements
        ' caluclate the buffer values: total, valid elements, ....
        Calculate_Total (intDimension)
        Calculate_ValidAverage (intDimension)
        Calculate_Result (intDimension)
    End If
    
    ' buffer is full
    mlngPosition = mlngElements
    
End Sub

' define type of buffer : integer, long, single, double
Public Function FIFO_Buffer(Optional ByVal intDim) As Double()
    'FIFO content
    
    ' define type of buffer : integer, long, single, double
    Dim Buffer() As Double
    Dim intDimension As Integer
    
    
    If IsMissing(intDim) Then
        FIFO_Buffer = mBuffer
    Else
        intDimension = intDim
        ReDim Buffer(0 To mlngElements - 1)
        CopyMemory Buffer(0), mBuffer(0, intDimension), Len(mBuffer(0, 0)) * mlngElements
        FIFO_Buffer = Buffer
    End If
    
End Function

Private Sub Class_Initialize()
    
    mlngElements = 1
    mintDimensions = 1
    mlngMaxNonValidElements = 0
    Call FIFO_Clear
    
End Sub

Public Sub FIFO_Clear()
    
    Dim intDimension As Integer
    
    ReDim mBuffer(0 To mlngElements - 1, 0 To mintDimensions - 1)
    ReDim mInput(0 To mintDimensions - 1)
    ReDim mOutput(0 To mintDimensions - 1)
    
    ReDim mdblTotal(0 To mintDimensions - 1)
    ReDim mlngValidElements(0 To mintDimensions - 1)
    ReDim mdblValidAverage(0 To mintDimensions - 1)
    ReDim mdblResult(0 To mintDimensions - 1)
    
    mlngPosition = 0   ' if buffer start from 0
    
    For intDimension = 0 To mintDimensions - 1
        mdblTotal(intDimension) = 0
        mlngValidElements(intDimension) = 0
        mdblValidAverage(intDimension) = 0
        mdblResult(intDimension) = 0
    Next intDimension
    
End Sub

' define type of buffer : integer, long, single, double
Public Function FIFO_Fill(ByRef NewValue) As Double()
    
    'the FIFO buffer is filled with from the output side
    ' when buffer is full the elements are moved
    
    Dim intDimensions As Integer
    
    ' input
    mInput = NewValue
        
    Select Case mlngPosition
            
        Case Is < mlngElements
            ' fill buffer with first elements
            For intDimensions = 0 To mintDimensions - 1
                mBuffer(mlngPosition, intDimensions) = mInput(intDimensions)
            Next intDimensions
            mlngPosition = mlngPosition + 1   'when not yet full
            mlngStartPosition = 0
            
        Case mlngElements
            ' output
            For intDimensions = 0 To mintDimensions - 1
                mOutput(intDimensions) = mBuffer(0, intDimensions)
                ' move/push values: destination, source, byte length  / and add input
                If mlngElements > 1 Then
                    CopyMemory mBuffer(0, intDimensions), mBuffer(1, intDimensions), Len(mBuffer(0, 0)) * (mlngElements - 1)
                End If
                ' input
                mBuffer(mlngElements - 1, intDimensions) = mInput(intDimensions)
            Next intDimensions
            mlngStartPosition = mlngStartPosition + 1
            
        Case Else
            FIFO_Fill = Array(0)
            Exit Function
        
    End Select
        
    For intDimensions = 0 To mintDimensions - 1
        
        mdblTotal(intDimensions) = mdblTotal(intDimensions) - CDbl(mOutput(intDimensions)) + CDbl(mInput(intDimensions))      'remove old value, add new value

        'number of valid elements means non zero elements
        If mInput(intDimensions) = 0 And mOutput(intDimensions) <> 0 Then mlngValidElements(intDimensions) = mlngValidElements(intDimensions) - 1
        If mInput(intDimensions) <> 0 And mOutput(intDimensions) = 0 Then mlngValidElements(intDimensions) = mlngValidElements(intDimensions) + 1
    
        Calculate_ValidAverage (intDimensions)
        Calculate_Result (intDimensions)

    Next intDimensions
    
    'Dim i As Integer
    'For i = 0 To mlngElements - 1
    '    Debug.Print mBuffer(i),
    'Next i
    'Debug.Print
    'Debug.Print mlngPosition, mInput, mOutput
    'Debug.Print mdblTotal, mlngValidElements, mdblValidAverage
    'Debug.Print mlngMaxNonValidElements, Result

    FIFO_Fill = mOutput       'return output

End Function

Private Sub Calculate_Total(ByVal intDimension As Integer)
    ' sum of all elements
    Dim i As Long
    
    For i = 0 To mlngElements - 1
        If mBuffer(i, intDimension) <> 0 Then
            mdblTotal(intDimension) = mdblTotal(intDimension) + CDbl(mBuffer(i, intDimension))
            mlngValidElements(intDimension) = mlngValidElements(intDimension) + 1
        End If
    Next i
        
End Sub
Private Sub Calculate_ValidAverage(ByVal intDimension As Integer)
    ' average of valid elements
    
    mdblValidAverage(intDimension) = 0
    If mlngValidElements(intDimension) > 0 Then mdblValidAverage(intDimension) = mdblTotal(intDimension) / mlngValidElements(intDimension)
    
End Sub
Private Sub Calculate_Result(ByVal intDimension As Integer)
    'result = 0 if there are more non valid elements than specified by MaxNonValidElements
    
    mdblResult(intDimension) = 0
    If mlngValidElements(intDimension) >= mlngElements - mlngMaxNonValidElements Then mdblResult(intDimension) = mdblValidAverage(intDimension)

End Sub

