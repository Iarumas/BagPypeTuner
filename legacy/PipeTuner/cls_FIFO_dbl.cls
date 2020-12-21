VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_FIFO_dbl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)
    
Private mdblBuffer() As Double
Private mdblInput As Double
Private mdblOutput As Double

Private mdblResult As Double
Private mdblTotal As Double
Private mdblValidAverage As Double
Private mlngElements As Long
Private mlngValidElements As Long
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
    ReDim mdblBuffer(0 To mlngElements - 1)
End Property
Public Property Get FIFO_Elements() As Long
    'number of elements
    FIFO_Elements = mlngElements
End Property
Public Property Get FIFO_ValidAverage() As Double
    'average of valid elements
    FIFO_ValidAverage = mdblValidAverage
End Property
Public Property Get FIFO_Total() As Double
    ' sum of all elements
    FIFO_Total = mdblTotal
End Property
Public Property Get FIFO_ValidElements() As Long
    ' number of valid elements
    FIFO_ValidElements = mlngValidElements
End Property
Public Property Get FIFO_Input() As Double
    'current input to FIFO
    FIFO_Input = mdblInput
End Property
Public Property Get FIFO_Result() As Double
    'result is the valid average if non valid elements < MaxNonValidElements
    FIFO_Result = mdblResult
End Property
Public Property Let FIFO_MaxNonValidElements(ByVal Value As Long)
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
Public Property Let FIFO_Position(ByRef Value As Long)
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
Public Property Let FIFO_Buffer(ByRef Value() As Double)
    'FIFO content
    mlngElements = UBound(Value) + 1
    FIFO_Clear
    mdblBuffer = Value
    
    If mlngMaxNonValidElements >= mlngElements Then
        'at least one elements must be valid -> MaxNonValidElements <= Elements -1
        mlngMaxNonValidElements = mlngElements - 1
        MsgBox ("Maximum is Number of Elements - 1")
    End If
    
    ' caluclate the buffer values: total, valid elements, ....
    Dim i As Long
    For i = 0 To mlngElements - 1
        If mdblBuffer(i) <> 0 Then
            mdblTotal = mdblTotal + mdblBuffer(i)
            mlngValidElements = mlngValidElements + 1
        End If
    Next i

    ' average of valid elements
    mdblValidAverage = 0
    If mlngValidElements > 0 Then mdblValidAverage = mdblTotal / mlngValidElements

    'result = 0 if there are more non valid elements than specified by MaxNonValidElements
    mdblResult = 0
    If mlngValidElements >= mlngElements - mlngMaxNonValidElements Then mdblResult = mdblValidAverage

    ' buffer is full
    mlngPosition = mlngElements
    
End Property
Public Property Get FIFO_Buffer() As Double()
    'FIFO content
    FIFO_Buffer = mdblBuffer
End Property

Private Sub Class_Initialize()
    
    mlngElements = 1
    mlngMaxNonValidElements = 0
    Call FIFO_Clear
    
End Sub

Public Sub FIFO_Clear()

    ReDim mdblBuffer(0 To mlngElements - 1)
    
    mdblInput = 0
    mdblOutput = 0
    mdblTotal = 0
    mlngValidElements = 0
    mdblValidAverage = 0
    mdblResult = 0
    mlngPosition = 0   ' if buffer start from 0
    mlngStartPosition = 0   ' if buffer start from 0
    
End Sub

Public Function FIFO_Fill(dblNewValue As Double) As Double
    
    'the FIFO buffer is filled with from the output side
    ' when buffer is full the elements are moved
    
    ' input
    mdblInput = dblNewValue
        
    Select Case mlngPosition
            
        Case Is < mlngElements
            ' fill buffer with first elements
            mdblBuffer(mlngPosition) = mdblInput
            mlngPosition = mlngPosition + 1   'when not yet full
            mlngStartPosition = 0
            
        Case mlngElements
            ' output
            mdblOutput = mdblBuffer(0)
            ' move/push values: destination, source, byte length  / and add input
            If mlngElements > 1 Then
                Call CopyMemory(mdblBuffer(0), mdblBuffer(1), (mlngElements - 1) * Len(mdblInput))
            End If
            ' input
            mdblBuffer(mlngElements - 1) = mdblInput
            mlngStartPosition = mlngStartPosition + 1
            
        Case Else
            FIFO_Fill = 0
            Exit Function
        
    End Select
         
    mdblTotal = mdblTotal - mdblOutput + mdblInput      'remove old value, add new value

    'number of valid elements means non zero elements
    If mdblInput = 0 And mdblOutput <> 0 Then mlngValidElements = mlngValidElements - 1
    If mdblInput <> 0 And mdblOutput = 0 Then mlngValidElements = mlngValidElements + 1
    
    ' average of valid elements
    mdblValidAverage = 0
    If mlngValidElements > 0 Then mdblValidAverage = mdblTotal / mlngValidElements
    
    'result = 0 if there are more non valid elements than specified by MaxNonValidElements
    mdblResult = 0
    If mlngValidElements >= mlngElements - mlngMaxNonValidElements Then mdblResult = mdblValidAverage
        
    'Dim i As Integer
    'For i = 0 To mlngElements - 1
    '    Debug.Print mdblBuffer(i),
    'Next i
    'Debug.Print
    'Debug.Print mlngPosition, mdblInput, mdblOutput
    'Debug.Print mdblTotal, mlngValidElements, mdblValidAverage
    'Debug.Print mlngMaxNonValidElements, mdblResult

    FIFO_Fill = mdblOutput       'return output

End Function


