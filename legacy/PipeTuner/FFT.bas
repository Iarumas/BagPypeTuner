Attribute VB_Name = "FFT"
Option Explicit

Private Declare Sub FFTDouble Lib "FFT.dll" Alias "fft_double" _
  (ByVal NumSamples As Long, ByVal InverseTransform As Boolean, _
  realin As Double, _
  imagin As Double, _
  realout As Double, _
  imagout As Double)
 
' Only Double Precision / Single not needed
'Private Declare Sub FFTSingle Lib "FFT.dll" Alias "fft_float" _
  (ByVal NumSamples As Long, ByVal InverseTransform As Boolean, _
  realin As Single, _
  imagin As Single, _
  realout As Single, _
  imagout As Single)
 
'Not needed (simply was in FFT.dll)
'Private Declare Function IndexToFrequency Lib "FFT.dll" _
  Alias "Index_to_frequency" _
  (ByVal NumSamples As Long, _
  ByVal index As Long) As Double

Private Const mdblTwoPi As Double = 6.28318530717959
Private Const mdblPi As Double = 3.14159265358979

Private mlngDataLength As Long

Private mstrWindowType As String
Private mintGaussOrder As Integer

Private mdblWindowFunction() As Double
Private mdblPowerSpectrum() As Double
Private mdblPhaseSpectrum() As Double
Private mdblRealIn() As Double
Private mdblImagIn() As Double
Private mdblRealOut() As Double
Private mdblImagOut() As Double

Public Sub SetWindowType(ByVal Value As String)
    mstrWindowType = Value
End Sub
Public Function WindowType() As String
    WindowType = mstrWindowType
End Function

Public Sub SetLength(ByVal Value As Long)
    mlngDataLength = Value
End Sub
Public Function Length() As Long
    Length = mlngDataLength
End Function

Public Function WindowFunction() As Double()
    WindowFunction = mdblWindowFunction
End Function
Public Sub SetGaussOrder(ByVal Value As Integer)
    mintGaussOrder = Value
End Sub
Public Function GaussOrder() As Integer
    GaussOrder = mintGaussOrder
End Function
Public Function GaussSigma() As Double
    GaussSigma = mdblGaussSigma
End Function

Public Function Real() As Double()
    Real = mdblRealOut
End Function
Public Function Imaginary() As Double()
    Imaginary = mdblImagOut
End Function

Public Sub Init()
' has to be run first: to check sample length, define arrays, calculate window function, ...
    
    Dim lngI As Long
    
    If mlngDataLength = 0 Then mlngDataLength = 8192        'Default sample length
    If mstrWindowType = "" Then mstrWindowType = "Gauss"    'Default Gaussian window
    If mintGaussOrder = 0 Then mintGaussOrder = 8           'Default order of Gauss function
    
    Select Case mlngDataLength
        Case 128, 256, 512, 1024, 2048, 4096, 8192, 16384, 32768, 65536
        Case Is < 128
            MsgBox ("size is too small")
            Exit Sub
        Case Is > 65536
            MsgBox ("size is too large")
            Exit Sub
        Case Else
            MsgBox ("size is not 2^n")
            Exit Sub
    End Select
            
    ReDim mdblWindowFunction(0 To mlngDataLength - 1)
    ReDim mdblRealOut(0 To mlngDataLength - 1)
    ReDim mdblImagOut(0 To mlngDataLength - 1)
    ReDim mdblRealIn(0 To mlngDataLength - 1)
    ReDim mdblImagIn(0 To mlngDataLength - 1)
    ReDim mdblPowerSpectrum(0 To mlngDataLength / 2)
    ReDim mdblPhaseSpectrum(0 To mlngDataLength / 2)
    
            
    If mstrWindowType = "Gauss" Then
        Select Case mintGaussOrder
            Case 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16
                
            Case Is < 1
                MsgBox ("min. Gauss Order is 1")
                Exit Sub
            Case Is > 16
                MsgBox ("max. Gauss Order is 16")
                Exit Sub
            Case Else
                MsgBox ("not an integer value")
                Exit Sub
        End Select
    End If
    
    ' 3 types of windowfuctions for FFT are supported: "Hanning", "Hamming" and most important "Gauss"
    Select Case mstrWindowType
        
        Case "Rectangle"             ' No window function, just 1
            For lngI = 0 To mlngDataLength - 1
                mdblWindowFunction(lngI) = 1
            Next lngI
            
        Case "Hanning"      ' Hanning
            For lngI = 0 To mlngDataLength - 1
                mdblWindowFunction(lngI) = 0.5 * (1 - Cos(mdblTwoPi * lngI / mlngDataLength))
            Next lngI
            
        Case "Hamming"      ' Hamming
            For lngI = 0 To mlngDataLength - 1
                mdblWindowFunction(lngI) = 0.54 - 0.46 * Cos(mdblTwoPi * lngI / mlngDataLength)
            Next lngI
            
        Case "Gauss"        ' Gaussian window with n. order
            For lngI = 0 To mlngDataLength - 1
                mdblWindowFunction(lngI) = Exp(-(lngI - mlngDataLength / 2) ^ 2 / (2 * (mlngDataLength / mintGaussOrder) ^ 2))
            Next lngI
        Case Else
            MsgBox ("Window Type Not Supported")
            Exit Sub
            
    End Select
    
End Sub

Public Function PowerSpectrum(ByRef data, Optional ByVal intChannel As Integer = 0) As Double()

    Dim lngI As Long
       
    'FFT with windowfunction
    For lngI = 0 To mlngDataLength - 1
        mdblRealIn(lngI) = CDbl(data(intChannel, lngI) * mdblWindowFunction(lngI))
    Next lngI
   
    '------------Call FFT ...............
    Call FFTDouble(mlngDataLength, False, mdblRealIn(0), mdblImagIn(0), mdblRealOut(0), mdblImagOut(0))

    ' PowerSpectrum = square root (Re^2 + Im^2)
    For lngI = 0 To mlngDataLength / 2
        mdblPowerSpectrum(lngI) = Sqr(mdblRealOut(lngI) ^ 2 + mdblImagOut(lngI) ^ 2) / mlngDataLength
    Next lngI
       
    PowerSpectrum = mdblPowerSpectrum

End Function

Public Function PhaseSpectrum() As Double()

    Dim lngI As Long
    ReDim mdblPhaseSpectrum(0 To UBound(mdblPowerSpectrum))
    
    For lngI = 0 To UBound(mdblPhaseSpectrum)
        Select Case mdblRealOut(lngI)
            ' Re = 0 then phase = +/- PI/2
            Case Is = 0: mdblPhaseSpectrum(lngI) = Sgn(mdblImagOut(lngI)) * mdblPi / 2
            ' Re  > 0 then -PI/2 < phase < PI/2
            Case Is > 0: mdblPhaseSpectrum(lngI) = Atn(mdblImagOut(lngI) / mdblRealOut(lngI))
            ' Re < 0 then -PI < phase < PI/2 or  PI/2 < phase < PI
            Case Is < 0: mdblPhaseSpectrum(lngI) = Atn(mdblImagOut(lngI) / mdblRealOut(lngI)) + Sgn(mdblImagOut(lngI)) * mdblPi
        End Select
    Next lngI
    
    PhaseSpectrum = mdblPhaseSpectrum
    
End Function

Public Function InverseFFT() As Double()
'not needed at all

    Dim lngI As Long
    
    For lngI = 0 To mlngDataLength / 2
        mdblRealIn(lngI) = mdblPowerSpectrum(lngI) * Cos(mdblPhaseSpectrum(lngI))
        mdblImagIn(lngI) = mdblPowerSpectrum(lngI) * Sin(mdblPhaseSpectrum(lngI))
    Next lngI
    
    Call FFTDouble(mlngDataLength, True, mdblRealIn(0), mdblImagIn(0), mdblRealOut(0), mdblImagOut(0))
    
    For lngI = 0 To mlngDataLength / 2
        mdblPowerSpectrum(lngI) = Sqr(mdblRealOut(lngI) ^ 2 + mdblImagOut(lngI) ^ 2) / mlngDataLength
    Next lngI

End Function




