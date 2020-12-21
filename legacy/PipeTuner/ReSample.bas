Attribute VB_Name = "ReSample"
Option Explicit


Public Offset As Long
Private Const mdblTwoPi As Double = 6.28318530717959

Public OrgTime() As Long
Public dblOrgTime() As Double
Public OrgSample() As Integer
Public dblOrgSample() As Double
Public ReSampled() As Integer
Public dblReSampled() As Double
Public SplineCoefficients() As Double

Private TestWaveIn As New cls_WavFile
Private TestWaveOut As New cls_WavFile

Private Declare Sub FFTDouble Lib "FFT.dll" Alias "fft_double" _
  (ByVal NumSamples As Long, ByVal InverseTransform As Boolean, _
  realin As Double, _
  imagin As Double, _
  realout As Double, _
  imagout As Double)
 

Public Function ReSample(ByRef OrgSample, ByVal lngFromSampleRate, ByVal lngToSampleRate, Optional ByVal Offset) As Integer()

    Dim i, j As Long
    Dim Z, p As Double
    Dim Factor As Double
    Dim Elements As Long
    Dim intReSampled() As Integer
    
    Factor = lngToSampleRate / lngFromSampleRate
    'Factor = 2 ^ (Cent / 1200)
    Elements = UBound(OrgSample) - LBound(OrgSample) + 1
    
    
    ReDim intReSampled(0 To Elements - 1)
    
    For i = 0 To Elements - 1
    
        j = (i - Offset + 1) * Factor + Offset - 1
        Z = Int(j)
        p = j - Z
        
        If Z >= 0 And Z < Elements - 1 Then
            intReSampled(i) = CInt(OrgSample(Z) * (1 - p) + OrgSample(Z + 1) * p)
        End If
    
    Next i
    
    ReSample = intReSampled

End Function

Public Sub ReSampleTest()

Dim i As Long

TestWaveIn.Init
TestWaveOut.Init
TestWaveOut.ReadWrite = True

TestWaveIn.FileName = "D:\Profiles\ABGRAUST\My Documents\Pipes\FFT\Samples 44kHz\boum mono.wav"

If TestWaveIn.Exists Then
    TestWaveIn.FileName = "D:\Profiles\ABGRAUST\My Documents\Pipes\FFT\Samples 44kHz\boum mono.wav"
    TestWaveOut.FileName = "D:\Profiles\ABGRAUST\My Documents\Pipes\FFT\Samples 44kHz\boum mono test.wav"
Else
    TestWaveIn.FileName = "C:\DATA\Pipes\FFT\Samples 44kHz\boum mono.wav"
    TestWaveOut.FileName = "C:\DATA\Pipes\FFT\Samples 44kHz\boum mono test.wav"
End If

TestWaveOut.Delete

TestWaveIn.OpenFile
TestWaveIn.GetWavFileInfo

TestWaveIn.SampleStart = 0
TestWaveIn.SampleLength = TestWaveIn.SampleRate * 20 '10s
'TestWaveIn.SampleLength = 300

ReDim OrgTime(0 To TestWaveIn.SampleLength - 1)
ReDim dblOrgTime(0 To TestWaveIn.SampleLength - 1)
ReDim OrgSample(0 To TestWaveIn.SampleLength - 1)
ReDim dblOrgSample(0 To TestWaveIn.SampleLength - 1)
ReDim ReSampled(0 To TestWaveIn.SampleLength - 1)
ReDim dblReSampled(0 To TestWaveIn.SampleLength - 1)

'TestWaveIn.OpenFile
TestWaveIn.ReadWavData
OrgSample = TestWaveIn.ReadData
TestWaveIn.CloseFile

For i = 0 To TestWaveIn.SampleLength - 1
    OrgTime(i) = i
    dblOrgTime(i) = CDbl(OrgTime(i))
    dblOrgSample(i) = CDbl(OrgSample(i))
Next i

Call BuildCubicSpline(dblOrgTime, dblOrgSample, TestWaveIn.SampleLength, 2, 0, 2, 0, SplineCoefficients)

For i = 0 To TestWaveIn.SampleLength - 1
    dblReSampled(i) = SplineInterpolation(SplineCoefficients, i * 0.99)
    ReSampled(i) = CInt(IIf(dblReSampled(i) > 32767, 32767, dblReSampled(i)))
Next i

'ReSampled = ReSample(OrgSample, TestWaveIn.SampleRate, 48000, 0)

TestWaveOut.BitsPerSample = TestWaveIn.BitsPerSample
TestWaveOut.Channels = TestWaveIn.Channels
TestWaveOut.SampleRate = TestWaveIn.SampleRate


TestWaveOut.IntegerData = ReSampled
TestWaveOut.OpenFile
TestWaveOut.WriteHeader
TestWaveOut.WriteWavData
TestWaveOut.CloseFile

End Sub

Public Sub test()

    Dim i As Long
    Dim nnn As Long
    Dim n0 As Double
    Dim dblFreq As Double
    Dim sig As Double
    Dim orgdata() As Double
    Dim realin() As Double
    Dim imagin() As Double
    Dim realout() As Double
    Dim imagout() As Double
    Dim dblSpectrum() As Double
    Dim dblPhase() As Double
    
    nnn = 128
    n0 = CDbl(nnn * 3 / 4)
    sig = CDbl(nnn / 8)
    
    ReDim orgdata(0 To 0, 0 To nnn - 1)
    ReDim realin(0 To nnn - 1)
    ReDim imagin(0 To nnn - 1)
    ReDim realout(0 To nnn - 1)
    ReDim imagout(0 To nnn - 1)

    dblFreq = 3.4
    For i = 0 To nnn - 1
        'orgdata(0, i) = 100 * Cos(mdblTwoPi * dblFreq * i / nnn)
        orgdata(0, i) = 100
        realin(i) = orgdata(0, i)
    Next i
    
    Call FFTDouble(nnn, False, realin(0), imagin(0), realout(0), imagout(0))
    
    FFT.SetWindowType "Rectangle"
    FFT.SetLength nnn
    FFT.Init
    dblSpectrum = FFT.PowerSpectrum(orgdata, 0)
    dblPhase = FFT.PhaseSpectrum
    
    For i = 0 To nnn / 2
        Debug.Print Format(dblSpectrum(i), "00000.000000"), _
                    Format(dblPhase(i) * 360 / mdblTwoPi, "000.00000")
    Next i
    Debug.Print
    
    For i = 1 To nnn / 2 - 1
        realout(nnn / 2 + i) = realout(nnn / 2 - i)
        imagout(nnn / 2 + i) = -1 * imagout(nnn / 2 - i)
    Next i
    
    Call FFTDouble(nnn, True, realout(0), imagout(0), realin(0), imagin(0))
    
    For i = 0 To nnn - 1
    Debug.Print Format(orgdata(0, i), "00000.000000"), _
                Format(realout(i), "00000.000000"), Format(imagout(i), "00000.0000"), _
                Format(realin(i), "00000.0000"), Format(imagin(i), "00000.0000")

    Next i
    Debug.Print
    
End Sub
