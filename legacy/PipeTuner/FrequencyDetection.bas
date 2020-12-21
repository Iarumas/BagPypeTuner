Attribute VB_Name = "FrequencyDetection"
Option Explicit

' global variables needed:
' gudtDrones
' WavFile.SampleLength
' WavFile.SampleRate

Private mlngSampleRate As Long                          ' Sample Rate
Private mlngSampleLength As Long                        ' Sample Length

Private mdblSpectrum() As Double                        ' Spectrum
Private mdblRefFreq As Double                           ' reference frequency
Private mdblDroneRatios() As Double                     ' ratio of drone freqs with reference to RefFreq
Private mdblFrequencies() As Double

Private mlngPeakRange As Long
Private mlngIndexRangeBasicFreq(0 To 1) As Long

Private mdblIndexFreqs() As Double
Private mdblAmplitudes() As Double


Public Sub Set_Spectrum(ByRef Value() As Double)
' only used if the spectrum has to be displayed
    mdblSpectrum = Value()
End Sub
Public Function Spectrum() As Double()
'returns array containing the power spectrum
    Spectrum = mdblSpectrum
End Function

Public Sub Set_RefFreq(ByVal Value As Double)
' sets the reference frequency
    mdblRefFreq = Value
End Sub
Public Function RefFreq() As Double
'returns the reference frequency
    RefFreq = mdblRefFreq
End Function

Public Sub Set_DroneRatios(ByRef Value() As Double)
' Sets ratio for drones (with reference to RefFreq)
    Set_DroneRatios = mdblDroneRatios()
End Sub

Public Function Frequencies() As Double()
' returns array with measured frequencies
    Frequencies = mdblFrequencies
End Function

Public Sub Init()
    
    'global variables:
    'gudtDrones
    
    Dim lngCounter As Long
       
    ' number of frequencies:
    '0: chanter
    '1: bass drone
    '2: tenor drone 1
    '3: tenor drone 2 : not used, only usefull with contact microphone
    ReDim mdblFrequencies(LBound(gudtDrones) To UBound(gudtDrones))
    ReDim mdblDroneRatios(LBound(gudtDrones) To UBound(gudtDrones))
       
    For lngCounter = LBound(mdblDroneRatios) To UBound(mdblDroneRatios)
        mdblDroneRatios(lngCounter) = gudtDrones(lngCounter).Ratio
    Next lngCounter
    
    
    ' set parameters for FFT
    mlngSampleLength = WavFile.SampleLength
    mlngSampleRate = WavFile.SampleRate
    
    FFT.SetWindowType "Gauss"
    FFT.SetGaussOrder 8
    FFT.SetLength mlngSampleLength
    FFT.Init
    
   
End Sub
Public Function MeasureFrequencies(ByRef intData) As Double()
' intDATA is currently an array with only 1 dimension: audio channels 0/1 left/rigth are mixed
' but this only allows 1 tenor drone to be running. If both tenors are running the frequencies cannot be measured properly
' if you have 2 audio channel you could also define channel 0 as chanter/bass/tenor1 and channel 1 as tenor 2
' if you have 4 audio chaneld you could also define channel 0,1,2,3 as chanter,bass,tenor1,tenor2
    
    ' create spectrum
    mdblSpectrum = FFT.PowerSpectrum(intData, 0)     ' chanter is channel 0
    mdblFrequencies(0) = MeasureChanterFrequency     ' frequncies(0) = chanter
    
    'mdblSpectrum = FFT.PowerSpectrum(intData, 1)    ' if bass is channel 1
    mdblFrequencies(1) = MeasureBassFrequency        ' frequencies(1) = bass
    
    'mdblSpectrum = FFT.PowerSpectrum(intData, 2)    ' if tenor1 is channel 2
    mdblFrequencies(2) = MeasureTenorFrequency       ' frequencies(2) = tenor 1
    
    'mdblSpectrum = FFT.PowerSpectrum(intData, 1)    ' if tenor2 is channel 1
    'mdblSpectrum = FFT.PowerSpectrum(intData, 3)    ' if tenor2 is channel 3
    'mdblFrequencies(3) = MeasureTenorFrequency      ' frequencies(3) = tenor 2
    
    MeasureFrequencies = mdblFrequencies
    
End Function

Public Function MeasureChanterFrequency() As Double
    ' returns chanter frequency

    ' modular variables
    ' mdblFreqRange
    ' mdblBasicFreqRangeMin
    ' mdblBasicFreqRangeMax

    Dim dblFreqRange As Double
    Dim dblBasicFreqRangeMin As Double
    Dim dblBasicFreqRangeMax As Double
    
    
    ' set Parameters for chanter frequency measurement
    
    ' range where to search for the peak
    dblFreqRange = 10 / 480 * mdblRefFreq            ' 5 Hz @ 480 Hz
    ' range where to search for a peak
    mlngPeakRange = Round(dblFreqRange / mlngSampleRate * mlngSampleLength + 0.5)
    
    ' range of expected chanter frequency
    dblBasicFreqRangeMin = 10 / 12 * mdblRefFreq    ' 400 Hz @ 480 Hz
    dblBasicFreqRangeMax = 25 / 12 * mdblRefFreq    ' 1000 Hz @ 480 Hz

    ' index range of basic frequency
    mlngIndexRangeBasicFreq(0) = Round(dblBasicFreqRangeMin / mlngSampleRate * mlngSampleLength)
    mlngIndexRangeBasicFreq(1) = Round(dblBasicFreqRangeMax / mlngSampleRate * mlngSampleLength)
       
    MeasureChanterFrequency = ChanterFreqByHarmonics

End Function

Public Function FindHarmonics(ByRef dblSpectrum, ByVal lngIndexBasicFreq As Long, ByVal dblLimit As Double, ByVal lngPeakRange) As Long
' returns the number of harmonics of basic freq.(index) found above limit
    
    'modular variables:
    
    'mdblIndexFreqs
    'mdbAmplitudes
    
    Dim blnDebug As Boolean
    Dim NoError As Boolean
    
    Dim lngCounter As Long
    Dim lngIndex(0 To 1) As Long                        ' lower/upper limit of indexes
    Dim lngCurrentIndexMax As Long                      ' index of current maximum
    
    Dim lngMaxFactor As Long                            ' max. number of harmonics
    Dim lngFoundElements As Long                        ' found harmonics

    
    blnDebug = False

    ' maximal number of harmonics
    lngMaxFactor = Round(UBound(mdblSpectrum) / 4 / lngIndexBasicFreq)
    'Debug.Print lngMaxFactor
            
    ' resize to number of harmonics
    ReDim mdblIndexFreqs(1 To lngMaxFactor)          'frequency indexes
    ReDim mdblAmplitudes(1 To lngMaxFactor)         'amplitudes
    ReDim mlngGoodHarmonics(1 To lngMaxFactor)      'harmonic (1,2,3,4,...) found above limit, 0 if none founde
    
    ' check all harmonics
    For lngCounter = 1 To lngMaxFactor
                    
        ' range
        lngIndex(0) = Round(lngCounter * (lngIndexBasicFreq - lngPeakRange))
        lngIndex(1) = Round(lngCounter * (lngIndexBasicFreq + lngPeakRange))

        ' current max in range
        lngCurrentIndexMax = IndexOfMaxInRange(mdblSpectrum, lngIndex(0), lngIndex(1))
        ' gauss fit
        NoError = GaussFit.Fit(lngCurrentIndexMax, mdblSpectrum)
                
        ' count harmonics that have certain amplitude
        If NoError And GaussFit.Amplitude > dblLimit Then

                ' harmonic
                mlngGoodHarmonics(lngCounter) = lngCounter
                mdblAmplitudes(lngCounter) = GaussFit.Amplitude
                mdblIndexFreqs(lngCounter) = GaussFit.Center / lngCounter
                ' count harmonics
                lngFoundElements = lngFoundElements + 1
                        
        End If
                
        blnDebug = False
        If blnDebug Then
            Debug.Print Format(lngCounter, "000"), _
                        Format(lngIndex(0), "0000"), Format(lngIndexBasicFreq, "0000"), Format(lngIndex(1), "0000"), _
                        Format(GaussFit.Center, "0000.000"), _
                        Format(mdblIndexFreqs(lngCounter), "000.000"), _
                        Format(mdblAmplitudes(lngCounter), "000.000"), _
                        Format(GaussFit.Sigma, "000.000")
        End If
        blnDebug = False
            
    Next lngCounter
    
    FindHarmonics = lngFoundElements
    
End Function

Public Function ChanterFreqByHarmonics() As Double

    'modular variables:
    
    'mdblIndexFreqs
    'mdblAmplitudes
    
    'mlngIndexRangeBasicFreq
    'mlngPeakRange
    
    'mdblSpectrum
    'mlngSampleRate
    'mlngSampleLength
    
    Dim lngI As Long
    
    Dim blnDebug As Boolean
    Dim NoError As Boolean
    Dim lngCounter As Long
    
    Dim lngIndex(0 To 1) As Long                        ' lower/upper limit of indexes
    
    Dim lngDivider(0 To 1) As Long                       ' range for factors n*fo to m*fo = f with max. amplitude
    Dim lngDividerValue As Long
    
    Dim lngBasicIndex() As Long                         ' index of basic freq.
    Dim lngCurrentIndexMax As Long                      ' index of current maximum
    
    Dim lngMaxFactor As Long                            ' max. harmonic
    Dim lngMultiBasicIndex As Long
    
    Dim dblMaxAmp As Double                             ' max. amplitude
    Dim dblMaxAmpPosition As Double                     ' position/index of max. amplitude (not lng but dbl)

    Dim lngBestDivider As Long
    Dim lngMaxElements As Long
    Dim lngElements() As Long
    
    Dim dblIndexReg() As Double
    Dim dblRegressionResults() As Double
    
    blnDebug = False
    
    lngCurrentIndexMax = IndexOfMaxInRange(mdblSpectrum, LBound(mdblSpectrum), UBound(mdblSpectrum) / 4)
    NoError = GaussFit.Fit(lngCurrentIndexMax, mdblSpectrum)
    
    If Not NoError Then
        ChanterFreqByHarmonics = 0
        Exit Function
    End If
    
    
    ' center position and amplitude
    dblMaxAmp = GaussFit.Amplitude
    dblMaxAmpPosition = GaussFit.Center

    ' max ratio of max. harmonic / min. basic frequency
    lngDivider(1) = Round(lngCurrentIndexMax / mlngIndexRangeBasicFreq(0))
    ' max ratio of max. harmonic / min. basic frequency
    lngDivider(0) = Round(lngCurrentIndexMax / mlngIndexRangeBasicFreq(1))
    If lngDivider(0) = 0 Then lngDivider(0) = 1
        
        
    ReDim dblIndexReg(lngDivider(0) To lngDivider(1))
    ReDim dblFreqReg(lngDivider(0) To lngDivider(1))
    ReDim lngElements(lngDivider(0) To lngDivider(1))
    ReDim lngBasicIndex(lngDivider(0) To lngDivider(1))
    
    ' look for frequencies which are a fraction of max. harmonic
    For lngDividerValue = lngDivider(0) To lngDivider(1)
        
        
        ' check if basic index is within range of basic frequency index
 '       If dblMaxAmpPosition / lngDividerValue >= mlngIndexRangeBasicFreq(0) And dblMaxAmpPosition / lngDividerValue <= mlngIndexRangeBasicFreq(1) Then
            
            lngBasicIndex(lngDividerValue) = Round(dblMaxAmpPosition / lngDividerValue)
                  
            ' check all harmonics
            lngElements(lngDividerValue) = FindHarmonics(mdblSpectrum, lngBasicIndex(lngDividerValue), 0.2 * dblMaxAmp, mlngPeakRange)

             ' do regression over all harmonics
            dblRegressionResults = RegressionAverage(mdblIndexFreqs, mdblAmplitudes)
            dblIndexReg(lngDividerValue) = dblRegressionResults(0)
            
  '      End If

    Next lngDividerValue
    
    ' check for which harmonic the most elements were found -> this is the best divider
    lngMaxElements = 0
    For lngDividerValue = lngDivider(0) To lngDivider(1)
        If lngElements(lngDividerValue) > lngMaxElements Then
            lngMaxElements = lngElements(lngDividerValue)
            lngBestDivider = lngDividerValue
        End If
    Next lngDividerValue
    
    ' no harmonics found (except 1 peak)
    If lngMaxElements = 1 Then
        ChanterFreqByHarmonics = 0
        Exit Function
    End If

    
    ' select the best harmonics / exclude some harmonics that don't match very good
    ' number of harmonics for best divider
    lngElements(lngBestDivider) = FindHarmonics(mdblSpectrum, Round(dblIndexReg(lngBestDivider)), 0.2 * dblMaxAmp, mlngPeakRange)
    lngCounter = lngElements(lngBestDivider)
    
    ' use the best harmonics (min 50% of all harmonics, but min 3)
    Do While lngCounter > Round(0.5 * lngElements(lngBestDivider)) And lngCounter >= 3
        dblRegressionResults = RegressionAverage(mdblIndexFreqs, mdblAmplitudes)
        mdblIndexFreqs(dblRegressionResults(3)) = 0         ' harmonic with max. sigma removed
        mdblAmplitudes(dblRegressionResults(3)) = 0         ' harmonic with max. sigma removed
        lngCounter = lngCounter - 1                         ' new number of harmonics
    Loop
        
    'calculate frequency from index
    ChanterFreqByHarmonics = dblRegressionResults(0) * mlngSampleRate / mlngSampleLength
    
End Function

Public Function MeasureBassFrequency() As Double
    ' returns bass frequency
    ' will be measured at expected peak (bass drone ratio = 1/4 reference frequency)
    
    'modular variables
    'mdblDroneRatios
    
    Dim dblBasicFreq As Double              ' expected basic freq. of drone
    Dim dblFreqRange(0 To 1) As Double      ' range where to look for peak
    Dim lngHarmonics(1 To 3) As Long        ' list of harmonics to be used for freq. detection
    
     ' set Parameters for frequency measurement
    lngHarmonics(1) = 1     'use basic frequency (1. harmonic) for freq. measurement
    lngHarmonics(2) = 0     ' don't use 2. harmonics: interference with tenor
    lngHarmonics(3) = 3     'use 3. harmonic for freq. measurement

    dblFreqRange(0) = -20           'lower range
    dblFreqRange(1) = 30            'upper range
    
    dblBasicFreq = mdblDroneRatios(1) * mdblRefFreq
    
    MeasureBassFrequency = MeasureDroneFrequency(dblBasicFreq, dblFreqRange, lngHarmonics)

End Function

Public Function MeasureTenorFrequency() As Double
    ' returns tenor frequency
    ' will be measured at expected peak freq. with specified range
    
    'modular variables
    'mdblDroneRatios
    
    Dim dblBasicFreq As Double              ' expected basic freq. of drone
    Dim dblFreqRange(0 To 1) As Double      ' range where to look for peak
    Dim lngHarmonics(1 To 1) As Long        ' list of harmonics to be used for freq. detection
    
    ' set Parameters for frequency measurement
    lngHarmonics(1) = 1             ' use only basic frequency (1. harmonic) for freq. measurement

    dblFreqRange(0) = -40           'lower range
    dblFreqRange(1) = 60            'upper range
    
    
    dblBasicFreq = mdblDroneRatios(2) * mdblRefFreq

    MeasureTenorFrequency = MeasureDroneFrequency(dblBasicFreq, dblFreqRange, lngHarmonics)
    
End Function

Public Function MeasureDroneFrequency(ByVal dblBasicFreq, ByRef dblBasicFreqRange, ByRef lngHarmonics) As Double
    ' returns drone frequency
    ' will be measured at expected peak freq. with specified range
       
    'modular variables
    'mdblSpectrum
    'mSampleLength
    'mSampleLength
   
    Dim lngBasicIndex As Long                ' position of expected drone freq.
    Dim lngBasicIndexRange(0 To 1) As Double ' range where to look for basic drone freq.
    Dim dblDronePosition As Double           ' position/index of drone frequency
    
    lngBasicIndex = Round(dblBasicFreq / mlngSampleRate * mlngSampleLength)
    lngBasicIndexRange(0) = Round(dblBasicFreqRange(0) / mlngSampleRate * mlngSampleLength)
    lngBasicIndexRange(1) = Round(dblBasicFreqRange(1) / mlngSampleRate * mlngSampleLength)
    
    dblDronePosition = DronePositionByHarmonics(mdblSpectrum, lngBasicIndex, lngBasicIndexRange, lngHarmonics)
    
    MeasureDroneFrequency = dblDronePosition * mlngSampleRate / mlngSampleLength

End Function

Public Function DronePositionByHarmonics(ByRef dblSpectrum() As Double, ByVal lngBasicIndex, ByRef lngBasicIndexRange, ByRef lngHarmonics) As Double
    ' returns drone frequency using the specified harmonics, the expected frequency, +/- range where to look
    
    Dim NoError As Boolean
    Dim lngCounter As Long                      ' counter
    Dim lngIndex(0 To 1) As Long                ' range of indexes to search for peak
    Dim lngCurrentIndexMax As Long              ' index of current max

    Dim dblIndexFreqs() As Double               ' position/index of measured frequencies
    Dim dblAmplitudes() As Double               ' amplitudes of peaks
    Dim dblRegressionResults() As Double        ' results of regression/average
    
    ReDim dblIndexFreqs(LBound(lngHarmonics) To UBound(lngHarmonics))
    ReDim dblAmplitudes(LBound(lngHarmonics) To UBound(lngHarmonics))
    
      
    For lngCounter = LBound(lngHarmonics) To UBound(lngHarmonics)
    ' run through the harmonics
       
        ' only use valid harmonics
        If lngHarmonics(lngCounter) <> 0 Then
        
            If lngCounter = LBound(lngHarmonics) Then
                'Limits for basic frequency
                lngIndex(0) = (lngBasicIndex + lngBasicIndexRange(0)) * lngHarmonics(lngCounter)
                lngIndex(1) = (lngBasicIndex + lngBasicIndexRange(1)) * lngHarmonics(lngCounter)
            Else
                'Limits for harmonics
                lngIndex(0) = Round(lngHarmonics(lngCounter) * dblIndexFreqs(1) - 2)
                lngIndex(1) = Round(lngHarmonics(lngCounter) * dblIndexFreqs(1) + 2)
            End If
        
            ' search maximum in range
            lngCurrentIndexMax = IndexOfMaxInRange(dblSpectrum, lngIndex(0), lngIndex(1))
        
            ' Gauss fit
            NoError = GaussFit.Fit(lngCurrentIndexMax, dblSpectrum)
        
            ' calculate frequencies and amplitude of harmonics
            If NoError Then
                dblAmplitudes(lngCounter) = GaussFit.Amplitude
                dblIndexFreqs(lngCounter) = GaussFit.Center / lngHarmonics(lngCounter)
            End If
        
        End If
                
    Next lngCounter
    
    ' calculate average freq from regression over all harmonics
    If UBound(lngHarmonics) <> 1 Then
        dblRegressionResults = RegressionAverage(dblIndexFreqs, dblAmplitudes)
        DronePositionByHarmonics = dblRegressionResults(0)
    Else
        ' if no harmonics, only basic frequency
        DronePositionByHarmonics = dblIndexFreqs(1)
    End If

End Function

Private Function IndexOfMaxInRange(dblArray() As Double, Optional ByVal LowerLimit As Long, Optional ByVal UpperLimit As Long) As Long
    ' returns the index of the maximum of the array within the specified range
    
    If LowerLimit <= 0 Then LowerLimit = LBound(dblArray)
    If UpperLimit <= 0 Then UpperLimit = UBound(dblArray)
    If LowerLimit > UBound(dblArray) Then LowerLimit = UBound(dblArray)
    If UpperLimit > UBound(dblArray) Then UpperLimit = UBound(dblArray)
    
    Dim lngIndex As Long
    Dim lngIndexMax As Long
    Dim dblMaximum As Double
    
    For lngIndex = LowerLimit To UpperLimit
    'detect maximum in array wihtin limits and store index of maximum
        If dblArray(lngIndex) > dblMaximum Then
            dblMaximum = dblArray(lngIndex)
            lngIndexMax = lngIndex
        End If
    Next lngIndex
    
    IndexOfMaxInRange = lngIndexMax

End Function

Private Function RegressionAverage(ByRef dblFreqs, ByRef dblAmplitudes) As Double()
    ' calculates and returns the frequency via weighted average
    '
    Dim dblSum_AmpWeightedFreqs                     ' sum Freq * Amp
    Dim dblSum_Amplitudes As Double                 ' sum Amp
    Dim dblSum_1 As Double                          ' sum 1
    
    Dim dblReg_AmpWeightedFreqs  As Double          ' regression Freq : sum F*A / sum A         ( weigthed with Amp )
    Dim dblVar_AmpWeightedFreqs As Double           ' variance Freq * Amp
    Dim dblSig_AmpWeightedFreqs As Double           ' sigma Freq * Amp
    Dim dblMaxVar_AmpWeightedFreqs As Double        ' max. variance Freq * Amp
    Dim dblMaxSigma_AmpWeightedFreqs As Double      ' max. sigma Freq * Amp
    Dim lngMaxVar_AmpWeightedFreqs As Long          ' harmonic with max. variance Freq * Amp
       
    Dim dblRegressionResults() As Double
    ReDim dblRegressionResults(0 To 3)
    
    Dim lngI As Long
    
    ' calulculate sums over all the specified harmoncis that are <> 0
    For lngI = LBound(dblFreqs) To UBound(dblFreqs)
        If dblFreqs(lngI) <> 0 Then
            dblSum_AmpWeightedFreqs = dblSum_AmpWeightedFreqs + dblFreqs(lngI) * dblAmplitudes(lngI)
            dblSum_Amplitudes = dblSum_Amplitudes + dblAmplitudes(lngI)
            dblSum_1 = dblSum_1 + 1
        End If
    Next lngI

    If dblSum_1 > 1 And dblSum_Amplitudes <> 0 Then

        dblReg_AmpWeightedFreqs = dblSum_AmpWeightedFreqs / dblSum_Amplitudes
             
        ' calulate variances and maximum variation
        For lngI = LBound(dblFreqs) To UBound(dblFreqs)
            If dblFreqs(lngI) <> 0 Then
                dblVar_AmpWeightedFreqs = dblVar_AmpWeightedFreqs + (dblFreqs(lngI) - dblReg_AmpWeightedFreqs) ^ 2
                If (dblFreqs(lngI) - dblReg_AmpWeightedFreqs) ^ 2 > dblMaxVar_AmpWeightedFreqs Then
                    dblMaxVar_AmpWeightedFreqs = (dblFreqs(lngI) - dblReg_AmpWeightedFreqs) ^ 2
                    lngMaxVar_AmpWeightedFreqs = lngI
                End If
            End If
        Next lngI
        
        ' calculate sigma
        dblSig_AmpWeightedFreqs = Sqr(dblVar_AmpWeightedFreqs) / (dblSum_1 - 1)
        dblMaxSigma_AmpWeightedFreqs = Sqr(dblMaxVar_AmpWeightedFreqs)
    
    End If
    
    dblRegressionResults(0) = dblReg_AmpWeightedFreqs           ' use regression for freq weigthed with amplitude
    dblRegressionResults(1) = dblSig_AmpWeightedFreqs           ' average simga
    dblRegressionResults(1) = dblMaxSigma_AmpWeightedFreqs      ' largest sigma
    dblRegressionResults(3) = CDbl(lngMaxVar_AmpWeightedFreqs)  ' harmonic that has largest sigma
    
    RegressionAverage = dblRegressionResults
    
End Function

