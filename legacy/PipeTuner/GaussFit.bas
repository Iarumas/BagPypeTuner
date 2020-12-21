Attribute VB_Name = "GaussFit"
Option Explicit
' performs a least square fit to Gauss function with 3 data points only (max and datapoints left/rigth)
' Gauss fit is a parabolic fit for logarithm of Gauss function
' you need at least 3 datapoints for a parabolic fit

Private Type GaussFitValues
    dblLeft As Double           ' value of datapoint i-1
    dblCenter As Double         ' value of datapoint i
    dblRight As Double          ' value of datapoint i+1
End Type

Private Type GaussFitResults
    dblCenter As Double         ' gauss fit result: center position
    dblSigma As Double          ' gauss fit result: sigma
    dblAmplitude As Double      ' gauss fit result: center amplitude
End Type

Private mdblParabolicResult(0 To 2) As Double   ' results of parabolic least square fit

Private mudtFitValues As GaussFitValues         ' are the input values for the parabolic fit
Private mudtFitResults As GaussFitResults       ' are the results for Gauss fit: 3 parameters for Gauss function


Public Function Center() As Double
'returns the center value for the Gauss fit function
    Center = mudtFitResults.dblCenter
End Function
Public Function Sigma() As Double
'returns the standard deviation for the Gauss fit function
    Sigma = mudtFitResults.dblSigma
End Function
Public Function Amplitude() As Double
' returns the amplitude (max) of the Gauss fit function
    Amplitude = mudtFitResults.dblAmplitude
End Function

Public Function Fit(ByVal PeakIndex As Long, CurveValues) As Boolean
' peak index = index of local maximum in the curve

Fit = False

' least square fit für Gauss-Kurve Amp*exp(-(X-Xcenter)^2/(2*Sigma^2))
' für 3 benachbarte Punkte um den peak_index aus den Daten von "curve"

' Input: peak_index und die Funktion "curve"
' peak_index ist lokales Maximum der Funktion "Curve" bzgl. seiner Nachbarn peak_index-1 und peak_index+1
' "curve" ist das Feld mit den Funktionswerten

If PeakIndex = LBound(CurveValues) Or PeakIndex = UBound(CurveValues) Then
    Exit Function
End If
' bei 3 benachbarten Punktem reduziert sich der least square fit ganz wesentlich
' Ausgewertet werden die Punkte:
'   peak_index-1, curve(peak_index-1)
'   peak_index  , curve(peak_index  )
'   peak_index+1, curve(peak_index+1)

    'falls peak_index am rand von curve liegt wird peak_index zurückgegeben
    'oder auch falls peak_index nicht der größte Werte ist
    ' bzw. falls der Wert an der Stelle peak_index = 0 ist
    If CurveValues(PeakIndex) < CurveValues(PeakIndex - 1) Or _
       CurveValues(PeakIndex) < CurveValues(PeakIndex + 1) Or _
       PeakIndex <= LBound(CurveValues) Or _
       PeakIndex >= UBound(CurveValues) Or _
       CurveValues(PeakIndex) = 0 _
       Then
            mudtFitResults.dblCenter = CDbl(PeakIndex)
            mudtFitResults.dblSigma = 0
            mudtFitResults.dblAmplitude = CDbl(CurveValues(PeakIndex))
            Exit Function
    End If
    
    ' for Gauss fit the logarithmic values of the curve are needed
    mudtFitValues.dblLeft = Log(CurveValues(PeakIndex - 1))
    mudtFitValues.dblCenter = Log(CurveValues(PeakIndex))
    mudtFitValues.dblRight = Log(CurveValues(PeakIndex + 1))

    
    ' parabolic least square fit
    mdblParabolicResult(0) = mudtFitValues.dblCenter
    mdblParabolicResult(1) = 0.5 * (mudtFitValues.dblRight - mudtFitValues.dblLeft)
    mdblParabolicResult(2) = 0.5 * (mudtFitValues.dblLeft + mudtFitValues.dblRight) - mudtFitValues.dblCenter
    
    ' gaussian least square fit: convert parabolic results to Gauss parameters
    mudtFitResults.dblCenter = PeakIndex - mdblParabolicResult(1) / (2 * mdblParabolicResult(2))
    mudtFitResults.dblSigma = Sqr(-1 / (2 * mdblParabolicResult(2)))
    mudtFitResults.dblAmplitude = Exp(mdblParabolicResult(0) - mdblParabolicResult(1) ^ 2 / (4 * mdblParabolicResult(2)))

    Fit = True
    
End Function



