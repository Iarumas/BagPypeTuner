Attribute VB_Name = "FFTLib"
'---------------------------------------------------------------------
'Declares for use with FFT.DLL -- Murphy McCauley 08/01/99
'---------------------------------------------------------------------

'You'll notice that the arrays (RealIn, ImagIn, RealOut, ImagOut) are
'represented as a single variable, not an array.  Pass the FIRST
'ELEMENT of the array and things work just right.
'For example...
'Dim RealIn(1 to 128) As Single, ImagIn(1 to 128) As Single
'Dim RealOut(1 to 128) As Single, ImagOut(1 to 128) As Single
' ...
'Call FFTSingle(128, False, RealIn(1), ImagIn(1), RealOut(1), ImagOut(1))

'Also, I aliased the functions so you can use the pretty VB-style
'names FFTDouble, FFTSingle, and IndexToFrequency.

 Declare Sub FFTDouble Lib "FFT.dll" Alias "fft_double" _
  (ByVal NumSamples As Long, ByVal InverseTransform As Boolean, _
  RealIn As Double, _
  ImagIn As Double, _
  RealOut As Double, _
  ImagOut As Double)
 
 Declare Sub FFTSingle Lib "FFT.dll" Alias "fft_float" _
  (ByVal NumSamples As Long, ByVal InverseTransform As Boolean, _
  RealIn As Single, _
  ImagIn As Single, _
  RealOut As Single, _
  ImagOut As Single)
 
 Declare Function IndexToFrequency Lib "FFT.dll" _
  Alias "Index_to_frequency" _
  (ByVal NumSamples As Long, _
  ByVal Index As Long) As Double

