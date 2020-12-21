Attribute VB_Name = "Display"
Option Explicit

Private Const mdblTwoPi As Double = 6.28318530717959

Private Const BI_RGB = 0&
Private Const CBM_INIT = &H4
Private Const DIB_RGB_COLORS = 0
Private Const SRCCOPY = &HCC0020

' BlendFunction BlendOp-Konstante
Private Const AC_SRC_OVER = &H0 ' die Quelle wird über dem Ziel gezeichnet
 
' 'BlendFunction AlphaFormat-Konstante
Private Const AC_SRC_ALPHA = &H1 ' das Quellbitmap wurde schon mit dem Alphawert multipliziert

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO_8Bit
    bmiHeader As BITMAPINFOHEADER
    bmiColors(0 To 255) As RGBQUAD
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type

' Benötigte API-Deklaration
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal cbCopy As Long)
Private Declare Sub MoveMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateDIBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                     lpInfoHeader As BITMAPINFOHEADER, _
                                                     ByVal dwUsage As Long, _
                                                     lpInitBits As Any, _
                                                     lpInitInfo As BITMAPINFO, _
                                                     ByVal wUsage As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hdc As Long, _
                                                       pBitmapInfo As BITMAPINFO, _
                                                       ByVal un As Long, _
                                                       ByVal lplpVoid As Long, _
                                                       ByVal handle As Long, _
                                                       ByVal dw As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, _
                                                 ByVal X As Long, _
                                                 ByVal Y As Long, _
                                                 ByVal nWidth As Long, _
                                                 ByVal nHeight As Long, _
                                                 ByVal hSrcDC As Long, _
                                                 ByVal xSrc As Long, _
                                                 ByVal ySrc As Long, _
                                                 ByVal nSrcWidth As Long, _
                                                 ByVal nSrcHeight As Long, _
                                                 ByVal dwRop As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, _
                                                       ByVal xDest As Long, _
                                                       ByVal yDest As Long, _
                                                       ByVal WidthDest As Long, _
                                                       ByVal HeightDest As Long, _
                                                       ByVal hdcSrc As Long, _
                                                       ByVal xSrc As Long, _
                                                       ByVal ySrc As Long, _
                                                       ByVal WidthSrc As Long, _
                                                       ByVal HeightSrc As Long, _
                                                       ByVal Blendfunc As Long) As Long



Private mudtPixels() As RGBQUAD             ' pixel data.
Private mudtBackGround() As RGBQUAD         ' backgorund pixel data
Private mudtColor As RGBQUAD                ' color rgb quadrupel
Private mbytUseColor As Boolean

Private bm_info As BITMAPINFO               ' DIB bitmap info.

Private hDIB As Long              ' Bitmap handle.
Private pBF As Long               ' Pointer Blendfunction

Private mlngOffSetCent() As Long
Private mlngOffSetDrones() As Long
Private mlngOffSetLeft As Long              ' left offset in pix
Private mlngOffSetStart As Long             ' start offset in pix
Private mlngOffSetStop As Long              ' stop offset in pix
Private mlngOffSetRight As Long             ' right offset in pix

Private mlngWidth As Long                   ' window width in pix
Private mlngHeight As Long                  ' window height in pix
Private msngZoom As Single                  ' zoom factor

Private mintScaleDistance As Integer        ' distance between 2 note lines
Private mintCentAtHalfScale As Integer      ' cents at half distance between to note lines
Private mintColorsPerCent As Integer        ' colors per cent
Private mbytTransparency As Byte            ' transparency

Private mbytMaxBGColor As Byte              ' max background color
Private mbytMinBGColor As Byte              ' min background color

Private mobjPicture As PictureBox           ' picture obj
Private mobjFormPicture As Form             ' form obj
Private mudtColorTable() As RGBQUAD         ' array with color table


Public Sub SetPicture(ByRef picValue As PictureBox, Optional ByRef formValue As Form)
    'If formValue = Nothing Then formValue = Me
    Set mobjPicture = picValue
    Set mobjFormPicture = formValue
    Init
End Sub
Public Sub SetScaleWidth(ByVal Value As Integer)
    mobjPicture.ScaleWidth = Value
    Init
End Sub
Public Sub SetScaleHeight(ByVal Value As Integer)
    mobjPicture.ScaleHeight = Value
    Init
End Sub
Public Sub SetTransparency(ByVal Value As Byte)
    mbytTransparency = Value
End Sub
Public Sub SetCentAtHalfScale(ByVal Value As Integer)
    mintCentAtHalfScale = Value
End Sub
Public Sub SetColorsPerCent(ByVal Value As Integer)
    mintColorsPerCent = Value
End Sub
Public Sub Set_Zoom(ByVal Value As Single)
    msngZoom = Value
End Sub
Public Function Zoom()
    Zoom = msngZoom
End Function

Public Sub Init()

    Dim i As Integer
     
    ReDim mlngOffSetCent(LBound(gudtNotes) To UBound(gudtNotes))
    ReDim mlngOffSetDrones(LBound(gudtDrones) To UBound(gudtDrones))

    mintColorsPerCent = 8
    
    mbytTransparency = 255
    mbytMaxBGColor = 180
    mbytMinBGColor = 20
    
    msngZoom = 2
    
    mlngWidth = Int(mobjPicture.ScaleWidth / msngZoom)
    'mlngOffSetLeft = 32 / msngZoom  ' 32 pixels
    'mlngOffSetStart = Int(0.01 * mlngWidth)
    mlngOffSetStop = 50 / msngZoom  ' 64 pixels 'Int(0.08 * mlngWidth)
    mlngOffSetRight = 10 / msngZoom 'Int(0.02 * mlngWidth)
    NoteBuffer.NoteBufferInit (mlngWidth - mlngOffSetLeft - mlngOffSetStart - mlngOffSetStop - mlngOffSetRight)
    
    mintScaleDistance = CInt(mobjPicture.ScaleHeight / (UBound(gudtNotes) + 2 + UBound(gudtDrones)))
    mlngHeight = (UBound(gudtNotes) + 2 + UBound(gudtDrones)) * mintScaleDistance
    
    For i = 1 To UBound(gudtNotes)
        mlngOffSetCent(i) = CInt((i + 1 + UBound(gudtDrones)) * mintScaleDistance)
    Next i
    
    For i = 1 To UBound(gudtDrones)
        mlngOffSetDrones(i) = CInt(i * mintScaleDistance)
    Next i
    
    Dim BF As BLENDFUNCTION
    
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = &HFF
        .AlphaFormat = AC_SRC_ALPHA
    End With
    MoveMemory pBF, BF, Len(BF)

    SetColors
    SetBackground
    'SetDefaultValues

End Sub

Public Sub Draw()

    Dim intAllNotes() As Integer
    Dim dblAllCents() As Double
    
    intAllNotes = gintBufferNote.FIFO_Buffer
    dblAllCents = gdblBufferCent.FIFO_Buffer

    SetPixels intAllNotes, dblAllCents
    CreateDIB
    DrawDIB
    
End Sub

Private Sub CreateDIB()

    With bm_info.bmiHeader
        .biSize = Len(bm_info.bmiHeader)
        .biWidth = mlngWidth    ' Width in pixels.
        .biHeight = mlngHeight  ' Height in pixels.
        .biPlanes = 1           ' 1 color plane.
        '.biBitCount = 8         ' 8 bits per pixel.
        .biBitCount = 32        ' 8 bits per pixel.
        .biCompression = BI_RGB ' No compression.
        .biSizeImage = 0        ' Unneeded with no compression.
        .biXPelsPerMeter = 0    ' Unneeded.
        .biYPelsPerMeter = 0    ' Unneeded.
        '.biClrUsed = 256        ' # colors in color table that are used by the image. 0 means all.
        .biClrUsed = 0          ' # colors in color table that are used by the image. 0 means all.
        '.biClrImportant = 256   ' # important colors. 0 means all.
        .biClrImportant = 0     ' # important colors. 0 means all.
    End With

    
    hDIB = CreateDIBitmap(mobjPicture.hdc, _
        bm_info.bmiHeader, CBM_INIT, mudtPixels(0, 0), _
        bm_info, DIB_RGB_COLORS)
    
    'MsgBox (hDIB & " / " & mobjPicture.hdc)
     
End Sub

Private Sub DrawDIB()
' Draw the DIB onto the form
    
    Dim retval As Boolean
    Dim compat_dc As Long
    
    compat_dc = CreateCompatibleDC(mobjFormPicture.hdc)
    SelectObject compat_dc, hDIB
    
    mobjPicture.Cls
    
    StretchBlt mobjPicture.hdc, 0, 0, mobjPicture.ScaleWidth, mobjPicture.ScaleHeight, _
            compat_dc, 0, 0, mlngWidth, mlngHeight, SRCCOPY
    
    'retval = AlphaBlend(mobjPicture.hdc, 0, 0, mobjPicture.Width, mobjPicture.Height, compat_dc, 0, 0, mlngWidth, mlngHeight, pBF)
    
    DeleteObject hDIB
    DeleteDC compat_dc
    'mobjPicture.Refresh
    
End Sub
 
Public Sub SetDefaultValues()
    
    Dim lngX As Integer
    Dim lngY As Integer
    Dim lngMaxCent As Long
    
    Dim dblDrone() As Double
    
    Dim intAllNotes() As Integer
    Dim dblAllCents() As Double
    Dim dblAllFreqs() As Double

    ReDim intAllNotes(0 To gintBufferNote.FIFO_Elements - 1)
    ReDim dblAllCents(0 To gintBufferNote.FIFO_Elements - 1, 0 To UBound(gudtDrones))
    ReDim dblAllFreqs(0 To gintBufferNote.FIFO_Elements - 1, 0 To UBound(gudtDrones))

    lngMaxCent = 10
    
    For lngX = 0 To gintBufferNote.FIFO_Elements - 1
        dblAllCents(lngX, 0) = lngMaxCent * Cos(2# * 6# * lngX / mlngWidth)
        intAllNotes(lngX) = Int(lngX / mlngWidth * 24) Mod UBound(mlngOffSetCent) + 1
        dblAllFreqs(lngX, 0) = gdblReferenceFrequency * gudtNotes(intAllNotes(lngX)).Ratio * 2 ^ (dblAllCents(lngX, 0) / 1200)
        For lngY = 1 To UBound(gudtDrones)
            dblAllCents(lngX, lngY) = lngMaxCent * Cos(2 * (3 + lngY) * lngX / mlngWidth)
            dblAllFreqs(lngX, lngY) = gdblReferenceFrequency * gudtDrones(lngY).Ratio * 2 ^ (dblAllCents(lngX, lngY) / 1200)
        Next lngY
    Next lngX

    gintBufferNote.FIFO_Buffer = intAllNotes
    gdblBufferCent.Set_FIFO_Buffer dblAllCents
    gdblBufferFrequency.Set_FIFO_Buffer dblAllFreqs

    SetPixels intAllNotes, dblAllCents
    Draw
    
End Sub

Private Sub SetPixels(ByRef intAllNotes, ByRef dblAllCents)

    Dim i As Integer
    Dim lngX As Long
    Dim lngNumberOfBytes As Long

    ReDim mudtPixels(0 To mlngWidth - 1, 0 To mlngHeight - 1)
    
    lngNumberOfBytes = 4 * CLng(mlngWidth) * CLng(mlngHeight)
    CopyMemory mudtPixels(0, 0), mudtBackGround(0, 0), lngNumberOfBytes
    
    For lngX = 0 To gintBufferNote.FIFO_Position - 1
        SetPixelRGB mlngOffSetLeft + mlngOffSetStart + lngX, dblAllCents(lngX, 0), mlngOffSetCent(intAllNotes(lngX))
        For i = 1 To UBound(gudtDrones)
            SetPixelRGB mlngOffSetLeft + mlngOffSetStart + lngX, dblAllCents(lngX, i), mlngOffSetDrones(i)
        Next i
    Next lngX

End Sub

Private Sub SetPixelRGB(ByVal lngX As Long, ByVal dblCent As Double, ByVal lngOffset As Long)
' set pixels for x , y from offset to offset + cent(converted to y = lngMaxY)

    Dim lngY As Long
    Dim lngMaxY As Long
    Dim intColorTableIndex As Integer
        
    lngMaxY = CInt(CentToY(dblCent, mintCentAtHalfScale, mintScaleDistance))
        
    If lngOffset = 0 Or lngMaxY = 0 Then Exit Sub
            
    Select Case dblCent

        Case Is > 100: intColorTableIndex = 100 * mintColorsPerCent
        Case -100 To 100: intColorTableIndex = dblCent * mintColorsPerCent
        Case Is < -100: intColorTableIndex = -100 * mintColorsPerCent
            
    End Select
            
    For lngY = lngOffset To lngOffset + lngMaxY Step Sgn(lngMaxY)
        mudtPixels(lngX, lngY) = mudtColorTable(intColorTableIndex)
    Next lngY

End Sub

Private Sub SetColors()
    
    ReDim mudtColorTable(mintColorsPerCent * -100 - 1 To mintColorsPerCent * 100 + 1)
    Dim i As Integer
    
    For i = 0 To mintColorsPerCent * 100
        mudtColorTable(i).rgbRed = CByte(Round(255 * 1 / (1 + ((100 - Abs(i) / mintColorsPerCent) / 98) ^ 80)) - 1 * Abs(i) / mintColorsPerCent)
        mudtColorTable(i).rgbGreen = CByte(Round(255 * (1 / (1 + (Abs(i) / mintColorsPerCent / 4) ^ 4))))
        mudtColorTable(i).rgbBlue = 0
        mudtColorTable(i).rgbReserved = mbytTransparency
    Next i
    For i = -1 To -100 * mintColorsPerCent Step -1
        mudtColorTable(i).rgbRed = 0
        mudtColorTable(i).rgbGreen = CByte(Round(255 * (1 / (1 + (Abs(i) / mintColorsPerCent / 4) ^ 4))))
        mudtColorTable(i).rgbBlue = CByte(Round(255 * 1 / (1 + ((100 - Abs(i) / mintColorsPerCent) / 98) ^ 80)) - 1 * Abs(i) / mintColorsPerCent)
        mudtColorTable(i).rgbReserved = mbytTransparency
    Next i
    
    For i = 0 To mintColorsPerCent * 100
        mudtColorTable(i).rgbRed = CByte(Round(255 * 1 / (1 + ((100 - Abs(i) / mintColorsPerCent) / 98) ^ 80)) - 1 * Abs(i) / mintColorsPerCent)
        mudtColorTable(i).rgbGreen = CByte(Round(255 * (1 / (1 + (Abs(i) / mintColorsPerCent / 4) ^ 4))))
        mudtColorTable(i).rgbBlue = 0
        mudtColorTable(i).rgbReserved = mbytTransparency
    Next i
    For i = -1 To -100 * mintColorsPerCent Step -1
        mudtColorTable(i).rgbRed = 0
        mudtColorTable(i).rgbGreen = CByte(Round(255 * (1 / (1 + (Abs(i) / mintColorsPerCent / 4) ^ 4))))
        mudtColorTable(i).rgbBlue = CByte(Round(255 * 1 / (1 + ((100 - Abs(i) / mintColorsPerCent) / 98) ^ 80)) - 1 * Abs(i) / mintColorsPerCent)
        mudtColorTable(i).rgbReserved = mbytTransparency
    Next i
    
    mudtColorTable(LBound(mudtColorTable)).rgbRed = 0
    mudtColorTable(LBound(mudtColorTable)).rgbGreen = 0
    mudtColorTable(LBound(mudtColorTable)).rgbBlue = 0
    mudtColorTable(LBound(mudtColorTable)).rgbReserved = 255
    
    mudtColorTable(UBound(mudtColorTable)).rgbRed = 255
    mudtColorTable(UBound(mudtColorTable)).rgbGreen = 255
    mudtColorTable(UBound(mudtColorTable)).rgbBlue = 255
    mudtColorTable(UBound(mudtColorTable)).rgbReserved = 255

End Sub
Private Sub SetBackground()
' Create a Background

    Dim lngX As Long
    Dim lngY As Long
    Dim lngYMax As Long
    Dim udtColor As RGBQUAD
    Dim dblCent As Double
    Dim i As Integer
    Dim intTicks() As Integer
    ReDim intTicks(0 To 7)
    
    intTicks(0) = 0
    intTicks(1) = 1
    intTicks(2) = 2
    intTicks(3) = 5
    intTicks(4) = 10
    intTicks(5) = 20
    intTicks(6) = 50
    intTicks(7) = 100
    
    ReDim mudtBackGround(0 To mlngWidth - 1, 0 To mlngHeight - 1)
        
    mbytMaxBGColor = 160
    mbytMinBGColor = 32
    
    For lngY = 0 To mlngHeight - 1

        mudtColor.rgbBlue = CByte((mbytMaxBGColor - mbytMinBGColor) / mlngHeight * lngY + mbytMinBGColor)
        'mudtColor.rgbBlue = 0
        mudtColor.rgbGreen = CByte((mbytMaxBGColor - mbytMinBGColor) / mlngHeight * lngY + mbytMinBGColor)
        mudtColor.rgbGreen = mudtColor.rgbGreen - mbytMinBGColor
        'mudtColor.rgbGreen = 0
        mudtColor.rgbRed = CByte((mbytMaxBGColor - mbytMinBGColor) / mlngHeight * lngY + mbytMinBGColor)
        'mudtColor.rgbRed = 0
        'mudtColor.rgbReserved = 0
        
        For lngX = 0 To mlngWidth - 1
            mudtBackGround(lngX, lngY) = mudtColor
        Next lngX
        
    Next lngY
          
    mbytUseColor = True
    mudtColor.rgbRed = 255
    mudtColor.rgbGreen = 255
    mudtColor.rgbBlue = 255
    
    For i = 1 To UBound(gudtDrones)
        DrawLines 0, mlngOffSetDrones(i), mbytUseColor
    Next i
    
    For i = 1 To UBound(gudtNotes)
        DrawLines 0, mlngOffSetCent(i), mbytUseColor
    Next i

    DrawLines 0, Round((mlngOffSetCent(1) + mlngOffSetDrones(UBound(gudtDrones))) / 2) - 1, mbytUseColor
    DrawLines 0, Round((mlngOffSetCent(1) + mlngOffSetDrones(UBound(gudtDrones))) / 2), mbytUseColor
    DrawLines 0, Round((mlngOffSetCent(1) + mlngOffSetDrones(UBound(gudtDrones))) / 2) + 1, mbytUseColor
        
    udtColor.rgbBlue = 255
    udtColor.rgbGreen = 255
    udtColor.rgbRed = 255
    
    For lngY = 0 To mlngHeight - 1
        For lngX = 0 To mlngOffSetLeft - 1
            mudtBackGround(lngX, lngY) = udtColor
        Next lngX
    Next lngY

    udtColor.rgbBlue = 0
    udtColor.rgbGreen = 0
    udtColor.rgbRed = 0
    
    For lngY = 0 To mlngHeight - 1
        For lngX = mlngWidth - 1 - mlngOffSetRight To mlngWidth - 1
            mudtBackGround(lngX, lngY) = udtColor
        Next lngX
    Next lngY
    
    
    'lngYMax = Int(mintScaleDistance / 2)
    'For lngY = -1 * lngYMax To lngYMax
    '    dblCent = YToCent(lngY, mintCentAtHalfScale, mintScaleDistance)
    '    If dblCent < -100 Then dblCent = -100
    '    If dblCent > 100 Then dblCent = 100
    '    udtColor = mudtColorTable(dblCent * mintColorsPerCent)
    '    For lngX = mlngWidth - 1 - mlngOffSetRight To mlngWidth - 1
    '        For i = 1 To UBound(gudtNotes)
    '            mudtBackGround(lngX, lngY + mlngOffSetCent(i)) = udtColor
    '        Next i
    '        For i = 1 To UBound(gudtDrones)
    '            mudtBackGround(lngX, lngY + mlngOffSetDrones(i)) = udtColor
    '        Next i
    '    Next lngX
    'Next lngY

    DrawTicks 0, 0, mlngOffSetLeft - 1
    DrawTicks 0, mlngWidth - 1 - mlngOffSetRight, mlngWidth - 1
    For i = 1 To UBound(intTicks)
        If intTicks(i) <= mintCentAtHalfScale / 2 And intTicks(i) >= mintCentAtHalfScale / 5 Then
'            DrawTicks intTicks(i), 0, mlngOffSetLeft - 1
            'DrawTicks intTicks(i), 0, mlngOffSetLeft - 2
            DrawTicks intTicks(i), mlngWidth - 1 - mlngOffSetRight, mlngWidth - 1
            'DrawTicks intTicks(i), mlngOffSetLeft, mlngWidth - 1 - mlngOffSetRight
        End If
    Next i
    
    For i = 1 To mintCentAtHalfScale * 10
            DrawTicks i / 10, mlngWidth - 1 - mlngOffSetRight, mlngWidth - 1
    Next i
    
    udtColor.rgbBlue = 255
    udtColor.rgbGreen = 255
    udtColor.rgbRed = 255
    
    For lngY = 0 To mlngHeight - 1
'        mudtBackGround(mlngOffSetLeft - 1, lngY) = udtColor
        'mudtBackGround(mlngOffSetLeft, lngY) = udtColor
        'mudtBackGround(mlngWidth - 1 - mlngOffSetRight - 1, lngY) = udtColor
        mudtBackGround(mlngWidth - 1 - mlngOffSetRight, lngY) = udtColor
    Next lngY
    
    'mobjPicture.AutoRedraw = False
    mobjPicture.AutoRedraw = True

End Sub

Private Sub DrawLines(ByVal dblCent As Double, ByVal lngOffset As Long, Optional ByVal UseColor As Boolean)

    Dim udtColor As RGBQUAD
    Dim lngX As Long
    Dim lngY As Long
    
    If Not UseColor Then
        udtColor = mudtColorTable(CInt(mintColorsPerCent * dblCent))
    Else
        udtColor = mudtColor
    End If
    
    If dblCent <> 0 Then lngY = CentToY(dblCent, mintCentAtHalfScale, mintScaleDistance)
    
    For lngX = mlngOffSetLeft + mlngOffSetStart To mlngWidth - 1
        mudtBackGround(lngX, lngY + lngOffset) = udtColor
    Next lngX
    
End Sub

Private Sub DrawTicks(ByVal dblCent As Double, ByVal lngLeft, ByVal lngRight)

    Dim lngX As Long
    Dim lngY_p As Long
    Dim lngY_n As Long
    Dim i As Integer
    Dim udtColor_p As RGBQUAD
    Dim udtColor_n As RGBQUAD
    
    lngY_p = CentToY(dblCent, mintCentAtHalfScale, mintScaleDistance)
    udtColor_p = mudtColorTable(CInt(dblCent) * mintColorsPerCent)
    
    lngY_n = CentToY(-1 * dblCent, mintCentAtHalfScale, mintScaleDistance)
    udtColor_n = mudtColorTable(CInt(-1 * dblCent) * mintColorsPerCent)
    
    For i = 1 To UBound(gudtNotes)
        For lngX = lngLeft To lngRight
            'mudtBackGround(lngX, lngY_p + mlngOffSetCent(i) - 1) = udtColor_p
            'mudtBackGround(lngX, lngY_n + mlngOffSetCent(i) - 1) = udtColor_n
            mudtBackGround(lngX, lngY_p + mlngOffSetCent(i)) = udtColor_p
            mudtBackGround(lngX, lngY_n + mlngOffSetCent(i)) = udtColor_n
            'mudtBackGround(lngX, lngY_p + mlngOffSetCent(i) + 1) = udtColor_p
            'mudtBackGround(lngX, lngY_n + mlngOffSetCent(i) + 1) = udtColor_n
        Next lngX
    Next i
    For i = 1 To UBound(gudtDrones)
        For lngX = lngLeft To lngRight
            'mudtBackGround(lngX, lngY_p + mlngOffSetDrones(i) - 1) = udtColor_p
            'mudtBackGround(lngX, lngY_n + mlngOffSetDrones(i) - 1) = udtColor_n
            mudtBackGround(lngX, lngY_p + mlngOffSetDrones(i)) = udtColor_p
            mudtBackGround(lngX, lngY_n + mlngOffSetDrones(i)) = udtColor_n
            'mudtBackGround(lngX, lngY_p + mlngOffSetDrones(i) + 1) = udtColor_p
            'mudtBackGround(lngX, lngY_n + mlngOffSetDrones(i) + 1) = udtColor_n
        Next lngX
    Next i
    
End Sub
Private Function CentToY(ByVal dblCent As Double, ByVal intCentAtHalfY As Integer, ByVal lngMaxY As Long) As Integer
    CentToY = Atn(dblCent / intCentAtHalfY) / mdblTwoPi * 4 * lngMaxY
End Function

Private Function YToCent(ByVal lngY As Long, ByVal intCentAtHalfY As Integer, ByVal lngMaxY As Long) As Double
    YToCent = Tan(lngY * mdblTwoPi / 4 / lngMaxY) * intCentAtHalfY
End Function

Public Sub SetAxis(ByRef nAxis As Axis)
    
    Dim sngSizeSec As Single            ' size/width (time)
    Dim sngStartSec As Single           ' start time (left)
    Dim sngCurrentSec As Single         ' current time
    
    Dim lAxis As Axis
    Set lAxis = nAxis
    
    ' size = width / measurements/s
    sngSizeSec = CSng(mlngWidth) / WavFile.SampleRate * WavFile.SampleInterval
    ' start = bufer start position / measurements/s
    sngStartSec = CSng(gdblBufferFrequency.FIFO_StartPosition) / WavFile.SampleRate * WavFile.SampleInterval
    ' add start time of wav file
    sngStartSec = sngStartSec + WavFile.StartTime
    
    lAxis.Left = mobjPicture.Left - 4
    lAxis.Width = mobjPicture.Width + 8

    
    If sngStartSec = 0 Then
        lAxis.OffsetBegin = 6
        lAxis.BeginValue = 0
    Else
        lAxis.OffsetBegin = 6 + (1 - sngStartSec + Int(sngStartSec)) / sngSizeSec * mlngWidth * msngZoom
        lAxis.BeginValue = Round(sngStartSec + 0.5)
    End If
    
    lAxis.endValue = Int(sngStartSec + sngSizeSec)      ' start plus size
    lAxis.OffsetEnd = 6 + (sngStartSec + sngSizeSec - lAxis.endValue) / sngSizeSec * mlngWidth * msngZoom

End Sub
