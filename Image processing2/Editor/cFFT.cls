VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFFT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Enum MODULAR_TRANSFER_FUNCTION
'Assuming: a - maximum (Niquist) frequency
'          a = NumSamples / 2
'          r = Sqr(x^2 + y^2) - distance from kernel center
'          Sigma - Gaussian standard deviation
    [_Min]
    eCylinder  'aka circular,pill-box
               'MTF =  PSF(Jinc) - for 2D
               'MTF =  PSF(Sinc) - for 1D
    eCone      'aka triangle-circular
    ePeak      'MTF = 1/f
    ePiramid   'aka triangle-rectangle
    eGaussian  'Gaussian distribution
               'PSF = Exp((r^2)/(2 * Sigma^2))
    eLaplacian 'Laplacian distribution
    eSinc      'aka sombrerro_1D
               'PSF = (Sin(PI * r)/(PI * r)
    eJinc      'aka sombrerro_2D (2D analog of Sinc)
               'PSF = (J1(PI * r)/(PI * r)
               'Where J1 is Besselian function of first kind
    eExpDecay  'MTF = (2*PI*a)/((w^2+a^2)^(3/2))
    eAiry      'aka difraction-circle,releygh
    eRectangle 'aka average,uniform,median
               'PSF = 1
    [_Max]
End Enum

Public Enum eFFT_MODE
    FFT_1D
    FFT_2D
End Enum

'Events
Public Event ProgressInit(ByVal sTitle As String, ByVal nMax As Long)
Public Event ProgressChanged(ByVal nValue As Long)
Public Event ProgressEnd(ByVal nTime As Long)

'local variable(s) to hold property value(s)
Private m_sRadius As Single
Private m_iFFTMode As Long
Private m_iBPP As eBPP
Private m_iFilterPower As Long

Private m_sMTF() As Single
Private m_iNiquistX As Long, m_iNiquistY As Long

Public Property Let FilterPower(ByVal vData As Long)
    m_iFilterPower = vData
    If m_iFilterPower = 0 Then m_iFilterPower = 1
End Property

Public Property Get FilterPower() As Long
    FilterPower = m_iFilterPower
End Property

Public Property Let FFT_Mode(ByVal vData As eFFT_MODE)
    m_iFFTMode = vData
End Property

Public Property Get FFT_Mode() As eFFT_MODE
    FFT_Mode = m_iFFTMode
End Property

Public Property Let BPP(ByVal vData As eBPP)
    m_iBPP = vData
End Property

Public Property Get BPP() As eBPP
    BPP = m_iBPP
End Property

Public Property Let Radius(ByVal vData As Single)
    m_sRadius = vData
End Property

Public Property Get Radius() As Single
    Radius = m_sRadius
End Property

Public Sub PerformFFT(abPicture() As Byte, MTF As MODULAR_TRANSFER_FUNCTION, Optional bReverse As Boolean)
    Dim xMax As Long, yMax As Long, nx As Long, ny As Long
    Dim imgSize As Long
    Dim lTime As Long
    
    lTime = GetTickCount
    RaiseEvent ProgressInit("Prepare to FFT - scanning image...", 0)
    If BPP > HighColor_16 Then
       xMax = UBound(abPicture, 1) \ 3
    Else
       xMax = UBound(abPicture, 1)
    End If
    yMax = UBound(abPicture, 2)
    imgSize = (xMax + 1) * (yMax + 1)
    If FFT_Mode = FFT_1D Then
       For x = 0 To 31
           If Power2(x) > imgSize Then
              nx = Power2(x)
              Exit For
           End If
       Next x
       m_iNiquistX = nx / 2
       PrepareMTF MTF, m_iNiquistX, , bReverse
       If BPP > HighColor_16 Then
          ProcessTrueColorFFT_1D abPicture, nx
       Else
          ProcessPaletteFFT_1D abPicture, nx
       End If
    Else
       For x = 0 To 31
           If Power2(x) > xMax Then
              If nx = 0 Then nx = Power2(x)
           End If
           If Power2(x) > yMax Then
              If ny = 0 Then ny = Power2(x)
           End If
           If (nx > 0) And (ny > 0) Then Exit For
       Next x
       m_iNiquistX = nx / 2
       m_iNiquistY = ny / 2
       PrepareMTF MTF, m_iNiquistX, m_iNiquistX, bReverse
       If BPP > HighColor_16 Then
          ProcessTrueColorFFT_2D abPicture, nx, ny
       Else
          ProcessPaletteFFT_2D abPicture, nx, ny
       End If
    End If
    RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Private Sub PrepareMTF(ByVal MTF As MODULAR_TRANSFER_FUNCTION, nx As Long, Optional ny As Long = 0, Optional bReverse As Boolean)
    Dim i As Long, j As Long
    Dim sVal As Single, MaxVal As Single
    If ny = 0 Then
       ReDim m_sMTF(nx)
       MaxVal = GetMTFValue(MTF, 0)
       sVal = GetMTFValue(MTF, nx)
       If sVal > MaxVal Then MaxVal = sVal
       For i = 0 To nx
           sVal = GetMTFValue(MTF, i) / MaxVal
           If bReverse Then
'              If sVal = 0 Then sVal = 0.0001
              m_sMTF(nx - i) = sVal
           Else
              m_sMTF(i) = sVal
           End If
       Next i
    Else
       ReDim m_sMTF(nx, ny)
       MaxVal = GetMTFValue(MTF, 0, 0)
       sVal = GetMTFValue(MTF, nx, ny)
       For i = 0 To nx
           For j = 0 To ny
               sVal = GetMTFValue(MTF, i, j) / MaxVal
               If bReverse Then
'                  If sVal = 0 Then sVal = 0.0001
                  m_sMTF(nx - i, ny - j) = sVal
               Else
                  m_sMTF(i, j) = sVal
               End If
           Next j
       Next i
    End If
End Sub

Private Function GetMTFValue(ByVal MTF As MODULAR_TRANSFER_FUNCTION, ByVal x As Long, Optional y As Long = 0) As Single
    Dim freqX As Single, freqY As Single
    Dim arg As Single
    freqX = x / m_iNiquistX
    If y > 0 Then freqY = y / m_iNiquistY
    Select Case MTF
        Case eCylinder
            arg = PI * Radius * Sqr(freqX * freqX + freqY * freqY)
            If arg = 0 Then arg = 0.0001
            GetMTFValue = 2 * Bessel_J1(2 * arg) / arg
        Case eCone
            arg = PI * Radius * Sqr(freqX * freqX + freqY * freqY)
            If arg = 0 Then arg = 0.0001
            arg = Bessel_J1(2 * arg) / arg
            GetMTFValue = 4 * arg * arg
        Case ePeak
            arg = Sqr(freqX * freqX + freqY * freqY)
            If arg = 0 Then arg = 0.0001
            GetMTFValue = 1 / arg
        Case Else
    End Select
End Function

'****************True Color routines**************************
Private Sub ProcessTrueColorFFT_1D(abPicture() As Byte, ByVal nx As Long)
   Dim x As Long, y As Long, dx As Long, imgSize As Long
   Dim lB As Long, lG As Long, lR As Long
   Dim lTime As Long
   
   Dim rgb() As RGBTRIPPLE
   Dim FFT_RED() As COMPLEX, FFT_GREEN() As COMPLEX, FFT_BLUE() As COMPLEX
   
   lTime = GetTickCount
'   RaiseEvent ProgressInit(x)
   ReDim rgb(nx)
   ReDim FFT_RED(nx)
   ReDim FFT_GREEN(nx)
   ReDim FFT_BLUE(nx)
   imgSize = (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1) \ 3
   CopyMemory rgb(0), abPicture(0, 0), imgSize * 3
   dx = nx / 2 - imgSize / 2
   For x = 0 To imgSize
       FFT_RED(x + dx).real = rgb(x).rgbRed
       FFT_GREEN(x + dx).real = rgb(x).rgbGreen
       FFT_BLUE(x + dx).real = rgb(x).rgbBlue
   Next x
   RaiseEvent ProgressInit("Processing FFT_1D - Red chanel...", 0)
   FFT_InPlace nx, FFT_RED
   RaiseEvent ProgressInit("Processing FFT_1D - Green chanel...", 0)
   FFT_InPlace nx, FFT_GREEN
   RaiseEvent ProgressInit("Processing FFT_1D - Blue chanel...", 0)
   FFT_InPlace nx, FFT_BLUE
   RaiseEvent ProgressInit("Filtering...", 0)
   For x = 0 To m_iNiquistX
       FFT_RED(x).real = FFT_RED(x).real * m_sMTF(x)
       FFT_GREEN(x).real = FFT_GREEN(x).real * m_sMTF(x)
       FFT_BLUE(x).real = FFT_BLUE(x).real * m_sMTF(x)
       FFT_RED(nx - x).real = FFT_RED(nx - x).real * m_sMTF(x)
       FFT_GREEN(nx - x).real = FFT_GREEN(nx - x).real * m_sMTF(x)
       FFT_BLUE(nx - x).real = FFT_BLUE(nx - x).real * m_sMTF(x)
   Next x
   RaiseEvent ProgressInit("Processing Inverse FFT_1D - Red chanel...", 0)
   FFT_InPlace nx, FFT_RED, True
   RaiseEvent ProgressInit("Processing Inverse FFT_1D - Green chanel...", 0)
   FFT_InPlace nx, FFT_GREEN, True
   RaiseEvent ProgressInit("Processing Inverse FFT_1D - Blue chanel...", 0)
   FFT_InPlace nx, FFT_BLUE, True
   RaiseEvent ProgressInit("Preparing result image...", 0)
   For x = 0 To imgSize
       lR = FFT_RED(x + dx).real
       lG = FFT_GREEN(x + dx).real
       lB = FFT_BLUE(x + dx).real
       ltob lR
       ltob lG
       ltob lB
       rgb(x).rgbRed = lR
       rgb(x).rgbGreen = lG
       rgb(x).rgbBlue = lB
   Next x
   CopyMemory abPicture(0, 0), rgb(0), imgSize * 3
   Erase rgb
   Erase FFT_RED
   Erase FFT_GREEN
   Erase FFT_BLUE
End Sub

Private Sub ProcessTrueColorFFT_2D(abPicture() As Byte, ByVal nx As Long, ByVal ny As Long)
   Dim x As Long, y As Long, dx As Long, dy As Long
   Dim imgSize As Long
   Dim xMax As Long, yMax As Long
   Dim lB As Long, lG As Long, lR As Long
   Dim lTime As Long
   
   Dim rgb() As RGBTRIPPLE
   Dim FFT_RED() As COMPLEX, FFT_GREEN() As COMPLEX, FFT_BLUE() As COMPLEX
   
   lTime = GetTickCount
   
   ReDim rgb(nx, ny)
   ReDim FFT_RED(nx, ny)
   ReDim FFT_GREEN(nx, ny)
   ReDim FFT_BLUE(nx, ny)
   
   xMax = UBound(abPicture, 1) \ 3
   yMax = UBound(abPicture, 2)
   CopyMemory rgb(0, 0), abPicture(0, 0), (xMax + 1) * (yMax + 1) * 3
   dx = nx / 2 - xMax / 2
   dy = nx / 2 - yMax / 2
   For x = 0 To xMax
       For y = 0 To yMax
           FFT_RED(x + dx, y + dy).real = rgb(x, y).rgbRed
           FFT_GREEN(x + dx, y + dy).real = rgb(x, y).rgbGreen
           FFT_BLUE(x + dx, y + dy).real = rgb(x, y).rgbBlue
       Next y
   Next x
   RaiseEvent ProgressInit("Processing FFT_2D - Red chanel...", 0)
   FFT_2D_InPlace nx, ny, FFT_RED
   RaiseEvent ProgressInit("Processing FFT_2D - Green chanel...", 0)
   FFT_2D_InPlace nx, ny, FFT_GREEN
   RaiseEvent ProgressInit("Processing FFT_2D - Blue chanel...", 0)
   FFT_2D_InPlace nx, ny, FFT_BLUE
   RaiseEvent ProgressInit("Filterring...", 0)
   For x = 0 To m_iNiquistX
       For y = 0 To m_iNiquistY
       'First quoter
           FFT_RED(x, y).real = FFT_RED(x, y).real * m_sMTF(x, y)
           FFT_GREEN(x, y).real = FFT_GREEN(x, y).real * m_sMTF(x, y)
           FFT_BLUE(x, y).real = FFT_BLUE(x, y).real * m_sMTF(x, y)
       'Second quoter
           FFT_RED(nx - x, y).real = FFT_RED(nx - x, y).real * m_sMTF(x, y)
           FFT_GREEN(nx - x, y).real = FFT_GREEN(nx - x, y).real * m_sMTF(x, y)
           FFT_BLUE(nx - x, y).real = FFT_BLUE(nx - x, y).real * m_sMTF(x, y)
       'Third quoter
           FFT_RED(nx - x, ny - y).real = FFT_RED(nx - x, ny - y).real * m_sMTF(x, y)
           FFT_GREEN(nx - x, ny - y).real = FFT_GREEN(nx - x, ny - y).real * m_sMTF(x, y)
           FFT_BLUE(nx - x, ny - y).real = FFT_BLUE(nx - x, ny - y).real * m_sMTF(x, y)
       'Forth quoter
           FFT_RED(nx, ny - y).real = FFT_RED(nx, ny - y).real * m_sMTF(x, y)
           FFT_GREEN(nx, ny - y).real = FFT_GREEN(nx, ny - y).real * m_sMTF(x, y)
           FFT_BLUE(nx, ny - y).real = FFT_BLUE(nx, ny - y).real * m_sMTF(x, y)
       Next y
   Next x
   RaiseEvent ProgressInit("Processing inverse FFT_2D - Red chanel...", 0)
   FFT_2D_InPlace nx, ny, FFT_RED, True
   RaiseEvent ProgressInit("Processing inverse FFT_2D - Green chanel...", 0)
   FFT_2D_InPlace nx, ny, FFT_GREEN, True
   RaiseEvent ProgressInit("Processing inverse FFT_2D - Blue chanel...", 0)
   FFT_2D_InPlace nx, ny, FFT_BLUE, True
   RaiseEvent ProgressInit("Preparing result image...", 0)
   For x = 0 To xMax
       For y = 0 To yMax
           lR = FFT_RED(x + dx, y + dy).real
           lG = FFT_GREEN(x + dx, y + dy).real
           lB = FFT_BLUE(x + dx, y + dy).real
           ltob lR
           ltob lG
           ltob lB
           rgb(x, y).rgbRed = lR
           rgb(x, y).rgbGreen = lG
           rgb(x, y).rgbBlue = lB
      Next y
   Next x
   CopyMemory abPicture(0, 0), rgb(0, 0), (xMax + 1) * (yMax + 1) * 3
End Sub


'****************Palette routines**************************
Private Sub ProcessPaletteFFT_1D(abPicture() As Byte, ByVal nx As Long)
   Dim x As Long, y As Long, dx As Long, imgSize As Long
   Dim lClr As Long
   Dim lTime As Long
   
   Dim clr() As Byte
   Dim FFT_CLR() As COMPLEX
   
   lTime = GetTickCount
'   RaiseEvent ProgressInit(x)
   ReDim clr(nx)
   ReDim FFT_CLR(nx)
   
   imgSize = (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
   CopyMemory clr(0), abPicture(0, 0), imgSize
   dx = nx / 2 - imgSize / 2
   For x = 0 To imgSize
       FFT_CLR(x + dx).real = clr(x)
   Next x
   FFT_InPlace nx, FFT_CLR
   For x = 0 To m_iNiquistX
       FFT_CLR(x).real = FFT_CLR(x).real * m_sMTF(x)
       FFT_CLR(nx - x).real = FFT_CLR(nx - x).real * m_sMTF(x)
   Next x
   FFT_InPlace nx, FFT_CLR, True
   For x = 0 To imgSize
       lClr = FFT_CLR(x + dx).real
       ltob lClr
       clr(x) = lClr
   Next x
   CopyMemory abPicture(0, 0), clr(0), imgSize
   Erase clr
   Erase FFT_CLR
End Sub

Private Sub ProcessPaletteFFT_2D(abPicture() As Byte, ByVal nx As Long, ByVal ny As Long)
   Dim x As Long, y As Long, dx As Long, dy As Long
   Dim imgSize As Long
   Dim xMax As Long, yMax As Long
   Dim lClr As Long
   Dim lTime As Long
   
   Dim clr() As Byte
   Dim FFT_CLR() As COMPLEX
   
   lTime = GetTickCount
   
   ReDim clr(nx, ny)
   ReDim FFT_CLR(nx, ny)
   
   xMax = UBound(abPicture, 1) \ 3
   yMax = UBound(abPicture, 2)
   CopyMemory clr(0, 0), abPicture(0, 0), (xMax + 1) * (yMax + 1)
   dx = nx / 2 - xMax / 2
   dy = nx / 2 - yMax / 2
   For x = 0 To xMax
       For y = 0 To yMax
           FFT_CLR(x + dx, y + dy).real = clr(x, y)
       Next y
   Next x
   FFT_2D_InPlace nx, ny, FFT_CLR
   For x = 0 To m_iNiquistX
       For y = 0 To m_iNiquistY
           FFT_CLR(x, y).real = FFT_CLR(x, y).real * m_sMTF(x, y)
           FFT_CLR(nx - x, y).real = FFT_CLR(nx - x, y).real * m_sMTF(x, y)
           FFT_CLR(nx - x, ny - y).real = FFT_CLR(nx - x, ny - y).real * m_sMTF(x, y)
           FFT_CLR(x, ny - y).real = FFT_CLR(x, ny - y).real * m_sMTF(x, y)
       Next y
   Next x
   FFT_2D_InPlace nx, ny, FFT_CLR, True
   For x = 0 To xMax
       For y = 0 To yMax
           lClr = FFT_CLR(x + dx, y + dy).real
           ltob lClr
           clr(x, y) = lClr
      Next y
   Next x
   CopyMemory abPicture(0, 0), clr(0, 0), (xMax + 1) * (yMax + 1)
End Sub

Private Sub Class_Initialize()
   Radius = 3
   BPP = TrueColor_24
End Sub



