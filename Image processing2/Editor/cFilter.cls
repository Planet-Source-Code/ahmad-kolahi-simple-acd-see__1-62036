VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum KernelFilterTypes
'Assuming: a - KernelSize/2
'          r = Sqr(x^2 + y^2) - distance from kernel center
'          Sigma - Gaussian standard deviation
    [_Min]
    eRectangle 'aka average,uniform,median
               'PSF = 1
    eCylinder  'aka circular,pill-box
               'PSF =  1   if r < a
               '       1/2 if r = a
               '       0 otherwise
    eGaussian  'Gaussian distribution
               'PSF = Exp((r^2)/(2 * Sigma^2))
               'or
               'PSF = Exp(-PI*(r/a)^2)
    eCone      'aka triangle-circular
    ePyramid   'aka triangle-rectangle
    eJinc      'aka sombrerro_2D (2D analog of Sinc)
               'PSF = (J1(PI * r)/(PI * r)
               'Where J1 is Besselian function of first kind
    eSinc      'aka sombrerro_1D
               'PSF = (Sin(PI * r)/(PI * r)
    ePeak      'PSF = 1/r
    eExpDecay  'PSF = Exp(-a * r)
    eAiry      'aka difraction-circle,releygh
    eLaplacian 'Laplacian distribution
    eCustom    'Use custom kernel
    eEdgeGradient
    eEdgeEmboss
    eEdgeSobel
    eEdgePrewitt
    [_Max]
End Enum

Public Enum RankFilterTypes
    [_Min]
    eMedian
    eMin
    eMax
    [_Max]
End Enum

Public Enum eEdgeDir
    NORTH
    NE
    EAST
    SE
    SOUTH
    SW
    WEST
    NW
End Enum

'Events
Public Event ProgressInit(ByVal nMax As Long)
Public Event ProgressChanged(ByVal nValue As Long)
Public Event ProgressEnd(ByVal nTime As Long)

'local variable(s) to hold property value(s)
Private m_iSize As Long
Private m_sGaussian_SD As Single
Private m_iFilterPower As Long
Private m_iBPP As eBPP

'Other local variables
Private m_iOffset As Long
Private m_iFilt() As Long
Private m_iWeight As Long
Private m_iBias As Long

Public Property Let BPP(ByVal vData As eBPP)
    m_iBPP = vData
End Property

Public Property Get BPP() As eBPP
    BPP = m_iBPP
End Property

Public Property Let FilterPower(ByVal vData As Long)
    m_iFilterPower = vData
    If m_iFilterPower = 0 Then m_iFilterPower = 1
End Property

Public Property Get FilterPower() As Long
    FilterPower = m_iFilterPower
End Property

Public Property Let KernelSize(ByVal vData As Long)
    If (vData And 1) = False Then
        m_iSize = vData + 1
    Else
        m_iSize = vData
    End If
    m_iOffset = m_iSize \ 2
End Property

Public Property Get KernelSize() As Long
    KernelSize = m_iSize
End Property

Public Property Let Gaussian_SD(ByVal vData As Single)
    m_sGaussian_SD = vData
    If m_sGaussian_SD = 0 Then m_sGaussian_SD = 0.7
End Property

Public Property Get Gaussian_SD() As Single
    Gaussian_SD = m_sGaussian_SD
End Property

'***********************Preparing filtering kernel************************
Public Sub PrepareKernel(ByVal fType As KernelFilterTypes, ByVal bReverse As Boolean, Optional Flt As Variant, Optional eDir As eEdgeDir)
    Dim i As Long, j As Long
    Select Case fType
        Case eCustom
             Call CustomFilter(bReverse, Flt)
        Case eGaussian
             Call GaussianFilter(bReverse)
        Case eCylinder
             Call CylFilter(bReverse)
        Case eRectangle
             Call RectFilter(bReverse)
        Case eJinc
             Call SombFilter(bReverse)
        Case eCone
             Call ConeFilter(bReverse)
        Case ePyramid
             Call PyramidFilter(bReverse)
        Case eSinc, eAiry
             Call SincFilter(bReverse)
        Case ePeak
             Call PeakFilter(bReverse)
        Case eExpDecay
             Call ExpDecayFilter(bReverse)
        Case eLaplacian
             Call LaplacianFilter(bReverse)
        Case eEdgeGradient, eEdgeEmboss, eEdgePrewitt, eEdgeSobel
        
    End Select
    If fType <> eCustom Then
       ReDim Flt(LBound(m_iFilt, 1) To UBound(m_iFilt, 1), LBound(m_iFilt, 2) To UBound(m_iFilt, 2))
       For i = LBound(m_iFilt, 1) To UBound(m_iFilt, 1)
           For j = LBound(m_iFilt, 2) To UBound(m_iFilt, 2)
               Flt(i, j) = m_iFilt(i, j)
           Next j
       Next i
    End If
    If m_iWeight = 0 Then m_iWeight = 1
End Sub

Private Sub CustomFilter(ByVal bReverse As Boolean, Flt As Variant)
    Dim i As Long, j As Long, x As Long, y As Long, v As Long
    m_iOffset = m_iSize \ 2
    ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
    m_iWeight = 0
    For i = -m_iOffset To m_iOffset
        x = i + m_iOffset
        For j = -m_iOffset To m_iOffset
            y = j + m_iOffset
            On Error Resume Next
            v = Flt(x * m_iSize + y)
            m_iWeight = m_iWeight + v
            If bReverse Then
               m_iFilt(i, j) = -v
            Else
               m_iFilt(i, j) = v
            End If
        Next j
    Next i
    If bReverse Then m_iFilt(0, 0) = m_iWeight * 2 + m_iFilt(0, 0)
End Sub

Private Sub ConeFilter(ByVal bReverse As Boolean)
    Dim i As Long, j As Long, r As Double, iScale As Long
    Dim v As Single
    m_iOffset = m_iSize \ 2
    ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
    m_iWeight = 0
    For i = -m_iOffset To m_iOffset
        For j = -m_iOffset To m_iOffset
            r = Sqr(i * i + j * j)
            If r < m_iOffset Then
               v = FilterPower * (1 - r / m_iOffset) * m_iSize
            Else
               v = 0
            End If
            m_iWeight = m_iWeight + v
            If bReverse Then
               m_iFilt(i, j) = -v
            Else
               m_iFilt(i, j) = v
            End If
        Next j
    Next i
    If bReverse Then m_iFilt(0, 0) = m_iWeight * 2 + m_iFilt(0, 0)
End Sub

Private Sub CylFilter(ByVal bReverse As Boolean)
    Dim i As Long, j As Long, r As Double, iScale As Long
    Dim v As Long
    m_iOffset = m_iSize \ 2
    ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
    m_iWeight = 0
    For i = -m_iOffset To m_iOffset
        For j = -m_iOffset To m_iOffset
            r = Sqr(i * i + j * j)
            If r < m_iOffset Then
               v = 2 * m_iFilterPower
            ElseIf r = m_iOffset Then
               v = m_iFilterPower
            Else
               v = 0
            End If
            m_iWeight = m_iWeight + v
            If bReverse Then
               m_iFilt(i, j) = -v
            Else
               m_iFilt(i, j) = v
            End If
        Next j
    Next i
    If bReverse Then m_iFilt(0, 0) = m_iWeight * 2 + m_iFilt(0, 0)
End Sub

Private Sub ExpDecayFilter(Optional bReverse As Boolean)
    Dim i As Long, j As Long, v As Double, iScale As Double
    Dim r As Single
    m_iOffset = m_iSize \ 2
    iScale = 1 / Exp(-(m_iOffset * m_iOffset))
    While iScale > &HFFFF&
        KernelSize = m_iSize - 2
        m_iOffset = m_iSize \ 2
        iScale = 1 / Exp(-(m_iOffset * m_iOffset))
    Wend
    ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
    m_iWeight = 0
    For i = -m_iOffset To m_iOffset
        For j = -m_iOffset To m_iOffset
            r = (i * i + j * j)
            v = Exp(-r * m_iOffset)
            m_iWeight = m_iWeight + v * iScale
            If bReverse Then
               m_iFilt(i, j) = -v * iScale
            Else
               m_iFilt(i, j) = v * iScale
            End If
        Next j
    Next i
    If bReverse Then m_iFilt(0, 0) = m_iWeight * 2 + m_iFilt(0, 0)
End Sub

Private Sub GaussianFilter(Optional bReverse As Boolean)
    Dim i As Long, j As Long, v As Double, iScale As Double
    Dim sd As Single
    sd = Gaussian_SD * Gaussian_SD * 2
    m_iOffset = m_iSize \ 2
    iScale = 1 / Exp(-(m_iOffset * m_iOffset) / sd)
    While iScale > &HFFFF&
        KernelSize = m_iSize - 2
        m_iOffset = m_iSize \ 2
        iScale = 1 / Exp(-(m_iOffset * m_iOffset) / sd)
    Wend
    ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
    m_iWeight = 0
    For i = -m_iOffset To m_iOffset
        For j = -m_iOffset To m_iOffset
            v = Exp(-(i * i + j * j) / sd)
            m_iWeight = m_iWeight + v * iScale
            If bReverse Then
               m_iFilt(i, j) = -v * iScale
            Else
               m_iFilt(i, j) = v * iScale
            End If
        Next j
    Next i
    If bReverse Then m_iFilt(0, 0) = m_iWeight * 2 + m_iFilt(0, 0)
End Sub

Private Sub LaplacianFilter(Optional bReverse As Boolean)
    Dim i As Long, j As Long, v As Double, iScale As Double
    Dim sd As Single
    sd = Gaussian_SD
    m_iOffset = m_iSize \ 2
    iScale = 1 / (sd * Exp(-(sd * m_iOffset)) / 2)
    While iScale > &HFFFF&
        KernelSize = m_iSize - 2
        m_iOffset = m_iSize \ 2
        iScale = 1 / (sd * Exp(-(sd * m_iOffset)) / 2)
    Wend
    ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
    m_iWeight = 0
    For i = -m_iOffset To m_iOffset
        For j = -m_iOffset To m_iOffset
            v = sd * Exp(-(sd * Sqr(i * i + j * j))) / 2
            m_iWeight = m_iWeight + v * iScale
            If bReverse Then
               m_iFilt(i, j) = -v * iScale
            Else
               m_iFilt(i, j) = v * iScale
            End If
        Next j
    Next i
    If bReverse Then m_iFilt(0, 0) = m_iWeight * 2 + m_iFilt(0, 0)
End Sub

Private Sub GaussianFilter_2(Optional bReverse As Boolean)
    Dim i As Long, j As Long, v As Double, iScale As Double
    Dim sd As Single
    m_iOffset = m_iSize \ 2
    iScale = 1 / Exp(-PI / 4)
    ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
    sd = m_iSize * m_iSize
    m_iWeight = 0
    For i = -m_iOffset To m_iOffset
        For j = -m_iOffset To m_iOffset
            v = Exp(-PI * (i * i + j * j) / (sd))
            m_iWeight = m_iWeight + v * iScale
            If bReverse Then
               m_iFilt(i, j) = -v * iScale
            Else
               m_iFilt(i, j) = v * iScale
            End If
        Next j
    Next i
    If bReverse Then m_iFilt(0, 0) = m_iWeight * 2 + m_iFilt(0, 0)
End Sub

Private Sub PeakFilter(Optional bReverse As Boolean)
    Dim i As Long, j As Long, r As Double, iScale As Long
    Dim v As Single
    m_iOffset = m_iSize \ 2
    ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
    m_iWeight = 0
    For i = -m_iOffset To m_iOffset
        For j = -m_iOffset To m_iOffset
            r = (i * i + j * j)
            If r = 0 Then r = 0.001
            v = 2 * m_iSize * FilterPower / r
            m_iWeight = m_iWeight + v
            If bReverse Then
               m_iFilt(i, j) = -v
            Else
               m_iFilt(i, j) = v
            End If
        Next j
    Next i
    If bReverse Then m_iFilt(0, 0) = m_iWeight * 2 + m_iFilt(0, 0)
End Sub

Private Sub PyramidFilter(Optional bReverse As Boolean)
    Dim i As Long, j As Long, r As Double, iScale As Long
    Dim v As Single
    m_iOffset = m_iSize \ 2
    ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
    m_iWeight = 0
    For i = -m_iOffset To m_iOffset
        For j = -m_iOffset To m_iOffset
            v = m_iSize * FilterPower * (m_iOffset - Abs(i)) * (m_iOffset - Abs(j)) / (m_iOffset * m_iOffset)
            m_iWeight = m_iWeight + v
            If bReverse Then
               m_iFilt(i, j) = -v
            Else
               m_iFilt(i, j) = v
            End If
        Next j
    Next i
    If bReverse Then m_iFilt(0, 0) = m_iWeight * 2 + m_iFilt(0, 0)
End Sub

Private Sub RectFilter(Optional bReverse As Boolean)
    Dim i As Long, j As Long, r As Double, iScale As Long
    Dim v As Single
    m_iOffset = m_iSize \ 2
    ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
    m_iWeight = 0
    For i = -m_iOffset To m_iOffset
        For j = -m_iOffset To m_iOffset
            v = FilterPower
            m_iWeight = m_iWeight + v
            If bReverse Then
               m_iFilt(i, j) = -v
            Else
               m_iFilt(i, j) = v
            End If
        Next j
    Next i
    If bReverse Then m_iFilt(0, 0) = m_iWeight * 2 + m_iFilt(0, 0)
End Sub

Private Sub SincFilter(Optional bReverse As Boolean)
    Dim i As Long, j As Long, r As Double
    Dim v As Single, a As Single, iScale As Double
    m_iOffset = m_iSize \ 2
    ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
    m_iWeight = 0
    iScale = 2 * m_iSize
    For i = -m_iOffset To m_iOffset
        For j = -m_iOffset To m_iOffset
            a = PI * Sqr(i * i + j * j) / m_iOffset
            If a = 0 Then a = 0.00001
            v = 2 * FilterPower * Sin(a) / a
            If v < 0 Then v = 0
            m_iWeight = m_iWeight + v * iScale
            If bReverse Then
               m_iFilt(i, j) = -v * iScale
            Else
               m_iFilt(i, j) = v * iScale
            End If
        Next j
    Next i
    If bReverse Then m_iFilt(0, 0) = m_iWeight * 2 + m_iFilt(0, 0)
End Sub

Private Sub SombFilter(Optional bReverse As Boolean)
    Dim i As Long, j As Long, r As Double
    Dim v As Single, a As Single, iScale As Double
    m_iOffset = m_iSize \ 2
    ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
    m_iWeight = 0
    iScale = 1 / Abs(2 * Bessel_J1(0.84 * PI) / (0.84 * PI))
    For i = -m_iOffset To m_iOffset
        For j = -m_iOffset To m_iOffset
            a = PI * Sqr(i * i + j * j) / m_iOffset
            If a = 0 Then a = 0.00001
            v = 2 * FilterPower * Bessel_J1(a) / a
            If v < 0 Then v = 0
            m_iWeight = m_iWeight + v * iScale
            If bReverse Then
               m_iFilt(i, j) = -v * iScale
            Else
               m_iFilt(i, j) = v * iScale
            End If
        Next j
    Next i
    If bReverse Then m_iFilt(0, 0) = m_iWeight * 2 + m_iFilt(0, 0)
End Sub

'******************Filtering procedures**********************
Public Sub Convolve(abPicture() As Byte, PSF As KernelFilterTypes, ParamArray Flt())
   If (PSF <= KernelFilterTypes.[_Min] Or PSF >= KernelFilterTypes.[_Max]) Then Exit Sub
   Dim fltArray As Variant
   fltArray = Flt
   Call PrepareKernel(PSF, False, fltArray)
   If BPP > HighColor_16 Then
      If bAPISupported And m_iSize > 3 Then
         DoStdFltTrueColor_New abPicture
      Else
         DoStdFltTrueColor abPicture
      End If
   Else
      DoStdFltPalette abPicture
   End If
End Sub

Public Sub DeConvolve(abPicture() As Byte, PSF As KernelFilterTypes, ParamArray Flt())
   If (PSF <= KernelFilterTypes.[_Min] Or PSF >= KernelFilterTypes.[_Max]) Then Exit Sub
   Dim fltArray As Variant
   fltArray = Flt
   Call PrepareKernel(PSF, True, fltArray)
   If BPP > HighColor_16 Then
      If bAPISupported And m_iSize > 3 Then
         DoStdFltTrueColor_New abPicture
      Else
         DoStdFltTrueColor abPicture
      End If
   Else
      DoStdFltPalette abPicture
   End If
End Sub

Public Sub RankFilter(abPicture() As Byte, fType As RankFilterTypes)
   If fType = eMedian Then
      Convolve abPicture, eRectangle
      Exit Sub
   End If
   If BPP > HighColor_16 Then
      If bAPISupported And m_iSize > 3 Then
         DoRankFltTrueColor_New abPicture, fType
      Else
         DoRankFltTrueColor abPicture, fType
      End If
   Else
      DoRankFltPalette abPicture, fType
   End If
End Sub

'********** Filters 24BPP and 32BPP***************
Private Sub DoStdFltTrueColor(abPicture() As Byte)
   Dim abBuffer() As Byte
   Dim x As Long, y As Long
   Dim rgbOffset As Long, xOffset As Long
   Dim r As Long, g As Long, b As Long
   Dim i As Long, j As Long, yMax As Long, xMax As Long
   Dim rR As Long, rB As Long, rG As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   ReDim abBuffer(UBound(abPicture, 1), UBound(abPicture, 2))
   CopyMemory abBuffer(0, 0), abPicture(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
   rgbOffset = m_iOffset * 3
   xMax = UBound(abPicture, 1) - 3
   xMax = xMax - rgbOffset
   yMax = UBound(abPicture, 2) - m_iOffset
   RaiseEvent ProgressInit(xMax - rgbOffset)
   For x = rgbOffset To xMax Step 3
       For y = m_iOffset To yMax
           r = 0: g = 0: b = 0
           For i = -m_iOffset To m_iOffset
               xOffset = i * 3
               For j = -m_iOffset To m_iOffset
                   r = r + m_iFilt(i, j) * abPicture(x + xOffset, y + j)
                   g = g + m_iFilt(i, j) * abPicture(x + 1 + xOffset, y + j)
                   b = b + m_iFilt(i, j) * abPicture(x + 2 + xOffset, y + j)
               Next j
           Next i
           rR = r \ m_iWeight + m_iBias: rG = g \ m_iWeight + m_iBias: rB = b \ m_iWeight + m_iBias
           ltob rR
           ltob rG
           ltob rB
           abBuffer(x, y) = rR: abBuffer(x + 1, y) = rG: abBuffer(x + 2, y) = rB
        Next y
        If (x Mod 45) = 0 Then
           RaiseEvent ProgressChanged(x)
        End If
    Next x
    CopyMemory abPicture(0, 0), abBuffer(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
    RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Private Sub DoRankFltTrueColor(abPicture() As Byte, fType As RankFilterTypes)
   Dim abBuffer() As Byte
   Dim x As Long, y As Long
   Dim rgbOffset As Long, xOffset As Long
   Dim i As Long, j As Long, yMax As Long, xMax As Long
   Dim rR As Long, rB As Long, rG As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   ReDim abBuffer(UBound(abPicture, 1), UBound(abPicture, 2))
   CopyMemory abBuffer(0, 0), abPicture(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
   rgbOffset = m_iOffset * 3
   xMax = UBound(abPicture, 1) - rgbOffset
   yMax = UBound(abPicture, 2) - m_iOffset
   RaiseEvent ProgressInit(xMax - rgbOffset)
   For x = rgbOffset To xMax Step 3
       For y = m_iOffset To yMax
           If fType = eMax Then
              rR = 0: rG = 0: rB = 0
           Else
              rR = 255: rG = 255: rB = 255
           End If
           For i = -m_iOffset To m_iOffset
               xOffset = i * 3
               For j = -m_iOffset To m_iOffset
                    Select Case fType
                    Case eMin
                        If abPicture(x + xOffset, y + j) < rB Then
                            rB = abPicture(x + xOffset, y + j)
                        End If
                        If abPicture(x + 1 + xOffset, y + j) < rG Then
                            rG = abPicture(x + 1 + xOffset, y + j)
                        End If
                        If abPicture(x + 2 + xOffset, y + j) < rR Then
                            rR = abPicture(x + 2 + xOffset, y + j)
                        End If
                    Case eMax
                        If abPicture(x + xOffset, y + j) > rB Then
                           rB = abPicture(x + xOffset, y + j)
                        End If
                        If abPicture(x + 1 + xOffset, y + j) > rG Then
                            rG = abPicture(x + 1 + xOffset, y + j)
                        End If
                        If abPicture(x + 2 + xOffset, y + j) > rR Then
                            rR = abPicture(x + 2 + xOffset, y + j)
                        End If
                        Case Else
                    End Select
               Next j
           Next i
           ltob rR
           ltob rG
           ltob rB
           abBuffer(x, y) = rB: abBuffer(x + 1, y) = rG: abBuffer(x + 2, y) = rR
        Next y
        If (x Mod 45) = 0 Then
           RaiseEvent ProgressChanged(x)
        End If
    Next x
    CopyMemory abPicture(0, 0), abBuffer(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
    RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

'*************Kernel Filter 24BPP and 32BPP with IE5 HLS2RGB support****************
Private Sub DoStdFltTrueColor_New(abPicture() As Byte)
   Dim abBuffer() As Byte
   Dim HLSOld() As HLSTRIPPLE
   Dim HLSNew() As HLSTRIPPLE
   Dim x As Long, y As Long, xPos As Long
   Dim i As Long, j As Long, yMax As Long, xMax As Long
   Dim clr As Long, luma As Double, lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   ReDim abBuffer(xMax, yMax)
   CopyMemory abBuffer(0, 0), abPicture(0, 0), (xMax + 1) * (yMax + 1)
   xMax = xMax \ 3
   ReDim HLSOld(xMax, yMax)
   ReDim HLSNew(xMax, yMax)
   RaiseEvent ProgressInit(xMax)
   For x = 0 To xMax
       xPos = x * 3
       For y = 0 To yMax
           CopyMemory clr, abPicture(xPos, y), 3
           Call ColorRGBToHLS(clr, HLSOld(x, y).H, HLSOld(x, y).L, HLSOld(x, y).s)
       Next y
       If (x Mod 15) = 0 Then
          RaiseEvent ProgressChanged((x) / 50)
       End If
   Next x
   CopyMemory HLSNew(0, 0), HLSOld(0, 0), (xMax + 1) * (yMax + 1) * 6
   For x = m_iOffset To xMax - m_iOffset
       xPos = x * 3
       For y = m_iOffset To yMax - m_iOffset
           luma = 0
           For i = -m_iOffset To m_iOffset
               For j = -m_iOffset To m_iOffset
                   luma = luma + m_iFilt(i, j) * HLSOld(x + i, y + j).L
               Next j
           Next i
           HLSNew(x, y).L = luma \ m_iWeight + m_iBias
           clr = ColorHLSToRGB(HLSNew(x, y).H, HLSNew(x, y).L, HLSNew(x, y).s)
           CopyMemory abBuffer(xPos, y), clr, 3
        Next y
        If (x Mod 15) = 0 Then
           RaiseEvent ProgressChanged(x)
        End If
    Next x
    CopyMemory abPicture(0, 0), abBuffer(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
    RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

'*************Rank Filter 24BPP and 32BPP with IE5 HLS2RGB support****************
Private Sub DoRankFltTrueColor_New(abPicture() As Byte, fType As RankFilterTypes)
   Dim abBuffer() As Byte
   Dim HLSOld() As HLSTRIPPLE
   Dim HLSNew() As HLSTRIPPLE
   Dim x As Long, y As Long, xPos As Long
   Dim i As Long, j As Long, yMax As Long, xMax As Long
   Dim clr As Long, luma As Long, lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   ReDim abBuffer(xMax, yMax)
   CopyMemory abBuffer(0, 0), abPicture(0, 0), (xMax + 1) * (yMax + 1)
   xMax = xMax \ 3
   ReDim HLSOld(xMax, yMax)
   ReDim HLSNew(xMax, yMax)
   RaiseEvent ProgressInit(xMax)
   For x = 0 To xMax
       xPos = x * 3
       For y = 0 To yMax
           CopyMemory clr, abPicture(xPos, y), 3
           Call ColorRGBToHLS(clr, HLSOld(x, y).H, HLSOld(x, y).L, HLSOld(x, y).s)
       Next y
       If (x Mod 15) = 0 Then
          RaiseEvent ProgressChanged(x / 50)
       End If
   Next x
   CopyMemory HLSNew(0, 0), HLSOld(0, 0), (xMax + 1) * (yMax + 1) * 6
   For x = m_iOffset To xMax - m_iOffset
       xPos = x * 3
       For y = m_iOffset To yMax - m_iOffset
           If fType = eMax Then luma = 0 Else luma = 255
           For i = -m_iOffset To m_iOffset
               For j = -m_iOffset To m_iOffset
                   If fType = eMax Then
                      If HLSOld(x + i, y + j).L > luma Then
                         luma = HLSOld(x + i, y + j).L
                      End If
                   Else
                      If HLSOld(x + i, y + j).L < luma Then
                         luma = HLSOld(x + i, y + j).L
                      End If
                   End If
               Next j
           Next i
           HLSNew(x, y).L = luma
           clr = ColorHLSToRGB(HLSNew(x, y).H, HLSNew(x, y).L, HLSNew(x, y).s)
           CopyMemory abBuffer(xPos, y), clr, 3
        Next y
        If (x Mod 15) = 0 Then
           RaiseEvent ProgressChanged(x)
        End If
    Next x
    CopyMemory abPicture(0, 0), abBuffer(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
    RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

'********** Filter Palletized DIBs (BPP < 24)***************

Private Sub DoStdFltPalette(abPicture() As Byte)
   Dim abBuffer() As Byte
   Dim x As Long, y As Long
   Dim clr As Long
   Dim i As Long, j As Long, yMax As Long, xMax As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   ReDim abBuffer(UBound(abPicture, 1), UBound(abPicture, 2))
   CopyMemory abBuffer(0, 0), abPicture(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
   xMax = UBound(abPicture, 1) - m_iOffset
   yMax = UBound(abPicture, 2) - m_iOffset
   RaiseEvent ProgressInit(xMax - m_iOffset)
   For x = m_iOffset To xMax
       For y = m_iOffset To yMax
           clr = 0
           For i = -m_iOffset To m_iOffset
               For j = -m_iOffset To m_iOffset
                   clr = clr + m_iFilt(i, j) * abPicture(x + i, y + j)
               Next j
           Next i
           ltob clr
           abBuffer(x, y) = clr
        Next y
        If (x Mod 45) = 0 Then
           RaiseEvent ProgressChanged(x)
        End If
    Next x
    CopyMemory abPicture(0, 0), abBuffer(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
    RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Private Sub DoRankFltPalette(abPicture() As Byte, fType As RankFilterTypes)
   Dim abBuffer() As Byte
   Dim x As Long, y As Long
   Dim clr As Long
   Dim i As Long, j As Long, yMax As Long, xMax As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   ReDim abBuffer(UBound(abPicture, 1), UBound(abPicture, 2))
   CopyMemory abBuffer(0, 0), abPicture(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
   xMax = UBound(abPicture, 1) - m_iOffset
   yMax = UBound(abPicture, 2) - m_iOffset
   RaiseEvent ProgressInit(xMax - m_iOffset)
   For x = m_iOffset To xMax
       For y = m_iOffset To yMax
           If fType = eMax Then clr = 0 Else clr = 255
           For i = -m_iOffset To m_iOffset
               For j = -m_iOffset To m_iOffset
                   If fType = eMax Then
                      If abPicture(x + i, y + j) > clr Then
                         clr = abPicture(x + i, y + j)
                      End If
                   Else
                      If abPicture(x + i, y + j) < clr Then
                         clr = abPicture(x + i, y + j)
                      End If
                   End If
               Next j
           Next i
           ltob clr
           abBuffer(x, y) = clr
        Next y
        If (x Mod 45) = 0 Then
           RaiseEvent ProgressChanged(x)
        End If
    Next x
    CopyMemory abPicture(0, 0), abBuffer(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
    RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub EdgeDetection(abPicture() As Byte, fltType As KernelFilterTypes, ByVal Direction As eEdgeDir, Optional bGray As Boolean)
   Dim abBuffer() As Byte
   Dim x As Long, y As Long
   Dim rgbOffset As Long, xOffset As Long
   Dim r As Long, g As Long, b As Long
   Dim i As Long, j As Long, yMax As Long, xMax As Long
   Dim rR As Long, rB As Long, rG As Long
   Dim lTime As Long
   If (fltType < eEdgeGradient) Or (fltType > eEdgePrewitt) Then Exit Sub
   lTime = GetTickCount
   EdgeFilter fltType, Direction
   If bGray Then m_iBias = 128
   ReDim abBuffer(UBound(abPicture, 1), UBound(abPicture, 2))
   CopyMemory abBuffer(0, 0), abPicture(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
   rgbOffset = m_iOffset * 3
   xMax = UBound(abPicture, 1) - 3
   xMax = xMax - rgbOffset
   yMax = UBound(abPicture, 2) - m_iOffset
   RaiseEvent ProgressInit(xMax - rgbOffset)
   For x = rgbOffset To xMax Step 3
       For y = m_iOffset To yMax
           r = 0: g = 0: b = 0
           For i = -m_iOffset To m_iOffset
               xOffset = i * 3
               For j = -m_iOffset To m_iOffset
                   r = r + m_iFilt(i, j) * abPicture(x + xOffset, y + j)
                   g = g + m_iFilt(i, j) * abPicture(x + 1 + xOffset, y + j)
                   b = b + m_iFilt(i, j) * abPicture(x + 2 + xOffset, y + j)
               Next j
           Next i
           rR = r \ m_iWeight + m_iBias: rG = g \ m_iWeight + m_iBias: rB = b \ m_iWeight + m_iBias
           ltob rR
           ltob rG
           ltob rB
           abBuffer(x, y) = rR: abBuffer(x + 1, y) = rG: abBuffer(x + 2, y) = rB
        Next y
        If (x Mod 45) = 0 Then
           RaiseEvent ProgressChanged(x)
        End If
    Next x
    m_iBias = 0
    CopyMemory abPicture(0, 0), abBuffer(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
    RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Private Sub EdgeFilter(fltType As KernelFilterTypes, eDir As eEdgeDir)
   Dim i As Long, j As Long, lFlt As Long
   Dim varFlt As Variant
   If (fltType < eEdgeGradient) Or (fltType > eEdgePrewitt) Then Exit Sub
   If fltType > eEdgeEmboss Then
      If eDir = NE Then eDir = WEST
   End If
   KernelSize = 3
   m_iOffset = m_iSize \ 2
   m_iWeight = 1
   ReDim m_iFilt(-m_iOffset To m_iOffset, -m_iOffset To m_iOffset)
   Select Case eDir
       Case NORTH: varFlt = Array(1, 1, 1, 0, 0, 0, -1, -1, -1)
       Case NE:    varFlt = Array(0, 1, 1, -1, 0, 1, -1, -1, 0)
       Case EAST:  varFlt = Array(-1, 0, 1, -1, 0, 1, -1, 0, 1)
       Case SE:    varFlt = Array(-1, -1, 0, -1, 0, 1, 0, 1, 1)
       Case SOUTH: varFlt = Array(-1, -1, -1, 0, 0, 0, 1, 1, 1)
       Case SW:    varFlt = Array(0, -1, -1, 1, 0, -1, 1, 1, 0)
       Case WEST:  varFlt = Array(1, 0, -1, 1, 0, -1, 1, 0, -1)
       Case NW:    varFlt = Array(1, 1, 0, 1, 0, -1, 0, -1, -1)
   End Select
   For i = -m_iOffset To m_iOffset
       For j = -m_iOffset To m_iOffset
           lFlt = varFlt((i + m_iOffset) * m_iSize + j + m_iOffset)
           If fltType = eEdgeGradient Then
              If lFlt = 0 Then lFlt = 1
              If i = 0 And j = 0 Then lFlt = -2
           End If
           m_iFilt(i, j) = lFlt
       Next j
   Next i
   If fltType = eEdgeSobel Then
      If eDir = NORTH Then
         m_iFilt(0, -1) = -2
         m_iFilt(0, 1) = 2
      Else
         m_iFilt(-1, 0) = -2
         m_iFilt(1, 0) = 2
      End If
   End If
End Sub

Private Sub Class_Initialize()
   KernelSize = 3
   FilterPower = 1
   Gaussian_SD = 0.7
   BPP = eBPP.TrueColor_24
   bAPISupported = IsAPIColorSupported
End Sub
