Attribute VB_Name = "mMath"
Public Type COMPLEX
    real As Double
    imag As Double
End Type

Public Const PI As Double = 3.14159265358979

Private aPower2(0 To 31) As Long
Private aSine(0 To 359) As Single
Private aCosine(0 To 359) As Single
Private aInvSine(0 To 359) As Single
Private aInvCosine(0 To 359) As Single

'**********Missing VB trigonomethric functions***************

Property Get Arccos(ByVal x As Double) As Double
   If Abs(x) <> 1 Then
       Arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
    Else
       Arccos = IIf(x = 1, 0, Atn(1) * 4)
    End If
End Property

Property Get Arcsin(ByVal x As Double) As Double
   If Abs(x) <> 1 Then
       Arcsin = Atn(x / Sqr(-x * x + 1))
    Else
       Arcsin = IIf(x = 1, Atn(1) * 2, Atn(1) * 6)
    End If
End Property

Property Get Atan2(ByVal y As Double, ByVal x As Double) As Double
   If x = 0 And y = 0 Then
      Atan2 = 0
   Else
      If x = 0 Then x = 0.0001
      Atan2 = Atn(y / x) - PI * (x < 0)
   End If
End Property

Property Get Rad(ByVal x As Double) As Double
  Rad = x * PI / 180
End Property

Property Get Deg(ByVal x As Double) As Double
  Deg = x * 180 / PI
End Property

'********Trigonomethric functions with argument in degrees*******

Property Get Sind(ByVal x As Double) As Double
   Sind = Sin(Rad(x))
End Property

Property Get Cosd(ByVal x As Double) As Double
   Cosd = Cos(Rad(x))
End Property

'*********Normalize angle within 0 - 360 degrees************
Property Get NormalizedAngle(ByVal x As Double) As Double
   Dim ret As Double
   ret = x - Int(x / 360#) * 360#
   If ret < 0 Then ret = ret + 360
   NormalizedAngle = ret
End Property

'******integer power of 2 wrapper*************
Property Get Power2(ByVal i As Integer) As Long
    If aPower2(0) = 0 Then
        aPower2(0) = &H1&
        aPower2(1) = &H2&
        aPower2(2) = &H4&
        aPower2(3) = &H8&
        aPower2(4) = &H10&
        aPower2(5) = &H20&
        aPower2(6) = &H40&
        aPower2(7) = &H80&
        aPower2(8) = &H100&
        aPower2(9) = &H200&
        aPower2(10) = &H400&
        aPower2(11) = &H800&
        aPower2(12) = &H1000&
        aPower2(13) = &H2000&
        aPower2(14) = &H4000&
        aPower2(15) = &H8000&
        aPower2(16) = &H10000
        aPower2(17) = &H20000
        aPower2(18) = &H40000
        aPower2(19) = &H80000
        aPower2(20) = &H100000
        aPower2(21) = &H200000
        aPower2(22) = &H400000
        aPower2(23) = &H800000
        aPower2(24) = &H1000000
        aPower2(25) = &H2000000
        aPower2(26) = &H4000000
        aPower2(27) = &H8000000
        aPower2(28) = &H10000000
        aPower2(29) = &H20000000
        aPower2(30) = &H40000000
        aPower2(31) = &H80000000
    End If
    Power2 = aPower2(i)
End Property

'Sine, Cosine and inverse wrappers with integer args in degrees******
Property Get Sine(ByVal x As Integer) As Single
   Dim i As Integer
   If aSine(0) = 0 Then
      For i = 0 To 359
          aSine(i) = Sind(i)
      Next i
   End If
   Sine = aSine(x)
End Property

Property Get Cosine(ByVal x As Integer) As Single
   Dim i As Integer
   If aCosine(90) = 0 Then
      For i = 0 To 359
          aCosine(i) = Cosd(i)
      Next i
   End If
   Cosine = aCosine(x)
End Property

Property Get InvSine(ByVal x As Integer) As Single
   Dim i As Integer
   If aInvSine(0) = 0 Then
      For i = 0 To 359
          aInvSine(i) = Sind(i) * (-1)
      Next i
   End If
   InvSine = aInvSine(x)
End Property

Property Get InvCosine(ByVal x As Integer) As Single
   Dim i As Integer
   If aInvCosine(90) = 0 Then
      For i = 0 To 359
          aInvCosine(i) = Cosd(i) * (-1)
      Next i
   End If
   InvCosine = aInvCosine(x)
End Property

'******Bessel function****************
Property Get Bessel_J1(ByVal x As Single) As Single
   Dim i As Long, b As Double, c As Double, r As Long
   Dim y As Double, e As Double, s As Double
   y = (x / 2) ^ 2
   b = 1: c = 1: e = 1: r = 1
   Do While Abs(e) > 0.00000001
      i = i + 1: b = b * i: c = (1 + i) * c
      r = -r
      e = r * (y ^ i) / b / c
      s = s + e
   Loop
   Bessel_J1 = (1 + s) * (x / 2)
End Property

'**********Fast fourier transform**************
'**************Helper routines*****************

Private Function NumberOfBitsNeeded(PowerOfTwo As Long) As Byte
    Dim i As Byte
    For i = 0 To 31
        If (PowerOfTwo And Power2(i)) <> 0 Then
            NumberOfBitsNeeded = i
            Exit Function
        End If
    Next
End Function

Private Function IsPowerOfTwo(x As Long) As Boolean
    If (x < 2) Then IsPowerOfTwo = False: Exit Function
    If (x And (x - 1)) = False Then IsPowerOfTwo = True
End Function

Private Function ReverseBits(ByVal Index As Long, NumBits As Byte) As Long
    Dim i As Byte, Rev As Long
    For i = 0 To NumBits - 1
        Rev = (Rev * 2) Or (Index And 1)
        Index = Index \ 2
    Next
    ReverseBits = Rev
End Function

'****************Main function*****************
Private Sub DoFFT(ByVal NumSamples As Long, cIn() As COMPLEX, cOut() As COMPLEX, Optional bReverse As Boolean)
    Dim AngleNumerator As Double
    Dim NumBits As Byte, i As Long, j As Long, k As Long, n As Long, BlockSize As Long, BlockEnd As Long
    Dim DeltaAngle As Double, DeltaAr As Double
    Dim Alpha As Double, Beta As Double
    Dim TR As Double, TI As Double, ar As Double, AI As Double
       
    If (IsPowerOfTwo(NumSamples) = False) Then
        Call MsgBox("Error in procedure Fourier:" + vbCrLf + " NumSamples is " + CStr(NumSamples) + ", which is not a positive integer power of two.", , "Error!")
        Exit Sub
    End If
    
    If bReverse Then
       AngleNumerator = -2# * PI
    Else
       AngleNumerator = 2# * PI
    End If
   
    NumBits = NumberOfBitsNeeded(NumSamples)
    For i = 0 To (NumSamples - 1)
        j = ReverseBits(i, NumBits)
        cOut(j) = cIn(i)
    Next
    
    BlockEnd = 1
    BlockSize = 2
    
    Do While BlockSize <= NumSamples
        DeltaAngle = AngleNumerator / BlockSize
        Alpha = Sin(0.5 * DeltaAngle)
        Alpha = 2# * Alpha * Alpha
        Beta = Sin(DeltaAngle)
        i = 0
        Do While i < NumSamples
            ar = 1#
            AI = 0#
            j = i
            For n = 0 To BlockEnd - 1
                k = j + BlockEnd
                TR = ar * cOut(k).real - AI * cOut(k).imag
                TI = AI * cOut(k).real + ar * cOut(k).imag
                cOut(k).real = cOut(j).real - TR
                cOut(k).imag = cOut(j).imag - TI
                cOut(j).real = cOut(j).real + TR
                cOut(j).imag = cOut(j).imag + TI
                DeltaAr = Alpha * ar + Beta * AI
                AI = AI - (Alpha * AI - Beta * ar)
                ar = ar - DeltaAr
                j = j + 1
            Next
            i = i + BlockSize
        Loop
        BlockEnd = BlockSize
        BlockSize = BlockSize * 2
    Loop
    If bReverse Then  'Normalize output for reverse transform
       For i = 0 To NumSamples - 1
          cOut(i).real = cOut(i).real / NumSamples
          cOut(i).imag = cOut(i).imag / NumSamples
       Next i
    End If
End Sub

'*******Class wrappers for main function***********
Public Sub FFT(ByVal NumSamples As Long, cIn() As COMPLEX, cOut() As COMPLEX, Optional bReverse As Boolean)
   Call DoFFT(NumSamples, cIn(), cOut(), bReverse)
End Sub

'Same as above, but replacing IN array with output values
Public Sub FFT_InPlace(ByVal NumSamples As Long, cIn() As COMPLEX, Optional bReverse As Boolean)
   Dim cOut() As COMPLEX
   ReDim cOut(UBound(cIn))
   Call DoFFT(NumSamples, cIn(), cOut(), bReverse)
   CopyMemory cIn(0), cOut(0), (UBound(cOut) + 1) * Len(cOut(0))
End Sub

Public Sub FFT_2D(ByVal nx As Long, ByVal ny As Long, cIn() As COMPLEX, cOut() As COMPLEX, Optional bReverse As Boolean)
   CopyMemory cOut(0, 0), cIn(0, 0), (UBound(cIn, 1) + 1) * (UBound(cIn, 2) + 1) * Len(cIn(0))
   FFT_2D_InPlace nx, ny, cOut, bReverse
End Sub

Public Sub FFT_2D_InPlace(ByVal nx As Long, ByVal ny As Long, cIn() As COMPLEX, Optional bReverse As Boolean)
   Dim c() As COMPLEX
   Dim i As Long, j As Long
'Transform rows
   ReDim c(nx)
   For j = 0 To ny - 1
      For i = 0 To nx - 1
          c(i) = cIn(i, j)
      Next i
      Call FFT_InPlace(nx, c, bReverse)
      For i = 0 To nx - 1
          cIn(i, j) = c(i)
      Next i
   Next j
'Transform columns
   ReDim c(ny)
   For i = 0 To nx - 1
       For j = 0 To ny - 1
           c(j) = cIn(i, j)
       Next j
       Call FFT_InPlace(ny, c, bReverse)
       For j = 0 To ny - 1
           cIn(i, j) = c(j)
       Next j
   Next i
End Sub


Public Sub itob(iVal As Integer)
   If iVal < 0 Then
      iVal = 0
   ElseIf iVal > 255 Then
      iVal = 255
   End If
End Sub

Public Sub ltob(lVal As Long)
   If lVal < 0 Then
      lVal = 0
   ElseIf lVal > 255 Then
      lVal = 255
   End If
End Sub



