VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cImgProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Type DOUBLERGB
   b1 As Byte
   g1 As Byte
   r1 As Byte
   b2 As Byte
   g2 As Byte
   r2 As Byte
End Type

Public Enum RotateDirection
   eLeft
   eRight
End Enum

Public Enum FlipDirection
   eVert
   eHorz
End Enum

Public Event ProgressInit(ByVal nMax As Long)
Public Event ProgressChanged(ByVal nValue As Long)
Public Event ProgressEnd(ByVal nTime As Long)


Dim bTable(0 To 100, 0 To 255) As Byte 'Brightness table
Dim cTable(0 To 100, 0 To 255) As Byte 'Contrast table
Dim gTable(0 To 100, 0 To 255) As Byte 'Gamma table

Public Sub BlackWhite(abPicture() As Byte)
' Converts to Black and WHite using Floyd-Steinberg error diffusion.
   Dim abBuffer() As Byte
   Dim x As Long, y As Long
   Dim i As Long, iCoeff As Long
   Dim xMax As Long, yMax As Long
   Dim lError As Long
   Dim lNew As Long
   Dim iC As Long, iC2 As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   ReDim abBuffer(xMax, yMax)
   CopyMemory abBuffer(0, 0), abPicture(0, 0), (xMax + 1) * (yMax + 1)
   Call GrayScale(abPicture)
   RaiseEvent ProgressInit(xMax)
   For x = 0 To xMax Step 3
       For y = 0 To yMax
            If (abPicture(x, y) > 128) Then
                iC = iC + 1
                abBuffer(x, y) = 255
                abBuffer(x + 1, y) = 255
                abBuffer(x + 2, y) = 255
                lError = (255 - abPicture(x, y)) - 128
            Else
                iC2 = iC2 + 1
                abBuffer(x, y) = 0
                abBuffer(x + 1, y) = 0
                abBuffer(x + 2, y) = 0
                ' Black tolerance:
                If (abPicture(x, y) > 16) Then
                    lError = abPicture(x, y)
                Else
                    lError = 0
                End If
            End If
            If (x < xEnd - 3) Then
                lNew = abPicture(x + 3, y) + (lError * 7) \ 16
                If (lNew > 255) Then lNew = 255
                If (lNew < 0) Then lNew = 0
                abPicture(x + 3, y) = lNew
                abPicture(x + 4, y) = lNew
                abPicture(x + 5, y) = lNew
            End If
            If (y < yMax) Then
                For i = -3 To 3 Step 3
                    If (x + i) > 0 And (x + i) < xMax Then
                        Select Case i
                           Case -3: iCoeff = 3
                           Case 0:  iCoeff = 5
                           Case 3:  iCoeff = 1
                        End Select
                        lNew = abPicture(x + i, y + 1) + (lError * iCoeff) \ 16
                        If (lNew > 255) Then lNew = 255
                        If (lNew < 0) Then lNew = 0
                        abPicture(x + i, y + 1) = lNew
                        abPicture(x + i + 1, y + 1) = lNew
                        abPicture(x + i + 2, y + 1) = lNew
                    End If
                Next i
            End If
        Next y
        If (x Mod 15) = 0 Then
           RaiseEvent ProgressChanged(x)
        End If
    Next x
    RaiseEvent ProgressEnd(GetTickCount - lTime)
    CopyMemory abPicture(0, 0), abBuffer(0, 0), (UBound(abPicture, 1) + 1) * (UBound(abPicture, 2) + 1)
End Sub

Public Sub GrayScale(abPicture() As Byte)
' Gray scale using standard intensity components.
   Dim x As Long, y As Long
   Dim xMax As Long, yMax As Long, xPos As Long
   Dim rgb As RGBTRIPPLE
'   Dim lB As Long, lG As Long, lR As Long
   Dim lGray As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1) \ 3
   yMax = UBound(abPicture, 2)
   RaiseEvent ProgressInit(xMax)
   For x = 0 To xMax - 1
       xPos = x * 3
       For y = 0 To yMax
           CopyMemory rgb, abPicture(xPos, y), 3
'           lB = abPicture(x, y)
'           lG = abPicture(x + 1, y)
'           lR = abPicture(x + 2, y)
           'Gray color according ITU standard:
'           lGray = (222 * lR + 707 * lG + 71 * lB) / 1000
           lGray = (222# * rgb.rgbRed + 707# * rgb.rgbGreen + 71# * rgb.rgbBlue) / 1000
           CopyMemory abPicture(xPos, y), VBA.rgb(lGray, lGray, lGray), 3
'           abPicture(xPos, y) = lGray
'           abPicture(xPos + 1, y) = lGray
'           abPicture(xPos + 2, y) = lGray
       Next y
       If (x Mod 15) = 0 Then
          RaiseEvent ProgressChanged(x)
       End If
   Next x
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub Colourise(abPicture() As Byte, ByVal fHue As Single, ByVal fSaturation As Single)
   Dim x As Long, y As Long
   Dim xMax As Long, yMax As Long
   Dim lB As Long, lG As Long, lR As Long
   Dim H As Single, s As Single, L As Single
   Dim lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   RaiseEvent ProgressInit(xMax)
   For x = 0 To xMax Step 3
        For y = 0 To yMax
            RGBToHLS_OLD abPicture(x + 2, y), abPicture(x + 1, y), abPicture(x, y), H, s, L
            If H = 0 Then s = 0.5
            H = fHue
            HLSToRGB_OLD H, s, L, lR, lG, lB
            abPicture(x, y) = lB
            abPicture(x + 1, y) = lG
            abPicture(x + 2, y) = lR
        Next y
        If (x Mod 15) = 0 Then
           RaiseEvent ProgressChanged(x)
        End If
    Next x
    RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub ChangeBrightness(abPicture() As Byte, ByVal lAmount As Long, Optional clrIdx As Integer)
   Dim x As Long, y As Long
   Dim xMax As Long, yMax As Long
   Dim iVal As Integer, iStep As Integer
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   If clrIdx Then
      iStep = 3
      xStart = 3 - clrIdx
   Else
      iStep = 1
   End If
   For x = xStart To xMax Step iStep
       For y = 0 To yMax
           abPicture(x, y) = bTable(lAmount, abPicture(x, y))
       Next y
   Next x
End Sub

Public Sub ChangeContrast(abPicture() As Byte, ByVal lAmount As Long, Optional clrIdx As Integer)
   Dim x As Long, y As Long
   Dim xMax As Long, yMax As Long
   Dim iStep As Integer
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   If clrIdx Then
      iStep = 3
      xStart = 3 - clrIdx
   Else
      iStep = 1
   End If
   For x = xStart To xMax Step iStep
       For y = 0 To yMax
           abPicture(x, y) = cTable(lAmount, abPicture(x, y))
       Next y
   Next x
End Sub

Public Sub ChangeGamma(abPicture() As Byte, ByVal lAmount As Long, Optional clrIdx As Integer)
   Dim x As Long, y As Long
   Dim xMax As Long, yMax As Long
   Dim iStep As Integer
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   If clrIdx Then
      iStep = 3
      xStart = 3 - clrIdx
   Else
      iStep = 1
   End If
   For x = xStart To xMax Step iStep
       For y = 0 To yMax
           abPicture(x, y) = gTable(lAmount, abPicture(x, y))
       Next y
   Next x
End Sub

Public Sub BathRoom(abPicture() As Byte)
   Dim abBuffer() As Byte
   Dim x As Long, y As Long
   Dim xMax As Long, yMax As Long
   Dim clr As Long, clr2 As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   RaiseEvent ProgressInit(xMax)
   ReDim abBuffer(xMax, yMax)
   CopyMemory abBuffer(0, 0), abPicture(0, 0), (xMax + 1) * (yMax + 1)
   For x = 48 To xMax - 48 Step 12
       For y = 0 To yMax
           CopyMemory abBuffer(x, y), abPicture(x + (x Mod 96) - 48, y), 12
       Next y
       If (x Mod 45) = 0 Then
          RaiseEvent ProgressChanged(x)
       End If
   Next x
   CopyMemory abPicture(0, 0), abBuffer(0, 0), (xMax + 1) * (yMax + 1)
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub AddNoise(abPicture() As Byte, ByVal lPercent As Long)
   Dim x As Long, y As Long
   Dim xMax As Long, yMax As Long
   Dim bContinue As Boolean
   Dim lB As Long, lG As Long, lR As Long
   Dim lA As Long, lA2 As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   lA = 128 * lPercent \ 100
   lA2 = lA \ 2
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   RaiseEvent ProgressInit(xMax)
   For x = 0 To xMax Step 3
       For y = 0 To yMax
           bContinue = False
           If (Rnd * 100 > lPercent) Then
               bContinue = True
           End If
           If bContinue Then
              lB = abPicture(x, y) - lA2 + (Rnd * lA)
              lG = abPicture(x + 1, y) - lA2 + (Rnd * lA)
              lR = abPicture(x + 2, y) - lA2 + (Rnd * lA)
              ltob lR
              ltob lG
              ltob lB
              abPicture(x, y) = lB
              abPicture(x + 1, y) = lG
              abPicture(x + 2, y) = lR
            End If
        Next y
        If (x Mod 45) = 0 Then
           RaiseEvent ProgressChanged(x)
        End If
    Next x
    RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub Fade(abPicture() As Byte, ByVal lAmount As Long)
   Dim x As Long, y As Long, val As Long
   Dim xMax As Long, yMax As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   RaiseEvent ProgressInit(xMax)
   For x = 0 To xMax
       For y = 0 To yMax
           val = (255 - lAmount) * abPicture(x, y) \ 255
           If val > 255 Then val = 255
           abPicture(x, y) = val
       Next y
       If (x Mod 45) = 0 Then
          RaiseEvent ProgressChanged(x)
       End If
   Next x
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub Caricature(abPicture() As Byte, Optional ByVal rMax As Long, Optional TranspColor As Long)
   Dim abBuffer() As Byte
   Dim x As Long, y As Long, dx As Long, dy As Long
   Dim xMax As Long, yMax As Long, xPos As Long
   Dim Radius As Single, Angle As Single, rNew As Single
   Dim lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   ReDim abBuffer(xMax, yMax)
   xMax = xMax \ 3
   RaiseEvent ProgressInit(xMax)
   If TranspColor < 0 Then TranspColor = TranslateColor(TranspColor)
   If rMax = 0 Then rMax = Sqr(xMax * xMax + yMax * yMax) / 2
   For x = 0 To xMax
       xPos = x * 3
       For y = 0 To yMax
           dx = xMax / 2 - x
           dy = yMax / 2 - y
           Radius = Sqr(dx * dx + dy * dy)
           If Radius < rMax Then
              Angle = Atan2(dy, dx)
              rNew = Sqr(Radius * rMax)
              dx = xMax / 2 + rNew * Cos(Angle)
              dy = yMax / 2 - rNew * Sin(Angle)
              If dx < 0 Then dx = 0
              If dy < 0 Then dy = 0
              If dx > xMax Then dx = xMax
              If dy > yMax Then dy = yMax
              CopyMemory abBuffer(xPos, y), abPicture(dx * 3, dy), 3
           Else
              CopyMemory abBuffer(xPos, y), TranspColor, 3
           End If
       Next y
       If (x Mod 15) = 0 Then
          RaiseEvent ProgressChanged(x)
       End If
   Next x
   CopyMemory abPicture(0, 0), abBuffer(0, 0), (UBound(abBuffer, 1) + 1) * (yMax + 1)
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub FishEye(abPicture() As Byte, Optional ByVal rMax As Long, Optional TranspColor As Long)
   Dim abBuffer() As Byte
   Dim x As Long, y As Long, dx As Long, dy As Long
   Dim xMax As Long, yMax As Long, xPos As Long
   Dim Radius As Single, Angle As Single, rNew As Single
   Dim lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   ReDim abBuffer(xMax, yMax)
   xMax = xMax \ 3
   RaiseEvent ProgressInit(xMax)
   If TranspColor < 0 Then TranspColor = TranslateColor(TranspColor)
   If rMax = 0 Then rMax = Sqr(xMax * xMax + yMax * yMax) / 2
   For x = 0 To xMax
       xPos = x * 3
       For y = 0 To yMax
           dx = xMax / 2 - x
           dy = yMax / 2 - y
           Radius = Sqr(dx * dx + dy * dy)
           If Radius < rMax Then
              Angle = Atan2(dy, dx)
              rNew = Radius * Radius / rMax
              dx = xMax / 2 + rNew * Cos(Angle)
              dy = yMax / 2 - rNew * Sin(Angle)
              If dx < 0 Then dx = 0
              If dy < 0 Then dy = 0
              If dx > xMax Then dx = xMax
              If dy > yMax Then dy = yMax
              CopyMemory abBuffer(xPos, y), abPicture(dx * 3, dy), 3
           Else
              CopyMemory abBuffer(xPos, y), TranspColor, 3
           End If
       Next y
       If (x Mod 15) = 0 Then
          RaiseEvent ProgressChanged(x)
       End If
   Next x
   CopyMemory abPicture(0, 0), abBuffer(0, 0), (UBound(abBuffer, 1) + 1) * (yMax + 1)
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub Swirle(abPicture() As Byte, Optional ByVal Koeff As Single, Optional ByVal rMax As Single, Optional ByVal TranspColor As Long)
   Dim abBuffer() As Byte
   Dim x As Long, y As Long, dx As Long, dy As Long
   Dim xMax As Long, yMax As Long, xPos As Long
   Dim Radius As Single, Angle As Single, AngleNew As Single
   Dim lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   ReDim abBuffer(xMax, yMax)
   xMax = xMax \ 3
   RaiseEvent ProgressInit(xMax)
   If TranspColor < 0 Then TranspColor = TranslateColor(TranspColor)
   If rMax = 0 Then rMax = Sqr(xMax * xMax + yMax * yMax) / 2
   If Koeff = 0 Then Koeff = 100
   For x = 0 To xMax
       xPos = x * 3
       For y = 0 To yMax
           dx = xMax / 2 - x
           dy = yMax / 2 - y
           Radius = Sqr(dx * dx + dy * dy)
           If Radius < rMax Then
              Angle = Atan2(dy, dx)
              AngleNew = Angle + Radius / Koeff
              dx = xMax / 2 + Radius * Cos(AngleNew)
              dy = yMax / 2 - Radius * Sin(AngleNew)
              If dx < 0 Then dx = 0
              If dy < 0 Then dy = 0
              If dx > xMax Then dx = xMax
              If dy > yMax Then dy = yMax
              CopyMemory abBuffer(xPos, y), abPicture(dx * 3, dy), 3
           Else
              CopyMemory abBuffer(xPos, y), TranspColor, 3
           End If
       Next y
       If (x Mod 15) = 0 Then
          RaiseEvent ProgressChanged(x)
       End If
   Next x
   CopyMemory abPicture(0, 0), abBuffer(0, 0), (UBound(abBuffer, 1) + 1) * (yMax + 1)
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub Flip(abPicture() As Byte, Optional Direction As FlipDirection = eVert)
   Dim abBuffer() As Byte
   Dim x As Long, y As Long
   Dim xMax As Long, yMax As Long, newVal As Long
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   ReDim abBuffer(xMax, yMax)
   CopyMemory abBuffer(0, 0), abPicture(0, 0), (xMax + 1) * (yMax + 1)
   If Direction = eVert Then
      For y = 0 To yMax - 1
          CopyMemory abBuffer(0, yMax - y - 1), abPicture(0, y), xMax + 1
      Next y
   Else
      xMax = (Int(xMax / 3)) * 3
      For x = 0 To xMax - 3 Step 3
          newVal = xMax - x
          For y = 0 To yMax
              CopyMemory abBuffer(newVal - 3, y), abPicture(x, y), 3
          Next y
      Next x
   End If
   CopyMemory abPicture(0, 0), abBuffer(0, 0), (xMax + 1) * (yMax + 1)
End Sub

Public Sub Melt(abPicture() As Byte)
   Dim x As Long, y As Long
   Dim xMax As Long, yMax As Long
   Dim clr As Long, clrNext As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   RaiseEvent ProgressInit(xMax)
   For x = 0 To xMax Step 3
       For y = 0 To yMax - 1 Step 2
           CopyMemory clr, abPicture(x, y), 3
           CopyMemory clrNext, abPicture(x, y + 1), 3
           If clrNext > clr Then
              CopyMemory abPicture(x, y), clrNext, 3
              CopyMemory abPicture(x, y + 1), clr, 3
           End If
       Next y
       If (x Mod 45) = 0 Then
          RaiseEvent ProgressChanged(x)
       End If
   Next x
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub Negative(abPicture() As Byte)
   Dim x As Long, y As Long
   Dim xMax As Long, yMax As Long
   Dim clr As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   RaiseEvent ProgressInit(xMax)
   For x = 0 To xMax - 3 Step 3
       For y = 0 To yMax
           CopyMemory clr, abPicture(x, y), 3
           clr = &HFFFFFF - clr
           CopyMemory abPicture(x, y), clr, 3
       Next y
       If (x Mod 45) = 0 Then
          RaiseEvent ProgressChanged(x)
       End If
   Next x
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub Pixelize(abPicture() As Byte, ByVal nSize As Long)
   Dim x As Long, y As Long
   Dim i As Long, j As Long
   Dim xMax As Long, yMax As Long, xPos As Long, iPos As Long
   Dim rgb As RGBTRIPPLE, averRGB As RGBTRIPPLE
   Dim R As Long, G As Long, B As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1) \ 3 - nSize
   yMax = UBound(abPicture, 2) - nSize
   If nSize = 0 Then nSize = 3
   RaiseEvent ProgressInit(xMax)
   For x = 0 To xMax Step nSize
       xPos = x * 3
       For y = 0 To yMax Step nSize
           For i = 0 To nSize - 1
               iPos = i * 3
               R = 0: G = 0: B = 0
               For j = 0 To nSize - 1
                   CopyMemory rgb, abPicture(xPos + iPos, y + j), 3
                   R = R + rgb.rgbRed
                   G = G + rgb.rgbGreen
                   B = B + rgb.rgbBlue
               Next j
           Next i
           averRGB.rgbRed = R / nSize
           averRGB.rgbGreen = G / nSize
           averRGB.rgbBlue = B / nSize
           For i = 0 To nSize - 1
               iPos = i * 3
               For j = 0 To nSize - 1
                   CopyMemory abPicture(xPos + iPos, y + j), averRGB, 3
               Next j
           Next i
       Next y
       If (x Mod 45) = 0 Then
          RaiseEvent ProgressChanged(x)
       End If
   Next x
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub ReduceColors(abPicture() As Byte, Optional ByVal DiffuseError As Boolean = True)
' Based on Steve McMahon "clrdepth" sample
    Dim abBuffer() As Byte
    Dim x As Long, y As Long
    Dim xMax As Long, yMax As Long
    Dim i As Long, iCoeff As Long, j As Long
    Dim lTime As Long
    Dim lErrorRed As Long, lErrorBlue As Long, lErrorGreen As Long
    Dim lNewRed As Long, lNewBlue As Long, lNewGreen As Long
    Dim lIndex As Long

    lTime = GetTickCount
    xMax = UBound(abPicture, 1)
    yMax = UBound(abPicture, 2)
    ReDim abBuffer(xMax, yMax)
    
    RaiseEvent ProgressInit(xMax)
    For x = 0 To xMax - 3 Step 3
        For y = 0 To yMax
            ' Get nearest colour:
            
            lIndex = GetClosestIndex(abPicture(x + 2, y), abPicture(x + 1, y), abPicture(x, y))
                                 
            abBuffer(x + 2, y) = SysPal(lIndex).rgbRed
            abBuffer(x + 1, y) = SysPal(lIndex).rgbGreen
            abBuffer(x, y) = SysPal(lIndex).rgbBlue
            If DiffuseError Then
               lErrorRed = abPicture(x + 2, y) - CLng(abBuffer(x + 2, y))
               lErrorGreen = abPicture(x + 1, y) - CLng(abBuffer(x + 1, y))
               lErrorBlue = abPicture(x, y) - CLng(abBuffer(x, y))
               
               ' Diffuse the error:
               If Abs(lErrorRed) + Abs(lErrorGreen) + Abs(lErrorBlue) > 3 Then
                  If (x < xMax - 6) And (y < yMax) Then
                      lNewBlue = abPicture(x + 3, y) + (lErrorBlue * 7) \ 16
                      lNewGreen = abPicture(x + 4, y) + (lErrorGreen * 7) \ 16
                      lNewRed = abPicture(x + 5, y) + (lErrorRed * 7) \ 16
                      ltob lNewBlue
                      ltob lNewGreen
                      ltob lNewRed
                      abPicture(x + 3, y) = lNewBlue
                      abPicture(x + 4, y) = lNewGreen
                      abPicture(x + 5, y) = lNewRed
                      For i = -3 To 3 Step 3
                          If (x + i) > 0 And (x + i) < xMax - 3 Then
                             Select Case i
                                 Case -3: iCoeff = 2
                                 Case 0:  iCoeff = 4
                                 Case 3:  iCoeff = 1
                             End Select
                             lNewBlue = abPicture(x + i, y + 1) + (lErrorBlue * iCoeff) \ 16
                             lNewGreen = abPicture(x + i + 1, y + 1) + (lErrorGreen * iCoeff) \ 16
                             lNewRed = abPicture(x + i + 2, y + 1) + (lErrorRed * iCoeff) \ 16
                             ltob lNewBlue
                             ltob lNewGreen
                             ltob lNewRed
                             abPicture(x + i, y + 1) = lNewBlue
                             abPicture(x + i + 1, y + 1) = lNewGreen
                             abPicture(x + i + 2, y + 1) = lNewRed
                          End If
                      Next i
                  End If
               End If
            End If
        Next y
        RaiseEvent ProgressChanged(x)
    Next x
    
    RaiseEvent ProgressEnd(GetTickCount - lTime)
    CopyMemory abPicture(0, 0), abBuffer(0, 0), (xMax + 1) * (yMax + 1)
End Sub

Public Sub Relief(abPicture() As Byte)
   Dim x As Long, y As Long
   Dim xMax As Long, yMax As Long
   Dim clr As Long, clr2 As Long
   Dim lTime As Long
   
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   RaiseEvent ProgressInit(xMax)
   For x = 0 To xMax - 6 Step 3
       For y = 0 To yMax - 2
           CopyMemory clr, abPicture(x, y), 3
           CopyMemory clr2, abPicture(x + 6, y + 2), 3
           clr = clr - clr2 + &H808080
           If clr < 0 Then clr = 0
           If clr > &HFFFFFF Then clr = &HFFFFFF
           CopyMemory abPicture(x, y), clr, 3
       Next y
       If (x Mod 45) = 0 Then
          RaiseEvent ProgressChanged(x)
       End If
   Next x
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub Resample(W As Long, H As Long, abFrom() As Byte)
'Based on Steve McMahon "ImgProcess" sample
   Dim abTo() As Byte
   
   Dim xScale As Single, yScale As Single
   Dim x As Long, y As Long, xEnd As Long, xOut As Long
   Dim fX As Single, fY As Single
   Dim ifY As Long, ifX As Long
   Dim dx As Single, dy As Single

   Dim DR1 As DOUBLERGB, DR2 As DOUBLERGB
   Dim R As Long, G As Long, B As Long
   
   Dim ir1 As Long, ig1 As Long, ib1 As Long
   Dim ir2 As Long, ig2 As Long, ib2 As Long
   Dim lTime As Long
      
   lTime = GetTickCount
   RaiseEvent ProgressInit(H)
   ReDim abTo(BytesPerScanLine(W) - 1, H - 1)
   xScale = UBound(abFrom) / 3 / W
   yScale = UBound(abFrom, 2) / H
   xEnd = W - 1
   For y = 0 To H - 1
       fY = y * yScale
       ifY = Int(fY)
       dy = fY - ifY
       For x = 0 To xEnd
           fX = x * xScale
           ifX = Int(fX)
           dx = fX - ifX
           ifX = ifX * 3
           ' Interpolate using the four nearest pixels in the source
           CopyMemory DR1, abFrom(ifX, ifY), 6
           CopyMemory DR2, abFrom(ifX, ifY + 1), 6
           ' Interplate in x direction:
           ir1 = DR1.r1 * (1 - dy) + DR2.r1 * dy: ig1 = DR1.g1 * (1 - dy) + DR2.g1 * dy: ib1 = DR1.b1 * (1 - dy) + DR2.b1 * dy
           ir2 = DR1.r2 * (1 - dy) + DR2.r2 * dy: ig2 = DR1.g2 * (1 - dy) + DR2.g2 * dy: ib2 = DR1.b2 * (1 - dy) + DR2.b2 * dy
           ' Interpolate in y:
           R = ir1 * (1 - dx) + ir2 * dx: G = ig1 * (1 - dx) + ig2 * dx: B = ib1 * (1 - dx) + ib2 * dx
           ' Set output:
           ltob R: ltob G: ltob B
           xOut = x * 3
           abTo(xOut, y) = B
           abTo(xOut + 1, y) = G
           abTo(xOut + 2, y) = R
       Next x
       If (y Mod 15) = 0 Then
          RaiseEvent ProgressChanged(y)
       End If
   Next y
   ReDim abFrom(BytesPerScanLine(W) - 1, H - 1)
   CopyMemory abFrom(0, 0), abTo(0, 0), BytesPerScanLine(W) * H
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Public Sub Rotate(abPicture() As Byte, Optional Direction As RotateDirection = eLeft)
   Dim abBuffer() As Byte
   Dim x As Long, y As Long
   Dim W As Long, H As Long
   Dim i_Val As Long
   Dim lTime As Long
      
   lTime = GetTickCount
   W = UBound(abPicture, 1) / 3
   H = UBound(abPicture, 2) + 1
   ReDim abBuffer(BytesPerScanLine(H) - 1, W - 1)
   If Direction = eLeft Then
      RaiseEvent ProgressInit(H)
      For y = 0 To H - 1
          i_Val = (H - y - 1) * 3
          For x = 0 To W - 1
              CopyMemory abBuffer(i_Val, x), abPicture(x * 3, y), 3
          Next x
          If (x Mod 15) = 0 Then
             RaiseEvent ProgressChanged(y)
          End If
      Next y
   Else
     RaiseEvent ProgressInit(W)
      For x = 0 To W - 1
          i_Val = (W - x - 1) * 3
          For y = 0 To H - 1
              CopyMemory abBuffer(y * 3, x), abPicture(i_Val, y), 3
          Next y
          If (x Mod 15) = 0 Then
             RaiseEvent ProgressChanged(x)
          End If
      Next x
   End If
   ReDim abPicture(UBound(abBuffer, 1), UBound(abBuffer, 2))
   CopyMemory abPicture(0, 0), abBuffer(0, 0), (UBound(abBuffer, 1) + 1) * (UBound(abBuffer, 2) + 1)
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Private Sub InitBCTables()
'Create brightness and contrast tables.
'Assuming changes in -100 to 100 percents range
   Dim i As Long, j As Long, lVal As Long
   Dim gamma As Single
   For i = 0 To 100
       For j = 0 To 255
'*************brightness***************************
           lVal = j + i * 127 / 50 - 127
           ltob lVal
           bTable(i, j) = lVal
'*************contrast***************************
           lVal = j + ((j - 127) * (i - 50) / (101 - i))
           ltob lVal
           cTable(i, j) = lVal
'****************gamma***************************
           gamma = 10 ^ ((50 - i) / 50)
           lVal = 255 * (j / 255) ^ gamma
           ltob lVal
           gTable(i, j) = lVal
       Next j
   Next i
End Sub

Private Sub Class_Initialize()
   InitBCTables
End Sub
