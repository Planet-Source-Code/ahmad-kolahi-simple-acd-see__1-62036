VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cHistogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Public Enum HistoType
   htRed
   htGreen
   htBlue
   htIntensity
   htValue
End Enum

Public Event ProgressInit(ByVal nMax As Long)
Public Event ProgressChanged(ByVal nValue As Long)
Public Event ProgressEnd(ByVal nTime As Long)

Private m_iBPP As eBPP

Public Property Let BPP(ByVal vData As eBPP)
    m_iBPP = vData
End Property

Public Property Get BPP() As eBPP
    BPP = m_iBPP
End Property

Public Sub StretchHisto(abPicture() As Byte, Optional delta As Long = 1, Optional StretchFactor As Single = 1, Optional htType As HistoType = htIntensity)
   Dim abHisto(0 To 255) As Long
   Dim xMin As Long, xMax As Long, picSize As Long
   Dim x As Long, y As Long, lVal As Long, n As Long
   Dim StretchedRange As Long
   Dim ScaleFactor As Single
   Dim min As Long, Max As Long
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   If StretchFactor > 1 Or StretchFactor < 0 Then StretchFactor = 1
   If BPP < TrueColor_24 Then
      n = GetHistoStatsPalette(abPicture, abHisto, htType)
   Else
     n = GetHistoStatsTrueColor(abPicture, abHisto, htType)
   End If
   min = GetHistoFreq(abHisto, delta, n)
   Max = GetHistoFreq(abHisto, 100 - delta, n)
   StretchedRange = Max - min + Int(StretchFactor * (255 - Max + min))
   If min >= Max Then
      ScaleFactor = 1
   Else
      ScaleFactor = StretchedRange / (Max - min)
   End If
   For x = 0 To xMax
       For y = 0 To yMax
           lVal = ScaleFactor * (abPicture(x, y) - min)
           ltob lVal
           abPicture(x, y) = lVal
       Next y
   Next x
End Sub

Public Sub EqualizeHisto(abPicture() As Byte, Optional htType As HistoType = htIntensity)
   Dim abHisto(0 To 255) As Long, abCDF(0 To 255) As Long
   Dim xMin As Long, xMax As Long, picSize As Long
   Dim x As Long, y As Long, lVal As Long, n As Long
   Dim lTime As Long
   lTime = GetTickCount
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   If BPP < TrueColor_24 Then
      n = GetHistoStatsPalette(abPicture, abHisto, htType)
   Else
      n = GetHistoStatsTrueColor(abPicture, abHisto, htType)
   End If
   RaiseEvent ProgressInit(xMax)
   abCDF(0) = abHisto(0)
   For x = 1 To 255
       abCDF(x) = abCDF(x - 1) + abHisto(x)
   Next x
   For x = 0 To xMax
       For y = 0 To yMax
           lVal = 255 * abCDF(abPicture(x, y)) / n
           ltob lVal
           abPicture(x, y) = lVal
       Next y
       If (x Mod 45) = 0 Then
          RaiseEvent ProgressChanged(x)
       End If
   Next x
   RaiseEvent ProgressEnd(GetTickCount - lTime)
End Sub

Private Function GetHistoStatsTrueColor(abPicture() As Byte, abHisto() As Long, Optional htType As HistoType = htIntensity) As Long
   Dim rgb() As RGBTRIPPLE
   Dim xMin As Long, xMax As Long, picSize As Long
   Dim lVal As Long, TotalCount As Long
   Dim i As Long
   xMax = UBound(abPicture, 1) \ 3
   yMax = UBound(abPicture, 2)
   picSize = (xMax + 1) * (yMax + 1)
   ReDim rgb(picSize - 1)
   CopyMemory rgb(0), abPicture(0, 0), picSize * 3
   For i = 0 To picSize - 1
       Select Case htType
              Case htRed:       lVal = rgb(i).rgbRed
              Case htGreen:     lVal = rgb(i).rgbGreen
              Case htBlue:      lVal = rgb(i).rgbBlue
              Case htIntensity: lVal = (CLng(rgb(i).rgbRed) + rgb(i).rgbGreen + rgb(i).rgbBlue) / 3
              Case htValue:     lVal = Maximum(rgb(i).rgbRed, rgb(i).rgbGreen, rgb(i).rgbBlue)
              Case Else:        lVal = 0
       End Select
       abHisto(lVal) = abHisto(lVal) + 1
       TotalCount = TotalCount + 1
   Next i
   GetHistoStatsTrueColor = TotalCount
End Function

Private Function GetHistoStatsPalette(abPicture() As Byte, abHisto() As Long, Optional htType As HistoType = htIntensity) As Long
   Dim xMin As Long, xMax As Long, picSize As Long
   Dim lVal As Long, TotalCount As Long
   Dim i As Long
   Dim clr() As Byte
   xMax = UBound(abPicture, 1)
   yMax = UBound(abPicture, 2)
   picSize = (xMax + 1) * (yMax + 1)
   ReDim clr(picSize - 1)
   CopyMemory clr(0), abPicture(0, 0), picSize
   For i = 0 To picSize - 1
       lVal = clr(i)
       abHisto(lVal) = abHisto(lVal) + 1
       TotalCount = TotalCount + 1
   Next i
   GetHistoStatsPalette = TotalCount
End Function

Private Function GetHistoFreq(abHisto() As Long, ByVal Percent As Long, ByVal Total As Long) As Long
   Dim i As Integer, nCount As Long
   Dim Target As Single, Current As Single
   Target = Percent / 100
   For i = 1 To 255
       nCount = nCount + abHisto(i)
       Current = nCount / Total
       If Target < Current Then
          GetHistoFreq = i: Exit For
       End If
   Next i
End Function

Private Sub Class_Initialize()
   BPP = eBPP.TrueColor_24
End Sub
