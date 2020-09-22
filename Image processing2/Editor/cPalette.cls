VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPalette"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type

Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long

Private m_tPal() As RGBTRIPPLE
Private m_iPalette As Long
Dim pe(255) As PALETTEENTRY

Dim Palette_256(1 To 256) As RGBTRIPPLE
Dim Palette_System(1 To 256) As RGBTRIPPLE
Dim Palette_WebSafe(1 To 256) As RGBTRIPPLE
Dim Palette_16(1 To 16) As RGBTRIPPLE
Dim Palette_Mono(1 To 2) As RGBTRIPPLE

Public Property Get ClosestIndex(ByVal R As Long, ByVal G As Long, ByVal B As Long) As Long
   Dim i As Long
   Dim lER As Long, lEB As Long, lEG As Long
   Dim lMinER As Long, lMinEB As Long, lMinEG As Long
   
   lMinER = 255: lMinEB = 255: lMinEG = 255
   For i = 1 To m_iPalette
      With m_tPal(i)
         If (R = .rgbRed) And (B = .rgbBlue) And (G = .rgbGreen) Then
            ClosestIndex = i
            Exit Property
         Else
            lER = Abs(R - .rgbRed)
            lEG = Abs(G - .rgbGreen)
            lEB = Abs(B - .rgbBlue)
            If (lER + lEB + lEG < lMinER + lMinEB + lMinEG) Then
               lMinER = lER: lMinEG = lEG: lMinEB = lEB
               ClosestIndex = i
            End If
         End If
      End With
   Next i
End Property

Public Property Get Count() As Long
' Number of colours in this palette:
   Count = m_iPalette
End Property

Public Property Get Red(ByVal iIndex) As Byte
' Red value for palette index
   Red = m_tPal(iIndex).rgbRed
End Property

Public Property Get Green(ByVal iIndex) As Byte
' Green value for palette index
   Green = m_tPal(iIndex).rgbGreen
End Property

Public Property Get Blue(ByVal iIndex) As Byte
' Blue value for palette index
   Blue = m_tPal(iIndex).rgbBlue
End Property

Public Sub CreateWebSafe()
   Dim lIndex As Long
   Dim R As Long, G As Long, B As Long
   Dim L As Long, i As Long
   Static bDone As Boolean

   ' Websafe (IE 216 colour) palette
   m_iPalette = 256
   ReDim m_tPal(1 To 256)
   If bDone Then GoTo AlreadyDone
   p16ColourLow8 1
   p16ColourHigh8 248
   lIndex = 8
   For B = 0 To &HFF Step &H33
      For G = 0 To &HFF Step &H33
         For R = 0 To &HFF Step &H33
            ' ignore if the output is any combination of 0 and FF
            L = R + G + B
            If L = 0 Or L = &H2FD Then
               ' ignore
            ElseIf L = &H1FE And (R = 0 Or G = 0 Or B = 0) Then
               ' ignore
            ElseIf L = &HFF And ((R = 0 And G = 0) Or (R = 0 And B = 0) Or (G = 0 And B = 0)) Then
               ' ignore
            Else
               ' add
               lIndex = lIndex + 1
               With m_tPal(lIndex)
                  .rgbRed = R: .rgbGreen = G: .rgbBlue = B
               End With
            End If
         Next
      Next
   Next
   ' Fill the remain entries with gray shades:
   R = 8: G = 8: B = 8
   For i = 217 To 247
      With m_tPal(lIndex)
         .rgbRed = R: .rgbGreen = G: .rgbBlue = B
         R = R + 8: G = G + 8: B = B + 8
      End With
   Next i
   CopyMemory Palette_WebSafe(1), m_tPal(1), UBound(Palette_WebSafe) * Len(Palette_WebSafe(0))
   bDone = True
AlreadyDone:
   CopyMemory m_tPal(1), Palette_WebSafe(1), UBound(Palette_WebSafe) * Len(Palette_WebSafe(0))
End Sub

Public Sub CreateHalfTone()
   Dim lIndex As Long
   Dim R As Long, G As Long, B As Long
   Dim rA As Long, gA As Long, bA As Long
   Dim L As Long, i As Long
   Static bDone As Boolean
   
   ' Halftone 256 colour palette
   m_iPalette = 256
   ReDim m_tPal(1 To 256)
   If bDone Then GoTo AlreadyDone
   For B = 0 To &H100 Step &H40
      If B = &H100 Then bA = B - 1 Else bA = B
      For G = 0 To &H100 Step &H40
         If G = &H100 Then gA = G - 1 Else gA = G
         For R = 0 To &H100 Step &H40
            If R = &H100 Then rA = R - 1 Else rA = R
            lIndex = lIndex + 1
            With m_tPal(lIndex)
               .rgbRed = rA: .rgbGreen = gA: .rgbBlue = bA
            End With
         Next R
      Next G
   Next B
   CopyMemory Palette_256(1), m_tPal(1), UBound(Palette_256) * Len(Palette_256(0))
   bDone = True
AlreadyDone:
   CopyMemory m_tPal(1), Palette_256(1), UBound(Palette_256) * Len(Palette_256(0))
End Sub

Public Sub CreateDefault(ByVal hdc As Long)
   Static bDone As Boolean
   Dim i As Long
   m_iPalette = 256
   ReDim m_tPal(1 To 256)
   If bDone Then GoTo AlreadyDone
   GetSystemPaletteEntries hdc, 0, 255, pe(0)
   For i = 0 To 255
       Palette_System(i + 1).rgbRed = pe(i).peRed
       Palette_System(i + 1).rgbGreen = pe(i).peGreen
       Palette_System(i + 1).rgbBlue = pe(i).peBlue
   Next i
   bDone = True
AlreadyDone:
   CopyMemory m_tPal(1), Palette_System(1), UBound(Palette_System) * Len(Palette_System(0))
End Sub

Public Sub CreateMono()
   ' Monochrome palette
   m_iPalette = 2
   ReDim m_tPal(1 To 2)
   With m_tPal(2)
      .rgbBlue = 255: .rgbGreen = 255: .rgbRed = 255
   End With
End Sub

Private Sub p16ColourLow8(ByVal lStartIndex As Long)
   lStartIndex = lStartIndex - 1
   With m_tPal(lStartIndex + 2)
      .rgbRed = &H80: .rgbGreen = 0: .rgbBlue = 0
   End With
   With m_tPal(lStartIndex + 3)
      .rgbRed = 0: .rgbGreen = &H80: .rgbBlue = 0
   End With
   With m_tPal(lStartIndex + 4)
      .rgbRed = &H80: .rgbGreen = &H80: .rgbBlue = 0
   End With
   With m_tPal(lStartIndex + 5)
      .rgbRed = 0: .rgbGreen = 0: .rgbBlue = &H80
   End With
   With m_tPal(lStartIndex + 6)
      .rgbRed = &H80: .rgbGreen = 0: .rgbBlue = &H80
   End With
   With m_tPal(lStartIndex + 7)
      .rgbRed = 0: .rgbGreen = &H80: .rgbBlue = &H80
   End With
   With m_tPal(lStartIndex + 8)
      .rgbRed = &HC0: .rgbGreen = &HC0: .rgbBlue = &HC0
   End With
End Sub

Private Sub p16ColourHigh8(ByVal lStartIndex As Long)
   lStartIndex = lStartIndex - 9
   With m_tPal(lStartIndex + 9)
      .rgbRed = &H80: .rgbGreen = &H80: .rgbBlue = &H80
   End With
   With m_tPal(lStartIndex + 10)
      .rgbRed = &HFF: .rgbGreen = 0: .rgbBlue = 0
   End With
   With m_tPal(lStartIndex + 11)
      .rgbRed = 0: .rgbGreen = &HFF: .rgbBlue = 0
   End With
   With m_tPal(lStartIndex + 12)
      .rgbRed = &HFF: .rgbGreen = &HFF: .rgbBlue = 0
   End With
   With m_tPal(lStartIndex + 13)
      .rgbRed = 0: .rgbGreen = 0: .rgbBlue = &HFF
   End With
   With m_tPal(lStartIndex + 14)
      .rgbRed = &HFF: .rgbGreen = 0: .rgbBlue = &HFF
   End With
   With m_tPal(lStartIndex + 15)
      .rgbRed = 0: .rgbGreen = &HFF: .rgbBlue = &HFF
   End With
   With m_tPal(lStartIndex + 16)
      .rgbRed = &HFF: .rgbGreen = &HFF: .rgbBlue = &HFF
   End With
End Sub

Public Sub Create16Colour()
   ' Standard EGA style 16 colour palette:
   Static bDone As Boolean
   m_iPalette = 16
   ReDim m_tPal(1 To 16)
   If bDone Then GoTo AlreadyDone
   p16ColourLow8 1
   p16ColourHigh8 9
   CopyMemory Palette_16(1), m_tPal(1), UBound(Palette_16) * Len(Palette_16(0))
   bDone = True
AlreadyDone:
   CopyMemory m_tPal(1), Palette_16(1), UBound(Palette_16) * Len(Palette_16(0))
End Sub

