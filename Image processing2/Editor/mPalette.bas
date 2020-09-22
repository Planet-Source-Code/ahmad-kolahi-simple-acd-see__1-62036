Attribute VB_Name = "mPalette"
Private Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
End Type

Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hdc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long

Public SysPal(1 To 256) As RGBTRIPPLE

Public Sub InitPalette(ByVal hdc As Long)
   Static bDone As Boolean
   If bDone Then Exit Sub
   Dim ret As Long, i As Long
   Dim pe(255) As PALETTEENTRY
   ret = GetSystemPaletteEntries(hdc, 0, 255, pe(0))
   If ret = 0 Then CreateHalfTone
   For i = 0 To ret
       SysPal(i + 1).rgbRed = pe(i).peRed
       SysPal(i + 1).rgbGreen = pe(i).peGreen
       SysPal(i + 1).rgbBlue = pe(i).peBlue
   Next i
   bDone = True
End Sub

Private Sub CreateHalfTone()
   Dim lIndex As Long
   Dim R As Long, G As Long, B As Long
   Dim rA As Long, gA As Long, bA As Long
   
   ' Halftone 256 colour palette
   For B = 0 To &H100 Step &H40
      If B = &H100 Then bA = B - 1 Else bA = B
      For G = 0 To &H100 Step &H40
         If G = &H100 Then gA = G - 1 Else gA = G
         For R = 0 To &H100 Step &H40
            If R = &H100 Then rA = R - 1 Else rA = R
            lIndex = lIndex + 1
            With SysPal(lIndex)
               .rgbRed = rA: .rgbGreen = gA: .rgbBlue = bA
            End With
         Next R
      Next G
   Next B
End Sub

Public Function GetClosestIndex(ByVal R As Long, ByVal G As Long, ByVal B As Long) As Long
   Dim i As Long
   Dim lMinErr As Long, lCurErr As Long
   lMinErr = 765
   GetClosestIndex = 1
   For i = 1 To 256
      With SysPal(i)
         If (R = .rgbRed) And (B = .rgbBlue) And (G = .rgbGreen) Then
            GetClosestIndex = i
            Exit Function
         Else
            lCurErr = Abs(R - .rgbRed) + Abs(G - .rgbGreen) + Abs(B - .rgbBlue)
            If (lCurErr < lMinErr) Then
               lMinErr = lCurErr
               GetClosestIndex = i
            End If
         End If
      End With
   Next i
End Function

