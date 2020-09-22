Attribute VB_Name = "mHLS_RGB"
Public Type HLSTRIPPLE
   H As Integer
   L As Integer
   s As Integer
End Type

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Public Declare Function ColorHLSToRGB Lib "SHLWAPI.DLL" (ByVal wHue As Integer, ByVal wLuminance As Integer, ByVal wSaturation As Integer) As Long
Public Declare Function ColorRGBToHLS Lib "SHLWAPI.DLL" (ByVal clrRGB As Long, pwHue As Integer, pwLuminance As Integer, pwSaturation As Integer) As Boolean
Public Declare Function ColorAdjustLuma Lib "SHLWAPI.DLL" (ByVal clrRGB As Long, ByVal n As Long, ByVal fScale As Long) As Long

Public Sub RGBToHLS_OLD(ByVal r As Long, ByVal g As Long, ByVal b As Long, H As Single, s As Single, L As Single)
   Dim max As Single, min As Single
   Dim delta As Single
   Dim rR As Single, rG As Single, rB As Single
   rR = r / 255: rG = g / 255: rB = b / 255
'{Given: rgb each in [0,1].
' Desired: h in [0,360] and s in [0,1], except if s=0, then h=UNDEFINED.}
        max = Maximum(rR, rG, rB)
        min = Minimum(rR, rG, rB)
        L = (max + min) / 2    '{This is the lightness}
        '{Next calculate saturation}
        If max = min Then
            'begin {Acrhomatic case}
            s = 0
            H = 0
           'end {Acrhomatic case}
        Else
           'begin {Chromatic case}
                '{First calculate the saturation.}
           If L <= 0.5 Then
               s = (max - min) / (max + min)
           Else
               s = (max - min) / (2 - max - min)
           End If
            '{Next calculate the hue.}
           delta = max - min
           If rR = max Then
                H = (rG - rB) / delta    '{Resulting color is between yellow and magenta}
           ElseIf rG = max Then
                H = 2 + (rB - rR) / delta '{Resulting color is between cyan and yellow}
           ElseIf rB = max Then
                H = 4 + (rR - rG) / delta '{Resulting color is between magenta and cyan}
           End If
      End If
'end {RGB_to_HLS}
End Sub

Public Sub HLSToRGB_OLD(ByVal H As Single, ByVal s As Single, ByVal L As Single, r As Long, g As Long, b As Long)
   Dim rR As Single, rG As Single, rB As Single
   Dim min As Single, max As Single

   If s = 0 Then
      ' Achromatic case:
      rR = L: rG = L: rB = L
   Else
      ' Chromatic case:
      ' delta = Max-Min
      If L <= 0.5 Then
         's = (Max - Min) / (Max + Min)
         ' Get Min value:
         min = L * (1 - s)
      Else
         's = (Max - Min) / (2 - Max - Min)
         ' Get Min value:
         min = L - s * (1 - L)
      End If
      ' Get the Max value:
      max = 2 * L - min
      
      ' Now depending on sector we can evaluate the h,l,s:
      If (H < 1) Then
         rR = max
         If (H < 0) Then
            rG = min
            rB = rG - H * (max - min)
         Else
            rB = min
            rG = H * (max - min) + rB
         End If
      ElseIf (H < 3) Then
         rG = max
         If (H < 2) Then
            rB = min
            rR = rB - (H - 2) * (max - min)
         Else
            rR = min
            rB = (H - 2) * (max - min) + rR
         End If
      Else
         rB = max
         If (H < 4) Then
            rR = min
            rG = rR - (H - 4) * (max - min)
         Else
            rG = min
            rR = (H - 4) * (max - min) + rG
         End If
         
      End If
            
   End If
   r = rR * 255: g = rG * 255: b = rB * 255
End Sub

Public Function Maximum(ByVal rR As Single, ByVal rG As Single, ByVal rB As Single) As Single
   If (rR > rG) Then
      If (rR > rB) Then Maximum = rR Else Maximum = rB
   Else
      If (rB > rG) Then Maximum = rB Else Maximum = rG
   End If
End Function

Public Function Minimum(ByVal rR As Single, ByVal rG As Single, ByVal rB As Single) As Single
   If (rR < rG) Then
      If (rR < rB) Then Minimum = rR Else Minimum = rB
   Else
      If (rB < rG) Then Minimum = rB Else Minimum = rG
   End If
End Function

Public Function IsAPIColorSupported() As Boolean
   Dim hLib As Long, lAddress As Long
   hLib = LoadLibrary("SHLWAPI.DLL")
   If hLib Then
      IsAPIColorSupported = GetProcAddress(hLib, "ColorAdjustLuma")
      FreeLibrary hLib
   End If
End Function

