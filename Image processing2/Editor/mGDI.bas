Attribute VB_Name = "mGDI"

Public Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Public Type RGBTRIPPLE
   rgbRed As Byte
   rgbGreen As Byte
   rgbBlue As Byte
End Type

Public Type RGBQUAD
     rgbBlue As Byte
     rgbGreen As Byte
     rgbRed As Byte
     rgbReserved As Byte
End Type

Public Type BITMAPINFOHEADER '40 bytes
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

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors() As RGBQUAD
End Type

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDc As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Public Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As Any, ByVal wUsage As Long) As Long
Public Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hDc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As Any, ByVal wUsage As Long) As Long

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Const LR_DEFAULTSIZE = &H40
Private Const LR_REALSIZE = &H80
Private Const LR_LOADMAP3DCOLORS = &H1000
Private Const LR_LOADTRANSPARENT = &H20
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Public Const BI_RGB = 0&
Public Const BI_RLE4 = 2&
Public Const BI_RLE8 = 1&
Public Const DIB_RGB_COLORS = 0

Public Function GetTrueBits(pb As PictureBox, abPicture() As Byte, BI As BITMAPINFO) As Boolean
   Dim bmp As BITMAP
   Call GetObjectAPI(pb.Picture, Len(bmp), bmp)
   ReDim BI.bmiColors(0)
   With BI.bmiHeader
       .biSize = Len(BI.bmiHeader)
       .biWidth = bmp.bmWidth
       .biHeight = bmp.bmHeight
       .biPlanes = 1
       .biBitCount = 24
       .biCompression = BI_RGB
       .biSizeImage = BytesPerScanLine(.biWidth) * .biHeight
       ReDim abPicture(BytesPerScanLine(.biWidth) - 1, .biHeight - 1)
   End With
   GetTrueBits = GetDIBits(pb.hDc, pb.Picture, 0, BI.bmiHeader.biHeight, abPicture(0, 0), BI, DIB_RGB_COLORS)
End Function

Public Function GetBits(pb As PictureBox, abPicture() As Byte, BI As BITMAPINFO, Optional clrDepth As Long) As Boolean
   Dim BuffSize As Long
   Dim biArray() As Byte
   Dim bih As BITMAPINFOHEADER
   
   ReDim BI.bmiColors(0)
   BI.bmiHeader.biSize = Len(BI.bmiHeader)
   Call GetDIBits(pb.hDc, pb.Picture, 0, 0, ByVal 0, BI.bmiHeader, DIB_RGB_COLORS)
   If clrDepth > 0 Then
      If clrDepth < BI.bmiHeader.biBitCount Then
         BI.bmiHeader.biBitCount = clrDepth
      End If
   End If
   BI.bmiHeader.biCompression = BI_RGB
   BuffSize = BI.bmiHeader.biWidth
   Select Case BI.bmiHeader.biBitCount
       Case 1
            BuffSize = Int((BuffSize + 7) / 8)
            ReDim biArray(Len(bih) + 4 * 2 - 1)
       Case 4
            BuffSize = Int((BuffSize + 1) / 2)
            ReDim biArray(Len(bih) + 4 * 16 - 1)
       Case 8
            BuffSize = BuffSize
            ReDim biArray(Len(bih) + 4 * 256 - 1)
       Case 16
            BuffSize = BuffSize * 2
            ReDim biArray(Len(bih) + 4 - 1)
       Case 24
            BuffSize = BuffSize * 3
            ReDim biArray(Len(bih) + 4 - 1)
       Case 32
            BuffSize = BuffSize * 3
            ReDim biArray(Len(bih) + 4 * 3 - 1)
   End Select
   ReDim BI.bmiColors((UBound(biArray) + 1 - Len(bih)) \ 4 - 1)
   BuffSize = (Int((BuffSize + 3) / 4)) * 4
   ReDim abPicture(BuffSize - 1, BI.bmiHeader.biHeight - 1)
   BuffSize = BuffSize * BI.bmiHeader.biHeight
   BI.bmiHeader.biSizeImage = BuffSize
   CopyMemory biArray(0), BI, Len(BI.bmiHeader)
   GetBits = GetDIBits(pb.hDc, pb.Picture, 0, BI.bmiHeader.biHeight, abPicture(0, 0), biArray(0), DIB_RGB_COLORS)
   CopyMemory BI, biArray(0), Len(bih)
   CopyMemory BI.bmiColors(0), biArray(Len(bih) + 1), UBound(biArray) - Len(bih) + 1
End Function

Public Function SetBits(pb As PictureBox, abPicture() As Byte, BI As BITMAPINFO) As Boolean
   SetBits = SetDIBitsToDevice(pb.hDc, 0, 0, BI.bmiHeader.biWidth, BI.bmiHeader.biHeight, 0, 0, 0, BI.bmiHeader.biHeight, abPicture(0, 0), BI, DIB_RGB_COLORS)
   pb.Refresh
End Function

Public Function BytesPerScanLine(ByVal lWidth As Long) As Long
    BytesPerScanLine = (lWidth * 3 + 3) And &HFFFFFFFC
End Function
' Convert Automation color to Windows color
Public Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function


