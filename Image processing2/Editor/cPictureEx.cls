VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPictureEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum FILTERTYPES
   eBlur
   eSoften
   eSharp
   eEdge
   eRankMedian
   eRankMin
   eRankMax
   eEnhDetails
   eEnhFocus
   eEnhEdges
   eCustom
End Enum

Public Enum PICTURE_FORMATS
   eBMP = 1
   eJPEG = 2
   eGIF = 3
End Enum

Public Enum BalanceType
    eChangeBrightness
    eChangeContrast
    eChangeGamma
End Enum

Public Enum ImgProcessTypes
    eAddNoise
    eAutoBalance
    eBathRoom
    eCaricature
    eEmBoss
    eFade
    eFishEye
    eSwirle
    eMelt
    eNegative
    ePixelize
    eRelief
    eFlip
    eRotate
    eBlackWhite
    eGrayScale
    eDither256
    eColourise
End Enum

Dim WithEvents Histo As cHistogram
Attribute Histo.VB_VarHelpID = -1
Dim WithEvents FFT As cFFT
Attribute FFT.VB_VarHelpID = -1
Dim WithEvents Filter As cFilter
Attribute Filter.VB_VarHelpID = -1
Dim WithEvents ImgProcess As cImgProcess
Attribute ImgProcess.VB_VarHelpID = -1
Dim MSF As MSGrphFilter

Public Event ProgressInit(ByVal nMax As Long)
Public Event ProgressChanged(ByVal nValue As Long)
Public Event ProgressEnd(ByVal nTime As Long)

Private m_abPicture() As Byte         '2D Picture bytes array
Private m_BI As BITMAPINFO            'BITMAPINFO
Private m_colPrevPic As Collection    'Collection to store prev pictures
Private m_colOperations As Collection 'Collection to store prev operations
Private m_LastOp As String            'Last operation name
Private m_UseResampling As Boolean
Private m_BoundedPB As PictureBox

Private m_bFromCode As Boolean
Private m_FilePath As String
Private m_FileType As PICTURE_FORMATS
Private m_bUseSelectedArea As Boolean
Private m_SelectedArea As RECT

Public Property Get MSF_ImportSupported() As Boolean
   MSF_ImportSupported = MSF.IsImportSupported
End Property

Public Property Get MSF_ExportSupported() As Boolean
   MSF_ExportSupported = MSF.IsExportSupported
End Property

Public Sub BoundPictureBox(pb As PictureBox)
   Set m_BoundedPB = pb
End Sub

Public Property Set Picture(ByVal vData As StdPicture)
   Erase m_abPicture
   Set m_BoundedPB.Picture = vData
   If vData = 0 Then Exit Property
   Call GetTrueBits(m_BoundedPB, m_abPicture, m_BI) ', m_SelectedArea)
   If m_bFromCode Then Exit Property
   m_LastOp = "Picture change"
   UpdateProps
'   If m_SelectedArea.Right = 0 Then SetSelectedArea 0, 0, 0, 0
End Property

Public Property Get Picture() As StdPicture
   Set Picture = m_BoundedPB.Picture
End Property

Public Property Get EnableUndo() As Boolean
   EnableUndo = (m_colPrevPic.Count > 1)
End Property

Public Property Get Operations() As Collection
   Set Operations = m_colOperations
End Property

Public Property Let UseResampling(ByVal vData As Boolean)
   m_UseResampling = vData
End Property

Public Property Get UseResampling() As Boolean
   UseResampling = m_UseResampling
End Property

'*************************** Loading and saving ******************************
Public Function LoadFromFile(Optional sPath As String, Optional bPromptUser As Boolean = True, Optional fTemplate As Form) As Long
'Return 0 on success, -1 if user cancel loading, error number otherwise
   LoadFromFile = -1
   On Error GoTo ErrHandler
   'If sPath = "" Then sPath = m_FilePath
   'If sPath = "" Or bPromptUser Then
   '   sPath = GetFileName(sPath, "Graphic files (*.bmp,*.gif,*.jpg,*.pcx)|*.bmp;*.gif;*.jpg;*.pcx", 0, "Open graphical file", True, fTemplate)
   '   If sPath = "" Then Exit Function
   'End If
   sPath = "test1.jpg"
   Set Picture = LoadPicture(sPath)
   m_FilePath = Dir(sPath)
   m_FileType = GetFileType(sPath)
ErrHandler:
   LoadFromFile = Err.Number
End Function

Public Function LoadFromFileMSF(Optional sPath As String, Optional bPromptUser As Boolean = True, Optional fTemplate As Form) As Long
   Dim s As String
   LoadFromFileMSF = -1
   On Error GoTo ErrHandler
   If sPath = "" Or bPromptUser Then
      sPath = GetFileName(sPath, MSF.ImportFilterStr, 0, "Open graphical file", True, fTemplate)
      If sPath = "" Then Exit Function
   End If
   MSF.MsoLoadPicture sPath, m_BoundedPB
   Set Picture = m_BoundedPB.Image
   m_FilePath = Dir(sPath)
   m_FileType = GetFileType(sPath)
ErrHandler:
   LoadFromFileMSF = Err.Number
End Function

Public Function SaveToFile(Optional sPath As String, Optional bPromptUser As Boolean = True, Optional fTemplate As Form) As Long
'Return 0 on success, -1 if user cancel loading, error number otherwise
   SaveToFile = -1
   On Error GoTo ErrHandler
   'If sPath = "" Then sPath = m_FilePath
   'If sPath = "" Or bPromptUser Then
   '   sPath = GetFileName(sPath, "Windows bitmap (*.bmp)|*.bmp", 0, "Save graphical file", False)
   '   If sPath = "" Then Exit Function
   'End If
   sPath = "test2.jpg"
   SavePicture m_BoundedPB.Picture, sPath
   m_FilePath = Dir(sPath)
   m_FileType = eBMP
ErrHandler:
   SaveToFile = Err.Number
End Function

Public Function SaveToFileMSF(Optional sPath As String, Optional bPromptUser As Boolean = True, Optional fTemplate As Form) As Long
'Return 0 on success, -1 if user cancel loading, error number otherwise
   SaveToFileMSF = -1
   On Error GoTo ErrHandler
   If sPath = "" Then sPath = m_FilePath
   If sPath = "" Or bPromptUser Then
      sPath = GetFileName(sPath, MSF.ExportFilterStr, 0, "Save graphical file", False, fTemplate)
      If sPath = "" Then Exit Function
      If GetFileType(sPath) = eJPEG Then
         MSF.JPEGQuality = CLng(fTemplate.lblQuality)
         If fTemplate.Check1.Value Then ProcessImage eGrayScale
      End If
   End If
   MSF.MsoSavePicture sPath, m_BoundedPB
   m_FilePath = Dir(sPath)
   m_FileType = GetFileType(sPath)
ErrHandler:
   SaveToFileMSF = Err.Number
End Function

Public Sub SetSelectedArea(ByVal lft As Long, ByVal tp As Long, ByVal wdt As Long, ByVal hgt As Long)
   m_bUseSelectedArea = wdt
   With m_SelectedArea
       .Left = lft
       .Top = tp
       .Right = lft + wdt
       .Bottom = tp + hgt
   End With
End Sub

Public Sub Balance(ByVal BalanceType As BalanceType, ByVal lAmount As Long, Optional ByVal clrIdx As Integer)
     Dim abBuffer() As Byte
     ReDim abBuffer(UBound(m_abPicture, 1), UBound(m_abPicture, 2))
     CopyMemory abBuffer(0, 0), m_abPicture(0, 0), (UBound(m_abPicture, 1) + 1) * (UBound(m_abPicture, 2) + 1)
     Select Case BalanceType
            Case eChangeBrightness
                 ImgProcess.ChangeBrightness abBuffer, lAmount, clrIdx
            Case eChangeContrast
                 ImgProcess.ChangeContrast abBuffer, lAmount, clrIdx
            Case eChangeGamma
                 ImgProcess.ChangeGamma abBuffer, lAmount, clrIdx
     End Select
     SetBits m_BoundedPB, abBuffer, m_BI ', m_SelectedArea
End Sub

Public Sub UpdateChanges(Optional bUpdate As Boolean)
   If Not bUpdate Then
      m_BoundedPB.Cls
   Else
      m_LastOp = "Balance"
   End If
   UpdateProps
   m_BoundedPB.Picture = m_BoundedPB.Image
End Sub

Public Sub ApplyFilter(ByVal fltType As FILTERTYPES, Optional ByVal KernelSize As Long, Optional ByVal fPower As Long, Optional ByVal GaussSD As Long)
    If fPower Then Filter.FilterPower = fPower
    If KernelSize Then Filter.KernelSize = KernelSize
    If GaussSD Then Filter.Gaussian_SD = GaussSD
    m_LastOp = "Filtering"
    Select Case fltType
        Case eBlur, eSoften
             Filter.Convolve m_abPicture, fltType + 1
        Case eSharp
             Filter.DeConvolve m_abPicture, fltType + 1
        Case eEdge
             Filter.EdgeDetection m_abPicture, eEdgeGradient + KernelFilterType(3), KernelFilterSize(3), True
        Case eRankMedian To eRankMax
             Filter.RankFilter m_abPicture, fltType - eRankMedian + 1
        Case eEnhDetails To eEnhEdges
             Filter.DeConvolve m_abPicture, KernelFilterTypes.eLaplacian
    End Select
    UpdateProps
    SetBits m_BoundedPB, m_abPicture, m_BI ', m_SelectedArea
    m_BoundedPB.Picture = m_BoundedPB.Image
End Sub

Public Sub Crop(ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long)
    Dim sPic As StdPicture
    Set sPic = m_BoundedPB.Picture
    m_bFromCode = True
    m_BoundedPB = LoadPicture()
    m_BoundedPB.PaintPicture sPic, 0, 0, dx, dy, x, y, dx, dy
    Set Picture = m_BoundedPB.Image
    m_bFromCode = False
    Set sPic = Nothing
End Sub

Public Sub ResizeImage(ByVal wdt As Long, ByVal hgt As Long)
    m_BoundedPB.Width = wdt * Screen.TwipsPerPixelX
    m_BoundedPB.Height = hgt * Screen.TwipsPerPixelY
    If m_UseResampling Then
       ImgProcess.Resample wdt, hgt, m_abPicture
       With m_BI.bmiHeader
           .biWidth = wdt
           .biHeight = hgt
           .biSizeImage = BytesPerScanLine(.biWidth) * .biHeight
       End With
       UpdateProps
       m_BoundedPB.Picture = LoadPicture()
       SetBits m_BoundedPB, m_abPicture, m_BI ', m_SelectedArea
       m_BoundedPB.Picture = m_BoundedPB.Image
       Exit Sub
    End If
    Dim sPic As StdPicture
    Set sPic = m_BoundedPB.Picture
    m_bFromCode = True
    m_BoundedPB = LoadPicture()
    m_BoundedPB.PaintPicture sPic, 0, 0, wdt * Screen.TwipsPerPixelX, hgt * Screen.TwipsPerPixelY
    Set Picture = m_BoundedPB.Image
    m_bFromCode = False
    Set sPic = Nothing
End Sub

Public Sub ProcessImage(ByVal ProcessType As ImgProcessTypes, Optional lAmount As Long, Optional ByVal FlipDirection As FlipDirection, Optional ByVal RotateDirection As RotateDirection)
    Dim lTemp As Long
    Dim x As Long, y As Long, xStart As Long, yStart As Long, dx As Long, dy As Long
    Dim abBuffer() As Byte
    If m_bUseSelectedArea Then
       ReDim abBuffer(UBound(m_abPicture, 1), UBound(m_abPicture, 2))
       CopyMemory abBuffer(0, 0), m_abPicture(0, 0), (UBound(m_abPicture, 1) + 1) * (UBound(m_abPicture, 2) + 1)
       xStart = m_SelectedArea.Left
       yStart = m_BI.bmiHeader.biHeight - m_SelectedArea.Bottom
       dx = m_SelectedArea.Right - m_SelectedArea.Left
       dy = m_SelectedArea.Bottom - m_SelectedArea.Top
       ReDim m_abPicture((dx - 1) * 3, dy - 1)
       For y = yStart To yStart + dy
           CopyMemory m_abPicture(0, y - yStart), abBuffer(xStart, y), dx * 3
       Next y
    End If
    Select Case ProcessType
           Case eAutoBalance
                m_LastOp = "Auto balance"
                Histo.EqualizeHisto m_abPicture
           Case eAddNoise
                m_LastOp = "Adding noise"
                ImgProcess.AddNoise m_abPicture, lAmount
           Case eBathRoom
                m_LastOp = "BathRoom"
                ImgProcess.BathRoom m_abPicture
           Case eFade
                m_LastOp = "Fade"
                ImgProcess.Fade m_abPicture, lAmount
           Case eMelt
                m_LastOp = "Melt"
                ImgProcess.Melt m_abPicture
           Case eNegative
                m_LastOp = "Negative"
                ImgProcess.Negative m_abPicture
           Case eRelief
                m_LastOp = "Relief map"
                ImgProcess.Relief m_abPicture
           Case eFlip
                m_LastOp = "Flipping"
                ImgProcess.Flip m_abPicture, FlipDirection
           Case eCaricature
                m_LastOp = "Caricature"
                ImgProcess.Caricature m_abPicture ', m_BoundedPB.Height / Screen.TwipsPerPixelY / 2, m_BoundedPB.BackColor
                ImgProcess.Pixelize m_abPicture, 5
           Case eFishEye
                m_LastOp = "FishEye"
                ImgProcess.FishEye m_abPicture ', m_BoundedPB.Height / Screen.TwipsPerPixelY / 2, m_BoundedPB.BackColor
           Case ePixelize
                m_LastOp = "Pixelize"
                ImgProcess.Pixelize m_abPicture, lAmount
           Case eSwirle
                m_LastOp = "Swirle"
                ImgProcess.Swirle m_abPicture, lAmount  ', m_BoundedPB.Height / Screen.TwipsPerPixelY / 2, m_BoundedPB.BackColor
           Case eRotate
                m_LastOp = "Rotation"
                ImgProcess.Rotate m_abPicture, RotateDirection
                With m_BI.bmiHeader
                     lTemp = .biWidth
                     .biWidth = .biHeight
                     .biHeight = lTemp
                     .biSizeImage = BytesPerScanLine(.biWidth) * .biHeight
                     m_BoundedPB.Width = .biWidth * Screen.TwipsPerPixelX
                     m_BoundedPB.Height = .biHeight * Screen.TwipsPerPixelY
                End With
           Case eBlackWhite
                m_LastOp = "Black_White"
                ImgProcess.BlackWhite m_abPicture
           Case eGrayScale
                m_LastOp = "Gray_Scale"
                ImgProcess.GrayScale m_abPicture
           Case eDither256
                m_LastOp = "Dithering"
                InitPalette m_BoundedPB.hdc
                ImgProcess.ReduceColors m_abPicture, CBool(lAmount)
           Case eColourise
                m_LastOp = "Colourize"
                ImgProcess.Colourise m_abPicture, lAmount, lAmount
    End Select
    If m_bUseSelectedArea Then
       For y = yStart To yStart + dy
           CopyMemory abBuffer(xStart, y), m_abPicture(0, y - yStart), dx * 3
       Next y
    End If
    UpdateProps
    m_BoundedPB.Picture = LoadPicture()
    SetBits m_BoundedPB, m_abPicture, m_BI ', m_SelectedArea
    m_BoundedPB.Picture = m_BoundedPB.Image
End Sub

Public Sub Undo()
   m_bFromCode = True
   Set Picture = m_colPrevPic.Item(m_colPrevPic.Count)
   m_colPrevPic.Remove m_colPrevPic.Count
   m_colOperations.Remove m_colOperations.Count
   m_bFromCode = False
End Sub

Private Sub UpdateProps()
   If Not m_bFromCode Then
      m_colOperations.Add m_LastOp
      m_colPrevPic.Add m_BoundedPB.Picture
   End If
End Sub

Private Sub Class_Initialize()
   Set Histo = New cHistogram
   Set FFT = New cFFT
   Set Filter = New cFilter
   Set ImgProcess = New cImgProcess
   Set MSF = New MSGrphFilter
   Set m_colPrevPic = New Collection
   Set m_colOperations = New Collection
End Sub

Private Sub Class_Terminate()
   Erase m_abPicture
   Set Histo = Nothing
   Set FFT = Nothing
   Set Filter = Nothing
   Set ImgProcess = Nothing
   Set MSF = Nothing
End Sub

'=============Events===========
Private Sub Filter_ProgressChanged(ByVal nValue As Long)
   RaiseEvent ProgressChanged(nValue)
End Sub

Private Sub Filter_ProgressEnd(ByVal nTime As Long)
   RaiseEvent ProgressEnd(nTime)
End Sub

Private Sub Filter_ProgressInit(ByVal nMax As Long)
   RaiseEvent ProgressInit(nMax)
End Sub

Private Sub Histo_ProgressChanged(ByVal nValue As Long)
   RaiseEvent ProgressChanged(nValue)
End Sub

Private Sub Histo_ProgressEnd(ByVal nTime As Long)
   RaiseEvent ProgressEnd(nTime)
End Sub

Private Sub Histo_ProgressInit(ByVal nMax As Long)
   RaiseEvent ProgressInit(nMax)
End Sub

Private Sub ImgProcess_ProgressChanged(ByVal nValue As Long)
   RaiseEvent ProgressChanged(nValue)
End Sub

Private Sub ImgProcess_ProgressEnd(ByVal nTime As Long)
   RaiseEvent ProgressEnd(nTime)
End Sub

Private Sub ImgProcess_ProgressInit(ByVal nMax As Long)
   RaiseEvent ProgressInit(nMax)
End Sub

Private Function GetFileType(ByVal sPath As String) As PICTURE_FORMATS
   Select Case Right$(UCase$(sPath), 3)
          Case "JPG", "PEG"
               GetFileType = eJPEG
          Case "GIF"
               GetFileType = eGIF
          Case "BMP"
               GetFileType = eBMP
          Case Else
               If m_FileType Then GetFileType = m_FileType Else GetFileType = eBMP
   End Select
End Function


