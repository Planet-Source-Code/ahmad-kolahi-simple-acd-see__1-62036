VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MSGrphFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long

Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Declare Function GetObjectType Lib "gdi32" (ByVal hgdiobj As Long) As Long
Private Declare Function SetViewportExtEx Lib "gdi32" (ByVal hdc As Long, ByVal nx As Long, ByVal ny As Long, lpSize As Any) As Long
Private Declare Function CreateMetaFile Lib "gdi32" Alias "CreateMetaFileA" (ByVal lpString As String) As Long
Private Declare Function GetMetaFileBitsEx Lib "gdi32" (ByVal hMF As Long, ByVal nSize As Long, lpvData As Any) As Long
Private Declare Function SetMetaFileBitsEx Lib "gdi32" (ByVal nSize As Long, lpData As Any) As Long
Private Declare Function PlayMetaFile Lib "gdi32" (ByVal hdc As Long, ByVal hMF As Long) As Long
Private Declare Function CloseMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Private Declare Function DeleteMetaFile Lib "gdi32" (ByVal hMF As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hdc As Long, ByVal nMapMode As Long) As Long
Private Const MM_ANISOTROPIC = 8
Private Const OBJ_METAFILE = 9
Private Const IBMFNSIZE = 256


'(* Data types defined by the Aldus Interface. *)

'typedef struct {
'   unsigned    slippery : 1;        /* TRUE if file may disappear. */
'   unsigned    write : 1;           /* TRUE if open for write. */
'   unsigned    unnamed : 1;         /* TRUE if unnamed. */
'   unsigned    linked : 1;          /* Linked to an FS FCB. */
'   unsigned    mark : 1;            /* Generic mark bit. */
'   FILETYPE    fType;               /* The file type. */
'#define IBMFNSIZE 124
'   short       handle;              /* MS-DOS open file handle. */
'   char        fullName[IBMFNSIZE]; /* Device, path, file names. */
'   DWORD       filePos;          /* Our current file posn. */
'} FILESPEC, FAR *LPFILESPEC;
'//Note: Although one-bit wide, the flags each occupy a full byte...


'Private Type ALDUS_FILESPEC_OLD
'   slippery As Byte                '/* TRUE if file may disappear. */
'   write As Byte                   '/* TRUE if open for write. */
'   unnamed As Byte                 '/* TRUE if unnamed. */
'   linked As Byte                  '/* Linked to an FS FCB. */
'   Mark As Byte                    '/* Generic mark bit. */
'   Flags As Integer
'   fType As Long                   '/* The file type. */
'   handle As Integer               '/* MS-DOS open file handle. */
'   fullName As String * IBMFNSIZE '/* Device, path, file names. */
'   fullName(IBMFNSIZE - 1) As Byte '/* Device, path, file names. */
'   filePos As Long                 '/* Our current file posn. */
'End Type

'Functions prototypes
'Public Declare Function ImportGr Lib "jpegim32.flt" (ByVal hIC As Integer, ByRef FileSpec As Any, ByRef PictInfo As Any, ByVal hMemPref As Long) As Long
'Public Declare Function ExportGr Lib "jpegim32.flt" (ByVal hIC As Long, ByRef FileSpec As Any, ByRef PictInfo As Any, ByVal hMemPref As Long) As Long
'Public Declare Function GetFilterInfo Lib "gifimp32.flt" (ByVal PageMakerVersion As Long, ByVal sIni As String, hPrefMem As Long, hFileTypes As Long) As Long
'Public Declare Function GetFilterPref Lib "gifimp32.flt" (ByVal hInst As Long, ByVal ahWnd As Long, ByVal hPrefMem As Long, ByVal wFlags As Long) As Long
'Public Declare Function SetFilterPref Lib "jpegim32.flt" (ByVal hMemPref As Long, ByVal szOption As String, ByVal szValue As String, ByVal dwReserved1 As Long, ByVal dwReserved2 As Long) As Long

Private Type ALDUS_FILESPEC
   flags As Long
   hFile As Long                   '/* The file handle. */
'   fullName As String * IBMFNSIZE '/* Device, path, file names. */
   fullName(IBMFNSIZE - 1) As Byte '/* Device, path, file names. */
   filePos As Long                 '/* Our current file posn. */
   Unknown As Long
End Type

'Flags
Const fSlippery = &H1    ' { file may disappear }
Const fWrite = &H2       ' { open for write }
Const fUnnamed = &H4     ' { unnamed }
Const fLinked = &H8      ' { linked to an FS FCB }
Const fMark = &H10       ' { Generic mark bit }

Const CommonFileTypes = &H170000

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Type ALDUS_PICTINFO
   hMF As Long
   rec As RECT
   inch As Long
End Type

Dim m_ImportSupported As Boolean
Dim m_ExportSupported As Boolean
Dim m_TransparentBk As Boolean
Dim m_JPEGQuality As Integer
Dim m_ShowProgress As Boolean
Dim m_ImportFilter As String
Dim m_ExportFilter As String

Dim hModule As Long, hMemPref As Long

Public Property Let ShowProgress(ByVal bShow As Boolean)
    m_ShowProgress = bShow
End Property

Public Property Get ShowProgress() As Boolean
    ShowProgress = m_ShowProgress
End Property

Public Property Let JPEGQuality(ByVal Quality As Integer)
    m_JPEGQuality = Quality
End Property

Public Property Get JPEGQuality() As Integer
    JPEGQuality = m_JPEGQuality
End Property

Public Property Let TransparentBackground(ByVal bTransparent As Boolean)
    m_TransparentBk = bTransparent
End Property

Public Property Get TransparentBackground() As Boolean
    TransparentBackground = m_TransparentBk
End Property

Public Property Get IsImportSupported() As Boolean
   IsImportSupported = m_ImportSupported
End Property

Public Property Get IsExportSupported() As Boolean
   IsExportSupported = m_ExportSupported
End Property

Public Property Get ImportFilterStr() As String
   Dim sTemp As String
   Dim i As Long
   If m_ImportFilter = "" Then
     For i = 0 To UBound(ImportFilters)
          m_ImportFilter = m_ImportFilter & ImportFilters(i).Name & " (*." & ImportFilters(i).Extension & " )" & "|*." & ImportFilters(i).Extension & "|"
          sTemp = sTemp & "*." & ImportFilters(i).Extension & ";"
      Next i
      m_ImportFilter = Left(m_ImportFilter, Len(m_ImportFilter) - 1)
      sTemp = Left(sTemp, Len(sTemp) - 1) & ";*.tiff;*.jpeg"
      m_ImportFilter = "All graphics (" & sTemp & ")" & "|" & sTemp & "|" & m_ImportFilter
   End If
   ImportFilterStr = m_ImportFilter
End Property

Public Property Get ExportFilterStr() As String
   Dim sTemp As String
   Dim i As Long
   If m_ExportFilter = "" Then
      For i = 0 To UBound(ExportFilters)
          m_ExportFilter = m_ExportFilter & ExportFilters(i).Name & " (*." & ExportFilters(i).Extension & " )" & "|*." & ExportFilters(i).Extension & "|"
          IsBMPExportSupported = IsBMPExportSupported Or (ExportFilters(i).Extension = "bmp")
      Next i
      m_ExportFilter = Left(m_ExportFilter, Len(m_ExportFilter) - 1)
      If Not IsBMPExportSupported Then
         m_ExportFilter = "Windows bitmap (*.bmp)|*.bmp|" & m_ExportFilter
      End If
   End If
   ExportFilterStr = m_ExportFilter
End Property

Private Sub ApplySettings(ByVal sExt As String)
   Dim pMem As Long, lTemp As Long, nSize As Long
   If hMemPref = 0 Then Exit Sub
' *******hMemPref settings for *.gif,*.jpeg,*.png********
' jpg occupies 32 bytes, while *.gif and *.png occupy 36 bytes
' Offset   Default_Value          Description
'   0           2               nPageMakerversion(1-4)
'   4           2               MetaFileType(1-EMW,2-WMF)
'                               EMF files(1) are supported by
'                               nPageMakerVersion > 2 only
'   6           1               ProgressDialog Template number
'  10           1               ShowProgressDialog (Boolean)
'..22      75(jpg), 0(gif,png)  JPEG Quality, or transparent background
'                               for gif/png
   pMem = GlobalLock(hMemPref)
   nSize = GlobalSize(hMemPref)
   If nSize = 0 Then Exit Sub
   CopyMemory ByVal pMem, 2, 4
   If nSize >= 4 Then CopyMemory ByVal pMem + 4, 2, 1
   lTemp = Abs(m_ShowProgress)
   If pMem >= 6 Then CopyMemory ByVal pMem + 6, lTemp, 4
   If pMem >= 10 Then CopyMemory ByVal pMem + 10, lTemp, 4
   If sExt = "jpg" Then
      lTemp = m_JPEGQuality
      If lTemp Then CopyMemory ByVal pMem + 22, lTemp, 2
   ElseIf (sExt = "gif") Or (sExt = "png") Then
      lTemp = Abs(m_TransparentBk)
      If nSize >= 22 Then CopyMemory ByVal pMem + 22, CInt(lTemp), 2
   End If
   GlobalUnlock hMemPref
End Sub

Private Function GetFilterSettings(ByVal sFilterPath As String) As Boolean
    Dim lpfnGetFilterInfo As Long
    Dim lpfnGetFilterPref As Long
    Dim ret As Long
    hModule = LoadLibrary(sFilterPath)

    lpfnGetFilterInfo = GetProcAddress(hModule, "GetFilterInfo")
    ret = CallApiByPointer(lpfnGetFilterInfo, 2, 0, VarPtr(hMemPref), CommonFileTypes)
    If (ret > 5000) Then 'ret is supported nPageMaker version. Currently 9. Error values started from 5000.
       ClearUp
       Exit Function
    End If
    GetFilterSettings = True
    lpfnGetFilterPref = GetProcAddress(hModule, "GetFilterPref")
    If lpfnGetFilterPref And hMemPref Then
       CallApiByPointer lpfnGetFilterPref, 0, 0, hMemPref, CommonFileTypes
    End If
    FreeLibrary hModule
End Function

Public Function MsoLoadPicture(ByVal sFilename As String, pb As PictureBox) As Boolean
    Dim sExt As String, sFilterPath As String
    Dim ret As Long, hDummy As Long, i As Long
    Dim picWidth As Single, picHeight As Single
    Dim pMem As Long, hMem As Long
    Dim lpfnImportGr As Long
    Dim PictInfo As ALDUS_PICTINFO
    Dim FileSpec As ALDUS_FILESPEC
    Dim sPic As StdPicture
    sExt = GetExtension(sFilename)
    sFilterPath = GetFilter(sExt, True).Path
    If GetFilterSettings(sFilterPath) = False Then Exit Function
    
    hModule = LoadLibrary(sFilterPath)
    If hModule = 0 Then Exit Function
    
    lpfnImportGr = GetProcAddress(hModule, "ImportGr")
    If lpfnImportGr = 0 Then lpfnImportGr = GetProcAddress(hModule, "ImportGR")
    If lpfnImportGr = 0 Then ClearUp
    
    CopyMemory FileSpec.fullName(0), ByVal sFilename, IBMFNSIZE
    If (sExt = "jpg") Or (sExt = "gif") Or (sExt = "png") Then ApplySettings sExt
    ret = CallApiByPointer(lpfnImportGr, pb.hdc, VarPtr(FileSpec), VarPtr(PictInfo), hMemPref)
    If PictInfo.hMF = 0 Then
       ClearUp
       Exit Function
    End If
'EMS and WPG filters return hMem value instead of hWMF, ie they return memory pointer to metafile
'bits instead of Metafile handle. Check this and convert memory block to MetaFile if nessesary
    If GetObjectType(PictInfo.hMF) <> OBJ_METAFILE Then
       hMem = PictInfo.hMF
       pMem = GlobalLock(hMem)
       PictInfo.hMF = SetMetaFileBitsEx(GlobalSize(hMem), ByVal pMem)
       PictInfo.inch = PictInfo.inch * 96 / 72
       GlobalUnlock hMem
       GlobalFree hMem
    End If
    If PictInfo.hMF = 0 Then
       ClearUp
       Exit Function
    End If
    picWidth = (PictInfo.rec.Right - PictInfo.rec.Left) / PictInfo.inch  'image width in inches
    picHeight = (PictInfo.rec.Bottom - PictInfo.rec.Top) / PictInfo.inch 'image height in inches
    'convert inches to pixels
    picWidth = pb.ScaleX(picWidth, vbInches, vbPixels)
    picHeight = pb.ScaleX(picHeight, vbInches, vbPixels)
    pb.Move 0, 0, picWidth * Screen.TwipsPerPixelX, picHeight * Screen.TwipsPerPixelY
'EMF filter call SetViewportExtEx internally, so second call lead to error.
    SetMapMode pb.hdc, MM_ANISOTROPIC
    If sExt <> "emf" Then
       SetViewportExtEx pb.hdc, picWidth, picHeight, ByVal 0&
    End If
    Call PlayMetaFile(pb.hdc, PictInfo.hMF)
    pb.Refresh
    DeleteMetaFile PictInfo.hMF
    FreeLibrary hModule
End Function

Public Function MsoSavePicture(ByVal sFilename As String, pb As PictureBox) As Boolean
    Dim sExt As String, sFilterPath As String
    Dim ret As Long, hMetaDC As Long, hDummy As Long, i As Long
    Dim lpfnGetFilterInfo As Long
    Dim lpfnExportGr As Long
'Actually, ExportGr function require forth parameter as ExportCallBack
'address but VB address of function doesn't work in this case because
'above function declared as _cdecl. Following trick allow bypass this
'restriction - we fill stack with dummy values.
    Dim sEatStackSpace As String * 64
    Dim PictInfo As ALDUS_PICTINFO
    Dim FileSpec As ALDUS_FILESPEC
    
    ReDim abFileSpec(LenB(FileSpec) - 1)
    
    sExt = GetExtension(sFilename)
    If sExt = "bmp" Then
       SavePicture pb.Picture, sFilename
       MsoSavePicture = True
       Exit Function
    End If
    sFilterPath = GetFilter(sExt, False).Path
    hModule = LoadLibrary(sFilterPath)
    If hModule = 0 Then Exit Function
    Call GetFilterSettings(sFilterPath)
'   Check out if ExportGr enabled
    lpfnExportGr = GetProcAddress(hModule, "ExportGr")
    If lpfnExportGr = 0 Then
       ClearUp
       Exit Function
    End If
'Transform picture(BMP) to WMF format
    hMetaDC = CreateMetaFile(vbNullString)
    Call BitBlt(hMetaDC, 0, 0, 0, 0, pb.hdc, 0, 0, vbSrcCopy)
    PictInfo.hMF = CloseMetaFile(hMetaDC)
'Apply some settings(Show_Progress, JPEG_Quality)
    If (sExt = "jpg") Or (sExt = "gif") Or (sExt = "png") Then ApplySettings sExt
    CopyMemory FileSpec.fullName(0), ByVal sFilename, IBMFNSIZE
'Fill up the stack with dummy values - see note above. Now callback function
'will just re-fill dummies
    sEatStackSpace = String(64, 1)
    ret = CallApiByPointer(lpfnExportGr, VarPtr(FileSpec), VarPtr(PictInfo), hMemPref) ',lpfn_cdeclCallBack)
    MsoSavePicture = (ret = 0)
    DeleteMetaFile PictInfo.hMF
    GlobalFree hMemPref
    FreeLibrary hModule
End Function

Private Sub Class_Initialize()
    InitFilters
    If m_ExportSupported Then
       m_JPEGQuality = Asc(GetRegValue(GetFilter("jpg", False).RegPath & "\Options", "Quality"))
       m_TransparentBk = (GetRegValue(GetFilter("gif", False).RegPath & "\Options", "SetMetafileBackgroundTransparent") = "Yes")
       m_ShowProgress = (GetRegValue(GetFilter("gif", False).RegPath & "\Options", "ShowProgressDialog") = "Yes")
    End If
End Sub

Private Sub Class_Terminate()
    ClearUp
End Sub

Private Sub InitFilters()
   Dim nBound As Long
   m_ImportSupported = EnumFilters(ImportFilters, eImport)
   m_ExportSupported = EnumFilters(ExportFilters, eExport)
End Sub

Private Function GetExtension(ByVal sFilename As String) As String
    Dim sExt As String
    sExt = Right(sFilename, 4)
    If Left(sExt, 1) = "." Then sExt = Mid(sExt, 2)
    If sExt = "tiff" Then sExt = "tif"
    If sExt = "jpeg" Then sExt = "jpg"
    GetExtension = LCase(sExt)
End Function

Private Function GetFilter(ByVal sExt As String, Optional bImport As Boolean = True) As FILTERINFO
    Dim i As Long
    If bImport Then
       For i = 0 To UBound(ImportFilters)
           If ImportFilters(i).Extension = sExt Then
              GetFilter = ImportFilters(i)
              Exit For
           End If
       Next i
    Else
       For i = 0 To UBound(ExportFilters)
           If ExportFilters(i).Extension = sExt Then
              GetFilter = ExportFilters(i)
              Exit For
           End If
       Next i
    End If
End Function

Private Sub ClearUp()
    GlobalFree hMemPref
    FreeLibrary hModule
End Sub
