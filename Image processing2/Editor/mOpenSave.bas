Attribute VB_Name = "mOpenSave"
Option Explicit

Private Type OPENFILENAME 'Open & Save Dialog
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' Hook and notification support:
Private Type NMHDR
    hwndFrom As Long
    idfrom As Long
    code As Long
End Type

Private Type OFNOTIFYshort
    hdr As NMHDR
    lpOFN As Long
End Type

Private Const LVM_FIRST = &H1000
Private Const LVM_GETITEMTEXT = LVM_FIRST + 45
Private Const LVM_GETNEXTITEM = LVM_FIRST + 12
Private Const LVNI_FOCUSED = &H1
Private Const LVNI_SELECTED = &H2

Private Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type


Private Type POINTAPI
     x As Long
     y As Long
End Type
 
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Private Const ID_OPEN = &H1  'Open or Save button
Private Const ID_CANCEL = &H2 'Cancel Button
Private Const ID_HELP = &H40E 'Help Button
Private Const ID_READONLY = &H410 'Read-only check box
Private Const ID_FILETYPELABEL = &H441 'FileType label
Private Const ID_FILELABEL = &H442 'FileName label
Private Const ID_FOLDERLABEL = &H443 'Folder label
Private Const ID_LIST = &H461 'Parent of file list
Private Const ID_FORMAT = &H470 'FileType combo box
Private Const ID_FOLDER = &H471 'Folder combo box
Private Const ID_FILETEXT = &H480 'FileName text box

Private Const OFN_HELPBUTTON = &H10
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_EXPLORER = &H80000
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXISTS = &H1000
'OFN_ENABLEHOOK OR OFN_HELPBUTTON  OR OFN_EXPLORER OR OFN_FILEMUSTEXISTS
Private Const OFN_OPENFLAGS = &H81030
'OFN_OPENFLAGS OR OFN_OVERWRITEPROMPT AND NOT OFN_FILEMUSTEXIST
Private Const OFN_SAVEFLAGS = &H80032

Private Const WM_INITDIALOG = &H110
Private Const WM_COMMAND = &H111

Private Const WM_DESTROY = &H2
Private Const WM_NOTIFY = &H4E
Private Const WM_SETICON = &H80

Private Const WM_USER = &H400
Private Const CDM_FIRST = (WM_USER + 100)
Private Const CDM_GETSPEC = (CDM_FIRST + &H0)
Private Const CDM_GETFILEPATH = (CDM_FIRST + &H1)
Private Const CDM_GETFOLDERPATH = (CDM_FIRST + &H2)
Private Const CDM_SETCONTROLTEXT = (CDM_FIRST + &H4)
Private Const CDM_HIDECONTROL = (CDM_FIRST + &H5)
Private Const CDM_SETDEFEXT = (CDM_FIRST + &H6)
Private Const CB_GETCURSEL = &H147

Private Const CDN_FIRST = -601&
Private Const CDN_INITDONE = CDN_FIRST
Private Const CDN_SELCHANGE = (CDN_FIRST - &H1)
Private Const CDN_FOLDERCHANGE = (CDN_FIRST - &H2)
Private Const CDN_HELP = (CDN_FIRST - &H4)
Private Const CDN_FILEOK = (CDN_FIRST - &H5)
Private Const CDN_TYPECHANGE = (CDN_FIRST - &H6)

Public Const MAX_PATH = 260
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private m_hDlg As Long
Private m_fExtraForm As Form
Private m_hOldParent As Long
Private m_bOpen As Boolean, m_bPreview As Boolean
Private m_DlgNormalRC As RECT, m_DlgExtendedRc As RECT
Private sOK As String, sCancel As String, sHelp As String
Private sDlgFilter As String, sCurExt As String

Public Function GetFileName(Optional sPath As String, Optional sFilter As String, Optional nFltIndex As Long, Optional sTitle As String, Optional bOpen As Boolean = True, Optional fExtra As Form) As String
   Dim OFN As OPENFILENAME
   Dim ret As Long, i As Long, sExt As String
   m_hOldParent = 0
   m_hDlg = 0
   m_bOpen = bOpen
   m_bPreview = False
'   Set m_fExtraForm = Nothing
   m_DlgNormalRC.Right = 0
   With OFN
       .hwndOwner = frmTest.hWnd
       .lStructSize = Len(OFN)
        For i = 1 To Len(sFilter)
            If Mid(sFilter, i, 1) = "|" Then
               Mid(sFilter, i, 1) = vbNullChar
            End If
        Next
        If Len(sFilter) < MAX_PATH Then
           sFilter = sFilter & String$(MAX_PATH - Len(sFilter), 0)
        Else
           sFilter = sFilter & Chr(0) & Chr(0)
        End If
        sDlgFilter = sFilter
        If sPath <> "" And (nFltIndex = 0) Then
           nFltIndex = GetFilterIndex(sPath)
        End If
        .lpstrFilter = sFilter
        .nFilterIndex = nFltIndex
        .lpstrTitle = sTitle
        .lpstrInitialDir = App.Path
        .hInstance = App.hInstance
        .lpstrFile = sPath & String(MAX_PATH - Len(sPath), 0)
        .nMaxFile = MAX_PATH
        .lpfnHook = lHookAddress(AddressOf DialogHookProcess)
   End With
   Set m_fExtraForm = fExtra
   If Not m_fExtraForm Is Nothing Then
      m_fExtraForm.picSaveOptions.Visible = Not m_bOpen
      m_fExtraForm.picPreview.Visible = m_bOpen
      m_fExtraForm.lblPreview.Visible = m_bOpen
   End If
   If m_bOpen Then
      OFN.flags = OFN.flags Or OFN_OPENFLAGS
      ret = GetOpenFileName(OFN)
   Else
      OFN.flags = OFN.flags Or OFN_SAVEFLAGS
      ret = GetSaveFileName(OFN)
   End If
   If ret Then
      GetFileName = TrimNull(OFN.lpstrFile)
      If (OFN.nFileExtension = 0) And Len(sCurExt) > 2 Then
         GetFileName = GetFileName & Mid$(sCurExt, 2)
      End If
   End If
End Function

Public Function lHookAddress(lPtr As Long) As Long
  lHookAddress = lPtr
End Function

Public Function DialogHookProcess(ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
   Dim tNMH As NMHDR
   Dim sPath As String, sExt As String
   Dim nPos As Long
   Select Case wMsg
         Case WM_NOTIFY
               CopyMemory tNMH, ByVal lParam, Len(tNMH)
               Select Case tNMH.code
                      Case CDN_FOLDERCHANGE
                           SendMessage m_hDlg, CDM_SETCONTROLTEXT, ID_FILETEXT, ByVal ""
                      Case CDN_SELCHANGE
                           sPath = GetSelItem
                           If sPath <> "" Then
                              SendMessage m_hDlg, CDM_SETCONTROLTEXT, ID_FILETEXT, ByVal sPath
                           End If
                           If m_bPreview Then PreviewPicture
                      Case CDN_FILEOK
                      Case CDN_TYPECHANGE
                           sPath = String(MAX_PATH, 0)
                           Call SendMessage(m_hDlg, CDM_GETSPEC, MAX_PATH, ByVal sPath)
                           sPath = TrimNull(sPath)
                           If (Len(sPath) > 4) Then
                              sExt = Right$(sPath, 5)
                              nPos = InStr(1, sExt, ".")
                              If nPos Then
                                 sPath = Left$(sPath, Len(sPath) - 6 + nPos)
                              End If
                           End If
                           sCurExt = GetExtension()
                           If Len(sCurExt) > 2 Then
                              SendMessage m_hDlg, CDM_SETDEFEXT, 0, ByVal Mid$(sCurExt, 3)
                           End If
                           SendMessage m_hDlg, CDM_SETCONTROLTEXT, ID_FILETEXT, ByVal sPath
                      Case CDN_INITDONE
                           m_hDlg = GetParent(hDlg)
                           ModifyDlg
                           CentreDlg
                      Case CDN_HELP
                           m_bPreview = Not m_bPreview
                           If m_bPreview Then
                              CentreDlg True
                              SendMessage m_hDlg, CDM_SETCONTROLTEXT, ID_HELP, ByVal "<<&Hide"
                              PreviewPicture
                           Else
                              CentreDlg False
                              SendMessage m_hDlg, CDM_SETCONTROLTEXT, ID_HELP, ByVal sHelp
                           End If
                      Case Else
               End Select
          Case WM_DESTROY
               If m_hOldParent Then
                  m_fExtraForm.Visible = False
                  SetParent m_fExtraForm.hWnd, m_hOldParent
               End If
               ' Here you can add user's notification
               ' before exiting
          Case Else
   End Select
End Function

Private Sub CentreDlg(Optional bShowExtra As Boolean)
   Dim lft As Long, tp As Long, wdt As Long, hgt As Long
   Dim rct As RECT
   Dim pt As POINTAPI
   If m_DlgNormalRC.Right = 0 Then
      GetWindowRect m_hDlg, rct
      m_DlgNormalRC = rct
      m_DlgExtendedRc = rct
      If Not m_fExtraForm Is Nothing Then
         m_DlgExtendedRc.Bottom = m_DlgExtendedRc.Bottom + m_fExtraForm.Height / Screen.TwipsPerPixelY
         pt.x = rct.Left
         pt.y = rct.Bottom
         ScreenToClient m_hDlg, pt
         m_fExtraForm.Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY
      End If
   End If
   If bShowExtra Then rct = m_DlgExtendedRc Else rct = m_DlgNormalRC
   wdt = rct.Right - rct.Left
   hgt = rct.Bottom - rct.Top
   lft = (Screen.Width / Screen.TwipsPerPixelX - wdt) / 2
   tp = (Screen.Height / Screen.TwipsPerPixelX - hgt) / 2
   MoveWindow m_hDlg, lft, tp, wdt, hgt, 1
End Sub

Private Sub ModifyDlg()
   SendMessage m_hDlg, WM_SETICON, 0&, ByVal CLng(frmTest.Icon)
   sCancel = "&Forget It!"
   If m_bOpen Then
      sOK = "&Get It!"
      sHelp = "&Preview >>"
   Else
      sOK = "&Save It!"
      sHelp = "&Options >>"
   End If
   If m_fExtraForm Is Nothing Then
      SendMessage m_hDlg, CDM_HIDECONTROL, ID_HELP, ByVal 0&
   Else
      m_hOldParent = SetParent(m_fExtraForm.hWnd, m_hDlg)
      m_fExtraForm.Visible = True
   End If
   SendMessage m_hDlg, CDM_HIDECONTROL, ID_READONLY, ByVal 0&
   SendMessage m_hDlg, CDM_SETCONTROLTEXT, ID_OPEN, ByVal sOK
   SendMessage m_hDlg, CDM_SETCONTROLTEXT, ID_CANCEL, ByVal sCancel
   SendMessage m_hDlg, CDM_SETCONTROLTEXT, ID_HELP, ByVal sHelp
End Sub

Private Sub PreviewPicture()
   If Not m_bOpen Then Exit Sub
   Dim sPath As String, sExt As String
   Dim sSize As String
   On Error Resume Next
   Set m_fExtraForm.picPreview.Picture = LoadPicture()
   m_fExtraForm.lblPreview = ""
   sPath = String(MAX_PATH, 0)
   Call SendMessage(m_hDlg, CDM_GETFILEPATH, MAX_PATH, ByVal sPath)
   sPath = TrimNull(sPath)
   If sPath = "" Then Exit Sub
   If (GetAttr(sPath) And vbDirectory) = vbDirectory Then Exit Sub
   sExt = Right(sPath, 5)
   If (InStr(sExt, ".") = 0) And (Len(sCurExt) > 2) Then
      sPath = sPath & Mid(sCurExt, 2)
   End If
   Set m_fExtraForm.picPreview.Picture = LoadPicture(sPath)
   If m_fExtraForm.picPreview.Picture = 0 Then
      m_fExtraForm.lblPreview = "Can't preview this picture"
   Else
      If FileLen(sPath) > 1024 Then
         sSize = " (" & FileLen(sPath) \ 1024 & " KB)"
      Else
         sSize = " (" & FileLen(sPath) & " Bytes)"
      End If
      m_fExtraForm.lblPreview = GetPicInfo(m_fExtraForm.picPreview, sPath) & sSize
   End If
End Sub

Public Function TrimNull(startstr As String) As String
   Dim pos As Integer
   pos = InStr(startstr, Chr$(0))
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
   TrimNull = startstr
End Function

Private Function GetSelItem() As String
   Dim LI As LV_ITEM
   Dim ret As Long, hFileList As Long
   Dim sPath As String, sNewPath As String
   Static sOldPath As String
   sNewPath = String(MAX_PATH, 0)
   Call SendMessage(m_hDlg, CDM_GETFILEPATH, MAX_PATH, ByVal sNewPath)
   sNewPath = TrimNull(sNewPath)
   If sNewPath <> sOldPath Then 'User selected another file
      sOldPath = sNewPath
      SendMessage m_hDlg, CDM_SETCONTROLTEXT, ID_OPEN, ByVal sOK
      Exit Function
   End If
'User selected a folder
'   SendMessage m_hDlg, CDM_SETCONTROLTEXT, ID_OPEN, ByVal "&Open it!"
   hFileList = GetDlgItem(GetDlgItem(m_hDlg, ID_LIST), 1)
   If hFileList = 0 Then Exit Function
   ret = SendMessage(hFileList, LVM_GETNEXTITEM, -1, ByVal LVNI_SELECTED)
   If ret = -1 Then Exit Function
   LI.cchTextMax = MAX_PATH
   LI.pszText = Space$(MAX_PATH)
   ret = SendMessage(hFileList, LVM_GETITEMTEXT, ret, LI)
   If ret > 1 Then sPath = Left$(LI.pszText, ret)
   GetSelItem = sPath
   sOldPath = sPath
End Function

Private Function GetPicInfo(pb As PictureBox, ByVal sPath As String) As String
  Dim bm As BITMAP
  Dim lWidth As Long, lHeight As Long, lBPP As Long, ret As Long
  Dim sType As String, sExt As String
  sExt = LCase(Right(sPath, 4))
  If Left(sExt, 1) = "." Then sExt = Mid(sExt, 2)
  If sExt = "jpeg" Then sExt = "jpg"
  ret = GetObjectAPI(pb.Picture, Len(bm), bm)
  If ret Then
     lWidth = bm.bmWidth
     lHeight = bm.bmHeight
     lBPP = bm.bmBitsPixel
  Else
     lWidth = pb.ScaleX(pb.Picture.Width, vbHimetric, vbPixels)
     lHeight = pb.ScaleY(pb.Picture.Height, vbHimetric, vbPixels)
  End If
  Select Case pb.Picture.Type
         Case vbPicTypeBitmap
              If sExt = "bmp" Then
                 sType = "Bitmap"
              Else
                 sType = UCase(sExt)
              End If
         Case vbPicTypeMetafile
              sType = "Metafile"
         Case vbPicTypeEMetafile
              sType = "EnhMetafile"
         Case vbPicTypeIcon
              sType = "Icon"
         Case Else
              sType = UCase(sExt)
  End Select
  GetPicInfo = sType & " " & lWidth & "x" & lHeight
  If lBPP <> 0 Then GetPicInfo = GetPicInfo & "x" & lBPP & "BPP"
End Function

Private Function GetExtension() As String
   Dim i As Long, nFilter As Long, nStart As Long, hCombo As Long
   Dim sFilter As String, sTemp As String
   hCombo = GetDlgItem(m_hDlg, ID_FORMAT)
   nFilter = SendMessage(hCombo, CB_GETCURSEL, 0, ByVal 0&)
   sFilter = sDlgFilter
   For i = 1 To nFilter * 2 + 1
       nStart = InStr(1, sFilter, Chr(0))
       If nStart Then
          sFilter = Mid(sFilter, nStart + 1)
       Else
          Exit For
       End If
   Next i
   sTemp = TrimNull(sFilter)
   If Len(sTemp) = 0 Then Exit Function
   If InStr(1, sTemp, ";") = 0 Then GetExtension = sTemp
End Function

Private Function GetFilterIndex(ByVal sPath As String) As Long
   Dim sExt As String
   Dim nIdx As Long, nStart As Long
   sExt = LCase(Right(sPath, 4))
   If Left(sExt, 1) = "." Then sExt = Mid(sExt, 2)
   If sExt = "jpeg" Then sExt = "jpg"
   If sExt = "tiff" Then sExt = "tif"
   sExt = "*." & sExt & Chr(0)
   nStart = 1
   Do While nStart
     nStart = InStr(nStart + 1, sDlgFilter, Chr(0), vbTextCompare)
     If Mid(sDlgFilter, nStart + 1, Len(sExt)) = sExt Then Exit Do
     nIdx = nIdx + 1
   Loop
   GetFilterIndex = nIdx \ 2 + 1
End Function
