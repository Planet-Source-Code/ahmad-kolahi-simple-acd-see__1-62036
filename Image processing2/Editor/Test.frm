VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   "Form1"
   ClientHeight    =   10185
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   15240
   Icon            =   "Test.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11000
   ScaleMode       =   0  'User
   ScaleWidth      =   16708.52
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton nextb 
      Caption         =   "Next -->"
      Height          =   280
      Left            =   5640
      TabIndex        =   9
      Top             =   3720
      Width           =   800
   End
   Begin VB.CommandButton backb 
      Caption         =   "<-- Back"
      Height          =   280
      Left            =   4800
      TabIndex        =   8
      Top             =   3720
      Width           =   800
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   2100
      TabIndex        =   6
      Top             =   3720
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   480
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3120
      Width           =   3375
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   3315
      Left            =   6180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Width           =   255
   End
   Begin VB.PictureBox picContainer 
      Height          =   3435
      Left            =   60
      ScaleHeight     =   3375
      ScaleWidth      =   5955
      TabIndex        =   0
      Top             =   180
      Width           =   6015
      Begin VB.PictureBox picImage 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   420
         ScaleHeight     =   3135
         ScaleWidth      =   3675
         TabIndex        =   1
         Top             =   180
         Width           =   3675
         Begin VB.Label lblGrip 
            BackColor       =   &H00FF0000&
            BorderStyle     =   1  'Fixed Single
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   1260
            Width           =   795
         End
         Begin VB.Label lblShape 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Height          =   915
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   2205
         End
      End
   End
   Begin VB.Label lblProgress 
      Caption         =   "Label1"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   3720
      Width           =   1815
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCut 
         Caption         =   "C&ut"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "&Copy"
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "&SelectArea"
      End
      Begin VB.Menu mnuCrop 
         Caption         =   "Cr&op"
      End
      Begin VB.Menu mnuRemoveSelection 
         Caption         =   "&Remove selection"
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Cle&ar"
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCBClear 
         Caption         =   "Cl&ear clipboard"
      End
   End
   Begin VB.Menu mnuImage 
      Caption         =   "I&mage"
      Begin VB.Menu mnuRotate 
         Caption         =   "Rotate &Left"
         Index           =   0
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuRotate 
         Caption         =   "Rotate &Right"
         Index           =   1
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuFlip 
         Caption         =   "Flip &Vert"
         Index           =   0
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuFlip 
         Caption         =   "Flip &Horz"
         Index           =   1
         Shortcut        =   ^H
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuResize 
         Caption         =   "Re&size"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutoBalance 
         Caption         =   "&AutoBalance"
      End
      Begin VB.Menu mnuBalance 
         Caption         =   "&Balance"
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFlt 
         Caption         =   "&Filters"
         Begin VB.Menu mnuStdFlt 
            Caption         =   "&Blur"
            Index           =   0
         End
         Begin VB.Menu mnuStdFlt 
            Caption         =   "&Soften"
            Index           =   1
         End
         Begin VB.Menu mnuStdFlt 
            Caption         =   "S&harpen"
            Index           =   2
         End
         Begin VB.Menu mnuStdFlt 
            Caption         =   "Edge &detection"
            Index           =   3
         End
         Begin VB.Menu sep14 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRnk 
            Caption         =   "&Rank filters"
            Begin VB.Menu mnuRank 
               Caption         =   "&Median"
               Index           =   0
            End
            Begin VB.Menu mnuRank 
               Caption         =   "M&in"
               Index           =   1
            End
            Begin VB.Menu mnuRank 
               Caption         =   "M&ax"
               Index           =   2
            End
         End
         Begin VB.Menu mnuEnh 
            Caption         =   "&Enhanced"
            Begin VB.Menu mnuEnhanced 
               Caption         =   "&Details"
               Index           =   0
            End
            Begin VB.Menu mnuEnhanced 
               Caption         =   "&Edges"
               Index           =   1
            End
            Begin VB.Menu mnuEnhanced 
               Caption         =   "&Focus"
               Index           =   2
            End
         End
         Begin VB.Menu sep13 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFltOptions 
            Caption         =   "&Filter options"
         End
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEff 
         Caption         =   "&Effects"
         Begin VB.Menu mnuEffects 
            Caption         =   "&Add noise"
            Index           =   0
         End
         Begin VB.Menu mnuEffects 
            Caption         =   "&Bath room"
            Index           =   1
         End
         Begin VB.Menu mnuEffects 
            Caption         =   "&Caricature"
            Index           =   2
         End
         Begin VB.Menu mnuEffects 
            Caption         =   "&Fade"
            Index           =   3
         End
         Begin VB.Menu mnuEffects 
            Caption         =   "Fish &Eye"
            Index           =   4
         End
         Begin VB.Menu mnuEffects 
            Caption         =   "&Melt"
            Index           =   5
         End
         Begin VB.Menu mnuEffects 
            Caption         =   "&Negative"
            Index           =   6
         End
         Begin VB.Menu mnuEffects 
            Caption         =   "&Pixelize"
            Index           =   7
         End
         Begin VB.Menu mnuEffects 
            Caption         =   "&Relief map"
            Index           =   8
         End
         Begin VB.Menu mnuEffects 
            Caption         =   "&Swirle"
            Index           =   9
         End
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClr 
         Caption         =   "&Colors"
         Begin VB.Menu mnuColor 
            Caption         =   "Black and &White"
            Index           =   0
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&Gray Scale"
            Index           =   1
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&System_256"
            Index           =   2
         End
         Begin VB.Menu mnuColor 
            Caption         =   "&Colourise"
            Index           =   3
         End
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents PicEx As cPictureEx
Attribute PicEx.VB_VarHelpID = -1

Dim colOperations As New Collection
Dim fResize As frmResize
Dim fTrack As frmTrack
Attribute fTrack.VB_VarHelpID = -1
Dim fDlg As frmDlgExtra

'========Sizing grip staff===========
Dim XGrip(2) As Long, YGrip(2) As Long
Dim bMoving As Boolean, bSizing As Boolean
Dim xStart As Long, yStart As Long
Const GripSize = 90
'====================================

Private Sub backb_Click()
  If Not fDlg Is Nothing Then Set fDlg = Nothing
   Set fDlg = New frmDlgExtra
   PicEx.SaveToFile , True, fDlg
   Set fDlg = Nothing
   frmTest.Hide
   frmCrop.Show
   frmPicViewer.picViewX.Picture = LoadPicture("test2.jpg")
   frmPicViewer.picCopy.Picture = LoadPicture("test2.jpg")
   frmTest.Hide
   frmPicViewer.Show
   'Unload frmTest
End Sub

Private Sub Form_Activate()
  
'  If Not fDlg Is Nothing Then Set fDlg = Nothing
'  Set fDlg = New frmDlgExtra
'  PicEx.SaveToFile , True, fDlg
'  Set fDlg = Nothing
'
'   PicEx.BoundPictureBox picImage
'   DoEvents
'
'
'  Dim ret As Long
'    If Not fDlg Is Nothing Then Set fDlg = Nothing
'    Set fDlg = New frmDlgExtra
'    ret = PicEx.LoadFromFile(, , fDlg)
'    If ret = 0 Then
'       Set picImage.Picture = PicEx.Picture
'    ElseIf ret <> -1 Then
'       Err.Raise ret
'    End If
'    Set fDlg = Nothing
End Sub

Private Sub Form_Load()
   HScroll1.Visible = False
   Set PicEx = New cPictureEx
   GetFilterSettings
   Set fResize = New frmResize
   Set fTrack = New frmTrack
   Set fDlg = New frmDlgExtra
   Caption = "Edit"
   lblProgress = ""
   VScroll1.SmallChange = 150
   VScroll1.LargeChange = 1500
   HScroll1.SmallChange = 150
   HScroll1.LargeChange = 1500
   InitGrip

   Me.Show 'force showing
   PicEx.BoundPictureBox picImage
   DoEvents

    Dim ret As Long
    If Not fDlg Is Nothing Then Set fDlg = Nothing
    Set fDlg = New frmDlgExtra
    ret = PicEx.LoadFromFile(, , fDlg)
    If ret = 0 Then
       Set picImage.Picture = PicEx.Picture
    ElseIf ret <> -1 Then
       Err.Raise ret
    End If
    Set fDlg = Nothing
   

End Sub

Private Sub Form_Resize()
   Dim dy As Long
   If HScroll1.Visible Then dy = HScroll1.Height
   If WindowState = vbMinimized Then Exit Sub
   picContainer.Move 0, 0, ScaleWidth, ScaleHeight - ProgressBar1.Height * 2
   picImage.Move 0, 0 ', picContainer.ScaleWidth, picContainer.ScaleHeight
   nextb.Move ScaleWidth - 900, ScaleHeight - 400
   backb.Move ScaleWidth - 1800, ScaleHeight - 400
   CentreForm fResize
   CentreForm fTrack
End Sub

Private Sub Form_Unload(Cancel As Integer)
   On Error Resume Next
   Set PicEx = Nothing
   Set colOperations = Nothing
   Unload fResize: Set fResize = Nothing
   Unload fTrack: Set fTrack = Nothing
   If Not fDlg Is Nothing Then
      Unload fDlg
   End If
   End
End Sub

Private Sub mnuAutoBalance_Click()
   MousePointer = vbHourglass
   PicEx.ProcessImage eAutoBalance
   MousePointer = vbDefault
End Sub

Private Sub mnuBalance_Click()
   fTrack.Show
End Sub

Private Sub mnuCBClear_Click()
   Clipboard.Clear
End Sub

Private Sub mnuClear_Click()
   Set PicEx.Picture = Nothing
   Set picImage.Picture = Nothing
   ShowGrip False
End Sub

Private Sub mnuColor_Click(Index As Integer)
   MousePointer = vbHourglass
   Select Case Index
       Case 0: PicEx.ProcessImage eBlackWhite
       Case 1: PicEx.ProcessImage eGrayScale
       Case 2: PicEx.ProcessImage eDither256, 1
       Case 3: PicEx.ProcessImage eColourise, 2
   End Select
   MousePointer = vbDefault
End Sub

Private Sub mnuCopy_Click()
   Clipboard.Clear
   If lblShape.Visible Then
      PicEx.Crop lblShape.Left / Screen.TwipsPerPixelX, lblShape.Top / Screen.TwipsPerPixelY, lblShape.Width / Screen.TwipsPerPixelX, lblShape.Height / Screen.TwipsPerPixelY
      Clipboard.SetData PicEx.Picture
      Set PicEx.Picture = picImage.Picture
   Else
      Clipboard.SetData PicEx.Picture
   End If
End Sub

Private Sub mnuCrop_Click()
   PicEx.Crop lblShape.Left, lblShape.Top, lblShape.Width, lblShape.Height
   ShowGrip False
   picImage.Picture = PicEx.Picture
End Sub

Private Sub mnuCut_Click()
   Clipboard.Clear
   If lblShape.Visible Then
      PicEx.Crop lblShape.Left, lblShape.Top, lblShape.Width, lblShape.Height
      Clipboard.SetData PicEx.Picture
      picImage.Line (lblShape.Left, lblShape.Top)-Step(lblShape.Width, lblShape.Height), picImage.BackColor, BF
      Set picImage.Picture = picImage.Image
   Else
      Clipboard.SetData PicEx.Picture
      Set picImage.Picture = Nothing
   End If
   Set PicEx.Picture = picImage.Picture
End Sub

Private Sub mnuEdit_Click()
   UpdateEditMenu
End Sub

Private Sub mnuEffects_Click(Index As Integer)
   MousePointer = vbHourglass
   Select Case Index
       Case 0: PicEx.ProcessImage eAddNoise, 80
       Case 1: PicEx.ProcessImage eBathRoom
       Case 2: PicEx.ProcessImage eCaricature
       Case 3: PicEx.ProcessImage eFade, 40
       Case 4: PicEx.ProcessImage eFishEye
       Case 5: PicEx.ProcessImage eMelt
       Case 6: PicEx.ProcessImage eNegative
       Case 7: PicEx.ProcessImage ePixelize, 5
       Case 8: PicEx.ProcessImage eRelief
       Case 9: PicEx.ProcessImage eSwirle, 70
       Case Else
   End Select
   MousePointer = vbDefault
End Sub

Private Sub mnuEnhanced_Click(Index As Integer)
   MousePointer = vbHourglass
   PicEx.ApplyFilter eEnhDetails + Index, EnhancedFilterSize(Index), Index + 1, 0.7 + Index / 10
   MousePointer = vbDefault
End Sub

Private Sub mnuFlip_Click(Index As Integer)
   MousePointer = vbHourglass
   PicEx.ProcessImage eFlip, , Index
   MousePointer = vbDefault
End Sub

Private Sub mnuFltOptions_Click()
  frmFltOptions.Show
End Sub

Private Sub mnuImage_Click()
   UpdateImageMenu
End Sub

Private Sub mnuLoad_Click()
    Dim ret As Long
    If Not fDlg Is Nothing Then Set fDlg = Nothing
    Set fDlg = New frmDlgExtra
    ret = PicEx.LoadFromFile(, , fDlg)
    If ret = 0 Then
       Set picImage.Picture = PicEx.Picture
    ElseIf ret <> -1 Then
       Err.Raise ret
    End If
    Set fDlg = Nothing
End Sub

Private Sub mnuOpenMSF_Click()
    Dim ret As Long
    If Not fDlg Is Nothing Then Set fDlg = Nothing
    Set fDlg = New frmDlgExtra
    ret = PicEx.LoadFromFileMSF(, , fDlg)
    If ret = 0 Then
       Set picImage.Picture = PicEx.Picture
    ElseIf ret <> -1 Then
       Err.Raise ret
    End If
    Set fDlg = Nothing
End Sub

Private Sub mnuPaste_Click()
   If lblShape.Visible Then
      picImage.PaintPicture Clipboard.GetData, lblShape.Left, lblShape.Top, lblShape.Width, lblShape.Height, 0, 0
      picImage.Picture = picImage.Image
      Set PicEx.Picture = picImage.Picture
   Else
      Set PicEx.Picture = Clipboard.GetData
      Set picImage.Picture = PicEx.Picture
   End If
End Sub

Private Sub mnuRank_Click(Index As Integer)
   MousePointer = vbHourglass
   PicEx.ApplyFilter eRankMedian + Index, 3 + RankFilterSize(Index) * 2
   MousePointer = vbDefault
End Sub

Private Sub mnuRemoveSelection_Click()
   ShowGrip False
End Sub

Private Sub mnuResize_Click()
   fResize.lWidth = picImage.ScaleX(PicEx.Picture.Width, vbHimetric, vbPixels)
   fResize.lHeight = picImage.ScaleY(PicEx.Picture.Height, vbHimetric, vbPixels)
   fResize.bUseResampling = PicEx.UseResampling
   fResize.Show vbModal
   PicEx.UseResampling = fResize.bUseResampling
   MousePointer = vbHourglass
   PicEx.ResizeImage fResize.lWidth, fResize.lHeight
   MousePointer = vbDefault
End Sub

Private Sub mnuRotate_Click(Index As Integer)
   MousePointer = vbHourglass
   PicEx.ProcessImage eRotate, , , Index
   MousePointer = vbDefault
End Sub

Private Sub mnuSave_Click()
   If Not fDlg Is Nothing Then Set fDlg = Nothing
   Set fDlg = New frmDlgExtra
   PicEx.SaveToFile , True, fDlg
   Set fDlg = Nothing
End Sub

Private Sub mnuSaveMSF_Click()
   If Not fDlg Is Nothing Then Set fDlg = Nothing
   Set fDlg = New frmDlgExtra
   PicEx.SaveToFileMSF , True, fDlg
   Set fDlg = Nothing
End Sub

Private Sub mnuSelect_Click()
   ShowGrip True
End Sub

Private Sub mnuStdFlt_Click(Index As Integer)
   MousePointer = vbHourglass
   PicEx.ApplyFilter Index, 3 + KernelFilterSize(Index) * 2, KernelFilterPower(Index)
   MousePointer = vbDefault
End Sub

Private Sub mnuUndo_Click()
   PicEx.Undo
   picImage.Picture = PicEx.Picture
End Sub

Private Sub nextb_Click()
   If Not fDlg Is Nothing Then Set fDlg = Nothing
   Set fDlg = New frmDlgExtra
   PicEx.SaveToFile , True, fDlg
   Set fDlg = Nothing
   frmPicViewer.picViewX.Picture = LoadPicture("test2.jpg")
   frmPicViewer.picCopy.Picture = LoadPicture("test2.jpg")
   frmTest.Hide
   'Unload Me
   frmCrop.Show
   'Unload Me
End Sub

Private Sub PicEx_ProgressChanged(ByVal nValue As Long)
   ProgressBar1.Value = nValue
   DoEvents
End Sub

Private Sub PicEx_ProgressEnd(ByVal nTime As Long)
   ProgressBar1.Value = 0
   lblProgress = "Complete in " & Format(nTime / 1000, "#0.000") & " sec."
   DoEvents
End Sub

Private Sub PicEx_ProgressInit(ByVal nMax As Long)
   ProgressBar1.Max = nMax
   lblProgress = "Processing image..."
   DoEvents
End Sub

Private Sub picContainer_Resize()
   Dim nMax As Long, dy As Long
   On Error Resume Next
   nMax = picImage.Height - picContainer.Height
   If nMax < 0 Then nMax = 0
   VScroll1.Max = nMax
   If nMax > 0 Then
      VScroll1.Visible = True
      picContainer.Width = ScaleWidth - VScroll1.Width
      VScroll1.Move picContainer.Left + picContainer.Width, picContainer.Top, VScroll1.Width, picContainer.Height
   Else
      VScroll1.Visible = False
   End If
   nMax = picImage.Width - picContainer.Width
   If nMax < 0 Then nMax = 0
   HScroll1.Max = nMax
   If nMax > 0 Then
      HScroll1.Visible = True
      picContainer.Height = ScaleHeight - HScroll1.Height - ProgressBar1.Height * 2
      HScroll1.Move picContainer.Left, picContainer.Top + picContainer.Height, picContainer.Width
      dy = HScroll1.Height
   Else
      HScroll1.Visible = False
   End If
   lblProgress.Move 0, picContainer.Height + ProgressBar1.Height / 2 + dy
   ProgressBar1.Move lblProgress.Width, lblProgress.Top, ScaleWidth - lblProgress.Width - 2000
End Sub

Private Sub HScroll1_Change()
   picImage.Left = -HScroll1.Value
End Sub

Private Sub picImage_Resize()
   picContainer_Resize
End Sub

Private Sub VScroll1_Change()
   picImage.Top = -VScroll1.Value
End Sub

Private Sub UpdateEditMenu()
   UpdateUndo
   mnuPaste.Enabled = (Clipboard.GetFormat(vbCFBitmap) Or Clipboard.GetFormat(vbCFMetafile) Or Clipboard.GetFormat(vbCFDIB))
   mnuCopy.Enabled = (PicEx.Picture <> 0)
   mnuCut.Enabled = (PicEx.Picture <> 0)
   mnuClear.Enabled = (PicEx.Picture <> 0)
   mnuCBClear.Enabled = mnuPaste.Enabled Or (Clipboard.GetFormat(vbCFText)) Or (Clipboard.GetFormat(vbCFLink)) Or (Clipboard.GetFormat(vbCFPalette))
   mnuSelect.Enabled = (PicEx.Picture <> 0) And Not lblShape.Visible
   mnuCrop.Enabled = (PicEx.Picture <> 0) And lblShape.Visible
   mnuRemoveSelection.Enabled = lblShape.Visible
End Sub

Private Sub UpdateImageMenu()
   Dim i As Integer
   For i = 0 To 1
       mnuRotate(i).Enabled = (PicEx.Picture <> 0) And Not lblShape.Visible
       mnuFlip(i).Enabled = (PicEx.Picture <> 0)
   Next i
   mnuAutoBalance.Enabled = (PicEx.Picture <> 0)
   mnuBalance.Enabled = (PicEx.Picture <> 0)
   mnuResize.Enabled = (PicEx.Picture <> 0) And Not lblShape.Visible
   mnuFlt.Enabled = (PicEx.Picture <> 0)
   mnuEff.Enabled = (PicEx.Picture <> 0)
   mnuClr.Enabled = (PicEx.Picture <> 0)
End Sub

Private Sub UpdateUndo()
   Set colOperations = PicEx.Operations
   mnuUndo.Enabled = PicEx.EnableUndo
   If colOperations.Count > 1 Then
      mnuUndo.Caption = "&Undo " & colOperations.Item(colOperations.Count)
   Else
      mnuUndo.Caption = "&Undo"
   End If
End Sub

Private Sub CentreForm(frm As Form)
   frm.Move Me.Left + Me.Width / 2 - frm.Width / 2, Me.Top + Me.Height / 2 - frm.Height / 2
End Sub

Public Sub DoBalance(idx As Integer, clrIdx As Integer, nValue As Long)
   Screen.MousePointer = vbHourglass
   PicEx.Balance idx, nValue, clrIdx
   Screen.MousePointer = vbDefault
End Sub

Public Sub UpdateChanges(bUpdate As Boolean)
    PicEx.UpdateChanges bUpdate
End Sub

'=============Sizing grip staff==============
Private Sub InitGrip()
   Dim i As Integer
   lblGrip(0).Width = GripSize
   lblGrip(0).Height = GripSize
   For i = 1 To 7
      Load lblGrip(i)
      lblGrip(i).MousePointer = i + 4 * Int((9 - i) / 4)
   Next i
   lblGrip(0).MousePointer = 8
   ShowGrip False
End Sub

Private Sub lblGrip_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      bSizing = True
      xStart = x: yStart = y
      lblShape.Enabled = False
   End If
End Sub

Private Sub lblGrip_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lft As Long, tp As Long, wdt As Long, hgt As Long
   If bSizing Then
      Select Case Index
         Case 0
              lft = lblShape.Left + x - xStart
              tp = lblShape.Top + y - yStart
              wdt = lblShape.Width - x + xStart
              hgt = lblShape.Height - y + yStart
         Case 1
              lft = lblShape.Left + x - xStart
              tp = lblShape.Top
              wdt = lblShape.Width - x + xStart
              hgt = lblShape.Height
         Case 2
              lft = lblShape.Left + x - xStart
              tp = lblShape.Top
              wdt = lblShape.Width - x + xStart
              hgt = lblShape.Height + y - yStart
         Case 3
              lft = lblShape.Left
              tp = lblShape.Top
              wdt = lblShape.Width
              hgt = lblShape.Height + y - yStart
         Case 4
              lft = lblShape.Left
              tp = lblShape.Top
              wdt = lblShape.Width + x - xStart
              hgt = lblShape.Height + y - yStart
         Case 5
              lft = lblShape.Left
              tp = lblShape.Top
              wdt = lblShape.Width + x - xStart
              hgt = lblShape.Height
         Case 6
              lft = lblShape.Left
              tp = lblShape.Top + y - yStart
              wdt = lblShape.Width + x - xStart
              hgt = lblShape.Height - y + yStart
         Case 7
              lft = lblShape.Left
              tp = lblShape.Top + y - yStart
              wdt = lblShape.Width
              hgt = lblShape.Height - y + yStart
      End Select
      If wdt < 0 Or hgt < 0 Or lft < 0 Or tp < 0 Or lft + wdt > picImage.Width Or tp + hgt > picImage.Height Then Exit Sub
      lblShape.Move lft, tp, wdt, hgt
      MoveGrips
   End If
End Sub

Private Sub lblGrip_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
   bSizing = False
   lblShape.Enabled = True
'   SetSelection True
End Sub

Private Sub lblShape_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = vbLeftButton Then
      bMoving = True
      xStart = x: yStart = y
      lblShape.MousePointer = 5
   ElseIf Button = vbRightButton Then
      UpdateEditMenu
      PopupMenu mnuEdit
   ElseIf Button = vbMiddleButton Then
      UpdateImageMenu
      PopupMenu mnuImage
   End If
End Sub

Private Sub lblShape_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim lft As Long, tp As Long
   If bMoving Then
      lft = lblShape.Left + x - xStart
      tp = lblShape.Top + y - yStart
      If lft <= 0 Then lft = 0
      If tp <= 0 Then tp = 0
      If lft > picImage.Width - lblShape.Width Then lft = picImage.Width - lblShape.Width
      If tp > picImage.Height - lblShape.Height Then tp = picImage.Height - lblShape.Height
      lblShape.Move lft, tp
      MoveGrips
   End If
End Sub

Private Sub lblShape_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   bMoving = False
   lblShape.MousePointer = 0
'   SetSelection True
End Sub

Private Sub MoveGrips()
   XGrip(0) = lblShape.Left - GripSize
   XGrip(1) = lblShape.Left + lblShape.Width / 2 - GripSize / 2
   XGrip(2) = lblShape.Left + lblShape.Width
   YGrip(0) = lblShape.Top - GripSize
   YGrip(1) = lblShape.Top + lblShape.Height / 2 - GripSize / 2
   YGrip(2) = lblShape.Top + lblShape.Height
   lblGrip(0).Move XGrip(0), YGrip(0)
   lblGrip(1).Move XGrip(0), YGrip(1)
   lblGrip(2).Move XGrip(0), YGrip(2)
   lblGrip(3).Move XGrip(1), YGrip(2)
   lblGrip(4).Move XGrip(2), YGrip(2)
   lblGrip(5).Move XGrip(2), YGrip(1)
   lblGrip(6).Move XGrip(2), YGrip(0)
   lblGrip(7).Move XGrip(1), YGrip(0)
End Sub

Private Sub ShowGrip(bShow As Boolean)
   Dim i As Integer
   lblShape.Move 100, 100, 600, 600
   lblShape.Visible = bShow
   For i = 0 To 7
      lblGrip(i).Visible = bShow
   Next i
   MoveGrips
'   SetSelection bShow
End Sub

'Private Sub SetSelection(Optional bSet As Boolean = True)
'   On Error Resume Next
'   If Not bSet Then
'      PicEx.SetSelectedArea 0, 0, 0, 0
'   Else
'      PicEx.SetSelectedArea lblShape.Left / Screen.TwipsPerPixelX, lblShape.Top / Screen.TwipsPerPixelY, lblShape.Width / Screen.TwipsPerPixelX, lblShape.Height / Screen.TwipsPerPixelY
'   End If
'End Sub
