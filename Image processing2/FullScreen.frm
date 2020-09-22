VERSION 5.00
Begin VB.Form frmFullScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5010
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   334
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   497
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar HScrollImg 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   6735
   End
   Begin VB.VScrollBar VScrollImg 
      Height          =   4815
      Left            =   6960
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Image imgFullScr 
      Height          =   4335
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   6015
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuNext 
         Caption         =   "&Next"
      End
      Begin VB.Menu mnuPrevious 
         Caption         =   "Previous"
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFitToScreen 
         Caption         =   "Fit to Screen"
      End
      Begin VB.Menu mnuZoomSel 
         Caption         =   "Zoom"
         Begin VB.Menu mnuZoom 
            Caption         =   "1x"
            Index           =   1
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "2x"
            Index           =   2
         End
         Begin VB.Menu mnuZoom 
            Caption         =   "3x"
            Index           =   3
         End
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmFullScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' You are free to copy, use and distribute the entire code of ImageBrowser as long as
'' you keep this copyright notice:
'' Author: Maurizio Fassina (maufass@tin.it)
'' This condition do not apply to small portion of code that you can
'' use freely
'' You CANNOT MODIFY THIS CODE AND DISTRIBUTE IT without an explicit
'' agreement of  the author Maurizio Fassina
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''
''  Form frmFullScreen
''
''  shows the use
''     - of a full screen form
''     - of resizing an image
''     - of shortcut menu
''''''''''''''''''''''''''''''''''

Option Explicit

''''''''''''''' Events ''''''''''''''''''''
Public Event QueryFileName(ByRef sFileName As String)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

''''''''''''''' Enum Image Size ''''''''''''
Enum eImageZoom
  eFit = 0
  eZoom1 = 1
  eZoom2 = 2
  eZoom3 = 3
End Enum
Dim st As String
Private iImageZoom As eImageZoom

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
     Me.Hide
  Else
    'delegates to parent to handle the event KeyDown
    RaiseEvent KeyDown(KeyCode, Shift)
    If KeyCode = vbKeyDown Or KeyCode = vbKeyUp Then
      LoadImage
    End If
  End If
 End Sub

''''''''''''''
'' Resizes ScrollBars
Private Sub Form_Resize()
   With HScrollImg
    .Top = Me.ScaleHeight - .Height
    .Width = Me.ScaleWidth - .Height
  End With
  With VScrollImg
    .Left = Me.ScaleWidth - .Width
    .Height = Me.ScaleHeight
  End With
  LoadImage
End Sub
Private Sub LoadImage()
  Dim sFileName As String
  On Error GoTo NoFile
   'BitBlt imgFullScr.Picture.hdc, 0, 0, frmPicViewer.picCopy.ScaleWidth, frmPicViewer.picCopy.ScaleHeight, frmPicViewer.picCopy.hdc, 0, 0, vbSrcCopy

   imgFullScr.Picture = frmPicViewer.picCopy 'LoadPicture(st)
   If iImageZoom = eFit Then
       ImageFit
   Else
       ImageZoom
   End If
   'imgFullScr.ToolTipText = frmFileDir.imgSmall.ToolTipText
   Exit Sub
NoFile:
  'MsgWarn "Unable to Open " + sFileName
  Unload Me
End Sub

Private Sub HScrollImg_Change()
 With HScrollImg
   If (.Max - .Value < .LargeChange) Then 'per fare stare tutta l'immagine
     imgFullScr.Left = .LargeChange - .Max
   Else
     imgFullScr.Left = -.Value
   End If
 End With
End Sub

Private Sub imgFullScr_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
   PopupMenu mnuPopup
 End If
End Sub
Private Sub ImageFit()
Dim rc As CRect
  imgFullScr.Visible = False
  'insert image in screen rect
  Set rc = New CRect
  rc.GetFormRect Me
  ImgInRect rc, imgFullScr, Me
  imgFullScr.Stretch = True
  VScrollImg.Visible = False
  HScrollImg.Visible = False
  imgFullScr.Visible = True
End Sub
''''''''''''''''''''''''''''''''''
'' Creates a zoomed image with scrollbars, if needed
''
Private Sub ImageZoom()
 With imgFullScr
   ''''''''''''
   '' Image size = Picture size * Zoom
   .Width = ScaleX(.Picture.Width * iImageZoom, vbHimetric, Me.ScaleMode)
   .Height = ScaleY(.Picture.Height * iImageZoom, vbHimetric, Me.ScaleMode)
   If .Width > Me.ScaleWidth Then
     .Left = 0
     HScrollImg.Max = .Width
     HScrollImg.Value = 0
     HScrollImg.SmallChange = Me.ScaleWidth / 15
     HScrollImg.LargeChange = Me.ScaleWidth
     HScrollImg.Visible = True
     Debug.Print "HZoom " & HScrollImg.Max & "  " & HScrollImg.LargeChange & " " & HScrollImg.SmallChange
   Else
     .Left = (Me.ScaleWidth - .Width) / 2
     HScrollImg.Visible = False
   End If
   If .Height > Me.ScaleHeight Then
     .Top = 0
     VScrollImg.Max = .Height
     VScrollImg.Value = 0
     VScrollImg.SmallChange = Me.ScaleHeight / 12
     VScrollImg.LargeChange = Me.ScaleHeight
     VScrollImg.Visible = True
     Debug.Print "VZoom " & VScrollImg.Max & "  " & VScrollImg.LargeChange & " " & VScrollImg.SmallChange
   Else
     .Top = (Me.ScaleHeight - .Height) / 2
     VScrollImg.Visible = False
   End If
  .Stretch = True
End With
End Sub

Private Sub mnuClose_Click()
  Me.Hide
End Sub

Private Sub mnuFitToScreen_Click()
 Dim I As Integer
  mnuFitToScreen.Checked = True
  For I = 1 To eZoom3
     mnuZoom.Item(I).Checked = False
  Next I
  iImageZoom = eFit
  ImageFit
End Sub

Private Sub mnuNext_Click()
  RaiseEvent KeyDown(vbKeyDown, 0)
  LoadImage
End Sub


Private Sub mnuPrevious_Click()
    RaiseEvent KeyDown(vbKeyUp, 0)
    LoadImage
End Sub


Private Sub mnuZoom_Click(Index As Integer)
 Dim I As Integer
  mnuFitToScreen.Checked = False
  For I = 1 To eZoom3
     mnuZoom.Item(I).Checked = (I = Index)
  Next I
  iImageZoom = Index
  ImageZoom
End Sub

Private Sub VScrollImg_Change()
 With VScrollImg
   If (.Max - .Value < .LargeChange) Then
     imgFullScr.Top = .LargeChange - .Max
   Else
     imgFullScr.Top = -.Value
   End If
 Debug.Print "VScroll  " & .Max & " " & .Value & " " & .LargeChange
 End With
End Sub
