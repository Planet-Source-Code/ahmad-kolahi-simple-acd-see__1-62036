VERSION 5.00
Begin VB.Form frmfullviewer 
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
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   1200
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   3960
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   8
      Top             =   2160
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   2640
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   1200
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   480
      Top             =   4440
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   3960
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   5
      Top             =   1200
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   2640
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   1200
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   4080
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   2640
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   1200
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   885
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
Attribute VB_Name = "frmfullviewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Public Event QueryFileName(ByRef sFileName As String)
'Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
    ByVal y As Long, ByVal mDestWidth As Long, ByVal mDestHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal mSrcWidth As Long, _
    ByVal mSrcHeight As Long, ByVal dwRop As Long) As Long
Dim frm As Single
Dim wp1, hp1, W, H As Long

Private Sub Form_Activate()
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    bb = val(frmPicViewer.Text1.Text)
    W = frmPicViewer.picCopy.Width
    H = frmPicViewer.picCopy.Height
    
    If bb = 1 Or bb = 2 Then
        picture1.Visible = True
        Picture2.Visible = True
        Picture3.Visible = True
        Picture4.Visible = True
        Picture5.Visible = True
        Picture6.Visible = True
        'For pr = 1 To 100
        '    If pr / 100 * W > 236 Then Exit For
        '    If pr / 100 * H > 315 Then Exit For
        'Next
        'wp1 = pr / 100 * W
        'hp1 = pr / 100 * H
        wp1 = 236
        hp1 = 315
        picture1.Left = 79: Picture2.Left = 394: Picture3.Left = 709
        Picture4.Left = 79: Picture5.Left = 394: Picture6.Left = 709
        picture1.Top = 28: Picture2.Top = 28: Picture3.Top = 28
        Picture4.Top = 371: Picture5.Top = 371: Picture6.Top = 371
        picture1.Width = wp1: Picture2.Width = wp1: Picture3.Width = wp1
        Picture4.Width = wp1: Picture5.Width = wp1: Picture6.Width = wp1
        picture1.Height = hp1: Picture2.Height = hp1: Picture3.Height = hp1
        Picture4.Height = hp1: Picture5.Height = hp1: Picture6.Height = hp1
        picture1.Picture = LoadPicture()
        Picture2.Picture = LoadPicture()
        Picture3.Picture = LoadPicture()
        Picture4.Picture = LoadPicture()
        Picture5.Picture = LoadPicture()
        Picture6.Picture = LoadPicture()
        gh = SetStretchBltMode(picture1.hdc, 3)
        gh = SetStretchBltMode(Picture2.hdc, 3)
        gh = SetStretchBltMode(Picture3.hdc, 3)
        gh = SetStretchBltMode(Picture4.hdc, 3)
        gh = SetStretchBltMode(Picture5.hdc, 3)
        gh = SetStretchBltMode(Picture6.hdc, 3)
        '**************************************
        Picture7.Left = 0
        Picture7.Top = 0
        Picture7.Width = wp1
        Picture7.Height = hp1
        Picture7.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.Picture7.hdc, 3)
        Picture8.Left = 0
        Picture8.Top = 0
        Picture8.Width = wp1
        Picture8.Height = hp1
        Picture8.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.Picture8.hdc, 3)
        Picture9.Left = 0
        Picture9.Top = 0
        Picture9.Width = wp1
        Picture9.Height = hp1
        Picture9.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.Picture9.hdc, 3)
        Picture10.Left = 0
        Picture10.Top = 0
        Picture10.Width = wp1
        Picture10.Height = hp1
        Picture10.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.Picture10.hdc, 3)
        '***********************************************
       For x = 0 To W - 1 Step 2
            For y = 0 To H - 1 Step 2
                SetPixelV Picture7.hdc, x / 2, y / 2, GetPixel(frmPicViewer.picCopy.hdc, x, y)
                SetPixelV Picture8.hdc, x / 2, y / 2, GetPixel(frmPicViewer.picCopy.hdc, x + 1, y)
                SetPixelV Picture9.hdc, x / 2, y / 2, GetPixel(frmPicViewer.picCopy.hdc, x, y + 1)
                SetPixelV Picture10.hdc, x / 2, y / 2, GetPixel(frmPicViewer.picCopy.hdc, x + 1, y + 1)
            Next
        Next
        Picture7.Refresh
        Picture8.Refresh
        Picture9.Refresh
        Picture10.Refresh
        frm = 0.3
        '*************************
        
    End If
    
    If bb = 4 Or bb = 3 Then
        picture1.Visible = True
        Picture2.Visible = True
        'For pr = 1 To 100
        '    If pr / 100 * w > 473 Then Exit For
        '    If pr / 100 * h > 768 Then Exit For
        'Next
        'wp1 = pr / 100 * w
        'hp1 = pr / 100 * h
        wp1 = 473
        hp1 = 768
        picture1.Left = 26: Picture2.Left = 525
        picture1.Top = 0: Picture2.Top = 0
        picture1.Width = wp1: Picture2.Width = wp1
        picture1.Height = hp1: Picture2.Height = hp1
        picture1.Picture = LoadPicture()
        Picture2.Picture = LoadPicture()
        gh = SetStretchBltMode(picture1.hdc, 3)
        gh = SetStretchBltMode(Picture2.hdc, 3)
        '**************************************
        Picture3.Left = 0
        Picture3.Top = 0
        Picture3.Width = wp1
        Picture3.Height = hp1
        Picture3.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.Picture3.hdc, 3)
        Picture4.Left = 0
        Picture4.Top = 0
        Picture4.Width = wp1
        Picture4.Height = hp1
        Picture4.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.Picture4.hdc, 3)
        Picture5.Left = 0
        Picture5.Top = 0
        Picture5.Width = wp1
        Picture5.Height = hp1
        Picture5.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.Picture5.hdc, 3)
        Picture6.Left = 0
        Picture6.Top = 0
        Picture6.Width = wp1
        Picture6.Height = hp1
        Picture6.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.Picture6.hdc, 3)
        '***********************************************
       For x = 0 To W - 1 Step 2
            For y = 0 To H - 1 Step 2
                SetPixelV Picture3.hdc, x / 2, y / 2, GetPixel(frmPicViewer.picCopy.hdc, x, y)
                SetPixelV Picture4.hdc, x / 2, y / 2, GetPixel(frmPicViewer.picCopy.hdc, x + 1, y)
                SetPixelV Picture5.hdc, x / 2, y / 2, GetPixel(frmPicViewer.picCopy.hdc, x, y + 1)
                SetPixelV Picture6.hdc, x / 2, y / 2, GetPixel(frmPicViewer.picCopy.hdc, x + 1, y + 1)
            Next
        Next
        Picture3.Refresh
        Picture4.Refresh
        Picture5.Refresh
        Picture6.Refresh
        frm = 0.4
        
        'mresult = StretchBlt(frmfullviewer.picture1.hdc, 0, 0, wp1, hp1, frmPicViewer.picCopy.hdc, 0, 0, w, h, vbSrcCopy)
        'mresult = StretchBlt(frmfullviewer.Picture2.hdc, 0, 0, wp1, hp1, frmPicViewer.picCopy.hdc, 0, 0, w, h, vbSrcCopy)
    End If
    
    If bb > 4 Then
        picture1.Visible = True
        'For pr = 1 To 100
        '    If pr / 100 * w > 1024 Then Exit For
        '    If pr / 100 * h > 768 Then Exit For
        'Next
        'wp1 = pr / 100 * w
        'hp1 = pr / 100 * h
        wp1 = 1024
        hp1 = 768
        
        Picture3.Left = 0
        Picture3.Top = 0
        Picture3.Width = wp1
        Picture3.Height = hp1
        Picture3.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.Picture3.hdc, 3)
        Picture4.Left = 0
        Picture4.Top = 0
        Picture4.Width = wp1
        Picture4.Height = hp1
        Picture4.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.Picture4.hdc, 3)
        Picture5.Left = 0
        Picture5.Top = 0
        Picture5.Width = wp1
        Picture5.Height = hp1
        Picture5.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.Picture5.hdc, 3)
        Picture6.Left = 0
        Picture6.Top = 0
        Picture6.Width = wp1
        Picture6.Height = hp1
        Picture6.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.Picture6.hdc, 3)
        
        picture1.Left = 0
        picture1.Top = 0
        picture1.Width = wp1
        picture1.Height = hp1
        picture1.Picture = LoadPicture()
        gh = SetStretchBltMode(frmfullviewer.picture1.hdc, 3)
        
        
        For x = 0 To W - 1 Step 2
            For y = 0 To H - 1 Step 2
                SetPixelV Picture3.hdc, x / 2, y / 2, GetPixel(frmPicViewer.picCopy.hdc, x, y)
                SetPixelV Picture4.hdc, x / 2, y / 2, GetPixel(frmPicViewer.picCopy.hdc, x + 1, y)
                SetPixelV Picture5.hdc, x / 2, y / 2, GetPixel(frmPicViewer.picCopy.hdc, x, y + 1)
                SetPixelV Picture6.hdc, x / 2, y / 2, GetPixel(frmPicViewer.picCopy.hdc, x + 1, y + 1)
            Next
        Next
        Picture3.Refresh
        Picture4.Refresh
        Picture5.Refresh
        Picture6.Refresh
        frm = 0.5
    End If

End Sub

Private Sub Form_Click()
    picture1.Visible = False
    Picture2.Visible = False
    Picture3.Visible = False
    Picture4.Visible = False
    Picture5.Visible = False
    Picture6.Visible = False
    Me.Hide
    frmPicViewer.Show
End Sub

Private Sub Timer1_Timer()
If frm = 0.5 Then
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture3.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    picture1.Refresh
End If
If frm = 1.5 Then
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture4.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    picture1.Refresh
End If
If frm = 2.5 Then
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture5.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    picture1.Refresh
End If
If frm = 3.5 Then
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture6.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    picture1.Refresh
End If
'*****************************************************************************************
If frm = 0.4 Then
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture3.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture2.hdc, 0, 0, wp1, hp1, Picture3.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    picture1.Refresh
    Picture2.Refresh
End If
If frm = 1.4 Then
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture4.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture2.hdc, 0, 0, wp1, hp1, Picture4.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    picture1.Refresh
    Picture2.Refresh
End If
If frm = 2.4 Then
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture5.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture2.hdc, 0, 0, wp1, hp1, Picture5.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    picture1.Refresh
    Picture2.Refresh
End If
If frm = 3.4 Then
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture6.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture2.hdc, 0, 0, wp1, hp1, Picture6.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    picture1.Refresh
    Picture2.Refresh
End If

'*****************************************************************************************
If frm = 0.3 Then
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture7.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture2.hdc, 0, 0, wp1, hp1, Picture7.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture3.hdc, 0, 0, wp1, hp1, Picture7.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture4.hdc, 0, 0, wp1, hp1, Picture7.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture5.hdc, 0, 0, wp1, hp1, Picture7.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture6.hdc, 0, 0, wp1, hp1, Picture7.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    picture1.Refresh
    Picture2.Refresh
    Picture3.Refresh
    Picture4.Refresh
    Picture5.Refresh
    Picture6.Refresh
End If
If frm = 1.3 Then
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture8.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture2.hdc, 0, 0, wp1, hp1, Picture8.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture3.hdc, 0, 0, wp1, hp1, Picture8.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture4.hdc, 0, 0, wp1, hp1, Picture8.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture5.hdc, 0, 0, wp1, hp1, Picture8.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture6.hdc, 0, 0, wp1, hp1, Picture8.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    picture1.Refresh
    Picture2.Refresh
    Picture3.Refresh
    Picture4.Refresh
    Picture5.Refresh
    Picture6.Refresh
End If
If frm = 2.3 Then
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture9.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture2.hdc, 0, 0, wp1, hp1, Picture9.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture3.hdc, 0, 0, wp1, hp1, Picture9.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture4.hdc, 0, 0, wp1, hp1, Picture9.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture5.hdc, 0, 0, wp1, hp1, Picture9.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture6.hdc, 0, 0, wp1, hp1, Picture9.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    picture1.Refresh
    Picture2.Refresh
    Picture3.Refresh
    Picture4.Refresh
    Picture5.Refresh
    Picture6.Refresh
End If
If frm = 3.3 Then
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture10.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture2.hdc, 0, 0, wp1, hp1, Picture10.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture3.hdc, 0, 0, wp1, hp1, Picture10.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture4.hdc, 0, 0, wp1, hp1, Picture10.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture5.hdc, 0, 0, wp1, hp1, Picture10.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    mresult = StretchBlt(Picture6.hdc, 0, 0, wp1, hp1, Picture10.hdc, 0, 0, Int(W / 2), Int(H / 2), vbSrcCopy)
    picture1.Refresh
    Picture2.Refresh
    Picture3.Refresh
    Picture4.Refresh
    Picture5.Refresh
    Picture6.Refresh
End If
'*****************************************************************************************

If frm > 4 Then
    Unload Me
    Unload frmPicViewer
    Unload frmTest
    Unload frmCrop
    Close
    'frmfullviewer.Hide
    'frmPicViewer.Show
End If
frm = frm + 1


End Sub
