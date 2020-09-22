VERSION 5.00
Begin VB.Form frmCrop 
   Caption         =   "Crop"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form3"
   ScaleHeight     =   477
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command12 
      Caption         =   "Update Image"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   6480
      Width           =   1455
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   3480
      ScaleHeight     =   107
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   185
      TabIndex        =   15
      Top             =   720
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.PictureBox picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3075
      Left            =   3480
      ScaleHeight     =   203
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   289
      TabIndex        =   14
      Top             =   720
      Width           =   4365
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Original Image"
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   7080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Crop Sizes"
      Height          =   5655
      Left            =   360
      TabIndex        =   2
      Top             =   600
      Width           =   1455
      Begin VB.CommandButton Command10 
         Caption         =   "30 x 45"
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   4920
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "20 x 25"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   4440
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "16 x 21"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "13 x 18"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "10 x 15"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   3000
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "9 x 13"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2520
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "6 x 9"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "4 x 6"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "3 x 4"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "2 x 3"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   855
      End
   End
   Begin VB.CommandButton Command19 
      Caption         =   "<-- Back"
      Height          =   375
      Left            =   12600
      TabIndex        =   1
      Top             =   10200
      Width           =   855
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Next -->"
      Height          =   375
      Left            =   13680
      TabIndex        =   0
      Top             =   10200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   4680
      TabIndex        =   16
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmCrop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, _
    ByVal y As Long, ByVal mDestWidth As Long, ByVal mDestHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal mSrcWidth As Long, _
    ByVal mSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, _
    ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Byte
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Dim x As Long
Dim y As Long
Dim pr As Long
Public hw
Dim X1 As Single, Y1 As Single, X2 As Single, Y2 As Single
Dim RegionFlag As Boolean
Dim W, H As Long
Dim wp1, hp1 As Single



Private Sub Command11_Click()
Picture2.Visible = False
picture1.Visible = True
Form_Activate
End Sub

Private Sub Command12_Click()

    picture1.Picture = LoadPicture()
    'gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    W = Picture2.ScaleWidth
    H = Picture2.ScaleHeight
    For pr = 1 To 100
        If pr / 100 * W > 750 Then Exit For
        If pr / 100 * H > 650 Then Exit For
    Next
    wp1 = pr / 100 * W
    hp1 = pr / 100 * H
    picture1.Width = wp1
    picture1.Height = hp1
    gh = SetStretchBltMode(picture1.hdc, 3)
    
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, Picture2.hdc, 0, 0, W, H, vbSrcCopy)
    frmCrop.Label1.Caption = "Scale " + Str(pr) + " %"
  
End Sub

Private Sub Command1_Click()
hw = 2 / 3
frmPicViewer.Text1.Text = 1
End Sub
Private Sub Command2_Click()
hw = 3 / 4
frmPicViewer.Text1.Text = 2
End Sub
Private Sub Command3_Click()
hw = 4 / 6
frmPicViewer.Text1.Text = 3
End Sub
Private Sub Command4_Click()
hw = 6 / 9
frmPicViewer.Text1.Text = 4
End Sub
Private Sub Command5_Click()
hw = 9 / 13
frmPicViewer.Text1.Text = 5
End Sub
Private Sub Command6_Click()
hw = 10 / 15
frmPicViewer.Text1.Text = 6
End Sub
Private Sub Command7_Click()
hw = 13 / 18
frmPicViewer.Text1.Text = 7
End Sub
Private Sub Command8_Click()
hw = 16 / 21
frmPicViewer.Text1.Text = 8
End Sub
Private Sub Command9_Click()
hw = 20 / 25
frmPicViewer.Text1.Text = 9
End Sub
Private Sub Command10_Click()
hw = 30 / 45
frmPicViewer.Text1.Text = 10
End Sub

Private Sub Form_Activate()
   'frmPicViewer.picViewX.Picture = LoadPicture("test2.jpg")
   'frmPicViewer.picCopy.Picture = LoadPicture("test2.jpg")
    hw = 2 / 3
    frmPicViewer.Text1.Text = 1
    picture1.Picture = LoadPicture()
    Picture2.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    W = frmPicViewer.picCopy.ScaleWidth
    H = frmPicViewer.picCopy.ScaleHeight
    For pr = 1 To 100
        If pr / 100 * W > 750 Then Exit For
        If pr / 100 * H > 650 Then Exit For
    Next
    wp1 = pr / 100 * W
    hp1 = pr / 100 * H
    
    picture1.Width = wp1
    picture1.Height = hp1
    gh = SetStretchBltMode(picture1.hdc, 3)
    Picture2.Width = W
    Picture2.Height = H
    gh = SetStretchBltMode(Picture2.hdc, 3)
    
    mresult = StretchBlt(Picture2.hdc, 0, 0, W, H, frmPicViewer.picCopy.hdc, 0, 0, W, H, vbSrcCopy)
    mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, frmPicViewer.picCopy.hdc, 0, 0, W, H, vbSrcCopy)
    frmCrop.Label1.Caption = "Scale " + Str(pr) + " %"
    picture1.Refresh

End Sub
Private Sub Command18_Click()  'Next
    
    W = Picture2.Width
    H = Picture2.Height
    'Picture2.Visible = True
    frmPicViewer.picViewX.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picViewX.hdc, 3)
    frmPicViewer.picViewX.Width = W
    frmPicViewer.picViewX.Height = H
    
    frmPicViewer.picCopy.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    frmPicViewer.picCopy.Width = W
    frmPicViewer.picCopy.Height = H
    
    BitBlt frmPicViewer.picViewX.hdc, 0, 0, W, H, Picture2.hdc, 0, 0, vbSrcCopy
    BitBlt frmPicViewer.picCopy.hdc, 0, 0, W, H, Picture2.hdc, 0, 0, vbSrcCopy
    
    frmPicViewer.picCopy.Refresh
    frmPicViewer.picViewX.Refresh
 '       Print frmPicViewer.picViewX.Width, frmPicViewer.picViewX.Height
    'BitBlt picture1.hdc, 0, 0, W, H, frmPicViewer.picCopy.hdc, 0, 0, vbSrcCopy
  'picture1.Refresh
    frmCrop.Hide
    'frmPicViewer.Show
    'frmdecomp.Show
    frmfullviewer.Show
    
End Sub

Private Sub Command19_Click()   'Back
        
   W = Picture2.Width
   H = Picture2.Height
   
    frmPicViewer.picViewX.Width = W
    frmPicViewer.picViewX.Height = H
    frmPicViewer.picViewX.Refresh
    frmPicViewer.picViewX.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picViewX.hdc, 3)
    
    frmPicViewer.picCopy.Width = W
    frmPicViewer.picCopy.Height = H
    frmPicViewer.picCopy.Refresh
    frmPicViewer.picCopy.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    
    BitBlt frmPicViewer.picViewX.hdc, 0, 0, W, H, Picture2.hdc, 0, 0, vbSrcCopy
    BitBlt frmPicViewer.picCopy.hdc, 0, 0, W, H, Picture2.hdc, 0, 0, vbSrcCopy
        
    Unload Me
    frmTest.Show '0, Me
End Sub
Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
         RegionFlag = True
         picture1.DrawMode = vbInvert
         X1 = x: X2 = x: Y1 = y: Y2 = y
         
         picture1.Cls
         gh = SetStretchBltMode(picture1.hdc, 3)
         mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, frmPicViewer.picCopy.hdc, 0, 0, W, H, vbSrcCopy)
         picture1.Refresh
         picture1.Line (x, y)-(x, y), , B
    End If
End Sub


Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not RegionFlag Then
         Exit Sub
    End If
    picture1.Line (X1, Y1)-(X2, Y2), , B
    X2 = x
    Y2 = y
    ww = Abs(X2 - X1)
    hh = Abs(Y2 - Y1)
    If ww <> 0 Then
        If hh / ww > hw Then
            X2 = X1 + 3 / 4 * (Y2 - Y1)
        Else
            Y2 = Y1 + 4 / 3 * (X2 - X1)
        End If
    End If
    picture1.Line (X1, Y1)-(X2, Y2), , B
End Sub


Private Sub picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not RegionFlag Then
        Exit Sub
    Else
        RegionFlag = False
        picture1.DrawMode = vbCopyPen
    End If
    
    gh = SetStretchBltMode(picture1.hdc, 3)
    picture1.ScaleMode = 3
    Picture2.ScaleMode = 3
    gh = SetStretchBltMode(Picture2.hdc, 3)
    X1 = X1 * 100 / pr
    X2 = X2 * 100 / pr
    Y1 = Y1 * 100 / pr
    Y2 = Y2 * 100 / pr
    
    Picture2.Width = Abs(X2 - X1)
    Picture2.Height = Abs(Y2 - Y1)
    Picture2 = LoadPicture()
    mresult = StretchBlt(Picture2.hdc, 0, 0, Picture2.Width, Picture2.Height, frmPicViewer.picCopy.hdc, X1, Y1, (X2 - X1), (Y2 - Y1), vbSrcCopy)
End Sub

