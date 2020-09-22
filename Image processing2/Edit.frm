VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmEdit 
   Caption         =   "Edit"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9885
   LinkTopic       =   "Form3"
   ScaleHeight     =   501
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   659
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   5055
      Left            =   3840
      ScaleHeight     =   333
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   341
      TabIndex        =   30
      Top             =   840
      Width           =   5175
   End
   Begin VB.CommandButton Command19 
      Caption         =   "<-- Back"
      Height          =   375
      Left            =   7440
      TabIndex        =   24
      Top             =   6840
      Width           =   855
   End
   Begin VB.CommandButton Command18 
      Caption         =   "Next -->"
      Height          =   375
      Left            =   8640
      TabIndex        =   23
      Top             =   6840
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Unsharpness Mask"
      Height          =   1455
      Left            =   360
      TabIndex        =   22
      Top             =   5880
      Width           =   3015
      Begin VB.CommandButton Command20 
         Caption         =   "Invert"
         Height          =   375
         Left            =   960
         TabIndex        =   25
         Top             =   840
         Width           =   1095
      End
      Begin ComctlLib.Slider Slider4 
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   360
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   661
         _Version        =   327682
         Max             =   100
         SelStart        =   50
         Value           =   50
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Color Adjustment"
      Height          =   3375
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   3015
      Begin VB.CommandButton Command17 
         Caption         =   "+3"
         Height          =   375
         Left            =   2280
         TabIndex        =   21
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command16 
         Caption         =   "+3"
         Height          =   375
         Left            =   2280
         TabIndex        =   20
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton Command15 
         Caption         =   "+1"
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command14 
         Caption         =   "+1"
         Height          =   375
         Left            =   1800
         TabIndex        =   18
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton Command13 
         Caption         =   "-1"
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command12 
         Caption         =   "-1"
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton Command11 
         Caption         =   "-3"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command10 
         Caption         =   "-3"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   2760
         Width           =   375
      End
      Begin VB.CommandButton Command7 
         Caption         =   "-1"
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command9 
         Caption         =   "+3"
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command8 
         Caption         =   "+1"
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command6 
         Caption         =   "-3"
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   720
         Width           =   375
      End
      Begin ComctlLib.Slider Slider1 
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   661
         _Version        =   327682
         Max             =   100
         SelStart        =   50
         Value           =   50
      End
      Begin ComctlLib.Slider Slider2 
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   1320
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   661
         _Version        =   327682
         Max             =   100
         SelStart        =   50
         Value           =   50
      End
      Begin ComctlLib.Slider Slider3 
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   2280
         Width           =   2550
         _ExtentX        =   4498
         _ExtentY        =   661
         _Version        =   327682
         Max             =   100
         SelStart        =   50
         Value           =   50
      End
      Begin VB.Label Label3 
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   13
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   12
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Flip-Rotate"
      Height          =   1575
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   3015
      Begin VB.CommandButton Command1 
         Caption         =   "- 90"
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Flip Horizontal"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Flip Vertical"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "180"
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton Command2 
         Caption         =   "+ 90"
         Height          =   375
         Left            =   2160
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ILLUMINANT_A = 1
Const HALFTONE = 4
Private Type COLORADJUSTMENT
        caSize As Integer
        caFlags As Integer
        caIlluminantIndex As Integer
        caRedGamma As Integer
        caGreenGamma As Integer
        caBlueGamma As Integer
        caReferenceBlack As Integer
        caReferenceWhite As Integer
        caContrast As Integer
        caBrightness As Integer
        caColorfulness As Integer
        caRedGreenTint As Integer
End Type

Public bright_red As Single
Public bright_green As Single
Public bright_blue As Single
Public invert As Boolean
Dim w, h As Single
Dim wp1, hp1 As Single
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Byte
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
    ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal mDestWidth As Long, ByVal mDestHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal mSrcWidth As Long, _
    ByVal mSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function GetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As COLORADJUSTMENT) As Long
Private Declare Function SetColorAdjustment Lib "gdi32" (ByVal hdc As Long, lpca As COLORADJUSTMENT) As Long
Private Declare Function GetStretchBltMode Lib "gdi32" (ByVal hdc As Long) As Long


    Dim CA As COLORADJUSTMENT

Private Sub Command1_Click()   'Rotate -90
MousePointer = vbHourglass

SinA = -1
CosA = 0

'rotate the picture
    frmPicViewer.picCopy = LoadPicture()
    frmPicViewer.picCopy.Cls
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    frmPicViewer.picCopy.Width = h
    frmPicViewer.picCopy.Height = w
    For X = 1 To w
        For Y = 1 To h
          SetPixel frmPicViewer.picCopy.hdc, (X * CosA) - (Y * SinA), (X * SinA) + (Y * CosA) + w, GetPixel(frmPicViewer.picViewX.hdc, X, Y)
        Next
    Next
    
    frmPicViewer.picViewX.Width = h
    frmPicViewer.picViewX.Height = w
    frmPicViewer.picViewX.Refresh
    frmPicViewer.picViewX.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picViewX.hdc, 3)

    frmPicViewer.picCopy.Refresh
    Picture1.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    gh = SetStretchBltMode(Picture1.hdc, 3)
    Picture1.Width = hp1
    Picture1.Height = wp1
    gh = SetStretchBltMode(Picture1.hdc, 3)
    j = w: w = h: h = j
    j = wp1: wp1 = hp1: hp1 = j
    
    BitBlt frmPicViewer.picViewX.hdc, 0, 0, frmPicViewer.picViewX.ScaleWidth, frmPicViewer.picViewX.ScaleHeight, frmPicViewer.picCopy.hdc, 0, 0, vbSrcCopy
    
    mresult = StretchBlt(Picture1.hdc, 0, 0, wp1, hp1, frmPicViewer.picCopy.hdc, 0, 0, w, h, vbSrcCopy)
MousePointer = vbDefault

End Sub

Private Sub Command10_Click()
bright_green = bright_green - 3
Call imageupdate
End Sub

Private Sub Command11_Click()
bright_blue = bright_blue - 3
Call imageupdate
End Sub

Private Sub Command12_Click()
bright_green = bright_green - 1
Call imageupdate
End Sub

Private Sub Command13_Click()
bright_blue = bright_blue - 1
Call imageupdate
End Sub

Private Sub Command14_Click()
bright_green = bright_green + 1
Call imageupdate
End Sub

Private Sub Command15_Click()
bright_blue = bright_blue + 1
Call imageupdate
End Sub

Private Sub Command16_Click()
bright_green = bright_green + 3
Call imageupdate
End Sub

Private Sub Command17_Click()
bright_blue = bright_blue + 3
Call imageupdate
End Sub

Private Sub Command18_Click()   'Next
        Unload Me
        frmCrop.Show
End Sub

Private Sub Command19_Click()   'Back
        Unload Me
        frmPicViewer.Show
End Sub

Private Sub imageupdate()

   ' For X = 0 To frmPicViewer.picCopy.ScaleWidth
   '     For Y = 0 To frmPicViewer.picCopy.ScaleHeight
   '         r = (GetPixel(frmPicViewer.picViewX.hdc, X, Y) Mod 256)
   '         b = (Int(GetPixel(frmPicViewer.picViewX.hdc, X, Y) / 65536))
   '         g = ((GetPixel(frmPicViewer.picViewX.hdc, X, Y) - (b * 65536) - r) / 256)
   '         r = r * bright_red
   '         b = b * bright_blue
   '         g = g * bright_green
   '         If r > 255 Then r = 255
   '         If r < 0 Then r = 0
   '         If b > 255 Then b = 255
   '         If b < 0 Then b = 0
   '         If g > 255 Then g = 255
   '         If g < 0 Then g = 0
   '
   '         If invert Then
   '             r = 255 - r
   '             b = 255 - b
   '             g = 255 - g
   '         End If
   '         SetPixelV frmPicViewer.picCopy.hdc, X, Y, RGB(r, g, b)
   '     Next Y
   ' Next X
   ' frmPicViewer.picCopy.Refresh
   '
   '     Screen.MousePointer = vbHourglass
   '     picture1.Picture = LoadPicture()
   '     gh = SetStretchBltMode(picture1.hdc, 3)
   '     mresult = StretchBlt(picture1.hdc, 0, 0, wp1, hp1, frmPicViewer.picCopy.hdc, 0, 0, w, h, vbSrcCopy)
   '     Screen.MousePointer = vbDefault

   ' CA.caGreenGamma = bright_green * 1000
   ' CA.caBlueGamma = bright_blue * 1000
    CA.caRedGamma = bright_red * 1000
   
   
   
   
   
    'set the brightness to darkest
    'set a new illuminant
    'CA.caIlluminantIndex = bright_green * 10
    Print bright_red
    'check if the current StretchMode is set to HALFTONE
    'update the old coloradjustment
    SetColorAdjustment Picture1.hdc, CA
    'API uses pixels
    'Picture1.ScaleMode = vbPixels
    'copy the picture from Picture1 to Picture2
    m = StretchBlt(Picture1.hdc, 0, 0, wp1, hp1, frmPicViewer.picCopy.hdc, 0, 0, frmPicViewer.picCopy.ScaleWidth, frmPicViewer.picCopy.ScaleHeight, vbSrcCopy)




End Sub

Private Sub Command2_Click()   'Rotate +90
MousePointer = vbHourglass

SinA = 1
CosA = 0

'rotate the picture
    frmPicViewer.picCopy = LoadPicture()
    frmPicViewer.picCopy.Cls
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    frmPicViewer.picCopy.Width = h
    frmPicViewer.picCopy.Height = w
    For X = 1 To w
        For Y = 1 To h
          SetPixel frmPicViewer.picCopy.hdc, (X * CosA) - (Y * SinA) + h, (X * SinA) + (Y * CosA), GetPixel(frmPicViewer.picViewX.hdc, X, Y)
        Next
    Next
    
    frmPicViewer.picViewX.Width = h
    frmPicViewer.picViewX.Height = w
    frmPicViewer.picViewX.Refresh
    frmPicViewer.picViewX.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picViewX.hdc, 3)


    frmPicViewer.picCopy.Refresh
    Picture1.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    gh = SetStretchBltMode(Picture1.hdc, 3)
    Picture1.Width = hp1
    Picture1.Height = wp1
    gh = SetStretchBltMode(Picture1.hdc, 3)
    j = w: w = h: h = j
    j = wp1: wp1 = hp1: hp1 = j
    
    BitBlt frmPicViewer.picViewX.hdc, 0, 0, frmPicViewer.picViewX.ScaleWidth, frmPicViewer.picViewX.ScaleHeight, frmPicViewer.picCopy.hdc, 0, 0, vbSrcCopy

    mresult = StretchBlt(Picture1.hdc, 0, 0, wp1, hp1, frmPicViewer.picCopy.hdc, 0, 0, w, h, vbSrcCopy)
MousePointer = vbDefault
End Sub

Private Sub Command20_Click()
If invert = False Then invert = True Else invert = False
imageupdate
End Sub

Private Sub Command3_Click()   'Rotate 180
MousePointer = vbHourglass

SinA = 0
CosA = -1

'rotate the picture
    frmPicViewer.picCopy = LoadPicture()
    frmPicViewer.picCopy.Cls
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    frmPicViewer.picCopy.Width = w
    frmPicViewer.picCopy.Height = h
    For X = 1 To w
        For Y = 1 To h
          SetPixel frmPicViewer.picCopy.hdc, (X * CosA) - (Y * SinA) + w, (X * SinA) + (Y * CosA) + h, GetPixel(frmPicViewer.picViewX.hdc, X, Y)
        Next
    Next
    
    frmPicViewer.picViewX.Width = w
    frmPicViewer.picViewX.Height = h
    frmPicViewer.picViewX.Refresh
    frmPicViewer.picViewX.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picViewX.hdc, 3)


    frmPicViewer.picCopy.Refresh
    Picture1.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    gh = SetStretchBltMode(Picture1.hdc, 3)
    Picture1.Width = wp1
    Picture1.Height = hp1
    gh = SetStretchBltMode(Picture1.hdc, 3)
    
    BitBlt frmPicViewer.picViewX.hdc, 0, 0, frmPicViewer.picViewX.ScaleWidth, frmPicViewer.picViewX.ScaleHeight, frmPicViewer.picCopy.hdc, 0, 0, vbSrcCopy
    
    mresult = StretchBlt(Picture1.hdc, 0, 0, wp1, hp1, frmPicViewer.picCopy.hdc, 0, 0, w, h, vbSrcCopy)
MousePointer = vbDefault

End Sub

Private Sub Command4_Click()   'Flip Vertically
MousePointer = vbHourglass
    frmPicViewer.picCopy = LoadPicture()
    frmPicViewer.picCopy.Cls
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    frmPicViewer.picCopy.Width = w
    frmPicViewer.picCopy.Height = h
  
  For X = 0 To w
    For Y = 0 To h
      SetPixel frmPicViewer.picCopy.hdc, X, frmPicViewer.picCopy.Height - Y, GetPixel(frmPicViewer.picViewX.hdc, X, Y)
    Next
  Next
  
    frmPicViewer.picViewX.Width = w
    frmPicViewer.picViewX.Height = h
    frmPicViewer.picViewX.Refresh
    frmPicViewer.picViewX.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picViewX.hdc, 3)


    frmPicViewer.picCopy.Refresh
    Picture1.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    gh = SetStretchBltMode(Picture1.hdc, 3)
    Picture1.Width = wp1
    Picture1.Height = hp1
    gh = SetStretchBltMode(Picture1.hdc, 3)
      
    BitBlt frmPicViewer.picViewX.hdc, 0, 0, frmPicViewer.picViewX.ScaleWidth, frmPicViewer.picViewX.ScaleHeight, frmPicViewer.picCopy.hdc, 0, 0, vbSrcCopy
  
    mresult = StretchBlt(Picture1.hdc, 0, 0, wp1, hp1, frmPicViewer.picCopy.hdc, 0, 0, w, h, vbSrcCopy)
MousePointer = vbDefault
End Sub

Private Sub Command5_Click()    'Flip Horizontally
MousePointer = vbHourglass

    frmPicViewer.picCopy = LoadPicture()
    frmPicViewer.picCopy.Cls
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    frmPicViewer.picCopy.Width = w
    frmPicViewer.picCopy.Height = h

  For X = 0 To w
    For Y = 0 To h
      SetPixel frmPicViewer.picCopy.hdc, frmPicViewer.picCopy.Width - X, Y, GetPixel(frmPicViewer.picViewX.hdc, X, Y)
    Next
  Next
  
    frmPicViewer.picViewX.Width = w
    frmPicViewer.picViewX.Height = h
    frmPicViewer.picViewX.Refresh
    frmPicViewer.picViewX.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picViewX.hdc, 3)


    frmPicViewer.picCopy.Refresh
    Picture1.Picture = LoadPicture()
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    gh = SetStretchBltMode(Picture1.hdc, 3)
    Picture1.Width = wp1
    Picture1.Height = hp1
    gh = SetStretchBltMode(Picture1.hdc, 3)
        
    BitBlt frmPicViewer.picViewX.hdc, 0, 0, frmPicViewer.picViewX.ScaleWidth, frmPicViewer.picViewX.ScaleHeight, frmPicViewer.picCopy.hdc, 0, 0, vbSrcCopy
    
    mresult = StretchBlt(Picture1.hdc, 0, 0, wp1, hp1, frmPicViewer.picCopy.hdc, 0, 0, w, h, vbSrcCopy)
MousePointer = vbDefault
End Sub

Private Sub Command6_Click()
    bright_red = bright_red - 3
    Call imageupdate
End Sub

Private Sub Command7_Click()
    bright_red = bright_red - 1
    Call imageupdate
End Sub

Private Sub Command8_Click()
    bright_red = bright_red + 1
    Call imageupdate
End Sub

Private Sub Command9_Click()
    bright_red = bright_red + 3
    Call imageupdate
End Sub

Private Sub Form_Activate()
    gh = SetStretchBltMode(frmPicViewer.picCopy.hdc, 3)
    w = frmPicViewer.picCopy.Width
    h = frmPicViewer.picCopy.Height
    For pr = 1 To 100
        If pr / 100 * w > 350 Then Exit For
        If pr / 100 * h > 380 Then Exit For
    Next
    wp1 = pr / 100 * w
    hp1 = pr / 100 * h
    Picture1.Width = wp1
    Picture1.Height = hp1
    Picture1.Picture = LoadPicture()
    gh = SetStretchBltMode(Picture1.hdc, 3)
    
    mresult = StretchBlt(Picture1.hdc, 0, 0, wp1, hp1, frmPicViewer.picCopy.hdc, 0, 0, w, h, vbSrcCopy)
 
    invert = False
    frmEdit.Picture1.ScaleMode = 3
    bright_red = 1
    bright_blue = 1
    bright_green = 1

      
     DoEvents
    'retrieve the current color adjustment
    GetColorAdjustment Picture1.hdc, CA
    'initialize the type
    CA.caSize = Len(CA)
    If GetStretchBltMode(Picture1.hdc) <> HALFTONE Then
        'if it's not, set it to HALFTONE
        SetStretchBltMode Picture1.hdc, HALFTONE
    End If


   

End Sub

Private Sub Slider1_Click()
bright_red = frmEdit.Slider1.Value / 50
Call imageupdate
End Sub

Private Sub Slider2_Click()
bright_blue = frmEdit.Slider2.Value / 50
Call imageupdate
End Sub

Private Sub Slider3_Click()
bright_green = frmEdit.Slider3.Value / 50
Call imageupdate
End Sub

Private Sub Slider4_Click()
bright_red = frmEdit.Slider4.Value / 50
bright_green = frmEdit.Slider4.Value / 50
bright_blue = frmEdit.Slider4.Value / 50
Call imageupdate
End Sub
