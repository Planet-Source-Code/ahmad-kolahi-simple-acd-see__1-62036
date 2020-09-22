VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Classe rettangolo"
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' You are free to copy, use and distribute the entire code of ImageBrowser as long as
'' you keep this copyright notice:
'' Author: Maurizio Fassina (maufass@tin.it)
'' This condition do not apply to small portion of code that you can
'' use freely
'' You CANNOT MODIFY THIS CODE AND DISTRIBUTE IT without an explicit
'' agreement of  the author Maurizio Fassina
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''
'' Class Rect is a very simple, usefull (and classical) class
''
''''''''''''''''''''''''''''''''''
Option Explicit

Public Left As Long
Public Top As Long
Public Right As Long
Public Bottom As Long
Public Enum rPosition
    rLeftTop = 1
    rRightTop
    rLeftBottom
    rRightBottom
    rCenter
End Enum

Private Sub Class_Initialize()
    Left = 0
    Top = 0
    Right = 0
    Bottom = 0
End Sub

Public Property Get Height() As Long
    Height = (Bottom - Top)
End Property
Public Property Let Height(ByVal lNewHeight As Long)
  Bottom = Top + lNewHeight
End Property

Public Property Get Width() As Long
    Width = (Right - Left)
End Property
Public Property Let Width(ByVal lNewWidth As Long)
  Right = Left + lNewWidth
End Property
Public Sub Copy(Other As CRect)
  Left = Other.Left
  Right = Other.Right
  Top = Other.Top
  Bottom = Other.Bottom
End Sub
Public Sub SetRect(ByVal lLeft As Long, ByVal lTop As Long, ByVal lRight As Long, ByVal lBottom As Long)
  Left = lLeft
  Top = lTop
  Right = lRight
  Bottom = lBottom
End Sub

Public Sub SetRectWH(ByVal lLeft As Long, ByVal lTop As Long, ByVal lWidth As Long, ByVal lHeight As Long)
  Left = lLeft
  Top = lTop
  Right = Left + lWidth
  Bottom = Top + lHeight
End Sub


Public Sub Inflate(ByVal dx As Long, ByVal dy As Long)
    Left = Left + dx
    Right = Right + dx
    Top = Top + dy
    Bottom = Bottom + dy
End Sub


Public Sub GetFormRect(frm As Form)
        SetRect frm.ScaleLeft, frm.ScaleTop, frm.ScaleWidth, frm.ScaleHeight
End Sub

Public Sub GetControlRect(ctrl As Control)
        SetRectWH ctrl.Left, ctrl.Top, ctrl.Width, ctrl.Height
End Sub
Public Sub MoveControlTo(ctrl As Control, ByVal X As Long, ByVal Y As Long)
    GetControlRect ctrl
    MoveTo X, Y
    SetControlRect ctrl
End Sub
Public Sub MoveTo(ByVal lX As Long, ByVal lY As Long)
    Dim lW As Long, lH As Long
    lW = Width
    lH = Height
    Left = lX
    Top = lY
    Right = lX + lW
    Bottom = lY + lH
End Sub

Public Sub SetControlRect(ctrl As Control)
    ctrl.Left = Left
    ctrl.Top = Top
    ctrl.Width = Width
    ctrl.Height = Height
End Sub

Public Sub Reposition(ByVal pos As rPosition, ByVal outerRec As CRect)
    Select Case pos
        Case rLeftTop
          MoveTo outerRec.Left, outerRec.Top
        Case rRightTop
          MoveTo outerRec.Left + (outerRec.Width - Width), outerRec.Top
        Case rLeftBottom
          MoveTo outerRec.Left, outerRec.Top + (outerRec.Height - Height)
        Case rRightBottom
           MoveTo outerRec.Left + (outerRec.Width - Width), outerRec.Top + (outerRec.Height - Height)
        Case rCenter
           MoveTo outerRec.Left + (outerRec.Width - Width) / 2, outerRec.Top + (outerRec.Height - Height) / 2
    End Select
End Sub

Public Sub Normalize()
Dim Temp As Long
    If Left > Right Then
       Temp = Left
       Left = Right
       Right = Temp
    End If
    If Top > Bottom Then
       Temp = Top
       Top = Bottom
       Bottom = Temp
    End If
End Sub


