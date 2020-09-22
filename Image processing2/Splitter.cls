VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

''''''''''''''''''''''''''''''''''''''''''
'' CSplitter is a class that supports operations for manipulating
'' (move and resizing) controls contained in a form and splitters.
'' Splitters can be moved (thereby resizing the controls) by mouse input.
''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Event SplitRect(ByVal Percent As Single)

''''''''''''''''' Types ''''''''''''''''
Public Enum SplitType
  spVertical = 1
  spOrizontal
End Enum

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

''''''''''''''' API Costants ''''''''''''''''''

Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BF_BOTTOM = &H8
Private Const BF_MIDDLE = &H800    ' fill the central part
Private Const BF_TOP = &H2
Private Const BF_RIGHT = &H4
Private Const BF_LEFT = &H1
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)

''''''''''''''' API Functions ''''''''''''''''''''''''''''
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long

''''''''''''''''''''' Variables ''''''''''''''''''''''''
Private Const SplitterHeight As Integer = 80
Private bMouseDrag As Boolean
Private bMouseIn As Boolean
Private oldXY As Single
Private TypeOfSplitter As SplitType
Private Percent As Single
Private recSplitter As New CRect  'dimensioni dello splitter
Private recWindow As New CRect   'finestra da dividere
Private pForm As Form

''''''''''''''''' Methods ''''''''''''''''''''
Private Sub Class_Initialize()
    Percent = 0
    bMouseDrag = False
    bMouseIn = False
    oldXY = 0
End Sub
Public Sub Init(ByVal frm As Form, _
                ByVal rcWindow As CRect, _
                ByVal Perc As Single, _
                Optional ByVal sptype As SplitType = spVertical)
Attribute Init.VB_Description = "init the Splitter Geometry"
    TypeOfSplitter = sptype
    Percent = Perc
    Set pForm = frm
    If frm.ClipControls = True Then
      Err.Raise vbObjectError, "CSplitter.CSplitter", frm.Name + ".ClipControl should be False."
    End If
    Resize rcWindow
End Sub
Public Sub Resize(ByVal rcWindow As CRect)
    recWindow.Copy rcWindow
    
    If TypeOfSplitter = spVertical Then
      recSplitter.SetRectWH Int((rcWindow.Right * Percent) - (SplitterHeight / 2)), _
                            rcWindow.Top, _
                            SplitterHeight, _
                            rcWindow.Height
    Else
      recSplitter.SetRectWH rcWindow.Left, _
                            Int((rcWindow.Bottom * Percent) - (SplitterHeight / 2)), _
                            rcWindow.Right, _
                            SplitterHeight
    End If
    RaiseEvent SplitRect(Percent)
End Sub
Public Sub MouseDown()
    If bMouseIn Then bMouseDrag = True
End Sub

Public Sub MouseMove(X As Single, Y As Single)
  Dim bInside As Boolean
  
  'If Button <> vbLeftButton Then Exit Sub
  With recSplitter
    bInside = (X > .Left) And (X < .Right) And (Y > .Top) And (Y < .Bottom)
  End With
  'first entry in splitter area
  If bInside And bMouseIn = False Then
      bMouseIn = True
      SetCapture pForm.hwnd
      If TypeOfSplitter = spVertical Then
        Screen.MouseIcon = LoadResPicture(103, vbResCursor)
      Else
        Screen.MouseIcon = LoadResPicture(102, vbResCursor)
      End If
      Screen.MousePointer = 99
   'mouse exit from splitter area
   ElseIf (Not bInside) And bMouseIn Then
      bMouseIn = False
      If Not bMouseDrag Then
        Screen.MousePointer = vbDefault
        ReleaseCapture
      End If
    End If
    If bMouseDrag Then
        If oldXY > 0 Then DrawLine oldXY
        If TypeOfSplitter = spVertical Then
          DrawLine X
          oldXY = X
        Else
          DrawLine Y
          oldXY = Y
        End If
    End If
End Sub

Public Sub MouseUp(X As Single, Y As Single)

  If bMouseDrag Then
    bMouseDrag = False
    DrawLine oldXY
    oldXY = 0
    If Not bMouseIn Then
      Screen.MousePointer = vbDefault
      ReleaseCapture
    End If
    If TypeOfSplitter = spVertical Then
      Percent = X / recWindow.Width
      MoveSplitRect Int(X), recSplitter.Top
    Else
      Percent = Y / recWindow.Height
      MoveSplitRect recSplitter.Left, Int(Y)
    End If
    RaiseEvent SplitRect(Percent)
    'MouseUp = True  ' the Form should resize  the inner rectangles
   End If
End Sub
Public Sub GetRect(ByVal what As Integer, rec As CRect)
 With recWindow
  If what = 1 And TypeOfSplitter = spVertical Then
    rec.SetRect .Left, .Top, recSplitter.Left, .Bottom
  ElseIf what = 1 And TypeOfSplitter = spOrizontal Then
      rec.SetRect .Left, .Top, .Right, recSplitter.Top
  ElseIf what = 2 And TypeOfSplitter = spVertical Then
      rec.SetRect recSplitter.Right, .Top, .Right, .Bottom
  ElseIf what = 2 And TypeOfSplitter = spOrizontal Then
    rec.SetRect .Left, recSplitter.Bottom, .Right, .Bottom
  End If
 End With

End Sub

Public Sub ResizeFrame(what As Integer, frm As Control)
 Dim rec As CRect
 Set rec = New CRect
 GetRect what, rec
 rec.SetControlRect frm
End Sub
Public Sub DEdge(Optional ByVal EdgeType As Long = EDGE_RAISED)  'draws the edge
  Dim rcs As RECT
  With recSplitter
    rcs.Left = .Left
    rcs.Bottom = .Bottom
    rcs.Top = .Top
    rcs.Right = .Right
    pForm.DrawWidth = 1
    pForm.DrawMode = vbCopyPen
    DrawEdge pForm.hdc, rcs, EdgeType, BF_RECT
    pForm.Line (.Right - 2 * Screen.TwipsPerPixelX, .Top)-(.Right - 2 * Screen.TwipsPerPixelX, .Bottom), RGB(127, 127, 127)
    pForm.Line (.Right - Screen.TwipsPerPixelX, .Top)-(.Right - Screen.TwipsPerPixelX, .Bottom), RGB(0, 0, 0)
  End With
End Sub
Private Sub MoveSplitRect(ByVal X As Long, ByVal Y As Long)
    recSplitter.MoveTo X, Y
End Sub

Private Sub DrawLine(xy As Single)
  With pForm
    .DrawStyle = vbSolid
    .DrawWidth = 3
    .DrawMode = vbMergePenNot
    If TypeOfSplitter = spVertical Then
      pForm.Line (Int(xy), recWindow.Top)-(Int(xy), recWindow.Bottom), 0
    Else
      pForm.Line (recWindow.Left, Int(xy))-(recWindow.Right, Int(xy)), 0
    End If
    .DrawMode = vbNop
  End With
End Sub

Private Sub Class_Terminate()
  Set recWindow = Nothing
  Set recSplitter = Nothing
End Sub
