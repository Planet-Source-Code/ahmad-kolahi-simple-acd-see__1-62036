Attribute VB_Name = "Globals"
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
''  Module Globals
''
''  shows the use
''     - of registry in windows 95/98
''     - of Windows API calls
''''''''''''''''''''''''''''''


Option Explicit

''''''''''''''''''''''
''  Public variables
''
'Public dbTTip As Database  ' tooltip's database
'Public rsTTip As Recordset   ' tooltip's recordset


''''''''''''''''''''''''''
'' References to Resource Strings
''
Public Const idProgName As Long = 101

''''''''''''''''''''''''''''''''
'' Windows API Declarations
''
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Const REG_SZ = 1&
Const ERROR_SUCCESS = 0&
Const KEY_QUERY_VALUE = &H1&
Const KEY_SET_VALUE = &H2&
Const KEY_CREATE_SUB_KEY = &H4&
Const KEY_ENUMERATE_SUB_KEYS = &H8&
Const KEY_NOTIFY = &H10&
Const KEY_CREATE_LINK = &H20&
Const READ_CONTROL = &H20000
Const WRITE_DAC = &H40000
Const WRITE_OWNER = &H80000
Const SYNCHRONIZE = &H100000
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL
Const STANDARD_RIGHTS_EXECUTE = READ_CONTROL
Const KEY_READ = STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY
Const KEY_WRITE = STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY
Const KEY_EXECUTE = KEY_READ

''''''''''''''''''''''''''''''''''''''''''
''    Resizing
Public Sub ImgInRect(rc As CRect, img As Image, frm As Form)
 Dim inRec As CRect
    Set inRec = New CRect
    inRec.SetRectWH 0, 0, frm.ScaleX(img.Picture.Width, vbHimetric, frm.ScaleMode), _
                         frm.ScaleY(img.Picture.Height, vbHimetric, frm.ScaleMode)
    RectInRect rc, inRec
    inRec.SetControlRect img
End Sub

''''''''''''''''''''''''''''''''
'' RectInRect resizes inRec to fit in outRec without changing
'' the proportions of inRec
Private Sub RectInRect(ByRef outRec As CRect, inRec As CRect)
   Dim ratio As Double
    If (inRec.Height <= 0 Or inRec.Width <= 0) Then
      Exit Sub
    End If
    ratio = inRec.Height / inRec.Width
    While (inRec.Width > outRec.Width Or inRec.Height > outRec.Height)
      If (inRec.Width <= outRec.Width And inRec.Height >= outRec.Height) Then
        inRec.Width = Int(outRec.Height / ratio)
        inRec.Height = outRec.Height
      ElseIf (inRec.Width >= outRec.Width And inRec.Height <= outRec.Height) Then
        inRec.Width = outRec.Width
        inRec.Height = Int(outRec.Width * ratio)
      Else
       inRec.Width = outRec.Width
       inRec.Height = outRec.Width * ratio
      End If
    Wend
    inRec.Reposition rCenter, outRec
End Sub

''''''''''''''''''''''''''''''''''''''''''''''
''   Registry
Public Function GetStringValue(ByVal MainKeyHandle As Long, SubKey As String, Entry As String)
 Dim rtn As Long, lBufferSize As Long, hKey As Long
 Dim sBuffer As String
 Dim iPos As Integer

   rtn = RegOpenKeyEx(MainKeyHandle, SubKey, 0, KEY_READ, hKey) 'open the key
   If rtn = ERROR_SUCCESS Then 'if the key could be opened then
      sBuffer = Space(255)     'make a buffer
      lBufferSize = Len(sBuffer)
      rtn = RegQueryValueEx(hKey, Entry, 0, REG_SZ, sBuffer, lBufferSize) 'get the value from the registry
      If rtn = ERROR_SUCCESS Then 'if the value could be retreived then
         rtn = RegCloseKey(hKey)  'close the key
         sBuffer = Trim(sBuffer)
         iPos = InStr(sBuffer, Chr(0))
         If iPos <> 0 Then
            sBuffer = Left(sBuffer, iPos - 1)
        End If
         
         GetStringValue = sBuffer 'return the value to the user
      Else                        'otherwise, if the value couldnt be retreived
         GetStringValue = "Error" 'return Error to the user
      End If
    End If
End Function

''''''''''''''''''''''''''''''
''  Simple Utility Message Box
Public Sub MsgWarn(sMsg As String)
  MsgBox sMsg, vbOKOnly + vbExclamation, LoadResString(idProgName)
End Sub

