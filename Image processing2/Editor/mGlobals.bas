Attribute VB_Name = "mGlobals"

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Enum eBPP
    BlackWhite = 0
    GrayScale = 1
    Palette_256 = 2
    HighColor_16 = 3
    TrueColor_24 = 4
    TrueColor_32 = 5
End Enum

Public Const FilterSettingFile As String = "FltSet.dat"

Public KernelFilterSize(3) As Long
Public KernelFilterType(3) As Long
Public KernelFilterPower(3) As Long
Public RankFilterSize(2) As Long
Public EnhancedFilterSize(2) As Long
Public CustomFilterSize As Long
Public CustomFilter() As Long

Public bAPISupported As Boolean

Public Sub GetFilterSettings()
   Dim sPath As String
   sPath = NormalizePath(App.Path) & FilterSettingFile
   If Dir(sPath) = "" Then
      LoadDefaultSettings
      SaveFilterSettings
      Exit Sub
   End If
   Dim nFile As Integer
   nFile = FreeFile
   Open sPath For Binary As #nFile
        Get #nFile, , KernelFilterType
        Get #nFile, , KernelFilterSize
        Get #nFile, , KernelFilterPower
        Get #nFile, , RankFilterSize
        Get #nFile, , EnhancedFilterSize
        Get #nFile, , CustomFilterSize
        ReDim CustomFilter(-(3 + CustomFilterSize * 2) \ 2 To (3 + CustomFilterSize * 2) \ 2, -(3 + CustomFilterSize * 2) \ 2 To (3 + CustomFilterSize * 2) \ 2)
        Get #nFile, , CustomFilter
   Close #nFile
End Sub

Public Sub SaveFilterSettings()
   Dim sPath As String
   Dim nFile As Integer
   sPath = NormalizePath(App.Path) & FilterSettingFile
   nFile = FreeFile
   Open sPath For Binary As #nFile
        Put #nFile, , KernelFilterType
        Put #nFile, , KernelFilterSize
        Put #nFile, , KernelFilterPower
        Put #nFile, , RankFilterSize
        Put #nFile, , EnhancedFilterSize
        Put #nFile, , CustomFilterSize
        Put #nFile, , CustomFilter
   Close #nFile
End Sub

Private Sub LoadDefaultSettings()
   Dim i As Long
   For i = 0 To 3
       KernelFilterSize(i) = 0
       KernelFilterPower(i) = 1
       If i < 3 Then
          RankFilterSize(i) = 0
          EnhancedFilterSize(i) = 0
       End If
   Next i
   CustomFilterSize = 0
   ReDim CustomFilter(-1 To 1, -1 To 1)
   CustomFilter(0, 0) = 1
   KernelFilterType(0) = 1
   KernelFilterType(1) = 2
   KernelFilterType(2) = 2
   KernelFilterType(3) = 0
End Sub

Private Function NormalizePath(ByVal sPath As String) As String
  If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
  NormalizePath = sPath
End Function
