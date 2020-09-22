Attribute VB_Name = "mRegistry"
Public Type FILTERINFO
   Extension As String
   Name As String
   Path As String
   RegPath As String
End Type

Public Enum MS_FILTERTYPES
    eImport
    eExport
End Enum

Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const ERROR_SUCCESS = 0&
Private Const KEY_ALL_ACCESS = &HF003F
Private Const MAX_SIZE = 2048

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Const sRegImportFilters As String = "Software\Microsoft\Shared Tools\Graphics Filters\Import"
Const sRegExportFilters As String = "Software\Microsoft\Shared Tools\Graphics Filters\Export"
Public ImportFilters() As FILTERINFO
Public ExportFilters() As FILTERINFO
Public IsBMPExportSupported As Boolean

Public Function EnumFilters(arrFI() As FILTERINFO, Optional eFilterType As MS_FILTERTYPES) As Boolean
   Dim sKey As String, sTemp As String
   Dim hKey As Long, curidx As Long, KeyName As String, KeyValue As String
   Dim i As Long
   If eFilterType = eImport Then
      sKey = sRegImportFilters
   Else
      sKey = sRegExportFilters
   End If
   If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey, 0&, KEY_ALL_ACCESS, hKey) Then Exit Function
   Do
     KeyName = Space$(MAX_SIZE)
     KeyValue = Space$(MAX_SIZE)
     If RegEnumKey(hKey, curidx, KeyName, MAX_SIZE) <> ERROR_SUCCESS Then Exit Do
     ReDim Preserve arrFI(curidx)
     KeyName = TrimNull(KeyName)
     arrFI(curidx) = GetFilterInfo(sKey & "\" & KeyName)
     curidx = curidx + 1
   Loop
   RegCloseKey hKey
   EnumFilters = curidx
End Function

Private Function GetFilterInfo(ByVal sKey As String) As FILTERINFO
  Dim hKey As Long, sTemp As String, lSize As Long
  If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey, 0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then Exit Function
  GetFilterInfo.RegPath = sKey
  lSize = MAX_SIZE
  sTemp = String(lSize, 0)
  If RegQueryValueEx(hKey, "Path", 0, 1, ByVal sTemp, lSize) = ERROR_SUCCESS Then
     GetFilterInfo.Path = TrimNull(sTemp)
  End If
  lSize = MAX_SIZE
  sTemp = String(lSize, 0)
  If RegQueryValueEx(hKey, "Extensions", 0, 1, ByVal sTemp, lSize) = ERROR_SUCCESS Then
     GetFilterInfo.Extension = LCase(TrimNull(sTemp))
  End If
  lSize = MAX_SIZE
  sTemp = String(lSize, 0)
  If RegQueryValueEx(hKey, "Name", 0, 1, ByVal sTemp, lSize) = ERROR_SUCCESS Then
     GetFilterInfo.Name = TrimNull(sTemp)
  End If
End Function

Public Function GetRegValue(ByVal sKey As String, ByVal sSubKey As String) As String
  Dim hKey As Long, sTemp As String, lSize As Long, lTemp As Long, lType As Long
  If RegOpenKeyEx(HKEY_LOCAL_MACHINE, sKey, 0&, KEY_ALL_ACCESS, hKey) <> ERROR_SUCCESS Then Exit Function
  lType = 1
  lSize = MAX_SIZE
  sTemp = String(lSize, 0)
  If RegQueryValueEx(hKey, sSubKey, 0, lType, ByVal sTemp, lSize) = ERROR_SUCCESS Then
     GetRegValue = TrimNull(sTemp)
  End If
End Function

Public Function TrimNull(startstr As String) As String
   Dim pos As Integer
   pos = InStr(startstr, Chr$(0))
   If pos Then
      TrimNull = Left$(startstr, pos - 1)
      Exit Function
   End If
   TrimNull = startstr
End Function

