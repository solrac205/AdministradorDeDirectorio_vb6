Attribute VB_Name = "IniModuleOperations"
Private Declare Function GetPrivateProfileString Lib "kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    
Private Declare Function WritePrivateProfileString Lib "kernel32" _
    Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long
    

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' LEE EL INIFILE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReadStringINI(cINIFile As String, cSection As String, cKey As String, CDefault As String) As String
On Error GoTo Err_ReadStringIni

Dim cTemp As String
Dim IOK%
ReadStringINI = True

cTemp = Space$(255)
IOK% = GetPrivateProfileString(cSection, cKey, CDefault, cTemp, Len(cTemp), cINIFile)

If IOK% > 0 Then
   ReadStringINI = Left$(cTemp, IOK%)
Else
   ReadStringINI = CDefault
End If
Exit Function

Err_ReadStringIni:
  ReadStringINI = CDefault
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ESCRIBE EL INIFILE
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WriteStringINI(cINIFile As String, cSection As String, cKey As String, lpString As String)
On Error GoTo Err_WriteStringIni

Dim IOK%
WriteStringINI = True

IOK% = WritePrivateProfileString(cSection, cKey, lpString, cINIFile)

If IOK% > 0 Then
Else
   
   WriteStringINI = False
End If
Exit Function

Err_WriteStringIni:
  WriteStringINI = False
End Function

