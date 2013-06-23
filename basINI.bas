Attribute VB_Name = "basINI"
Option Compare Database
Option Explicit

' Retrieves a string from the specified section in an initialization file
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long

' Writes a string to the specified section in an initialization file
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long

Public Function GetINI(sINIFile As String, sSection As String, sKey As String, sDefault As String) As String
    
    Dim sTemp As String * 256
    Dim nLength As Integer
    
    sTemp = Space$(256)
    nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFile)
    GetINI = Left$(sTemp, nLength)
    
End Function

Public Sub WriteINI(sINIFile As String, sSection As String, sKey As String, sValue As String)
    
    Dim iCounter As Integer
    Dim sTemp As String
    
    sTemp = sValue
    
    'Replace any CR/LF characters with spaces
    For iCounter = 1 To Len(sValue)
        If Mid$(sValue, iCounter, 1) = vbCr Or Mid$(sValue, iCounter, 1) = vbLf Then Mid$(sValue, iCounter) = " "
    Next iCounter
    
    iCounter = WritePrivateProfileString(sSection, sKey, sTemp, sINIFile)
End Sub
