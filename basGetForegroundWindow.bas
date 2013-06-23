Attribute VB_Name = "basGetForegroundWindow"
Option Compare Database
Option Explicit

' Retrieves a handle to the foreground window _
(the window with which the user is currently working)
Declare Function GetForegroundWindow Lib "user32.dll" () As Long

'Copies the text of the specified window's title bar _
(if it has one) into a buffer
Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" ( _
    ByVal hWnd As Long, _
    ByVal lpString As String, _
    ByVal cch As Long) As Long
    
    ' Create string filled with null characters.
    strCaption = String$(255, vbNullChar)
    ' Return length of string.
    lngLen = Len(strCaption)
    
    ' Call GetActiveWindow to return handle to active window,
    ' and pass handle to GetWindowText, along with string and its length.
    If (GetWindowText(GetForegroundWindow, strCaption, lngLen) > 0) Then
        ' Return value that Windows has written to string.
        ForegroundWindowCaption = strCaption
    End If
End Function
