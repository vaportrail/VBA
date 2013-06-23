Attribute VB_Name = "basLockScreenStatus"
Option Compare Database
Option Explicit

' Makes the specified desktop visible and activates it
Private Declare Function SwitchDesktop Lib "user32" ( _
    ByVal hDesktop As Long) As Long

' Opens the specified desktop object
Private Declare Function OpenDesktop Lib "user32" Alias "OpenDesktopA" ( _
    ByVal lpszDesktop As String, _
    ByVal dwFlags As Long, _
    ByVal fInherit As Long, _
    ByVal dwDesiredAccess As Long) As Long
    
' Closes an open handle to a desktop object
Private Declare Function CloseDesktop Lib "user32" ( _
    ByVal hDesktop As Long) As Long
    
Private Const DESKTOP_SWITCHDESKTOP As Long = &H100

Public Function isLocked() As Boolean
    Dim p_lngHwnd As Long
    Dim p_lngRtn As Long
    Dim p_lngErr As Long
     
    p_lngHwnd = OpenDesktop(lpszDesktop:="Default", dwFlags:=0, fInherit:=False, dwDesiredAccess:=DESKTOP_SWITCHDESKTOP)
     
    If p_lngHwnd = 0 Then
        'System = "Error"
    Else
        p_lngRtn = SwitchDesktop(hDesktop:=p_lngHwnd)
        p_lngErr = Err.LastDllError
         
        If p_lngRtn = 0 Then
            If p_lngErr = 0 Then
                isLocked = True
            Else
                'isLocked = "Error"
            End If
        Else
            isLocked = False
        End If
         
        p_lngHwnd = CloseDesktop(p_lngHwnd)
    End If
End Function
