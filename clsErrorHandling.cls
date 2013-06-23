VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorHandling"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Public Function LogError(strCallingProc As String, Optional varParameters As Variant, _
                         Optional bolShowUser = False) As Boolean

    Dim strMsg As String ' String for display in MsgBox
    Dim strState As String ' Optional parameters to be displayed
    Dim lngErrorNumber As Long
    Dim strDescription As String
    Dim strHelpfile As String
    Dim lngHelpContext As Long

    ' This function will reset the Err.number due to its own ErrorHandler so save
    ' the calling function's error number, description, and helpfile
    lngErrorNumber = Err.number
    strDescription = Err.Description
    strHelpfile = Err.HelpFile
    lngHelpContext = Err.HelpContext
    
On Error GoTo ErrorHandler

    ' Build the error message to display
    strMsg = strDescription & " (" & lngErrorNumber & ")" _
             & vbCrLf & vbCrLf & "Calling Procedure: " & vbCrLf & strCallingProc
                    
    If Not IsMissing(varParameters) Then
        ' Add whatever parameters were passed
        strMsg = strMsg & vbCrLf & vbCrLf & "Parameters: " _
                 & vbCrLf & varParameters
    End If
    
    ' If debugging is supported, break using Debug.Assert.
    If Not DebuggingSupported() Then
        Select Case lngErrorNumber
            Case 0
                Debug.Print strCallingProc & " called error 0."
            Case 2501                ' Cancelled
                'Do nothing.
            Case 3314, 2101, 2115    ' Can't save.
                strMsg = "Record cannot be saved at this time." & vbCrLf & _
                         "Complete the entry, or press <Esc> to undo."
            Case Else
                Dim rsw As New clsRecordsetWrapper
                If rsw.OpenRecordset("tblLogError", , , , dbAppendOnly) Then
                    With rsw.Recordset
                        If Not rsw.AddNew Then Exit Function
                            ![ErrNumber] = lngErrorNumber
                            ![ErrDescription] = left$(strDescription, 255)
                            ![ErrDate] = Now()
                            ![CallingProc] = strCallingProc
                            ![UserName] = CurrentUser()
                            ![ShowUser] = bolShowUser
                            ![EmployeeID] = TempVars!EmployeeID
                            ' If parameters were passed then save them here
                            If Not IsMissing(varParameters) Then
                                ![Parameters] = left(varParameters, 255)
                            End If
                            ![Version] = g_intAppVersion
                        LogError = rsw.Update
                    End With
                End If
        End Select
        
        If bolShowUser Then
            ' Display the error
            MsgBox strMsg, vbCritical + vbMsgBoxHelpButton, strCallingProc, strHelpfile, lngHelpContext
        End If
    Else
        ' Display the error
        MsgBox strMsg, vbCritical + vbMsgBoxHelpButton, strCallingProc, strHelpfile, lngHelpContext
        Debug.Assert False  ' Stop code so that you can debug
        LogError = True     ' Step over this line if you don't want to resume
    End If

ExitHandler:
    Exit Function

ErrorHandler:
    strMsg = "An unexpected situation arose in your program." & vbCrLf & _
        "Please write down the following details:" & vbCrLf & vbCrLf & _
        "Calling Proc: " & strCallingProc & vbCrLf & _
        "Error Number " & lngErrorNumber & vbCrLf & strDescription & vbCrLf & vbCrLf & _
        "Unable to record because Error " & Err.number & vbCrLf & Err.Description
    MsgBox strMsg, vbCritical + vbMsgBoxHelpButton, "LogError()", Err.HelpFile, Err.HelpContext
    Resume ExitHandler
    
End Function

Public Function TryToCloseObject() As Boolean
    On Error GoTo ErrorHandler

    DoCmd.Close
    TryToCloseObject = True

ExitHandler:
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    Resume ExitHandler
    
End Function

Public Function TryToSaveRecord() As Boolean
    TryToSaveRecord = TryToRunCommand(acCmdSaveRecord)
End Function

Public Function TryToRunCommand(Command As AcCommand) As Boolean
    On Error GoTo ErrorHandler
    
    DoCmd.RunCommand Command
    TryToRunCommand = True

ExitHandler:
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    Resume ExitHandler
    
End Function

Function TryToGoToRecord(Record As AcRecord) As Boolean
    On Error GoTo ErrorHandler
    
    DoCmd.GoToRecord , , Record
    
ExitHandler:
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    Resume ExitHandler
    
End Function

