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

' Constants
Private Const conTableName = "tblLogError3" ' Where the error details are saved

' Retrieves the name of the user associated with the current thread
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
    ByVal lpBuffer As String, _
    nSize As Long) As Long

Private Sub Class_Initialize()
    Dim strSQL As String
    If Not DebuggingSupported() Then
        If Not TableExists(conTableName) Then
            ' Make the log table if compiled
            strSQL = "CREATE TABLE " & conTableName _
                   & " (ID AUTOINCREMENT(1, 1)," _
                   & " ErrNumber INTEGER," _
                   & " ErrDescription VARCHAR(255)," _
                   & " ErrDate DATE," _
                   & " CallingProc VARCHAR(100)," _
                   & " NetworkUserName VARCHAR(50)," _
                   & " ShowUser YESNO," _
                   & " Parameters VARCHAR(255));"
        
            CurrentDb.Execute strSQL
        End If
    End If
End Sub

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
                If rsw.OpenRecordset(conTableName, , , , dbAppendOnly) Then
                    With rsw.Recordset
                        If Not rsw.AddNew Then Exit Function
                            ![ErrNumber] = lngErrorNumber
                            ![ErrDescription] = left$(strDescription, 255)
                            ![ErrDate] = Now()
                            ![CallingProc] = strCallingProc
                            ![NetworkUserName] = NetworkUserName
                            ![ShowUser] = bolShowUser
                            ' If parameters were passed then save them here
                            If Not IsMissing(varParameters) Then
                                ![Parameters] = left(varParameters, 255)
                            End If
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

Private Function TableExists(strTableName As String) As Boolean
    Dim db As dao.Database
    Dim i As Integer
    
    Set db = CurrentDb
    TableExists = False
    db.TableDefs.Refresh
    For i = 0 To db.TableDefs.Count - 1
        If strTableName = db.TableDefs(i).Name Then
            'Table Exists
            TableExists = True
            Exit For
        End If
    Next i
    Set db = Nothing
End Function

Private Function NetworkUserName() As String
    On Error GoTo ErrorHandler
    
    'Purpose:   Returns the network login name.
    Dim lngLen As Long          'Length of string.
    Dim strUserName As String
    Const lngcMaxFieldSize As Long = 64& 'Length of field to store this data.
    
    'Initialize
    strUserName = String$(254, vbNullChar)
    lngLen = 255&
    
    'API returns a non-zero value if success.
    If GetUserName(strUserName, lngLen) <> 0& Then
        lngLen = lngLen - 1&    'Without null termination char.
        If lngLen > lngcMaxFieldSize Then  'Maximum field size
            lngLen = lngcMaxFieldSize
        End If
        NetworkUserName = left$(strUserName, lngLen)
    End If

ExitHandler:
    Exit Function
    
ErrorHandler:
    MsgBox Err.Description
    Resume ExitHandler
    
End Function

Private Function HasSourceCode() As Boolean
    On Error Resume Next
    HasSourceCode = (CurrentDb.Properties("MDE") <> "T")
    ' Property not found error
    If Err = 3270 Then HasSourceCode = True
    On Error GoTo 0
End Function

Private Function IsRuntime() As Boolean
    IsRuntime = SysCmd(acSysCmdRuntime)
End Function

Private Function DebuggingSupported() As Boolean
    DebuggingSupported = HasSourceCode() And Not IsRuntime()
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
