VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMailMerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private wApp As Object
Private wDoc As Object
Private wSingleDoc As Object
Private strMergeDataFile As String
Private m_MailMergeDestination As WdMailMergeDestination
Private m_FileName As String
Private m_FilePath As String
Private m_Query As String
Private m_Document As String

Private Const conClass = "clsMailMerge"

Private Sub Class_Initialize()
    Set wApp = CreateObject("Word.Application")
    strMergeDataFile = Environ("TEMP") & "\MergeData.txt"
End Sub

Private Sub Class_Terminate()
    Set wDoc = Nothing
    Set wApp = Nothing
End Sub

Public Property Get FileName() As String
    FileName = m_FileName
End Property

Public Property Let FileName(ByVal sNewValue As String)
    m_FileName = sNewValue
End Property

Public Property Get Document() As String
    Document = m_Document
End Property

Public Property Let Document(ByVal sNewValue As String)
    m_Document = sNewValue
End Property

Public Property Get FilePath() As Variant
    FilePath = m_FilePath
End Property

Public Property Let FilePath(ByVal sNewValue As Variant)
    m_FilePath = sNewValue
End Property

Private Property Get FullPath() As String
    FullPath = FilePath & "\" & FileName
End Property

Public Property Let Query(ByVal sNewValue As String)
    m_Query = sNewValue
End Property

Public Property Let MailMergeDestination(ByVal NewValue As WdMailMergeDestination)
    m_MailMergeDestination = NewValue
End Property

Public Sub MailMerge()
On Error GoTo Error_Handler

    Dim strActiveDoc As String
    
    Set wDoc = wApp.Documents.Open(Document)
    strActiveDoc = wApp.ActiveDocument.Name
    
    DoCmd.TransferText acExportDelim, , m_Query, Environ("TEMP") & "\MergeData.txt", True
    
    wApp.visible = False
    
    wDoc.MailMerge.OpenDataSource _
        Name:=strMergeDataFile, _
        ConfirmConversions:=False, _
        ReadOnly:=False, LinkToSource:=True, AddToRecentFiles:=False, _
        PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
        WritePasswordTemplate:="", Revert:=False, Format:=0, _
        Connection:="", SQLStatement:="", SQLStatement1:=""
    
    With wDoc.MailMerge
        .Destination = m_MailMergeDestination
        .MailAsAttachment = False
        .MailAddressFieldName = ""
        .MailSubject = ""
        .SuppressBlankLines = True
        
        With .DataSource
            .FirstRecord = 1
        End With
            
        .Execute Pause:=False
    End With
    
    Set wSingleDoc = wApp.ActiveDocument
   
    wDoc.Close (False)
    
Exit_Handler:
    Exit Sub

Error_Handler:
    eh.LogError conClass & ".MailMerge()"
    Resume Exit_Handler
    
End Sub

Public Sub SaveAs(ByVal SaveFormat As WdSaveFormat, OpenDoc As Boolean)
On Error GoTo Error_Handler
    Dim strFileName As String
    Dim strDestinationPath As String
    
    wSingleDoc.Protect Password:="ICAT", Type:=wdAllowOnlyFormFields, NoReset:=True, EnforceStyleLock:=True
    wSingleDoc.SaveAs FileName:=FullPath, FileFormat:=SaveFormat
    wSingleDoc.Close (False)

    strFileName = Split(FullPath, "\")(UBound(Split(FullPath, "\")))
    strDestinationPath = CaseDataPath & TempVars!AssignRefNum & "\"

    FileCopy FullPath, strDestinationPath & strFileName
    
    If OpenDoc Then
        GoHyperlink CaseDataPath & TempVars!AssignRefNum & "\" & m_FileName
    End If
    
    wApp.Quit
    
Exit_Handler:
    Exit Sub

Error_Handler:
    eh.LogError conClass & ".SaveAs(ByVal SaveFormat As WdSaveFormat, OpenDoc As Boolean)"
    Resume Exit_Handler
    
End Sub

Public Sub View()
On Error GoTo Error_Handler
    
    wApp.visible = True
    wSingleDoc.Activate
    wSingleDoc.Protect Password:="ICAT", Type:=wdAllowOnlyFormFields, NoReset:=True, EnforceStyleLock:=True
    
Exit_Handler:
    Exit Sub

Error_Handler:
    eh.LogError conClass & ".View()"
    Resume Exit_Handler
    
End Sub

Public Sub Edit()
   On Error GoTo CreateWordApp
   Set wApp = GetObject(, "Word.Application")
    
   On Error GoTo 0
   
   Set wDoc = wApp.Documents.Open(Document)
   wApp.visible = True

   wDoc.MailMerge.MainDocumentType = wdFormLetters

    wDoc.MailMerge.OpenDataSource _
        Name:=strMergeDataFile, _
        ConfirmConversions:=False, _
        ReadOnly:=False, LinkToSource:=True, AddToRecentFiles:=False, _
        PasswordDocument:="", PasswordTemplate:="", WritePasswordDocument:="", _
        WritePasswordTemplate:="", Revert:=False, Format:=0, _
        Connection:="", SQLStatement:="", SQLStatement1:=""

   wApp.Activate
   wApp.WindowState = wdWindowStateMaximize
   
   Exit Sub
   
CreateWordApp:
   
   Set wApp = CreateObject("Word.Application")
   Resume Next
End Sub

