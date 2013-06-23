VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsListFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_strFolderPath As String
Private m_strFileSpec As String
Private m_lst As ListBox
Private m_strTable As String

Private Sub Class_Initialize()
    Me.FileSpec = "*.*"
End Sub

Public Property Get FileSpec() As String
    FileSpec = m_strFileSpec
End Property

Public Property Let FileSpec(ByVal strFileSpec As String)
    m_strFileSpec = strFileSpec
End Property

Public Property Get FolderPath() As String
    FolderPath = m_strFolderPath
End Property

Public Property Let FolderPath(ByVal strFolderName As String)
    m_strFolderPath = strFolderName
End Property

Public Property Get ListBox() As ListBox
    Set ListBox = m_lst
End Property

Public Property Let ListBox(ByRef lst As ListBox)
    Set m_lst = lst
End Property

Public Property Get Table() As String
    Table = m_strTable
End Property

Public Property Let Table(ByVal strTable As String)
    m_strTable = strTable
End Property

Public Function ListFiles(Optional bIncludeSubfolders As Boolean)
    'Purpose:   List the files in the path.
    'Arguments: strPath = the path to search.
    '           strFileSpec = "*.*" unless you specify differently.
    '           bIncludeSubfolders: If True, returns results from subdirectories of strPath as well.
    '           lst: if you pass in a list box, items are added to it. If not, files are listed to immediate window.
    '               The list box must have its Row Source Type property set to Value List.
    'Method:    FilDir() adds items to a collection, calling itself recursively for subfolders.
    Dim colDirList As New Collection
    Dim varItem As Variant

    If FolderExists(FolderPath) Then
    
        Call FillDir(colDirList, FolderPath, Me.FileSpec, bIncludeSubfolders)
        
        'Add the files to a list box if one was passed in. Otherwise list to the Immediate Window.
        If Me.ListBox Is Nothing Then
            For Each varItem In colDirList
                Debug.Print varItem
            Next
        Else
            ClearListBox Me.ListBox
            For Each varItem In colDirList
                Me.ListBox.AddItem varItem
            Next
        End If
        
        If Len(Me.Table & vbNullString) <> 0 Then
            If TableExists(Me.Table) Then
                ClearTable Me.Table
                For Each varItem In colDirList
                    AddToTable varItem, CStr(Me.Table)
                Next
            Else
                CreateTable Me.Table
                For Each varItem In colDirList
                    AddToTable varItem, CStr(Me.Table)
                Next
            End If
        End If
    Else
            MsgBox "Unable to locate " & FolderPath, vbCritical, "Folder Not Found"
    End If
End Function

Private Function CreateTable(strTableName As String)
    Dim strSQL As String
    
    strSQL = "CREATE TABLE " & strTableName _
           & " (FPath VARCHAR(255)," _
           & " FName VARCHAR(255)," _
           & " FDateModified DATE," _
           & " FType VARCHAR(255)," _
           & " FSize LONG);"

    CurrentDb.Execute strSQL
End Function

Private Function AddToTable(vItem As Variant, strTableName As String)
    Dim strSQL As String
    Dim fso As Variant
    Dim f As Variant
    Dim strName As String
    Dim strFolder As String
    Dim lngFileSize As Double
    Dim dteDateModified As Date
    Dim strFileType As String
    
    strFolder = left(vItem, InStrRev(vItem, "\"))
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(vItem)
    strName = fso.GetFileName(vItem)
    lngFileSize = -Int(-(f.size / 1024))
    dteDateModified = f.DateLastModified
    strFileType = f.Type
                
    strSQL = "INSERT INTO " & strTableName _
           & " (FPath, FName, FDateModified, FType , FSize) " _
           & " SELECT """ & Trim(strFolder) & """" _
           & ", """ & Trim(strName) & """" _
           & ", """ & dteDateModified & """" _
           & ", """ & Trim(strFileType) & """" _
           & ", """ & lngFileSize & """;"
    CurrentDb.Execute strSQL
End Function

Private Function FillDir(colDirList As Collection, ByVal strFolder As String, strFileSpec As String, _
    bIncludeSubfolders As Boolean)
    'Build up a list of files, and then add add to this list, any additional folders
    Dim strTemp As String
    Dim colFolders As New Collection
    Dim vFolderName As Variant

    'Add the files to the folder.
    strFolder = TrailingSlash(strFolder)
    strTemp = Dir(strFolder & Me.FileSpec)
    Do While strTemp <> vbNullString
        colDirList.Add strFolder & strTemp
        strTemp = Dir
    Loop

    If bIncludeSubfolders Then
        'Build collection of additional subfolders.
        strTemp = Dir(strFolder, vbDirectory)
        Do While strTemp <> vbNullString
            If (strTemp <> ".") And (strTemp <> "..") Then
                If (GetAttr(strFolder & strTemp) And vbDirectory) <> 0& Then
                    colFolders.Add strTemp
                End If
            End If
            strTemp = Dir
        Loop
        'Call function recursively for each subfolder.
        For Each vFolderName In colFolders
            Call FillDir(colDirList, strFolder & TrailingSlash(vFolderName), Me.FileSpec, True)
        Next vFolderName
    End If
End Function

Private Function TrailingSlash(varIn As Variant) As String
    If Len(varIn) > 0& Then
        If Right(varIn, 1&) = "\" Then
            TrailingSlash = varIn
        Else
            TrailingSlash = varIn & "\"
        End If
    End If
End Function

Private Function FolderExists(strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
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

Private Function ClearListBox(lst As ListBox)
    lst.RowSource = ""
End Function

Private Function ClearTable(strTable As String)
    DoCmd.SetWarnings False
        DoCmd.RunSQL "DELETE FROM " & strTable & ";"
    DoCmd.SetWarnings True
End Function
