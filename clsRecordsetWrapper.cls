VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRecordsetWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_rs As dao.Recordset2

Public Function GetRecordsetClone(rs As dao.Recordset2) As dao.Recordset2
    If Not m_rs Is Nothing Then
        Debug.Assert False ' This is only designed to be used once
    Else
        Set m_rs = rs.Clone
        Set GetRecordsetClone = m_rs
    End If
End Function

Public Function OpenRecordset(Domain As String, _
                              Optional Criteria As String = "1=1", _
                              Optional OrderBy As String, _
                              Optional RecordsetType As dao.RecordsetTypeEnum = dbOpenDynaset, _
                              Optional RecordsetOptions As dao.RecordsetOptionEnum _
                              ) As Boolean
    
    
    If Not m_rs Is Nothing Then
        ' Close the recordset so it can be re-used
        CloseRecordset
    End If
    
    Dim strSQL As String
    
    If left(Domain, 6) = "SELECT" Then
        strSQL = Domain
    Else
        strSQL = "SELECT * FROM [" & Domain & "] WHERE " & Criteria
        
        If OrderBy <> "" Then
            strSQL = strSQL & " ORDER BY " & OrderBy
        End If
    End If
    
    Set m_rs = CurrentDb.OpenRecordset(strSQL, RecordsetType, RecordsetOptions)
    OpenRecordset = True
    
End Function

Public Function Delete() As Boolean
    m_rs.Delete
    Delete = True
End Function

Public Function AddNew() As Boolean
    m_rs.AddNew
    AddNew = True
End Function

Public Function Edit() As Boolean
    m_rs.Edit
    Edit = True
End Function

Public Function Update() As Boolean
    m_rs.Update
    Update = True
End Function

Public Function MoveNext() As Boolean
    m_rs.MoveNext
    MoveNext = True
End Function

Public Function CloseRecordset() As Boolean
    m_rs.Close
    CloseRecordset = True
End Function

Public Function RecordCount() As Integer
    RecordCount = m_rs.RecordCount
End Function

Public Property Get Recordset() As dao.Recordset2
    Set Recordset = m_rs
End Property

Private Sub Class_Terminate()
    If Not m_rs Is Nothing Then
        m_rs.Close
        Set m_rs = Nothing
    End If
End Sub


