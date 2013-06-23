VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTableProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private m_strPath As String

Public Property Get Path() As String
    Path = m_strPath
End Property

Public Property Let Path(ByVal strPath As String)
    m_strPath = strPath
End Property

Public Sub LinkTables()
    Dim objLocalDatabase As dao.Database
    Dim objRemoteDatabase As dao.Database
    Dim LocalTableDef As TableDef
    Dim RemoteTableDef As TableDef

    Set objRemoteDatabase = OpenDatabase(Me.Path) ' Path to database
    ' List the table names
    For Each RemoteTableDef In objRemoteDatabase.TableDefs
    ' if you want to display also the system tables, replace the line
        If RemoteTableDef.Attributes = 0 Then
            If RemoteTableDef.Name <> "USysRibbons" Then ' Just in case there's a custom ribbon table
                Set objLocalDatabase = CurrentDb
                Set LocalTableDef = objLocalDatabase.CreateTableDef(RemoteTableDef.Name)
            
                LocalTableDef.Connect = ";DATABASE=" & Me.Path
                LocalTableDef.SourceTableName = RemoteTableDef.Name
            
                objLocalDatabase.TableDefs.Append LocalTableDef
                objLocalDatabase.Close
    
                Set objLocalDatabase = Nothing
            End If
       End If
    Next RemoteTableDef
    
    objRemoteDatabase.Close

End Sub

Public Sub RemoveTable(strTable As String)
    DoCmd.DeleteObject acTable, strTable
End Sub

Public Sub RemoveTables(TableType As TableDefAttributeEnum)

    Dim db As dao.Database
    Dim tdf As TableDef
    Dim intCounter As Integer

    Set db = CurrentDb()

    For intCounter = 0 To db.TableDefs.Count - 1
        Set tdf = db.TableDefs(intCounter)
        If (tdf.Attributes And TableType) = TableType Then
            Debug.Print "Removing table " & tdf.Name
            DoCmd.DeleteObject acTable, tdf.Name ' Delete the linked table
        End If
    Next intCounter
    
    db.Close
    
    Set db = Nothing
    Set tdf = Nothing

End Sub

Sub Attributes()
 
    Dim db As dao.Database
    Dim fldLoop As Field
    Dim relLoop As Relation
    Dim tdfloop As TableDef
 
    Set db = OpenDatabase(Me.Path)
 
    With db
 
       ' Display the attributes of a TableDef object's
       ' fields.
       Debug.Print "Attributes of fields in " & _
       .TableDefs(0).Name & " table:"
    
       For Each fldLoop In .TableDefs(0).Fields
           Debug.Print " " & fldLoop.Name & " = " & _
           fldLoop.Attributes
       Next fldLoop
    
       ' Display the attributes of the Northwind database's
       ' relations.
       Debug.Print "Attributes of relations in " & _
       .Name & ":"
       For Each relLoop In .Relations
           Debug.Print " " & relLoop.Name & " = " & _
           relLoop.Attributes
       Next relLoop
    
       ' Display the attributes of the Northwind database's
       ' tables.
       Debug.Print "Attributes of tables in " & .Name & ":"
       For Each tdfloop In .TableDefs
           Debug.Print " " & tdfloop.Name & " = " & _
           tdfloop.Attributes
       Next tdfloop
 
    .Close
    End With
 
End Sub

