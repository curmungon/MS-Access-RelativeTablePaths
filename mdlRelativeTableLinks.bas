Attribute VB_Name = "mdlRelativeTableLinks"
Option Compare Database
Option Explicit

' collect all of the linked tables here before using LinkAssociatedTables
' the helper function CollectLinkedTalbes can print the currently linked tables into a format for pasting below
' CollectLinkedTalbes can be run from the "Immediate" window accessible by pressing "Ctrl+G" or going to "View -> Immediate Window"
' CollectLinkedTalbes also prints to the Immediate window (if emit is set to true)
Public Function LinkedTablesCollection() As Collection
    Set LinkedTablesCollection = New Collection
    With LinkedTablesCollection
        ' allow different tables and paths for dev "accdb" and compiled "accde"
        If Right(CurrentProject.Name, 1) = "e" Then
            .Add Array("Table1", "Table1", dbConnectPath(".\New Microsoft Access Database_BE.accdb"))
            .Add Array("Table1", "Table2", dbConnectPath(".\New Microsoft Access Database_BE.accdb"))
        ElseIf Right(CurrentProject.Name, 1) = "b" Then
            .Add Array("Table1", "Table1", dbConnectPath(".\New Microsoft Access Database_BE.accdb"))
            .Add Array("Table1", "Table2", dbConnectPath(".\New Microsoft Access Database_BE.accdb"))
        End If
    End With
End Function

' resolves relative paths, if needed, and prefixes with ;Database=
Public Function dbConnectPath(path) As String
    dbConnectPath = ";Database=" & ResloveRelativePath(path)
End Function

' correctly resolves paths relative to the current database's path
Public Function ResloveRelativePath(path) As String
    Dim ResloveRelativePathFSO As Object
    Set ResloveRelativePathFSO = CreateObject("Scripting.FileSystemObject")
    ResloveRelativePath = ResloveRelativePathFSO.GetAbsolutePathName(CurrentProject.path & "\" & path)
End Function

' removes existing linked tables and links whatever is supplied in the LinkedTablesCollection
' this could also be used to refresh links by forcing a drop and reconnecting
Public Function LinkAssociatedTables()
    KillTableLinks False
    LinkTables LinkedTablesCollection
    CurrentDb.TableDefs.Refresh
End Function


' the passed tableCollection must contain elements in an Array(),
' index 0 is the source table's name
' index 1 is the table's name for use in the current database
' index 2 is the connection string with ";Database=" in front of the source database's path
Public Function LinkTables(ByRef tableCollection As Collection)
  Dim link As Variant
  Dim tbl As TableDef
  
  For Each link In tableCollection
    Set tbl = New TableDef
    tbl.SourceTableName = link(0)
    tbl.Name = link(1)
    tbl.Connect = link(2)
    CurrentDb.TableDefs.Append tbl
  Next link
  
  CurrentDb.TableDefs.Refresh
End Function

' removes all linked tables from the database
' defaults to confirming before removing the tables
Public Function KillTableLinks(Optional ByVal confirm As Boolean = True)
    Dim lResult As Long
    If confirm = True Then
        lResult = MsgBox("All linked tables will be disconnected from the database!" & vbCrLf & vbCrLf & _
                         "Would you like to continue?", _
                         vbYesNo, "Disconnect Linked Tables")
        If lResult = vbNo Then
            Exit Function
        End If
    End If
    
    Dim tbl As TableDef
    For Each tbl In CurrentDb.TableDefs
        If tbl.Connect <> vbNullString Then
            CurrentDb.TableDefs.Delete tbl.Name
        End If
    Next tbl
    CurrentDb.TableDefs.Refresh
End Function


' CollectLinkedTables will return a collection of the currently linked tables
' it can also print the linked tables into the immediate window
' the emitted format can easily be used with LinkedTablesCollection to create dynamic links
Public Function CollectLinkedTalbes(Optional emit As Boolean = True, _
                             Optional TrimConnection As Boolean = True) As Collection
    Dim tbl As TableDef
    Dim tblPath As String
    Dim ind As Long
    Dim item As Variant
    
    Set CollectLinkedTalbes = New Collection
    
    For Each tbl In CurrentDb.TableDefs
        If tbl.Connect <> vbNullString Then
            tblPath = tbl.Connect
            If TrimConnection = True Then
                ind = InStr(1, tblPath, "=")
                tblPath = Mid$(tblPath, ind + 1, Len(tblPath) - ind)
                tblPath = "dbConnectPath(" & Chr(34) & tblPath & Chr(34) & ")"
            Else
                tblPath = Chr(34) & tblPath & Chr(34)
            End If
            CollectLinkedTalbes.Add Array(tbl.SourceTableName, tbl.Name, tblPath), tbl.Name
        End If
    Next tbl
    If emit = True Then
    For Each item In CollectLinkedTalbes
        Debug.Print ".Add Array(" & Chr(34) & item(0) & Chr(34) & ", " & Chr(34) & item(1) & Chr(34) & ", " & item(2) & ")"
    Next
    End If
    
End Function

