Attribute VB_Name = "ExamplesStubRecordset"
'@Folder "SecureADODB.Examples"
'@IgnoreModule
Option Explicit
Option Private Module

Private Const LIB_NAME As String = "SecureADODB"
Private Const PATH_SEP As String = "\"
Private Const REL_PREFIX As String = "Library" & PATH_SEP & LIB_NAME & PATH_SEP


Public Sub FabricateRecordset()
    Dim objRs As New ADODB.Recordset
    With objRs.Fields
        .Append "StudentID", adChar, 11, adFldUpdatable
        .Append "FullName", adVarChar, 50, adFldUpdatable
        .Append "PhoneNmbr", adVarChar, 20, adFldUpdatable
    End With
    With objRs
        .Open
        .AddNew
        .Fields(0) = "123-45-6789"
        .Fields(1) = "John Doe"
        .Fields(2) = "(425) 555-5555"
        .Update
        .AddNew
        .Fields(0) = "123-45-6780"
        .Fields(1) = "Jane Doe"
        .Fields(2) = "(615) 555-1212"
        .Update
    End With
    
    Dim FileName As String
    FileName = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & "FabricatedRecordset.xml"
    On Error Resume Next
    Kill FileName
    On Error GoTo 0
    objRs.Save FileName, adPersistXML
    objRs.Close
End Sub


Public Sub FabricateRecordsetWithNull()
    Dim objRs As New ADODB.Recordset
    With objRs.Fields
        .Append "StudentID", adChar, 11, adFldUpdatable
        .Append "FullName", adVarChar, 50, adFldUpdatable
        .Append "PhoneNmbr", adVarChar, 20, adFldUpdatable
    End With
    With objRs
        .Open
        .AddNew
        .Fields(0) = "123-45-6789"
        .Fields(1) = "John Doe"
        .Fields(2) = "(425) 555-5555"
        .Update
        .AddNew
        .Fields(0) = "123-45-6780"
        .Fields(1) = "Jane Doe"
        .Fields(2) = "(615) 555-1212"
        .Update
    End With
    
    Dim FileName As String
    FileName = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & "FabricatedRecordset.xml"
    On Error Resume Next
    Kill FileName
    On Error GoTo 0
    objRs.Save FileName, adPersistXML
    objRs.Close
End Sub


Private Sub SaveRestore()
    Dim FileName As String
    FileName = REL_PREFIX & LIB_NAME & ".db"

    Dim TableName As String
    TableName = "people"
    Dim SQLQuery As String
    SQLQuery = "SELECT * FROM " & TableName & " WHERE age >= ? AND country = ?"
    
    Dim dbm As IDbManager
    Set dbm = DbManager.CreateFileDb("sqlite", FileName, vbNullString, LoggerTypeEnum.logPrivate)

    Dim Log As ILogger
    Set Log = dbm.LogController

    Dim conn As IDbConnection
    Set conn = dbm.Connection
    Dim connAdo As ADODB.Connection
    Set connAdo = conn.AdoConnection
    
    Dim cmd As IDbCommand
    Set cmd = dbm.Command
    Dim cmdAdo As ADODB.Command
    Set cmdAdo = cmd.AdoCommand(SQLQuery, 45, "South Korea")
    
    Dim rst As IDbRecordset
    Set rst = dbm.Recordset(Disconnected:=True, CacheSize:=10)
    Dim rstAdo As ADODB.Recordset
    Set rstAdo = rst.OpenRecordset(SQLQuery, 45, "South Korea")
    
    rst.RecordsetToQT Buffer.Range("A1")
    
    FileName = ThisWorkbook.Path & PATH_SEP & REL_PREFIX & "PersistedRecordset.xml"
    On Error Resume Next
    Kill FileName
    On Error GoTo 0
    rstAdo.Save FileName, adPersistXML
    
    Dim RstFromFile As ADODB.Recordset
    Set RstFromFile = New ADODB.Recordset
    RstFromFile.Open Source:=FileName, Options:=adCmdFile
    
    DbRecordset.RecordsetToQT Buffer.Range("K1"), RstFromFile
End Sub
