VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbMetaData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Database introspection functionality"
'@Folder "SecureADODB.DbManager.DbMetaData"
'@ModuleDescription "Database introspection functionality"
'@PredeclaredId
'@Exposed
Option Explicit

Private Type TDbMeta
    DbConnStr As DbConnectionString
End Type
'@Ignore MoveFieldCloserToUsage: Follow the standard pattern
Private this As TDbMeta


'@Description "Default factory"
Public Function Create(ByVal DbConnStr As DbConnectionString) As DbMetaData
Attribute Create.VB_Description = "Default factory"
    Dim Instance As DbMetaData
    Set Instance = New DbMetaData
    Instance.Init DbConnStr
    Set Create = Instance
End Function


'@Description "Default constructor"
Friend Sub Init(ByVal DbConnStr As DbConnectionString)
Attribute Init.VB_Description = "Default constructor"
    Guard.NullReference DbConnStr
    Set this.DbConnStr = DbConnStr
End Sub


'@Ignore ParameterCanBeByVal: False positive
'@Description "Query table field metadata via the ADOX library"
Public Sub QueryTableADOXMeta(ByVal TableName As String, _
                              ByRef FieldNames As Variant, _
                              ByRef FieldTypes As Variant, _
                              ByVal FieldMap As Scripting.Dictionary)
Attribute QueryTableADOXMeta.VB_Description = "Query table field metadata via the ADOX library"
    Dim Catalog As ADOX.Catalog
    Set Catalog = New ADOX.Catalog
    Catalog.ActiveConnection = this.DbConnStr.ConnectionString
    Dim Table As ADOX.Table
    '@Ignore IndexedDefaultMemberAccess
    Set Table = Catalog.Tables(TableName)
    Dim FieldCount As Long
    FieldCount = Table.Columns.Count
    
    ReDim FieldNames(1 To FieldCount)
    ReDim FieldTypes(1 To FieldCount)
    
    Dim Column As ADOX.Column
    Dim FieldIndex As Long
    For FieldIndex = 1 To FieldCount
        '@Ignore IndexedDefaultMemberAccess
        Set Column = Table.Columns(FieldIndex - 1)
        FieldNames(FieldIndex) = Column.Name
        FieldTypes(FieldIndex) = Column.Type
        '@Ignore IndexedDefaultMemberAccess
        FieldMap(Column.Name) = FieldIndex
    Next FieldIndex
End Sub
