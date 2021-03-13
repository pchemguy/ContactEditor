VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder "ContactEditor.Storage.Table.SQL"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

Private Type TSQL
    TableName As String
End Type
Private this As TSQL


Public Property Get TableName() As String
    TableName = this.TableName
End Property


Public Function Create(ByVal TableName As String) As SQLlib
    Dim SQLTool As SQLlib
    Set SQLTool = New SQLlib
    SQLTool.Init TableName
    Set Create = SQLTool
End Function


Public Sub Init(ByVal TableName As String)
    Guard.EmptyString TableName
    this.TableName = TableName
End Sub


Public Function SelectAll(Optional ByVal FieldArray As Variant = Empty) As String
    Dim Fields As String
    If (VarType(FieldArray) And VBA.VbVarType.vbArray) = VBA.VbVarType.vbArray Then
        Fields = Join(FieldArray, ", ")
    ElseIf Len(FieldArray) > 0 Then
        Fields = FieldArray
    Else
        Fields = "*"
    End If
    SelectAll = "SELECT " & Fields & " FROM """ & this.TableName & """"
End Function


Public Function SelectOne() As String
    SelectOne = "SELECT * FROM """ & this.TableName & """ LIMIT 1"
End Function


Public Function AsText(ByVal FieldName As String) As String
    Guard.EmptyString FieldName
    AsText = "CAST(" & FieldName & " AS TEXT) AS " & FieldName
End Function


Public Function SelectIdAsText(ByVal FieldNames As Variant) As String
    Dim TypeCheck As Boolean
    TypeCheck = ((VarType(FieldNames) And VBA.VbVarType.vbArray) = VBA.VbVarType.vbArray)
    Guard.Expression TypeCheck, "SQLlib", "FieldNames is not array"
    
    Dim Fields As Variant: Fields = FieldNames
    Dim IdIndex As Long: IdIndex = LBound(FieldNames, 1)
    Fields(IdIndex) = AsText(Fields(IdIndex))
    SelectIdAsText = SelectAll(Fields)
End Function


Public Function SelectAllAsText(ByVal FieldNames As Variant, ByVal FieldTypes As Variant) As String
    Dim TypeCheck As Boolean
    TypeCheck = ((VarType(FieldNames) And VBA.VbVarType.vbArray) = VBA.VbVarType.vbArray)
    Guard.Expression TypeCheck, "SQLlib", "FieldNames is not array"
    TypeCheck = ((VarType(FieldTypes) And VBA.VbVarType.vbArray) = VBA.VbVarType.vbArray)
    Guard.Expression TypeCheck, "SQLlib", "FieldTypes is not array"
    Guard.Expression LBound(FieldNames) = LBound(FieldTypes), "SQLlib", "FieldNames and FieldTypes array base mismatch"
    Guard.Expression UBound(FieldNames) = UBound(FieldTypes), "SQLlib", "FieldNames and FieldTypes array size mismatch"
    
    Dim Fields() As Variant: ReDim Fields(LBound(FieldNames) To UBound(FieldNames))
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        Select Case FieldTypes(FieldIndex)
            Case ADODB.DataTypeEnum.adVarWChar
                Fields(FieldIndex) = FieldNames(FieldIndex)
            Case Else
                Fields(FieldIndex) = AsText(FieldNames(FieldIndex))
        End Select
    Next FieldIndex
    SelectAllAsText = SelectAll(Fields)
End Function


Public Function GetSQLiteConnectionString(Optional ByVal DBPath As String = vbNullString, _
                                          Optional ByVal Features As String = vbNullString) As Scripting.Dictionary
    Dim Driver As String
    Driver = "SQLite3 ODBC Driver"
    
    Dim Database As String
    Database = Right$(DBPath, Abs(Len(DBPath) - 7))
    Database = VerifyOrGetDefaultPath(Database, Array("db", "sqlite"))
    
    Dim Options As String
    Options = IIf(Len(Features) > 0, Features, "SyncPragma=NORMAL;LongNames=True;NoCreat=True;FKSupport=True;OEMCP=True;")
    
    Dim ConnectionString As Scripting.Dictionary
    Set ConnectionString = New Scripting.Dictionary
    ConnectionString.CompareMode = TextCompare
    
    ConnectionString("ADO") = "Driver=" & Driver & ";" & "Database=" & Database & ";" & Options
    ConnectionString("QT") = "OLEDB;" + ConnectionString("ADO")
    
    Set GetSQLiteConnectionString = ConnectionString
End Function

