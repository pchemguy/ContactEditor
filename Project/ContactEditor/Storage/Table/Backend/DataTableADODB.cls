VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTableADODB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Abstracts ADODB backend."
'@Folder "ContactEditor.Storage.Table.Backend"
'@ModuleDescription "Abstracts ADODB backend."
'@PredeclaredId
'@Exposed
'@IgnoreModule FunctionReturnValueDiscarded, IndexedDefaultMemberAccess
Option Explicit

Private Const SQLITE_CONNSTR_PREFIX As String = "sqlite:"

Implements IDataTableStorage

Private Type TDataTable
    Model As DataTableModel
    SQL As SQLlib
    AdoCommand As ADODB.Command
    ConnectionString As String
    TableName As String
    FieldNames() As String
    FieldTypes() As ADODB.DataTypeEnum
    FieldMap As Scripting.Dictionary
    Ids() As String
End Type
Private this As TDataTable


Private Sub Class_Initialize()
    Set this.FieldMap = New Scripting.Dictionary
    this.FieldMap.CompareMode = TextCompare
End Sub


Private Sub Class_Terminate()
    Set this.Model = Nothing
    Set this.SQL = Nothing
    Set this.FieldMap = Nothing
    On Error Resume Next
    this.AdoCommand.ActiveConnection.Close
    On Error GoTo 0
End Sub


'@Ignore ProcedureNotUsed
Public Property Get SelfIDataTableStorage() As IDataTableStorage
    Set SelfIDataTableStorage = Me
End Property


'@Description("Returns class reference.")
'@Ignore ProcedureNotUsed
Public Property Get Class() As DataTableADODB
Attribute Class.VB_Description = "Returns class reference."
    Set Class = DataTableADODB
End Property


Public Property Get AdoCommand() As ADODB.Command
    If this.AdoCommand Is Nothing Then
        Set AdoCommand = AdoCommandInit
    Else
        Set AdoCommand = this.AdoCommand
    End If
End Property


Public Property Get FieldNames() As Variant
    FieldNames = this.FieldNames
End Property
    
    
Public Property Get FieldTypes() As Variant
    FieldTypes = this.FieldTypes
End Property
    
    
Public Property Get FieldMap() As Scripting.Dictionary
    Set FieldMap = this.FieldMap
End Property


'@Ignore ProcedureNotUsed
'@Description "Returns a new IDataTable object."
Public Function Create(ByVal Model As DataTableModel, ByVal ConnectionString As String, ByVal TableName As String) As IDataTableStorage
Attribute Create.VB_Description = "Returns a new IDataTable object."
    Dim result As DataTableADODB
    Set result = New DataTableADODB
    result.Init Model, ConnectionString, TableName
    Set Create = result
End Function


'''' Creates a DataTableADODB instance with default interface, on which .Self can be use to access IDataTableStorage
'@Description "Returns a new DataTableADODB object."
Public Function CreateDefault(ByVal Model As DataTableModel, ByVal ConnectionString As String, ByVal TableName As String) As DataTableADODB
Attribute CreateDefault.VB_Description = "Returns a new DataTableADODB object."
    Dim result As DataTableADODB
    Set result = New DataTableADODB
    result.Init Model, ConnectionString, TableName
    Set CreateDefault = result
End Function


Public Sub Init(ByVal Model As DataTableModel, ByVal ConnectionString As String, ByVal TableName As String)
    Guard.NullReference Model
    Guard.EmptyString ConnectionString
    
    Set this.Model = Model
    Set this.SQL = SQLlib.Create(TableName)
    If LCase(Left$(ConnectionString, 7)) = SQLITE_CONNSTR_PREFIX Then
        this.ConnectionString = this.SQL.GetSQLiteConnectionString(ConnectionString)("ADO")
    Else
        this.ConnectionString = ConnectionString
    End If
    this.TableName = TableName
    
    '@Ignore FunctionReturnValueDiscarded
    AdoCommandInit this.SQL.SelectOne, ADODB.CursorLocationEnum.adUseClient
    If AdoCommand.ActiveConnection.State <> ADODB.ObjectStateEnum.adStateOpen Then
        Guard.Expression False, "DataTableADODB", "Test connection failed"
    End If
    
    CollectTableMetadata
End Sub


Private Sub CollectTableMetadata()
    Dim Fields As Fields: Set Fields = AdoRecordset.Fields
    Dim FieldCount As Long: FieldCount = Fields.Count
    Dim FieldIndex As Long
    With this
        ReDim .FieldNames(1 To FieldCount)
        ReDim .FieldTypes(1 To FieldCount)
        For FieldIndex = 1 To FieldCount
            .FieldNames(FieldIndex) = Replace(Fields(FieldIndex - 1).Name, .TableName & ".", vbNullString, 1, 1, vbTextCompare)
            .FieldTypes(FieldIndex) = Fields(FieldIndex - 1).Type
            .FieldMap(.FieldNames(FieldIndex)) = FieldIndex
        Next FieldIndex
    End With
End Sub


Public Function AdoCommandInit(Optional ByVal SQLQuery As String = vbNullString, _
                               Optional ByVal CursorLocation As ADODB.CursorLocationEnum = adUseClient) As ADODB.Command
    If Not this.AdoCommand Is Nothing Then
        On Error Resume Next
        this.AdoCommand.ActiveConnection.Close
        On Error GoTo 0
    End If
            
    Dim CommandText As String
    CommandText = IIf(Len(SQLQuery) > 0, SQLQuery, this.SQL.SelectAll)
    
    Set this.AdoCommand = New ADODB.Command
    With this.AdoCommand
        .CommandType = ADODB.CommandTypeEnum.adCmdText
        .CommandText = CommandText
        .Prepared = True
        .ActiveConnection = this.ConnectionString
        .ActiveConnection.CursorLocation = CursorLocation
    End With
    Set AdoCommandInit = this.AdoCommand
End Function


Public Function AdoRecordset(Optional ByVal SQLQuery As String = vbNullString) As ADODB.Recordset
    Dim Rst As ADODB.Recordset
    Set Rst = New ADODB.Recordset
    With Rst
        Set .source = IIf(SQLQuery = vbNullString, this.AdoCommand, AdoCommandInit(SQLQuery))
        .CursorLocation = this.AdoCommand.ActiveConnection.CursorLocation
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .CacheSize = 10
        .Open Options:=adAsyncFetch
    
        If .CursorLocation = ADODB.CursorLocationEnum.adUseClient Then
            Set .ActiveConnection = Nothing
        End If
    End With
    Set AdoRecordset = Rst
End Function


Public Function Records(Optional ByVal SQLQuery As String = vbNullString) As Variant
    Dim Rst As ADODB.Recordset
    Set Rst = AdoRecordset(SQLQuery)
    Records = Application.WorksheetFunction.Transpose(Rst.GetRows)
End Function


Public Function RecordsAsText() As Variant
    Dim Rst As ADODB.Recordset
    Set Rst = AdoRecordset(this.SQL.SelectAllAsText(this.FieldNames, this.FieldTypes))
    RecordsAsText = Application.WorksheetFunction.Transpose(Rst.GetRows)
End Function


Private Sub IDataTableStorage_LoadDataIntoModel()
    With this.Model
        .FieldIndices.RemoveAll
        
        .Values = RecordsAsText
        .FieldNames = this.FieldNames
        Dim FieldName As Variant
        For Each FieldName In this.FieldMap.Keys
            .FieldIndices(FieldName) = this.FieldMap(FieldName)
        Next FieldName
        
        Dim RecordCount As Long: RecordCount = UBound(.Values, 1)
        ReDim this.Ids(LBound(.Values, 1) To UBound(.Values, 1))
        Dim IdFieldIndex As Long: IdFieldIndex = 1
        Dim RecordIndex As Long
        For RecordIndex = 1 To RecordCount
            this.Ids(RecordIndex) = .Values(RecordIndex, IdFieldIndex)
            .IdIndices(this.Ids(RecordIndex)) = RecordIndex
        Next RecordIndex
    End With
End Sub


Private Function IDataTableStorage_GetIds() As Variant
    IDataTableStorage_GetIds = this.Ids
End Function


Private Function IDataTableStorage_GetColumnValues(ByVal FieldName As String) As Variant
    Dim FieldIndex As Long: FieldIndex = this.FieldMap(FieldName)
    Dim RecordIndex As Long
    Dim ColumnValues() As Variant
    ReDim ColumnValues(LBound(this.Model.Values, 1) To UBound(this.Model.Values, 1))
    For RecordIndex = LBound(ColumnValues) To UBound(ColumnValues)
        ColumnValues(RecordIndex) = this.Model.Values(RecordIndex, FieldIndex)
    Next RecordIndex
    IDataTableStorage_GetColumnValues = ColumnValues
End Function


Private Sub IDataTableStorage_SaveDataFromModel()
End Sub
