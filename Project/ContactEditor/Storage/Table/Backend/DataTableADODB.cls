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
    ADO As ADOlib
    AdoCommand As ADODB.Command
    AdoRst As ADODB.Recordset
    ConnectionString As String
    TableName As String
    FieldNames() As String
    FieldTypes() As ADODB.DataTypeEnum
    FieldMap As Scripting.Dictionary
    IDs As Variant
    TypeCast As FieldFormat
End Type
Private this As TDataTable


Private Sub Class_Initialize()
    Set this.FieldMap = New Scripting.Dictionary
    this.FieldMap.CompareMode = TextCompare
End Sub


Private Sub Class_Terminate()
    Set this.Model = Nothing
    Set this.SQL = Nothing
    Set this.ADO = Nothing
    Set this.FieldMap = Nothing
    Set this.AdoRst = Nothing
    On Error Resume Next
    this.AdoCommand.ActiveConnection.Close
    On Error GoTo 0
End Sub


'''' Note, the factory intentionally returns the default interface. The
'''' interface of the instance's reference is switched to IDataTableStorage
'''' when the returned value is assigned in the calling code to appropriately
'''' declared variable.
'@Description "Returns a new DataTableADODB object."
Public Function Create(ByVal Model As DataTableModel, _
                       ByVal ConnectionString As String, _
                       ByVal TableName As String) As DataTableADODB
    Dim Instance As DataTableADODB
    Set Instance = New DataTableADODB
    Instance.Init Model, ConnectionString, TableName
    Set Create = Instance
End Function


Friend Sub Init(ByVal Model As DataTableModel, _
                ByVal ConnectionString As String, _
                ByVal TableName As String)
    Guard.NullReference Model
    Guard.EmptyString ConnectionString
    
    Set this.Model = Model
    Set this.SQL = SQLlib.Create(TableName)
    Set this.ADO = ADOlib.Create
    If LCase$(Left$(ConnectionString, 7)) = SQLITE_CONNSTR_PREFIX Then
        this.ConnectionString = this.ADO.GetSQLiteConnectionString(ConnectionString)("ADO")
    Else
        this.ConnectionString = ConnectionString
    End If
    this.ADO.SetConnectionString this.ConnectionString
    this.TableName = TableName
    
    this.ADO.GetTableMeta this.TableName, this.FieldNames, this.FieldTypes, this.FieldMap
End Sub


'@Ignore ProcedureNotUsed
Public Property Get SelfIDataTableStorage() As IDataTableStorage
    Set SelfIDataTableStorage = Me
End Property


'@Ignore ProcedureNotUsed
'@Description("Returns class reference.")
Public Property Get Class() As DataTableADODB
Attribute Class.VB_Description = "Returns class reference."
    Set Class = DataTableADODB
End Property


'@Ignore ProcedureNotUsed
Public Property Get AdoCommand() As ADODB.Command
    If this.AdoCommand Is Nothing Then
        Set AdoCommand = AdoCommandInit(this.SQL.SelectAll)
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
Public Sub SetTypeCast(Optional ByVal TypeCast As FieldFormat = FieldFormat.CastAsIs)
    this.TypeCast = TypeCast
End Sub


Public Function AdoCommandInit(ByVal SQLQuery As String, _
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


Public Function GetAdoRecordset(Optional ByVal SQLQuery As String = vbNullString) As ADODB.Recordset
    Set this.AdoRst = Nothing
    Set this.AdoRst = New ADODB.Recordset
    With this.AdoRst
        Set .Source = IIf(SQLQuery = vbNullString, this.AdoCommand, AdoCommandInit(SQLQuery))
        .CursorLocation = this.AdoCommand.ActiveConnection.CursorLocation
        .CursorType = adOpenStatic
        .LockType = adLockBatchOptimistic
        .CacheSize = 10
        .Open Options:=adAsyncFetch
    
        If .CursorLocation = ADODB.CursorLocationEnum.adUseClient Then
            Set .ActiveConnection = Nothing
        End If
    End With
    ADOlib.RecordsetToQT this.AdoRst, DataTableADODBxData.Range("A1")
    Set GetAdoRecordset = this.AdoRst
End Function


Public Function Records(Optional ByVal SQLQuery As String = vbNullString) As Variant
    Dim Rst As ADODB.Recordset
    Set Rst = GetAdoRecordset(SQLQuery)
    Records = ArrayLib.TransposeArray(Rst.GetRows, OutputArrBase:=1)
End Function


Public Function RecordsAsText() As Variant
    Dim Rst As ADODB.Recordset
    Set Rst = GetAdoRecordset(this.SQL.SelectAllAsText(this.FieldNames, this.FieldTypes))
    RecordsAsText = ArrayLib.TransposeArray(Rst.GetRows, OutputArrBase:=1)
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
        
        Dim IDs As Variant
        IDs = ArrayLib.GetColumn(.Values, ColumnNumber:=1, OutputArrBase:=1)
        this.IDs = IDs
        
        Dim RecordCount As Long
        RecordCount = UBound(IDs)
        Dim RecordIndex As Long
        For RecordIndex = 1 To RecordCount
            .IdIndices(IDs(RecordIndex)) = RecordIndex
        Next RecordIndex
    End With
End Sub


Private Function IDataTableStorage_GetIds() As Variant
    IDataTableStorage_GetIds = this.IDs
End Function


Private Function IDataTableStorage_GetColumnValues(ByVal FieldName As String) As Variant
    IDataTableStorage_GetColumnValues = ArrayLib.GetColumn( _
        this.Model.Values, _
        ColumnNumber:=this.FieldMap(FieldName), _
        OutputArrBase:=1 _
    )
End Function


Private Sub IDataTableStorage_SaveDataFromModel()
    If Not this.Model.IsDirty Then Exit Sub
    
    Dim AdoCmd As ADODB.Command
    Set AdoCmd = AdoCommandInit(this.SQL.UpdateSingleRecord(this.FieldNames))
    this.ADO.MakeAdoParamsForRecordUpdate this.FieldNames, this.FieldTypes, AdoCmd

    Dim RecordsAffected As Long: RecordsAffected = 0
    Dim Record As Scripting.Dictionary
    Set Record = New Scripting.Dictionary
    Record.CompareMode = TextCompare
    With this.Model
        AdoCmd.ActiveConnection.BeginTrans
        Dim RecordId As Variant
        For Each RecordId In .DirtyRecords.Keys
            .CopyRecordToDictionary Record, RecordId
            this.ADO.RecordToAdoParams Record, AdoCmd
            AdoCmd.Execute RecordsAffected, , adExecuteNoRecords
        Next RecordId
        AdoCmd.ActiveConnection.CommitTrans
        .IsNotDirty
    End With
End Sub
