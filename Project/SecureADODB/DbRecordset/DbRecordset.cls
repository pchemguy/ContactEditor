VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A thin wrapper around an ADODB recordset."
'@Folder "SecureADODB.DbRecordset"
'@ModuleDescription "A thin wrapper around an ADODB recordset."
'@PredeclaredId
'@Exposed
Option Explicit

Implements IDbRecordset

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents AdoRecordset As ADODB.Recordset
Attribute AdoRecordset.VB_VarHelpID = -1
Private Type TRecordset
    cmd As IDbCommand
    OpenOptions As Long
    DataQT As Excel.QueryTable
End Type
Private this As TRecordset


Public Function Create(ByVal cmd As IDbCommand, _
              Optional ByVal Disconnected As Boolean = True, _
              Optional ByVal CacheSize As Long = 10, _
              Optional ByVal CursorType As ADODB.CursorTypeEnum = -1, _
              Optional ByVal LockType As ADODB.LockTypeEnum = adLockReadOnly, _
              Optional ByVal AsyncMode As Boolean = False, _
              Optional ByVal AsyncOption As ADODB.ExecuteOptionEnum = 0) As IDbRecordset
    Dim Instance As DbRecordset
    Set Instance = New DbRecordset
    Instance.Init cmd, Disconnected, CacheSize, CursorType, LockType, AsyncMode, AsyncOption
    Set Create = Instance
End Function


'''' For updatable recordset use LockType = adLockBatchOptimistic
'@Description("Default constructor")
Friend Sub Init(ByVal cmd As IDbCommand, _
            Optional ByVal Disconnected As Boolean = True, _
            Optional ByVal CacheSize As Long = 10, _
            Optional ByVal CursorType As ADODB.CursorTypeEnum = -1, _
            Optional ByVal LockType As ADODB.LockTypeEnum = adLockReadOnly, _
            Optional ByVal AsyncMode As Boolean = False, _
            Optional ByVal AsyncOption As ADODB.ExecuteOptionEnum = 0)
Attribute Init.VB_Description = "Default constructor"
    Set AdoRecordset = New ADODB.Recordset
    
    Set this.cmd = cmd
    
    If Disconnected Then
        AdoRecordset.CursorLocation = adUseClient
        AdoRecordset.CursorType = adOpenStatic
    Else
        AdoRecordset.CursorLocation = adUseServer
        AdoRecordset.CursorType = adOpenForwardOnly
    End If
    
    AdoRecordset.LockType = LockType
    
    AdoRecordset.CacheSize = CacheSize
    If CursorType > 0 Then
        AdoRecordset.CursorType = CursorType
    End If
    this.OpenOptions = AsyncOption Or (adAsyncFetch And AsyncMode)
End Sub


'@Description "Outputs Recordset to Excel Worksheet via QueryTable"
Friend Function RecordsetToQT(ByVal OutputRange As Excel.Range, _
                     Optional ByVal AdoRst As ADODB.Recordset = Nothing) As Excel.QueryTable
Attribute RecordsetToQT.VB_Description = "Outputs Recordset to Excel Worksheet via QueryTable"
    Guard.NullReference OutputRange
    
    Dim rst As ADODB.Recordset
    Set rst = IIf(AdoRst Is Nothing, AdoRecordset, AdoRst)
    rst.MoveFirst
    
    Dim QTs As Excel.QueryTables
    Set QTs = OutputRange.Worksheet.QueryTables
    
    '''' Cleans up target area before binding the data.
    '''' Provided range reference used to indicate the left column and
    '''' Recordset.Fields.Count determines the width.
    '''' If EntireColumn.Delete method is used, Range object becomes invalid, so
    '''' a textual address must be saved to reset the Range reference.
    '''' However, when multiple QTs are bound to the same worksheet,
    '''' EntireColumn.Delete shifts columns to the left, so the target range
    '''' may not be clear. EntireColumn.Clear clears the contents.
    Dim FieldsCount As Long
    FieldsCount = rst.Fields.Count
    Dim QTRangeAddress As String
    QTRangeAddress = OutputRange.Address(External:=True)
    Dim QTRange As Excel.Range
    '@Ignore ImplicitActiveSheetReference: Fully qualified range object is expected
    Set QTRange = Range(QTRangeAddress)
    QTRange.Resize(1, FieldsCount).EntireColumn.Clear
    '@Ignore ImplicitActiveSheetReference: Fully qualified range object is expected
    Set QTRange = Range(QTRangeAddress)
    
    Dim WSQueryTable As Excel.QueryTable
    For Each WSQueryTable In QTs
        WSQueryTable.Delete
    Next WSQueryTable
    
    Dim NamedRange As Excel.Name
    For Each NamedRange In QTRange.Worksheet.Names
        NamedRange.Delete
    Next NamedRange
    
    Set WSQueryTable = QTs.Add(Connection:=rst, Destination:=QTRange.Range("A1"))
    With WSQueryTable
        .FieldNames = True
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .EnableEditing = True
    End With
    WSQueryTable.Refresh
    '@Ignore IndexedDefaultMemberAccess
    QTRange.Worksheet.UsedRange.Rows(1).HorizontalAlignment = xlCenter
    
    '''' The same recordset object cannot be reused on the same worksheet:
    '''' outputs headers only, but no data. The source of the issue is not clear.
    '''' If this.DataQT is not set, set it. If set, update the reference and copy
    '''' the data (as a workaround).
    If Not this.DataQT Is Nothing Then
        rst.MoveFirst
        WSQueryTable.ResultRange.Range("A2").CopyFromRecordset rst
    End If
    WSQueryTable.ResultRange.CurrentRegion.Columns.AutoFit
    
    Set this.DataQT = WSQueryTable
    Set RecordsetToQT = this.DataQT
End Function


'@Description "Sets AdoCommand as to AdoRecordset.Source in preparation for .OpenXXX"
Friend Sub SetSource( _
            ByVal SQL As String, _
            ParamArray ADODBParamsValues() As Variant)
Attribute SetSource.VB_Description = "Sets AdoCommand as to AdoRecordset.Source in preparation for .OpenXXX"
    Dim localArgs() As Variant
    localArgs = UnfoldParamArray(ADODBParamsValues)
    If AdoRecordset.State <> adStateClosed Then AdoRecordset.Close
    Set AdoRecordset.Source = this.cmd.AdoCommand(SQL, localArgs)
End Sub


Private Function IDbRecordset_RecordsetToQT(ByVal OutputRange As Excel.Range) As Excel.QueryTable
    Set IDbRecordset_RecordsetToQT = RecordsetToQT(OutputRange)
End Function


Private Property Get IDbRecordset_cmd() As IDbCommand
    Set IDbRecordset_cmd = this.cmd
End Property


Private Property Get IDbRecordset_AdoRecordset() As ADODB.Recordset
    Set IDbRecordset_AdoRecordset = AdoRecordset
End Property


Private Function IDbRecordset_GetAdoRecordset( _
            ByVal SQL As String, _
            ParamArray ADODBParamsValues() As Variant) As ADODB.Recordset
    Dim localArgs() As Variant
    localArgs = UnfoldParamArray(ADODBParamsValues)
    SetSource SQL, localArgs
    Set IDbRecordset_GetAdoRecordset = AdoRecordset
End Function


' Execute and ExecuteScalar can be combined into one method returning Variant, where distinction
' is made based on the AdoRecordset.MaxRecords value (1 - Scalar, regular otherwise)
Private Function IDbRecordset_OpenRecordset( _
            ByVal SQL As String, _
            ParamArray ADODBParamsValues() As Variant) As ADODB.Recordset
    Dim localArgs() As Variant
    localArgs = UnfoldParamArray(ADODBParamsValues)
    SetSource SQL, localArgs
    
    On Error GoTo RecordsetOpenError
    With AdoRecordset
        .MaxRecords = 0
        .Open Options:=this.OpenOptions
        If .CursorLocation = adUseClient Then Set .ActiveConnection = Nothing
    End With
    On Error GoTo 0
    Set IDbRecordset_OpenRecordset = AdoRecordset
    Exit Function
    
RecordsetOpenError:
    Err.Raise Err.Number, _
              Err.Source, _
              "IDbRecordset->OpenRecordset->AdoRecordset.Open: " & Err.Description, _
              Err.HelpFile, _
              Err.HelpContext
End Function


Private Function IDbRecordset_OpenScalar( _
            ByVal SQL As String, _
            ParamArray ADODBParamsValues() As Variant) As Variant
    Dim localArgs() As Variant
    localArgs = UnfoldParamArray(ADODBParamsValues)
    SetSource SQL, localArgs
        
    On Error GoTo RecordsetOpenError
    With AdoRecordset
        .MaxRecords = 1
        .Open Options:=this.OpenOptions
        If .CursorLocation = adUseClient Then Set .ActiveConnection = Nothing
    End With
    On Error GoTo 0
    
    IDbRecordset_OpenScalar = AdoRecordset.Fields.Item(0).Value
    Exit Function

RecordsetOpenError:
    Err.Raise Err.Number, _
              Err.Source, _
              "DbRecordset->IDbRecordset_OpenScalar->AdoRecordset.Open: " & Err.Description, _
              Err.HelpFile, _
              Err.HelpContext
End Function


'''' Updates record values
''''
'''' Args:
''''   AbsolutePosition (Long):
''''     Recordset.AbsolutePosition identifying the target record
''''
''''   ValuesDict (Dictionary):
''''     FieldName -> Value map
''''
'@Description "Updates modified records"
Private Sub IDbRecordset_UpdateRecord( _
            ByVal AbsolutePosition As Long, _
            ByVal ValuesDict As Scripting.Dictionary)
Attribute IDbRecordset_UpdateRecord.VB_Description = "Updates modified records"
    Dim NumRecords As Long
    
    With AdoRecordset
        Guard.NullReference ValuesDict
        Guard.ExpressionErr .LockType = adLockBatchOptimistic, _
                            AdoFeatureNotAvailableErr, _
                            "DbRecordset", _
                            "Set LockType = adLockBatchOptimistic"
        Guard.ExpressionErr .State = adStateOpen, _
                            IncompatibleStatusErr, _
                            "DbRecordset", _
                            "Expected AdoRecordset.Status = adStateOpen"
        '@Ignore ValueRequired: false positive
        Guard.ExpressionErr AbsolutePosition <= .RecordCount, _
                            InvalidParameterErr, _
                            "DbRecordset", _
                            "AbsolutePosition must be <= AdoRecordset.RecordCount"
        
        '@Ignore ValueRequired: False positive
        NumRecords = AbsolutePosition - .AbsolutePosition
        '@Ignore ArgumentWithIncompatibleObjectType: False positive
        .Move NumRecords
        
        Dim FieldName As Variant
        For Each FieldName In ValuesDict.Keys
            '@Ignore ImplicitDefaultMemberAccess, IndexedDefaultMemberAccess
            .Fields(FieldName) = ValuesDict(FieldName)
        Next FieldName
    End With
End Sub


'@Description "Updates AdoRecordset data in preparation for a database update"
Friend Sub UpdateRecordsetData(ByRef AbsolutePositions() As Long, _
                               ByRef RecordsetData() As Variant)
Attribute UpdateRecordsetData.VB_Description = "Updates AdoRecordset data in preparation for a database update"
    Dim FieldCount As Long
    FieldCount = UBound(RecordsetData, 2) - LBound(RecordsetData, 2) + 1
    Dim RecordCount As Long
    RecordCount = UBound(RecordsetData, 1) - LBound(RecordsetData, 1) + 1
    
    Guard.ExpressionErr AdoRecordset.Fields.Count = FieldCount, _
                        IncompatibleArraysErr, "DbRecordset", _
                        "Field count mismatch"
    '@Ignore ValueRequired: false positive
    Guard.ExpressionErr AdoRecordset.RecordCount = RecordCount, _
                        IncompatibleArraysErr, "DbRecordset", _
                        "Record count mismatch"
    Guard.ExpressionErr LBound(RecordsetData, 1) = 1, IncompatibleArraysErr, _
                        "DbRecordset", "Records dimension should be 1-based"
    Guard.ExpressionErr LBound(RecordsetData, 2) = 1, IncompatibleArraysErr, _
                        "DbRecordset", "Fields dimension should be 1-based"
    Guard.ExpressionErr AbsolutePositions(UBound(AbsolutePositions)) <= RecordCount, _
                        SubscriptOutOfRange, "DbRecordset", _
                        "Record position out of range"

    Dim RecordPos As Long  '''' Current AbsolutePosition
    Dim FieldIndex As Long
    Dim RecordPosIndex As Long '''' Index of AbsolutePosition in dirty records
    Dim NumRecords As Long '''' Relative cursor shift for the .Move method
    
    '''' Initialize RecordPos to current recordset position. Since setting
    '''' the .AbsolutePosition attribute directly invalidates cache, use
    '''' this value to calculate relative shift for the .Move method
    AdoRecordset.MoveFirst
    '@Ignore ValueRequired: false positive
    RecordPos = AdoRecordset.AbsolutePosition
    '''' Loop through the list of dirty record indices
    For RecordPosIndex = LBound(AbsolutePositions) To UBound(AbsolutePositions)
        NumRecords = AbsolutePositions(RecordPosIndex) - RecordPos
        RecordPos = AbsolutePositions(RecordPosIndex)
        '@Ignore ArgumentWithIncompatibleObjectType: False positive
        AdoRecordset.Move NumRecords
        '''' Update field values
        For FieldIndex = 1 To FieldCount
            '@Ignore ImplicitDefaultMemberAccess, IndexedDefaultMemberAccess
            AdoRecordset.Fields(FieldIndex - 1) = RecordsetData(RecordPos, FieldIndex)
        Next FieldIndex
    Next RecordPosIndex
End Sub


'@Description "Executes transaction wrapped batch update of the recordset."
Friend Sub PersistRecordsetChanges(ByVal DirtyRecordCount As Long)
Attribute PersistRecordsetChanges.VB_Description = "Executes transaction wrapped batch update of the recordset."
    With AdoRecordset
        Guard.ExpressionErr .State = adStateOpen, _
                            IncompatibleStatusErr, _
                            "DbRecordset", _
                            "Expected AdoRecordset.Status = adStateOpen"
                            
        Dim db As IDbConnection
        Set db = this.cmd.Connection
        '''' Marshal dirty records only
'        .MarshalOptions = adMarshalModifiedOnly
        .MarshalOptions = adMarshalAll
        Set .ActiveConnection = this.cmd.Connection.AdoConnection
        On Error GoTo Rollback
        '''' Set the expected count of affected rows in the DbConnection object
        db.ExpectedRecordsAffected = DirtyRecordCount
        '''' Wrap update in a transaction
        db.BeginTransaction
        .UpdateBatch
        db.CommitTransaction
        On Error GoTo 0
        If .CursorLocation = adUseClient Then Set .ActiveConnection = Nothing
    End With
    
    Exit Sub
    
Rollback:
    this.cmd.Connection.RollbackTransaction
    With Err
        .Raise .Number, .Source, .Description, .HelpFile, .HelpContext
    End With
End Sub


'''' Factory for file-based databases.
''''
'''' Args:
''''   AbsolutePositions (1D array of long):
''''     A 1D 1-based array containing absolute positions of dirty records in
''''     the recordset object (as in AdoRecordset.AbsolutePosition). The caller
''''     should sort it in ascending order
''''
''''   RecordsetData (2D variant array):
''''     2D record-wise 1-based array of recordset data
''''
'@Description "Updates recordset data from 2D array of field values and an array of dirty record indices"
Private Sub IDbRecordset_UpdateRecordset(ByRef AbsolutePositions() As Long, _
                                         ByRef RecordsetData() As Variant)
Attribute IDbRecordset_UpdateRecordset.VB_Description = "Updates recordset data from 2D array of field values and an array of dirty record indices"
    UpdateRecordsetData AbsolutePositions, RecordsetData
    Dim DirtyRecordsCount As Long
    DirtyRecordsCount = UBound(AbsolutePositions) - LBound(AbsolutePositions) + 1
    PersistRecordsetChanges DirtyRecordsCount
End Sub


'================================================================================'
'========================= AdoRecordset EVENT HANDLERS =========================='
'================================================================================'


' https://docs.microsoft.com/en-us/sql/ado/guide/data/ado-event-instantiation-visual-basic
' https://docs.microsoft.com/en-us/sql/ado/guide/data/ado-event-handler-summary
' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/fetchprogress-event-ado
'
'@Ignore ProcedureNotUsed, ParameterNotUsed, ParameterCanBeByVal
Private Sub AdoRecordset_FetchProgress(ByVal Progress As Long, _
                                       ByVal MaxProgress As Long, _
                                       ByRef adStatus As ADODB.EventStatusEnum, _
                                       ByVal pRecordset As ADODB.Recordset)
    Debug.Print "ADODB.Recordset fetch in progress. " & CStr(Progress) & _
                " records of " & CStr(MaxProgress) & "have been retrived..."
End Sub


' https://docs.microsoft.com/en-us/sql/ado/reference/ado-api/fetchcomplete-event-ado
'
'@Ignore ProcedureNotUsed, ParameterNotUsed, ParameterCanBeByVal
Private Sub AdoRecordset_FetchComplete(ByVal pError As ADODB.Error, _
                                       ByRef adStatus As ADODB.EventStatusEnum, _
                                       ByVal pRecordset As ADODB.Recordset)
    Debug.Print "ADODB.Recordset fetch completed. "
    If adStatus = ADODB.EventStatusEnum.adStatusErrorsOccurred And pError Is Nothing Then
        Debug.Print "Error: " & pError.Number, pError.Description, pError.Source, pError.Source
    End If
End Sub


'''@Ignore ProcedureNotUsed, ParameterNotUsed, ParameterCanBeByVal
''Private Sub AdoRecordset_RecordsetChangeComplete( _
''                ByVal adReason As ADODB.EventReasonEnum, _
''                ByVal pError As ADODB.Error, _
''                ByRef adStatus As ADODB.EventStatusEnum, _
''                ByVal pRecordset As ADODB.Recordset)
''    Debug.Print "RecordsetChangeComplete"
''End Sub
