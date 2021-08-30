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
'@ModuleDescription("A thin wrapper around an ADODB recordset.")
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
End Type
Private this As TRecordset


'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As DbRecordset
Attribute Class.VB_Description = "Returns class reference"
    Set Class = DbRecordset
End Property


Public Function Create(ByVal cmd As IDbCommand, _
              Optional ByVal Scalar As Boolean = False, _
              Optional ByVal Disconnected As Boolean = True, _
              Optional ByVal CacheSize As Long = 10, _
              Optional ByVal CursorType As ADODB.CursorTypeEnum = -1, _
              Optional ByVal AsyncMode As Boolean = False, _
              Optional ByVal AsyncOption As ADODB.ExecuteOptionEnum = 0) As IDbRecordset
    Dim Instance As DbRecordset
    Set Instance = New DbRecordset
    Instance.Init cmd, Scalar, Disconnected, CacheSize, CursorType, AsyncMode, AsyncOption
    Set Create = Instance
End Function


'@Description("Default constructor")
Friend Sub Init(ByVal cmd As IDbCommand, _
       Optional ByVal Scalar As Boolean = False, _
       Optional ByVal Disconnected As Boolean = True, _
       Optional ByVal CacheSize As Long = 10, _
       Optional ByVal CursorType As ADODB.CursorTypeEnum = -1, _
       Optional ByVal AsyncMode As Boolean = False, _
       Optional ByVal AsyncOption As ADODB.ExecuteOptionEnum = 0)
Attribute Init.VB_Description = "Default constructor"
    
    Set this.cmd = cmd
    If Scalar Then AdoRecordset.MaxRecords = 1
    
    If Disconnected Then
        AdoRecordset.CursorLocation = adUseClient
        AdoRecordset.CursorType = adOpenStatic
    Else
        AdoRecordset.CursorLocation = adUseServer
        AdoRecordset.CursorType = adOpenForwardOnly
    End If
    
    AdoRecordset.CacheSize = CacheSize
    If CursorType > 0 Then
        AdoRecordset.CursorType = CursorType
    End If
    this.OpenOptions = AsyncOption Or (adAsyncFetch And AsyncMode)
End Sub


Private Sub Class_Initialize()
    Set AdoRecordset = New ADODB.Recordset
End Sub


Private Sub Class_Terminate()
    Set AdoRecordset = Nothing
End Sub


'@Description "Outputs Recordset to Excel Worksheet via QueryTable"
Public Sub RecordsetToQT(ByVal OutputRange As Excel.Range)
Attribute RecordsetToQT.VB_Description = "Outputs Recordset to Excel Worksheet via QueryTable"
    Guard.NullReference OutputRange
    
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
    FieldsCount = AdoRecordset.Fields.Count
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
    
    Set WSQueryTable = QTs.Add(Connection:=AdoRecordset, Destination:=QTRange.Range("A1"))
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
End Sub


Private Sub IDbRecordset_RecordsetToQT(ByVal OutputRange As Excel.Range)
    RecordsetToQT OutputRange
End Sub


Private Property Get IDbRecordset_Self() As IDbRecordset
    Set IDbRecordset_Self = Me
End Property


'@Description("Returns class reference")
Private Property Get IDbRecordset_Class() As DbRecordset
Attribute IDbRecordset_Class.VB_Description = "Returns class reference"
    Set IDbRecordset_Class = DbRecordset
End Property


Private Function IDbRecordset_AdoRecordset(ByVal SQL As String, ParamArray args() As Variant) As ADODB.Recordset
    If SQL <> vbNullString And AdoRecordset.ActiveCommand Is Nothing Then
        Dim localArgs() As Variant
        localArgs = UnfoldParamArray(args)
        
        Set AdoRecordset.Source = this.cmd.AdoCommand(SQL, localArgs)
    End If
    Set IDbRecordset_AdoRecordset = AdoRecordset
End Function


' Execute and ExecuteScalar can be combined into one method returning Variant, where distinction
' is made based on the AdoRecordset.MaxRecords value (1 - Scalar, regular otherwise)
Private Function IDbRecordset_OpenRecordset(ByVal SQL As String, ParamArray args() As Variant) As ADODB.Recordset
    '@Ignore FunctionReturnValueDiscarded
    IDbRecordset_AdoRecordset SQL, args
    
    On Error GoTo RecordsetOpenError
    AdoRecordset.Open Options:=this.OpenOptions
    On Error GoTo 0
    If AdoRecordset.CursorLocation = adUseClient Then Set AdoRecordset.ActiveConnection = Nothing
    Set IDbRecordset_OpenRecordset = AdoRecordset
    Exit Function
    
RecordsetOpenError:
    Err.Raise Err.Number, _
              Err.Source, _
              "DbRecordset->IDbRecordset_OpenRecordset->AdoRecordset.Open: " & Err.Description, _
              Err.HelpFile, _
              Err.HelpContext
End Function


Private Function IDbRecordset_OpenScalar(ByVal SQL As String, ParamArray args() As Variant) As Variant
    '@Ignore FunctionReturnValueDiscarded
    IDbRecordset_AdoRecordset SQL, args
    
    On Error GoTo RecordsetOpenError
    AdoRecordset.Open Options:=this.OpenOptions
    On Error GoTo 0
    If AdoRecordset.CursorLocation = adUseClient Then Set AdoRecordset.ActiveConnection = Nothing
    IDbRecordset_OpenScalar = AdoRecordset.Fields.Item(0).Value
    Exit Function

RecordsetOpenError:
    Err.Raise Err.Number, _
              Err.Source, _
              "DbRecordset->IDbRecordset_OpenScalar->AdoRecordset.Open: " & Err.Description, _
              Err.HelpFile, _
              Err.HelpContext
End Function


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
