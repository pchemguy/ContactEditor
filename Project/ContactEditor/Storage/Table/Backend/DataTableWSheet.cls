VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTableWSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Abstracts Excel worksheet"
'@Folder "ContactEditor.Storage.Table.Backend"
'@ModuleDescription "Abstracts Excel worksheet"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess
'@Exposed
Option Explicit

Implements IDataTableStorage

Private Type TDataTable
    Model As DataTableModel
    TableRange As Excel.Range
    TableBodyRange As Excel.Range
    TableIdColumn As Excel.Range
End Type
Private this As TDataTable


'@Description "Returns a new IDataTable object."
Public Function Create(ByVal Model As DataTableModel, ByVal ConnectionString As String, ByVal TableName As String) As IDataTableStorage
Attribute Create.VB_Description = "Returns a new IDataTable object."
    Dim result As DataTableWSheet
    Set result = New DataTableWSheet
    result.Init Model, ConnectionString, TableName
    Set Create = result
End Function


Public Sub Init(ByVal Model As DataTableModel, ByVal ConnectionString As String, ByVal TableName As String)
    Set this.Model = Model
    Dim BookSheet As Variant: BookSheet = Split(ConnectionString, "!")
    Guard.Expression UBound(BookSheet) = 1, "DataRecordWSheet", "Invalid connection string: " & ConnectionString
    On Error Resume Next
    Dim WSheet As Excel.Worksheet
    Set WSheet = Application.Workbooks(BookSheet(0)).Worksheets(BookSheet(1))
    Select Case Err.Number
        Case ErrNo.SubscriptOutOfRange
            Err.Clear
            Guard.Expression False, "DataTableWSheet constructor", "Invalid connection string"
            Exit Sub
        Case Is > 0
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
    With WSheet
        Set this.TableRange = .Range(TableName)
        Set this.TableBodyRange = .Range(TableName & "Body")
        Set this.TableIdColumn = .Range(TableName & "Id")
    End With
    Select Case Err.Number
        Case ErrNo.PassedNoErr
            On Error GoTo 0
        Case 1004
            Err.Clear
            Guard.Expression False, "DataTableWSheet constructor", "Verify named ranges"
            Exit Sub
        Case Else
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
End Sub


Private Sub Class_Terminate()
    Set this.Model = Nothing
End Sub


Private Sub IDataTableStorage_LoadDataIntoModel()
    With this.Model
        '''' The top row of the table should contain field names.
        '''' Loop through the first row to construct a 1D array of field names
        '''' and a map FieldName -> ColumnIndex
        .FieldIndices.RemoveAll
        
        Const HEADER_ROW_INDEX As Long = 1
        .FieldNames = Application.WorksheetFunction.Index(this.TableRange.Value, HEADER_ROW_INDEX)
    
        Dim ColumnCount As Long: ColumnCount = UBound(.FieldNames, 1)
        Dim ColumnIndex As Long
        For ColumnIndex = 1 To ColumnCount
            .FieldIndices(CStr(.FieldNames(ColumnIndex))) = ColumnIndex
        Next ColumnIndex
        
        Dim Ids As Variant: Ids = Application.WorksheetFunction.Transpose(this.TableIdColumn.Value)
        Dim RecordCount As Long: RecordCount = UBound(Ids, 1)
        Dim RecordIndex As Long
        For RecordIndex = 1 To RecordCount
            .IdIndices(CStr(Ids(RecordIndex))) = RecordIndex
        Next RecordIndex
        
        '''' Data records start from the second row. Get a range reference to the
        '''' data area of the table and set the records field on the model as 2D
        '''' array (ArrayOfRecords(ArrayOfFields))
        .Values = this.TableBodyRange.Value
        'Dim TopLeftCell As Excel.Range: Set TopLeftCell = this.TableRange.Offset(1, 0)
        'Dim height As Long: height = this.TableRange.Rows.Count - 1
        'Dim width As Long: width = this.TableRange.Columns.Count
        'this.Model.Records = TopLeftCell.Resize(height, width).Value
    End With
End Sub


Private Sub IDataTableStorage_SaveDataFromModel()
    With this.Model
        If Not .IsDirty Then Exit Sub
        Dim RecordId As Variant
        Dim RecordIndex As Long
        Dim RecordRange As Excel.Range
        Dim FieldIndex As Long
        For Each RecordId In .DirtyRecords.Keys
            RecordIndex = this.Model.RecordIndexFromId(RecordId)
            Set RecordRange = this.TableBodyRange.Rows(RecordIndex)
            For FieldIndex = 1 To UBound(this.Model.FieldNames, 1)
                RecordRange.Cells(1, FieldIndex) = .Values(RecordIndex, FieldIndex)
            Next FieldIndex
        Next RecordId
        .IsNotDirty
    End With
End Sub


Private Function IDataTableStorage_GetIds() As Variant
    '''' Ids should be strings (cast if necessary), not numeric
    Dim Ids() As Variant
    ReDim Ids(1 To this.Model.IdIndices.Count)
    Dim IdColumn As Variant: IdColumn = Application.WorksheetFunction.Transpose(this.TableIdColumn)
    Dim IdIndex As Long
    For IdIndex = 1 To UBound(Ids)
        Ids(IdIndex) = CStr(IdColumn(IdIndex))
    Next IdIndex
    IDataTableStorage_GetIds = Ids
End Function


Private Function IDataTableStorage_GetColumnValues(ByVal FieldName As String) As Variant
    Dim FieldIndex As Long: FieldIndex = this.Model.FieldIndices(FieldName)
    Dim FieldRange As Excel.Range
    Set FieldRange = this.TableBodyRange.Offset(0, FieldIndex - 1).Resize(this.TableBodyRange.Rows.Count, 1)
    IDataTableStorage_GetColumnValues = Application.WorksheetFunction.Transpose(FieldRange)
End Function
