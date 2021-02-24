VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Abstracts Excel worksheet"
'@Folder "Storage.Table.Backend"
'@ModuleDescription "Abstracts Excel worksheet"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess
'@Exposed
Option Explicit

Implements IDataTable

Private Type TDataTable
    Model As DataTableModel
    TableRange As Excel.Range
End Type
Private this As TDataTable


'@Description "Returns a new IDataTable object."
Public Function Create(ByVal Model As DataTableModel) As IDataTable
Attribute Create.VB_Description = "Returns a new IDataTable object."
    Dim result As SheetTable
    Set result = New SheetTable
    result.Init Model
    Set Create = result
End Function


Public Sub Init(ByVal Model As DataTableModel)
    Set this.Model = Model
End Sub


Private Sub Class_Terminate()
    Set this.Model = Nothing
End Sub


'''' The top row of the table should contain field names.
'''' Loop through the first row to construct a 1D array of field names
'''' and a map FieldName -> ColumnIndex
Private Sub GetFields()
    this.Model.FieldMap.RemoveAll
    
    Dim ColumnCount As Long: ColumnCount = this.TableRange.Columns.Count
    Dim FieldNames() As Variant: ReDim FieldNames(1 To ColumnCount)
    Dim ColumnIndex As Long
    For ColumnIndex = 1 To ColumnCount
        FieldNames(ColumnIndex) = this.TableRange(1, ColumnIndex)
        this.Model.FieldMap(FieldNames(ColumnIndex)) = ColumnIndex
    Next ColumnIndex
    this.Model.FieldNames = FieldNames
End Sub


'''' Data records start from the second row. Get a range reference to the
'''' data area of the table and set the records field on the model as 2D
'''' array (ArrayOfRecords(ArrayOfFields))
Private Sub GetRecords()
    Dim TopLeftCell As Excel.Range: Set TopLeftCell = this.TableRange.Offset(1, 0)
    Dim height As Long: height = this.TableRange.Rows.Count - 1
    Dim width As Long: width = this.TableRange.Columns.Count
    this.Model.Records = TopLeftCell.Resize(height, width).Value
End Sub


Private Sub IDataTable_GetData(ByVal ConnectionString As String, ByVal TableName As String)
    Dim BookSheet As Variant: BookSheet = Split(ConnectionString, "!")
    Set this.TableRange = Excel.Workbooks(BookSheet(0)).Worksheets(BookSheet(1)).Range(TableName)
    
    GetFields
    GetRecords
End Sub
