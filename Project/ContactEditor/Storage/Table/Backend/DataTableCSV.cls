VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTableCSV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Abstracts CSV file."
'@Exposed
'@Folder("ContactEditor.Storage.Table.Backend")
'@ModuleDescription "Abstracts CSV file."
'@PredeclaredId
Option Explicit

Implements IDataTableStorage

Private Type TDataTable
    Model As DataTableModel
    CSVFileName As String
    FieldSeparator As String
End Type
Private this As TDataTable


' ConnectionString - "<Path to CSV>"; TableName - "<CSV file name>!sep=<FieldSeparator>"
'@Description "Returns a new IDataTable object."
Public Function Create(ByVal Model As DataTableModel, ByVal ConnectionString As String, ByVal TableName As String) As IDataTableStorage
    Dim result As DataTableCSV
    Set result = New DataTableCSV
    result.Init Model, ConnectionString, TableName
    Set Create = result
End Function


Public Sub Init(ByVal Model As DataTableModel, ByVal ConnectionString As String, ByVal TableName As String)
    Guard.EmptyString ConnectionString
    Guard.EmptyString TableName
    
    Set this.Model = Model
    Dim FileName As String
    If Left$(Right$(TableName, 6), 5) = "!sep=" Then
        this.FieldSeparator = Right$(TableName, 1)
        FileName = Left$(TableName, Len(TableName) - 6)
    Else
        this.FieldSeparator = ","
        FileName = TableName
    End If
    FileName = ConnectionString & Application.PathSeparator & FileName
    
    Dim FileExist As Variant
    On Error Resume Next
    FileExist = Dir$(FileName)
    On Error GoTo 0
    If Len(FileExist) = 0 Then
        VBA.Err.Raise number:=53, source:="DataTableCSV", description:="CSV file <" & FileName & "> not found!"
    End If
    this.CSVFileName = FileName
End Sub


Private Sub Class_Terminate()
    Set this.Model = Nothing
End Sub


Private Sub IDataTableStorage_LoadDataIntoModel()
    Dim FileHandle As Long: FileHandle = FreeFile
    Open this.CSVFileName For Input As #FileHandle
    Dim ReadBuffer As String
    ReadBuffer = Input$(LOF(FileHandle), #FileHandle)
    Close #FileHandle
    Guard.EmptyString ReadBuffer
    
    With this.Model
        .FieldIndices.RemoveAll
                
        Dim TableRows() As String
        TableRows = Split(ReadBuffer, vbNewLine)
        Dim RecordCount As Long: RecordCount = UBound(TableRows, 1) ' 0-based array with header row in position zero.
        If Len(TableRows(RecordCount)) = 0 Then
            VBA.Err.Raise number:=59, source:="CSVParser.BasicParse", description:="Empty lines at the end of the file must be removed for the basic parser!"
        End If
        
        Dim FieldBuffer As Variant
        FieldBuffer = Split(TableRows(0), this.FieldSeparator)
        Dim FieldCount As Long: FieldCount = UBound(FieldBuffer, 1) + 1
        
        Dim FieldIndex As Long
        Dim FieldName As String
        Dim Buffer As Variant
        ReDim Buffer(1 To FieldCount)
        
        .FieldNames = Buffer
        For FieldIndex = 1 To FieldCount
            FieldName = FieldBuffer(FieldIndex - 1)
            .FieldNames(FieldIndex) = FieldName
            .FieldIndices(FieldName) = FieldIndex
        Next FieldIndex
        
        Dim RecordIndex As Long
        ReDim Buffer(1 To RecordCount, 1 To FieldCount)
        .Values = Buffer
        For RecordIndex = 1 To RecordCount
            FieldBuffer = Split(TableRows(RecordIndex), this.FieldSeparator)
            For FieldIndex = 1 To FieldCount
                .Values(RecordIndex, FieldIndex) = FieldBuffer(FieldIndex - 1)
            Next FieldIndex
        Next RecordIndex
    End With
End Sub


Private Sub IDataTableStorage_SaveDataFromModel()
'    With this.Model
'        If Not .IsDirty Then Exit Sub
'        Dim RecordId As Variant
'        Dim RecordIndex As Long
'        Dim RecordRange As Excel.Range
'        Dim FieldIndex As Long
'        For Each RecordId In .DirtyRecords.Keys
'            RecordIndex = this.Model.RecordIndexFromId(RecordId)
'            Set RecordRange = this.TableBodyRange.Rows(RecordIndex)
'            For FieldIndex = 1 To UBound(this.Model.FieldNames, 1)
'                RecordRange.Cells(1, FieldIndex) = .Values(RecordIndex, FieldIndex)
'            Next FieldIndex
'        Next RecordId
'        .IsNotDirty
'    End With
End Sub


Private Function IDataTableStorage_GetIds() As Variant
    '''' Ids should be strings (cast if necessary), not numeric
    Dim Ids() As Variant
    ReDim Ids(1 To this.Model.IdIndices.Count)
    Dim IdIndex As Long
    For IdIndex = 1 To UBound(Ids)
        Ids(IdIndex) = this.Model.Values(IdIndex, 1)
    Next IdIndex
    IDataTableStorage_GetIds = Ids
End Function


Private Function IDataTableStorage_GetColumnValues(ByVal FieldName As String) As Variant
    Dim FieldIndex As Long: FieldIndex = this.Model.FieldIndices(FieldName)
    Dim ColumnValues() As Variant
    ReDim ColumnValues(1 To this.Model.IdIndices.Count)
    Dim RecordIndex As Long
    For RecordIndex = 1 To UBound(ColumnValues)
        ColumnValues(RecordIndex) = this.Model.Values(RecordIndex, FieldIndex)
    Next RecordIndex
    IDataTableStorage_GetColumnValues = ColumnValues
End Function
