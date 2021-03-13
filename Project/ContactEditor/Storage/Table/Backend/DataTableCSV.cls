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
'@Folder "ContactEditor.Storage.Table.Backend"
'@ModuleDescription "Abstracts CSV file."
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

Implements IDataTableStorage

Private Type TDataTable
    Model As DataTableModel
    CSVFileName As String
    FieldSeparator As String
End Type
Private this As TDataTable


' ConnectionString - "<Path to CSV>" or empty to use ThisWorkbook.Path
' TableName        - "<CSV file name>!sep=<FieldSeparator>" or empty to use <ThisWorkbook.Name>(without ext)".xsv!sep=,"
'@Description "Returns a new IDataTable object."
Public Function Create(ByVal Model As DataTableModel, ByVal ConnectionString As String, ByVal TableName As String) As IDataTableStorage
Attribute Create.VB_Description = "Returns a new IDataTable object."
    Dim result As DataTableCSV
    Set result = New DataTableCSV
    result.Init Model, ConnectionString, TableName
    Set Create = result
End Function


Public Sub Init(ByVal Model As DataTableModel, ByVal ConnectionString As String, ByVal TableName As String)
    Dim CSVName As String
    If Len(TableName) > 0 Then
        CSVName = TableName
    Else
        CSVName = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - 5 + InStr(Right$(ThisWorkbook.Name, 5), ".")) & "xsv!sep=,"
    End If
    
    Set this.Model = Model
    Dim Prefix As String
    Prefix = IIf(Len(ConnectionString) = 0, ThisWorkbook.Path, ConnectionString) & Application.PathSeparator
    Dim FileName As String
    If Left$(Right$(CSVName, 6), 5) = "!sep=" Then
        this.FieldSeparator = Right$(CSVName, 1)
        FileName = Prefix & Left$(CSVName, Len(CSVName) - 6)
    Else
        this.FieldSeparator = ","
        FileName = Prefix & CSVName
    End If
    
    Dim FileExist As Variant
    On Error Resume Next
    FileExist = Dir$(FileName)
    On Error GoTo 0
    If Len(FileExist) = 0 Then
        VBA.Err.Raise number:=ErrNo.FileNotFoundErr, source:="DataTableCSV", description:="CSV file <" & FileName & "> not found!"
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
        Dim Buffer() As Variant
        ReDim Buffer(1 To FieldCount)
        
        For FieldIndex = 1 To FieldCount
            FieldName = FieldBuffer(FieldIndex - 1)
            Buffer(FieldIndex) = FieldName
            .FieldIndices(FieldName) = FieldIndex
        Next FieldIndex
        .FieldNames = Buffer
        
        Dim RecordIndex As Long
        ReDim Buffer(1 To RecordCount, 1 To FieldCount)
        For RecordIndex = 1 To RecordCount
            FieldBuffer = Split(TableRows(RecordIndex), this.FieldSeparator)
            For FieldIndex = 1 To FieldCount
                Buffer(RecordIndex, FieldIndex) = FieldBuffer(FieldIndex - 1)
            Next FieldIndex
            .IdIndices(FieldBuffer(0)) = RecordIndex
        Next RecordIndex
        .Values = Buffer
    End With
End Sub


Private Sub IDataTableStorage_SaveDataFromModel()
    Dim FieldIndex As Long
    Dim FieldCount As Long: FieldCount = UBound(this.Model.FieldNames, 1)
    Dim RecordIndex As Long
    Dim RecordCount As Long: RecordCount = UBound(this.Model.Values, 1)
    
    Dim TableRows() As Variant: ReDim TableRows(0 To RecordCount)
    Dim Records() As Variant: ReDim Records(1 To FieldCount)
    
    TableRows(0) = Join(this.Model.FieldNames, this.FieldSeparator)
    For RecordIndex = 1 To RecordCount
        For FieldIndex = 1 To FieldCount
            Records(FieldIndex) = this.Model.Values(RecordIndex, FieldIndex)
        Next FieldIndex
        TableRows(RecordIndex) = Join(Records, this.FieldSeparator)
    Next RecordIndex
    Dim Buffer As Variant
    Buffer = Join(TableRows, vbNewLine)
    Dim FileHandle As Long: FileHandle = FreeFile
    Open this.CSVFileName For Output As #FileHandle
    Print #FileHandle, Buffer;
    Close #FileHandle
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
