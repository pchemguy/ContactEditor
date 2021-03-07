VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSVParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Parses CSV files using plain VBA and Excel assisted VBA. Optionally copies data onto a worksheet."
'@Folder "Common.CSV Parser"
'@ModuleDescription "Parses CSV files using plain VBA and Excel assisted VBA. Optionally copies data onto a worksheet."
'@PredeclaredId
'@IgnoreModule ProcedureNotUsed, IndexedDefaultMemberAccess
'@Exposed
Option Explicit

Public Enum CSVParserClass
    CSVParserNotSpecified
    CSVParserBasicVBA
    CSVParserExcelOpen
    CSVParserExcelOpenText
End Enum

Private Type TCSVParser
    Parser As CSVParserClass
    CSVFileName As String
    FieldSeparator As String
    TableName As String
    TableRange As Excel.Range
    
    FieldNames As Variant
    FieldMap As Scripting.Dictionary
    Records As Variant
End Type
Private this As TCSVParser


Private Sub Class_Initialize()
    Set this.FieldMap = New Scripting.Dictionary
End Sub


Private Sub Class_Terminate()
    Set this.FieldMap = Nothing
    RestoreExcelApp Application
End Sub


' In practice, this optimization can only provide advantage
' if used outside the class in combination with bulk file
' parsing. Even then for small to moderate file sizes, basic
' parser, if sufficient, provides and order of magnitude higher
' performance (compared on a file with a header, 8 fields and
' 1000 records).
' When used inside the class, performance improvement
' is barely significant, if at all.
Private Sub RestrictExcelApp(ByVal App As Excel.Application)
    App.ScreenUpdating = False
    'App.Visible = False ' This only slow things down
    App.AutomationSecurity = msoAutomationSecurityForceDisable
    App.EnableEvents = False
    App.Calculation = xlCalculationManual
End Sub


Private Sub RestoreExcelApp(ByVal App As Excel.Application)
    App.Calculation = xlCalculationAutomatic
    App.EnableEvents = True
    App.AutomationSecurity = msoAutomationSecurityLow
    'App.Visible = True
    App.ScreenUpdating = True
End Sub


'@Description("Returns instance reference")
Public Property Get Self() As CSVParser
Attribute Self.VB_Description = "Returns instance reference"
    Set Self = Me
End Property


'@Description("Returns class reference")
Public Property Get Class() As CSVParser
Attribute Class.VB_Description = "Returns class reference"
    Set Class = CSVParser
End Property


Public Property Get FieldNames() As Variant
    FieldNames = this.FieldNames
End Property


Public Property Get FieldMap() As Scripting.Dictionary
    Set FieldMap = this.FieldMap
End Property
    
    
Public Property Get Records() As Variant
    Records = this.Records
End Property


Public Function Create(ByVal CSVFileName As String, _
              Optional ByVal FieldSeparator As String = vbNullString, _
              Optional ByVal Parser As CSVParserClass = CSVParserClass.CSVParserNotSpecified, _
              Optional ByVal TableRange As Excel.Range = Nothing, _
              Optional ByVal TableName As Variant) As CSVParser
    Guard.NonDefaultInstance Me

    Dim result As CSVParser
    Set result = New CSVParser
    result.Init CSVFileName, FieldSeparator, Parser, TableRange, TableName
    Set Create = result
End Function


Public Sub Init(ByVal CSVFileName As String, _
       Optional ByVal FieldSeparator As String, _
       Optional ByVal Parser As CSVParserClass = CSVParserClass.CSVParserNotSpecified, _
       Optional ByVal TableRange As Excel.Range = Nothing, _
       Optional ByVal TableName As Variant)
       
    Guard.EmptyString CSVFileName
    
    Dim FileName As String
    On Error Resume Next
    Dim FileExist As Variant
    FileName = CSVFileName
    FileExist = Dir$(FileName)
    If Len(FileExist) = 0 Then
        FileName = ThisWorkbook.Path & Application.PathSeparator & CSVFileName
        FileExist = Dir$(FileName)
    End If
    On Error GoTo 0
    If Len(FileExist) = 0 Then
        VBA.Err.Raise number:=53, source:="CSVParser", description:="CSV file <" & CSVFileName & "> not found!"
    End If
    this.CSVFileName = FileName
       
    this.Parser = IIf(Parser <> CSVParserClass.CSVParserNotSpecified, Parser, CSVParserClass.CSVParserBasicVBA)
    
    this.FieldSeparator = FieldSeparator
    Select Case VarType(TableName)
        Case vbEmpty
            this.TableName = vbNullString
        Case vbBoolean
            this.TableName = Dir(FileName)
            this.TableName = Left$(this.TableName, Len(this.TableName) - 4)
        Case vbString
            this.TableName = TableName
    End Select
    
    Set this.TableRange = TableRange
End Sub


Public Sub Parse()
    Select Case this.Parser
        Case CSVParserClass.CSVParserBasicVBA
            BasicParse
        Case CSVParserClass.CSVParserExcelOpen
            WBOpenParse
        Case CSVParserClass.CSVParserExcelOpenText
            WBOpenTextParse
    End Select
End Sub


'@Description "Pure VBA basic parser. Splits file contents into records and field using provided field and record separator strings."
Private Sub BasicParse()
Attribute BasicParse.VB_Description = "Pure VBA basic parser. Splits file contents into records and field using provided field and record separator strings."
    Dim FileHandle As Long: FileHandle = FreeFile
    Open this.CSVFileName For Input As #FileHandle
    Dim ReadBuffer As String
    ReadBuffer = Input$(LOF(FileHandle), #FileHandle)
    Close #FileHandle
    
    Guard.EmptyString ReadBuffer
    
    Dim TableRows() As String
    TableRows = Split(ReadBuffer, vbNewLine)
    Dim RecordCount As Long: RecordCount = UBound(TableRows, 1) ' 0-based array with header row in position zero.
    If Len(TableRows(RecordCount)) = 0 Then
        VBA.Err.Raise number:=59, source:="CSVParser.BasicParse", description:="Empty lines at the end of the file must be removed for the basic parser!"
    End If
    
    Dim FieldSeparator As String
    FieldSeparator = IIf(Len(this.FieldSeparator) > 0, this.FieldSeparator, ",")
    Dim FieldBuffer As Variant
    FieldBuffer = Split(TableRows(0), FieldSeparator)
    Dim FieldCount As Long: FieldCount = UBound(FieldBuffer, 1) + 1
    
    Dim FieldIndex As Long
    Dim FieldName As String
    ReDim this.FieldNames(1 To FieldCount)
    For FieldIndex = 1 To FieldCount
        FieldName = FieldBuffer(FieldIndex - 1)
        this.FieldNames(FieldIndex) = FieldName
        this.FieldMap(FieldName) = FieldIndex
    Next FieldIndex
    
    Dim RecordIndex As Long
    ReDim this.Records(1 To RecordCount, 1 To FieldCount)
    For RecordIndex = 1 To RecordCount
        FieldBuffer = Split(TableRows(RecordIndex), FieldSeparator)
        For FieldIndex = 1 To FieldCount
            this.Records(RecordIndex, FieldIndex) = FieldBuffer(FieldIndex - 1)
        Next FieldIndex
    Next RecordIndex
    
    WSheetTable
End Sub


'@Description "Parses file via the Workbook.Open routine."
Private Sub WBOpenParse()
Attribute WBOpenParse.VB_Description = "Parses file via the Workbook.Open routine."
    Const HEADER_ROW_INDEX As Long = 1
    
    Dim FieldSeparator As String
    FieldSeparator = IIf(Len(this.FieldSeparator) > 0, this.FieldSeparator, ",")
    
    Dim TempFile As String
    If Right$(this.CSVFileName, 4) = ".csv" Then
        TempFile = this.CSVFileName
        this.CSVFileName = Left$(TempFile, Len(TempFile) - 4) & "._csv"
        Name TempFile As this.CSVFileName
    End If
    
    RestrictExcelApp Application
    Dim TableRange As Excel.Range
    Dim CSVWBook As Excel.Workbook
    Set CSVWBook = Application.Workbooks.Open(FileName:=this.CSVFileName, Format:=XlFileFormat.xlCSV, Delimiter:=FieldSeparator)
    With CSVWBook
        Set TableRange = .Sheets(1).UsedRange
        this.FieldNames = Application.WorksheetFunction.Index(.Sheets(1).UsedRange.Value, HEADER_ROW_INDEX)
        this.Records = TableRange.Rows(2).Resize(TableRange.Rows.Count - 1, TableRange.Columns.Count)
        .Close SaveChanges:=False
    End With
    RestoreExcelApp Application

    If TempFile <> vbNullString Then
        Name this.CSVFileName As TempFile
        this.CSVFileName = TempFile
    End If

    Dim FieldIndex As Long
    For FieldIndex = 1 To UBound(this.FieldNames, 1)
        this.FieldMap(this.FieldNames(FieldIndex)) = FieldIndex
    Next FieldIndex
    
    WSheetTable
End Sub


'@Description "Parses file via the Workbook.OpenText routine."
Private Sub WBOpenTextParse()
Attribute WBOpenTextParse.VB_Description = "Parses file via the Workbook.OpenText routine."
    Const HEADER_ROW_INDEX As Long = 1
    
    Dim TempFile As String
    If Right$(this.CSVFileName, 4) = ".csv" Then
        TempFile = this.CSVFileName
        this.CSVFileName = Left$(TempFile, Len(TempFile) - 4) & ".$sv"
        Name TempFile As this.CSVFileName
    End If
    
    RestrictExcelApp Application
    Dim CSVWBook As Excel.Workbook
    If this.FieldSeparator = vbNullString Then
        Application.Workbooks.OpenText FileName:=this.CSVFileName, _
                                       DataType:=XlTextParsingType.xlDelimited, _
                                       ConsecutiveDelimiter:=True, _
                                       Tab:=True, _
                                       Semicolon:=True, _
                                       Comma:=True
    Else
        Application.Workbooks.OpenText FileName:=this.CSVFileName, _
                                       DataType:=XlTextParsingType.xlDelimited, _
                                       ConsecutiveDelimiter:=True, _
                                       OtherChar:=this.FieldSeparator
    End If
    Set CSVWBook = ActiveWorkbook
    Dim TableRange As Excel.Range
    With CSVWBook
        Set TableRange = .Sheets(1).UsedRange
        this.FieldNames = Application.WorksheetFunction.Index(.Sheets(1).UsedRange.Value, HEADER_ROW_INDEX)
        this.Records = TableRange.Rows(2).Resize(TableRange.Rows.Count - 1, TableRange.Columns.Count)
        .Close SaveChanges:=False
    End With
    RestoreExcelApp Application
    
    If TempFile <> vbNullString Then
        Name this.CSVFileName As TempFile
        this.CSVFileName = TempFile
    End If
    
    Dim FieldIndex As Long
    For FieldIndex = 1 To UBound(this.FieldNames, 1)
        this.FieldMap(this.FieldNames(FieldIndex)) = FieldIndex
    Next FieldIndex
    
    WSheetTable
End Sub


'@Description "Places data table onto an Excel worksheet."
Private Sub WSheetTable()
Attribute WSheetTable.VB_Description = "Places data table onto an Excel worksheet."
    If this.TableRange Is Nothing Then Exit Sub
    Dim RecordCount As Long: RecordCount = UBound(this.Records, 1)
    Dim FieldCount As Long: FieldCount = UBound(this.FieldNames, 1)
    
    Dim TableRange As Excel.Range
    Dim TableHeaderRange As Excel.Range
    Dim TableBodyRange As Excel.Range
    Set TableRange = this.TableRange.Range("A1").Resize(RecordCount + 1, FieldCount)
    Set TableBodyRange = TableRange.Range("A2").Resize(RecordCount, FieldCount)
    Set TableHeaderRange = TableRange.Rows(1)
    Set this.TableRange = TableRange
    TableHeaderRange.Value = this.FieldNames
    TableBodyRange.Value = this.Records
    
    TableHeaderRange.Cells.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
    TableHeaderRange.Cells.Font.Bold = True
    TableRange.Columns.AutoFit
    
    If this.TableName = vbNullString Then Exit Sub
    
    TableRange.Name = this.TableName
    TableHeaderRange.Name = this.TableName & "Header"
    TableBodyRange.Name = this.TableName & "Body"
End Sub
