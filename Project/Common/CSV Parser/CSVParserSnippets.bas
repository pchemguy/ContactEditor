Attribute VB_Name = "CSVParserSnippets"
'@Folder("Common.CSV Parser")
'@IgnoreModule ProcedureNotUsed
Option Explicit


Private Sub BasicParse()
    Dim CSVFileParser As CSVParser
    Dim CSVParserType As CSVParserClass: CSVParserType = CSVParserClass.CSVParserBasicVBA
    Dim CSVFileName As String: CSVFileName = "Contacts.xsv"
    Dim FieldSeparator As String: FieldSeparator = ","
    Dim TableRange As Excel.Range: Set TableRange = Nothing
    Dim TableName As String: TableName = vbNullString
    
    Set CSVFileParser = CSVParser.Create(CSVFileName, FieldSeparator, CSVParserType, TableRange, TableName)
    CSVFileParser.Parse
End Sub


Private Sub BasicParseExcelTable()
    Dim CSVFileParser As CSVParser
    Dim CSVParserType As CSVParserClass: CSVParserType = CSVParserClass.CSVParserBasicVBA
    Dim CSVFileName As String: CSVFileName = "Contacts.xsv"
    Dim FieldSeparator As String: FieldSeparator = ","
    Dim TableRange As Excel.Range: Set TableRange = Playground.Range("C3:D4")
    Dim TableName As Variant: TableName = True
    
    Set CSVFileParser = CSVParser.Create(CSVFileName, FieldSeparator, CSVParserType, TableRange, TableName)
    CSVFileParser.Parse
End Sub


Private Sub OpenParse()
    Dim CSVFileParser As CSVParser
    Dim CSVParserType As CSVParserClass: CSVParserType = CSVParserClass.CSVParserExcelOpen
    Dim CSVFileName As String: CSVFileName = "Contacts.xsv"
    Dim FieldSeparator As String: FieldSeparator = ","
    Dim TableRange As Excel.Range: Set TableRange = Nothing
    Dim TableName As String: TableName = vbNullString
    
    Set CSVFileParser = CSVParser.Create(CSVFileName, FieldSeparator, CSVParserType, TableRange, TableName)
    CSVFileParser.Parse
End Sub


Private Sub OpenParseExcelTable()
    Dim CSVFileParser As CSVParser
    Dim CSVParserType As CSVParserClass: CSVParserType = CSVParserClass.CSVParserExcelOpen
    Dim CSVFileName As String: CSVFileName = "Contacts.xsv"
    Dim FieldSeparator As String: FieldSeparator = ","
    Dim TableRange As Excel.Range: Set TableRange = Playground.Range("C3:D4")
    Dim TableName As Variant: TableName = True
    
    Set CSVFileParser = CSVParser.Create(CSVFileName, FieldSeparator, CSVParserType, TableRange, TableName)
    CSVFileParser.Parse
End Sub

