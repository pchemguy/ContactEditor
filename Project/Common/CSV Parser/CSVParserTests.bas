Attribute VB_Name = "CSVParserTests"
'@Folder("Common.CSV Parser")
'@TestModule
'@IgnoreModule LineLabelNotUsed, IndexedDefaultMemberAccess
Option Explicit
Option Private Module
Option Compare Text


#Const LateBind = LateBindTests
#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'This method runs once per module.
'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
End Sub


'This method runs once per module.
'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='

'@TestMethod("BasicParser")
Private Sub ztcBasicParser_ValidatesCorrectParsing()
    On Error GoTo TestFail

Arrange:
    Dim CSVFileParser As CSVParser
    Dim CSVParserType As CSVParserClass: CSVParserType = CSVParserClass.CSVParserBasicVBA
    Dim CSVFileName As String: CSVFileName = "Contacts.xsv"
    Dim FieldSeparator As String: FieldSeparator = ","
    Dim TableRange As Excel.Range: Set TableRange = Nothing
    Dim TableName As String: TableName = vbNullString
    
Act:
    Set CSVFileParser = CSVParser.Create(CSVFileName, FieldSeparator, CSVParserType, TableRange, TableName)
    CSVFileParser.Parse
Assert:
    Assert.IsNotNothing CSVFileParser.FieldMap, "FieldMap not set"
    Assert.IsFalse IsEmpty(CSVFileParser.FieldNames), "FieldNames is empty"
    Assert.IsFalse IsEmpty(CSVFileParser.Records), "Records is empty"
    Assert.AreEqual 8, CSVFileParser.FieldMap.Count, "FieldMap - wrong count"
    Assert.AreEqual 1, LBound(CSVFileParser.FieldNames, 1), "FieldNames - wrong base"
    Assert.AreEqual 8, UBound(CSVFileParser.FieldNames, 1), "FieldNames - wrong count"
    Assert.AreEqual 1, LBound(CSVFileParser.Records, 1), "Records - wrong base"
    Assert.AreEqual 1000, UBound(CSVFileParser.Records, 1), "Records - wrong count"
    Assert.AreEqual "id", CSVFileParser.FieldNames(1), "FieldNames - value mismatch"
    Assert.AreEqual "domain", CSVFileParser.FieldNames(8), "FieldNames - value mismatch"
    Assert.AreEqual 1, CSVFileParser.FieldMap("id"), "FieldMap - value mismatch"
    Assert.AreEqual 8, CSVFileParser.FieldMap("domain"), "FieldMap - value mismatch"
    Assert.AreEqual "Edna.Jennings@neuf.fr", CSVFileParser.Records(4, 6), "Records - value mismatch"
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("OpenParser")
Private Sub ztcOpenParser_ValidatesCorrectParsing()
    On Error GoTo TestFail

Arrange:
    Dim CSVFileParser As CSVParser
    Dim CSVParserType As CSVParserClass: CSVParserType = CSVParserClass.CSVParserExcelOpen
    Dim CSVFileName As String: CSVFileName = "Contacts.csv"
    Dim FieldSeparator As String: FieldSeparator = Chr$(9) '","
    Dim TableRange As Excel.Range: Set TableRange = Nothing
    Dim TableName As String: TableName = vbNullString
    
Act:
    Set CSVFileParser = CSVParser.Create(CSVFileName, FieldSeparator, CSVParserType, TableRange, TableName)
    CSVFileParser.Parse
Assert:
    Assert.IsNotNothing CSVFileParser.FieldMap, "FieldMap not set"
    Assert.IsFalse IsEmpty(CSVFileParser.FieldNames), "FieldNames is empty"
    Assert.IsFalse IsEmpty(CSVFileParser.Records), "Records is empty"
    Assert.AreEqual 8, CSVFileParser.FieldMap.Count, "FieldMap - wrong count"
    Assert.AreEqual 1, LBound(CSVFileParser.FieldNames, 1), "FieldNames - wrong base"
    Assert.AreEqual 8, UBound(CSVFileParser.FieldNames, 1), "FieldNames - wrong count"
    Assert.AreEqual 1, LBound(CSVFileParser.Records, 1), "Records - wrong base"
    Assert.AreEqual 1000, UBound(CSVFileParser.Records, 1), "Records - wrong count"
    Assert.AreEqual "id", CSVFileParser.FieldNames(1), "FieldNames - value mismatch"
    Assert.AreEqual "domain", CSVFileParser.FieldNames(8), "FieldNames - value mismatch"
    Assert.AreEqual 1, CSVFileParser.FieldMap("id"), "FieldMap - value mismatch"
    Assert.AreEqual 8, CSVFileParser.FieldMap("domain"), "FieldMap - value mismatch"
    Assert.AreEqual "Edna.Jennings@neuf.fr", CSVFileParser.Records(4, 6), "Records - value mismatch"
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub


'@TestMethod("OpenTextParser")
Private Sub ztcOpenTextParser_ValidatesCorrectParsing()
    On Error GoTo TestFail

Arrange:
    Dim CSVFileParser As CSVParser
    Dim CSVParserType As CSVParserClass: CSVParserType = CSVParserClass.CSVParserExcelOpenText
    Dim CSVFileName As String: CSVFileName = "Contacts.csv"
    Dim FieldSeparator As String: FieldSeparator = Chr$(9)
    Dim TableRange As Excel.Range: Set TableRange = Nothing
    Dim TableName As String: TableName = vbNullString
    
Act:
    Set CSVFileParser = CSVParser.Create(CSVFileName, FieldSeparator, CSVParserType, TableRange, TableName)
    CSVFileParser.Parse
Assert:
    Assert.IsNotNothing CSVFileParser.FieldMap, "FieldMap not set"
    Assert.IsFalse IsEmpty(CSVFileParser.FieldNames), "FieldNames is empty"
    Assert.IsFalse IsEmpty(CSVFileParser.Records), "Records is empty"
    Assert.AreEqual 8, CSVFileParser.FieldMap.Count, "FieldMap - wrong count"
    Assert.AreEqual 1, LBound(CSVFileParser.FieldNames, 1), "FieldNames - wrong base"
    Assert.AreEqual 8, UBound(CSVFileParser.FieldNames, 1), "FieldNames - wrong count"
    Assert.AreEqual 1, LBound(CSVFileParser.Records, 1), "Records - wrong base"
    Assert.AreEqual 1000, UBound(CSVFileParser.Records, 1), "Records - wrong count"
    Assert.AreEqual "id", CSVFileParser.FieldNames(1), "FieldNames - value mismatch"
    Assert.AreEqual "domain", CSVFileParser.FieldNames(8), "FieldNames - value mismatch"
    Assert.AreEqual 1, CSVFileParser.FieldMap("id"), "FieldMap - value mismatch"
    Assert.AreEqual 8, CSVFileParser.FieldMap("domain"), "FieldMap - value mismatch"
    Assert.AreEqual "Edna.Jennings@neuf.fr", CSVFileParser.Records(4, 6), "Records - value mismatch"
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.number & " - " & Err.description
End Sub
