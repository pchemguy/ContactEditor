Attribute VB_Name = "CSVParserSnippetsTiming"
'@Folder("Common.CSV Parser")
Option Explicit


Private Sub OpenParseTime()
    Dim TimeMan As IExecutionTimer
    Set TimeMan = New BasicTimer
    
    Dim CSVFileParser As CSVParser
    Dim CSVParserType As CSVParserClass: CSVParserType = CSVParserClass.CSVParserExcelOpen
    Dim CSVFileName As String: CSVFileName = "Contacts.xsv"
    Dim FieldSeparator As String: FieldSeparator = ","
    Dim TableRange As Excel.Range: Set TableRange = Nothing
    Dim TableName As String: TableName = vbNullString
    
    Set CSVFileParser = CSVParser.Create(CSVFileName, FieldSeparator, CSVParserType, TableRange, TableName)
    
    Const RepeatCount As Long = 20
    Dim RepeatIndex As Long
    Dim TimeDelta As Double
    
    TimeMan.Start
    For RepeatIndex = 1 To RepeatCount
        CSVFileParser.Parse
    Next RepeatIndex
    TimeDelta = TimeMan.TimeElapsed
    
    Debug.Print "Time elapsed: " & TimeDelta / RepeatCount & " s."
End Sub


Private Sub BasicParseTime()
    Dim TimeMan As IExecutionTimer
    Set TimeMan = New BasicTimer
    
    Dim CSVFileParser As CSVParser
    Dim CSVParserType As CSVParserClass: CSVParserType = CSVParserClass.CSVParserBasicVBA
    Dim CSVFileName As String: CSVFileName = "Contacts.xsv"
    Dim FieldSeparator As String: FieldSeparator = ","
    Dim TableRange As Excel.Range: Set TableRange = Nothing
    Dim TableName As String: TableName = vbNullString
    
    Set CSVFileParser = CSVParser.Create(CSVFileName, FieldSeparator, CSVParserType, TableRange, TableName)
    
    Const RepeatCount As Long = 20
    Dim RepeatIndex As Long
    Dim TimeDelta As Double
    
    TimeMan.Start
    For RepeatIndex = 1 To RepeatCount
        CSVFileParser.Parse
    Next RepeatIndex
    TimeDelta = TimeMan.TimeElapsed
    
    Debug.Print "Time elapsed: " & TimeDelta / RepeatCount & " s."
End Sub


Private Sub OpenParseTimeFast()
    Dim TimeMan As IExecutionTimer
    Set TimeMan = New BasicTimer
    
    Dim CSVFileParser As CSVParser
    Dim CSVParserType As CSVParserClass: CSVParserType = CSVParserClass.CSVParserExcelOpen
    Dim CSVFileName As String: CSVFileName = "Contacts.xsv"
    Dim FieldSeparator As String: FieldSeparator = ","
    Dim TableRange As Excel.Range: Set TableRange = Nothing
    Dim TableName As String: TableName = vbNullString
    
    Set CSVFileParser = CSVParser.Create(CSVFileName, FieldSeparator, CSVParserType, TableRange, TableName)
    
    Const RepeatCount As Long = 20
    Dim RepeatIndex As Long
    Dim TimeDelta As Double
    
    Application.ScreenUpdating = False
    Application.Visible = False
    Application.AutomationSecurity = msoAutomationSecurityForceDisable
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    TimeMan.Start
    For RepeatIndex = 1 To RepeatCount
        CSVFileParser.Parse
    Next RepeatIndex
    TimeDelta = TimeMan.TimeElapsed
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.AutomationSecurity = msoAutomationSecurityLow
    Application.Visible = True
    Application.ScreenUpdating = True
    
    Debug.Print "Time elapsed: " & TimeDelta / RepeatCount & " s."
End Sub


Private Sub ResetApp()
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.AutomationSecurity = msoAutomationSecurityLow
    Application.Visible = True
    Application.ScreenUpdating = True
    Application.Workbooks("Contacts.xsv").Windows(1).Visible = True
End Sub
