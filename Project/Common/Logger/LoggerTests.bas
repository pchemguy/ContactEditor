Attribute VB_Name = "LoggerTests"
Attribute VB_Description = "Tests for the Logger class."
'@Folder "Common.Logger"
'@TestModule
'@ModuleDescription("Tests for the Logger class.")
'@IgnoreModule VariableNotUsed, AssignmentNotUsed, LineLabelNotUsed, UnhandledOnErrorResumeNext
Option Explicit
Option Private Module


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


'===================================================='
'==================== TEST CASES ===================='
'===================================================='

'@TestMethod("Factory Guard")
Private Sub ztcCreate_PassesIfInvokedFromDefaultInstance()
    On Error Resume Next
    Dim AdoLogger As ILogger: Set AdoLogger = Logger.Create
    Guard.AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub


'@TestMethod("Factory Guard")
Private Sub ztcCreate_ThrowsIfNotInvokedFromDefaultInstance()
    On Error Resume Next
    Dim stubLogger As Logger: Set stubLogger = New Logger
    Dim stubILogger As Logger: Set stubILogger = stubLogger.Create
    Assert.IsNothing stubILogger
    Guard.AssertExpectedError Assert, ErrNo.NonDefaultInstanceErr
End Sub


'@TestMethod("Log Database")
Private Sub ztcGetLogDatabase_VerifyReturnsDictionary()
    On Error GoTo TestFail
    
Arrange:
    Dim instanceVar As Object: Set instanceVar = Logger.Create
Act:
    Dim stubILogger As Logger: Set stubILogger = Logger.Create
Assert:
    Assert.AreEqual "Dictionary", TypeName(Logger.LogDatabase)
    Assert.AreEqual "Dictionary", TypeName(stubILogger.LogDatabase)

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Log Database")
Private Sub ztcLog_VerifyItemCountOnInstance()
    On Error GoTo TestFail
    
Arrange:
    Dim AdoLogger As ILogger: Set AdoLogger = Logger.Create
Act:
    AdoLogger.Log "AAA"
    AdoLogger.Log "AAA"
Assert:
    Assert.AreEqual 2, AdoLogger.LogDatabase.Count

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Log Database")
Private Sub ztcLog_VerifyItemCountOnInstanceWithClear()
    On Error GoTo TestFail
    
Arrange:
    Dim AdoLogger As ILogger: Set AdoLogger = Logger.Create
Act:
    AdoLogger.Log "AAA"
    AdoLogger.Log "AAA"
    AdoLogger.ClearLog
    AdoLogger.Log "AAA"
    AdoLogger.Log "AAA"
    AdoLogger.Log "AAA"
Assert:
    Assert.AreEqual 3, AdoLogger.LogDatabase.Count

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Log Database")
Private Sub ztcLog_VerifyItemCountOnGlobalWithCustomDb()
    On Error GoTo TestFail
    
Arrange:
    Dim LogDb As Scripting.Dictionary
    Set LogDb = New Scripting.Dictionary
    LogDb.CompareMode = TextCompare
Act:
    Logger.Log "AAA", LogDb
    Logger.Log "AAA", LogDb
Assert:
    Assert.AreEqual 2, LogDb.Count

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Log Database")
Private Sub ztcLog_VerifyItemCountOnGlobalWithClearWithCustomDb()
    On Error GoTo TestFail
    
Arrange:
    Dim LogDb As Scripting.Dictionary
    Set LogDb = New Scripting.Dictionary
    LogDb.CompareMode = TextCompare
Act:
    Logger.Log "AAA", LogDb
    Logger.Log "AAA", LogDb
    Logger.ClearLog LogDb
    Logger.Log "AAA", LogDb
    Logger.Log "AAA", LogDb
    Logger.Log "AAA", LogDb
Assert:
    Assert.AreEqual 3, LogDb.Count

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Self")
Private Sub ztcSelf_CheckAvailability()
    On Error GoTo TestFail
    
Arrange:
    Dim instanceVar As Object: Set instanceVar = Logger.Create
Act:
    Dim selfVar As Object: Set selfVar = instanceVar.Self
Assert:
    Assert.AreEqual TypeName(instanceVar), TypeName(selfVar), "Error: type mismatch: " & TypeName(selfVar) & " type."
    Assert.AreSame instanceVar, selfVar, "Error: bad Self pointer"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Class")
Private Sub ztcClass_CheckAvailability()
    On Error GoTo TestFail
    
Arrange:
    Dim classVar As Object: Set classVar = Logger
Act:
    Dim classVarReturned As Object: Set classVarReturned = classVar.Create.Class
Assert:
    Assert.AreEqual TypeName(classVar), TypeName(classVarReturned), "Error: type mismatch: " & TypeName(classVarReturned) & " type."
    Assert.AreSame classVar, classVarReturned, "Error: bad Class pointer"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'===================================================='
'============== INTERACTIVE TEST CASES =============='
'===================================================='

'@TestMethod("PrintLog")
'@Description "Log some items, check the contents using Logger instance"
Private Sub ziPrintLogInstanceTest()
Attribute ziPrintLogInstanceTest.VB_Description = "Log some items, check the contents using Logger instance"
    Dim AdoLogger As ILogger
    Set AdoLogger = Logger.Create
    
    AdoLogger.Log "AAA"
    AdoLogger.Log "BBB"
    AdoLogger.PrintLog
End Sub


'@TestMethod("PrintLog")
'@Description "Log some items, check the contents using global Logger"
Private Sub ziPrintLogGlobalTest()
Attribute ziPrintLogGlobalTest.VB_Description = "Log some items, check the contents using global Logger"
    Logger.Log "AAA"
    Logger.Log "BBB"
    'Logger.ClearLog
    Logger.PrintLog
End Sub


'@TestMethod("PrintLog")
'@Description "Log some items, check the contents using global Logger and custom database"
Private Sub ziPrintLogGlobalCustomDatabaseTest()
Attribute ziPrintLogGlobalCustomDatabaseTest.VB_Description = "Log some items, check the contents using global Logger and custom database"
    Dim LogDb As Scripting.Dictionary
    Set LogDb = New Scripting.Dictionary
    LogDb.CompareMode = TextCompare
    
    Logger.Log "AAA", LogDb
    Logger.Log "BBB", LogDb
    'Logger.ClearLog
    Logger.PrintLog LogDb
End Sub
