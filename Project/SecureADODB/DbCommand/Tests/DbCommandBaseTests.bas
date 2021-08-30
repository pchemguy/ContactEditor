Attribute VB_Name = "DbCommandBaseTests"
'@Folder "SecureADODB.DbCommand.Tests"
'@TestModule
'@IgnoreModule
Option Explicit
Option Private Module

#Const LateBind = LateBindTests

#If LateBind Then
    Private Assert As Object
#Else
    Private Assert As Rubberduck.PermissiveAssertClass
#End If


'@ModuleInitialize
Private Sub ModuleInitialize()
    #If LateBind Then
        Set Assert = CreateObject("Rubberduck.PermissiveAssertClass")
    #Else
        Set Assert = New Rubberduck.PermissiveAssertClass
    #End If
End Sub


'@ModuleCleanup
Private Sub ModuleCleanup()
    Set Assert = Nothing
End Sub


Private Function GetParameterProvider() As IParameterProvider
    Set GetParameterProvider = AdoParameterProvider.Create(AdoTypeMappings.Default)
End Function



'@TestMethod("Factory Guard")
Private Sub Create_ThrowsGivenNullParameterProvider()
    On Error GoTo CleanFail
    Dim sut As DbCommandBase
    Set sut = DbCommandBase.Create(Nothing)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ErrNo.ObjectNotSetErr Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub


'@TestMethod("Guard Clauses")
Private Sub CreateCommand_ThrowsGivenNullConnection()
    On Error GoTo TestFail
    
    Dim sut As IDbCommandBase
    Set sut = DbCommandBase.Create(GetParameterProvider)
    
    Dim args() As Variant
    
    On Error GoTo CleanFail
    Dim cmd As ADODB.Command
    Set cmd = sut.CreateCommand(Nothing, adCmdText, "SQL", args)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ErrNo.ObjectNotSetErr Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub


'@TestMethod("Guard Clauses")
Private Sub CreateCommand_ThrowsGivenClosedConnection()
    On Error GoTo TestFail
    
    Dim sut As IDbCommandBase
    Set sut = DbCommandBase.Create(GetParameterProvider)
    
    Dim args() As Variant
    
    Dim db As StubDbConnection
    Set db = New StubDbConnection
    db.State = adStateClosed
    
    On Error GoTo CleanFail
    Dim cmd As ADODB.Command
    Set cmd = sut.CreateCommand(db, adCmdText, "SQL", args)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ErrNo.CustomErr Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub


'@TestMethod("Guard Clauses")
Private Sub CreateCommand_ThrowsGivenEmptyCommandString()
    On Error GoTo TestFail
    
    Dim sut As IDbCommandBase
    Set sut = DbCommandBase.Create(GetParameterProvider)
    
    Dim args() As Variant
    
    Dim db As StubDbConnection
    Set db = New StubDbConnection
    db.State = adStateOpen
    
    On Error GoTo CleanFail
    Dim cmd As ADODB.Command
    Set cmd = sut.CreateCommand(db, adCmdText, vbNullString, args)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ErrNo.EmptyStringErr Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub
