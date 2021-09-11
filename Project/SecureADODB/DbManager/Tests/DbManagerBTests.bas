Attribute VB_Name = "DbManagerBTests"
'@Folder "SecureADODB.DbManager.Tests"
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


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("Factory Guard")
Private Sub ztcCreate_ThrowsGivenNullConnection()
    On Error Resume Next
    Dim sut As IDbManager: Set sut = DbManager.Create(Nothing, New StubDbCommandFactory)
    Guard.AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub


'@TestMethod("Factory Guard")
Private Sub ztcCreate_ThrowsGivenNullCommandFactory()
    On Error Resume Next
    Dim sut As IDbManager: Set sut = DbManager.Create(New StubDbConnection, Nothing)
    Guard.AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub


'@TestMethod("Create")
Private Sub ztcCommand_CreatesDbCommandWithFactory()
    Dim stubCommandFactory As StubDbCommandFactory
    Set stubCommandFactory = New StubDbCommandFactory
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(New StubDbConnection, stubCommandFactory)
    
    Dim Result As IDbCommand
    Set Result = sut.Command
    
    Assert.AreEqual 1, stubCommandFactory.CreateCommandInvokes
End Sub


'@TestMethod("Transaction")
Private Sub ztcCreate_StartsTransaction()
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    sut.Begin
    
    Assert.IsTrue stubConnection.DidBeginTransaction
End Sub


'@TestMethod("Transaction")
Private Sub ztcCommit_CommitsTransaction()
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Begin
    sut.Commit
    
    Assert.IsTrue stubConnection.DidCommitTransaction
End Sub


'@TestMethod("Transaction")
Private Sub ztcCommit_ThrowsIfAlreadyCommitted()
    On Error Resume Next
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Commit
    sut.Commit
    Guard.AssertExpectedError Assert, ErrNo.AdoInvalidTransactionErr
End Sub


'@TestMethod("Transaction")
Private Sub ztcCommit_ThrowsIfAlreadyRolledBack()
    On Error Resume Next
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Rollback
    sut.Commit
    Guard.AssertExpectedError Assert, ErrNo.AdoInvalidTransactionErr
End Sub


'@TestMethod("Transaction")
Private Sub ztcRollback_ThrowsIfAlreadyCommitted()
    On Error Resume Next
    
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Commit
    sut.Rollback
    Guard.AssertExpectedError Assert, ErrNo.AdoInvalidTransactionErr
End Sub


'@TestMethod("Transaction")
Private Sub ztcRollback_RollbacksTransaction()
    Dim stubConnection As StubDbConnection
    Set stubConnection = New StubDbConnection
    
    Dim sut As IDbManager
    Set sut = DbManager.Create(stubConnection, New StubDbCommandFactory)
    
    sut.Begin
    sut.Rollback
    
    Assert.IsTrue stubConnection.DidRollBackTransaction
End Sub
