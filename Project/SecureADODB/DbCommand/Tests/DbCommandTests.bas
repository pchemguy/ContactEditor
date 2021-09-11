Attribute VB_Name = "DbCommandTests"
'@Folder "SecureADODB.DbCommand.Tests"
'@TestModule
'@IgnoreModule
Option Explicit
Option Private Module

Private Const ERR_INVALID_WITHOUT_LIVE_CONNECTION = 3709

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


Private Function GetSUT(ByRef stubBase As StubDbCommandBase, ByRef stubConnection As StubDbConnection) As IDbCommand
    Set stubConnection = New StubDbConnection
    Set stubBase = New StubDbCommandBase
    
    Dim Result As DbCommand
    Set Result = DbCommand.Create(stubConnection, stubBase)
    
    Set GetSUT = Result
End Function


Private Function GetSingleParameterSelectSql() As String
    GetSingleParameterSelectSql = "SELECT * FROM [dbo].[Table1] WHERE [Field1] = ?;"
End Function


Private Function GetTwoParameterSelectSql() As String
    GetTwoParameterSelectSql = "SELECT * FROM [dbo].[Table1] WHERE [Field1] = ? AND [Field2] = ?;"
End Function


Private Function GetSingleParameterInsertSql() As String
    GetSingleParameterInsertSql = "INSERT INTO [dbo].[Table1] ([Timestamp], [Value]) VALUES (GETDATE(), ?);"
End Function


Private Function GetTwoParameterInsertSql() As String
    GetTwoParameterInsertSql = "INSERT INTO [dbo].[Table1] ([Timestamp], [Value], [ThingID]) VALUES (GETDATE(), ?, ?);"
End Function


Private Function GetStubParameter() As ADODB.Parameter
    Dim stubParameter As ADODB.Parameter
    Set stubParameter = New ADODB.Parameter
    stubParameter.Value = 42
    stubParameter.Type = adInteger
    stubParameter.Direction = adParamInput
    Set GetStubParameter = stubParameter
End Function
