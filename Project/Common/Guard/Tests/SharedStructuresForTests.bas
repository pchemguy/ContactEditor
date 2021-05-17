Attribute VB_Name = "SharedStructuresForTests"
'@Folder "Common.Guard.Tests"
'@TestModule
'@IgnoreModule VariableNotAssigned, UnassignedVariableUsage
Option Explicit
Option Private Module


Const MsgExpectedErrNotRaised As String = "Expected error was not raised."
Const MsgUnexpectedErrRaised As String = "Unexpected error was raised."


Public Sub AssertExpectedError(ByVal Assert As Rubberduck.PermissiveAssertClass, Optional ByVal ExpectedErrorNo As ErrNo = ErrNo.PassedNoErr)
    Debug.Assert TypeOf Assert Is Rubberduck.PermissiveAssertClass
    
    Dim ActualErrNo As Long
    ActualErrNo = VBA.Err.Number
    Dim errorDetails As String
    errorDetails = " Error: #" & ActualErrNo & " - " & VBA.Err.Description
    VBA.Err.Clear
    
    Select Case ActualErrNo
        Case ExpectedErrorNo
            Assert.Succeed
        Case ErrNo.PassedNoErr
            Assert.Fail MsgExpectedErrNotRaised
        Case Else
            Assert.Fail MsgUnexpectedErrRaised & errorDetails
    End Select
End Sub

