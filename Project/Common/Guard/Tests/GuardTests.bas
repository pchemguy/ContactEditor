Attribute VB_Name = "GuardTests"
Attribute VB_Description = "Tests for the Guard class."
'@Folder "Common.Guard.Tests"
''@TestModule
'@ModuleDescription("Tests for the Guard class.")
'@IgnoreModule LineLabelNotUsed, UnhandledOnErrorResumeNext
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
    Set Guard = Nothing
End Sub


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub

'==================================================
'==================================================

'@TestMethod("Guard.EmptyString")
Private Sub EmptyString_Pass()
    On Error Resume Next
    Guard.EmptyString "Non-empty string"
    AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.EmptyString")
Private Sub EmptyString_ThrowsIfNotString()
    On Error Resume Next
    Guard.EmptyString True
    AssertExpectedError Assert, ErrNo.TypeMismatchErr
End Sub

'@TestMethod("Guard.EmptyString")
Private Sub EmptyString_ThrowsIfEmptyString()
    On Error Resume Next
    Guard.EmptyString vbNullString
    AssertExpectedError Assert, ErrNo.EmptyStringErr
End Sub

'@TestMethod("Guard.Singleton")
Private Sub Singleton_Pass()
    On Error Resume Next
    Guard.Singleton Guard
    AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.Singleton")
Private Sub Singleton_DefaultFactoryPass()
    On Error Resume Next
    Guard.Singleton Guard.Create
    AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.Singleton")
Private Sub Singleton_NothingCheck()
    On Error Resume Next
    Guard.Singleton Nothing
    AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub

'@TestMethod("Guard.Singleton")
Private Sub Singleton_ThrowsIfNewingObject()
    On Error Resume Next
    '@Ignore VariableNotUsed, AssignmentNotUsed
    Dim GuardInstance As Guard: Set GuardInstance = New Guard
    AssertExpectedError Assert, ErrNo.SingletonErr
End Sub

'@TestMethod("Guard.ObjectNotSet")
Private Sub ObjectNotSet_Pass()
    On Error Resume Next
    Guard.NullReference Guard
    AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.ObjectNotSet")
Private Sub ObjectNotSet_ThrowsIfNotObject()
    On Error Resume Next
    Guard.NullReference Empty
    AssertExpectedError Assert, ErrNo.ObjectRequiredErr
End Sub

'@TestMethod("Guard.ObjectNotSet")
Private Sub ObjectNotSet_ThrowsIfNothing()
    On Error Resume Next
    Guard.NullReference Nothing
    AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub

'@TestMethod("Guard.ObjectSet")
Private Sub ObjectSet_Pass()
    On Error Resume Next
    Guard.NonNullReference Nothing
    AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.ObjectSet")
Private Sub ObjectSet_ThrowsIfNotObject()
    On Error Resume Next
    Guard.NonNullReference Empty
    AssertExpectedError Assert, ErrNo.ObjectRequiredErr
End Sub

'@TestMethod("Guard.ObjectSet")
Private Sub ObjectSet_ThrowsIfNotNothing()
    On Error Resume Next
    Guard.NonNullReference Guard
    AssertExpectedError Assert, ErrNo.ObjectSetErr
End Sub

'@TestMethod("Guard.NonDefaultInstance")
Private Sub NonDefaultInstance_Pass()
    On Error Resume Next
    Guard.NonDefaultInstance Guard
    AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.NonDefaultInstance")
Private Sub NonDefaultInstance_ThrowsIfNothing()
    On Error Resume Next
    Guard.NonDefaultInstance Nothing
    AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub

'@TestMethod("Guard.DefaultInstance")
Private Sub DefaultInstance_ThrowsIfDefaultInstance()
    On Error Resume Next
    Guard.DefaultInstance Guard
    AssertExpectedError Assert, ErrNo.DefaultInstanceErr
End Sub


'@TestMethod("IsFalsy")
Private Sub IsFalsy_VerifiesFalsiness()
    On Error GoTo TestFail
    
Arrange:
    '@Ignore VariableNotAssigned
    Dim TestVar As Variant
    '@Ignore VariableNotAssigned
    Dim TestObj As Object
    Dim TestColl As VBA.Collection
    Set TestColl = New VBA.Collection
Act:
    
Assert:
    Assert.IsTrue IsFalsy(Empty), "Empty should be falsy"
    Assert.IsTrue IsFalsy(Null), "Null should be falsy"
    Assert.IsTrue IsFalsy(Nothing), "Nothing should be falsy"
    Assert.IsTrue IsFalsy(False), "False should be falsy"
    Assert.IsFalse IsFalsy(True), "True should be truthy"
    Assert.IsTrue IsFalsy(vbNullString), "vbNullString should be falsy"
    '@Ignore EmptyStringLiteral
    Assert.IsTrue IsFalsy(""), "Empty string literal should be falsy"
    Assert.IsFalse IsFalsy("Some text"), "Non-empty should be truthy"
    '@Ignore UnassignedVariableUsage
    Assert.IsTrue IsFalsy(TestVar), "Empty variant should be falsy"
    Assert.IsTrue IsFalsy(0&), "0 should be falsy"
    Assert.IsTrue IsFalsy(0#), "0.0 should be falsy"
    '@Ignore UnassignedVariableUsage
    Assert.IsTrue IsFalsy(TestObj), "Not set object should be falsy"
    Assert.IsFalse IsFalsy(TestColl), "Set object should be truthy"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("Guard.Self")
Private Sub Self_CheckAvailability()
    On Error GoTo TestFail
    
Arrange:
    Dim instanceVar As Object: Set instanceVar = Guard.Create
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


'@TestMethod("Guard.Class")
Private Sub Class_CheckAvailability()
    On Error GoTo TestFail
    
Arrange:
    Dim classVar As Object: Set classVar = Guard
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
