Attribute VB_Name = "GuardTests"
Attribute VB_Description = "Tests for the Guard class."
'@Folder "Common.Guard"
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
End Sub


'==================================================
'==================================================

'@TestMethod("Guard.EmptyString")
Private Sub EmptyString_Pass()
    On Error Resume Next
    Guard.EmptyString "Non-empty string"
    Guard.AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.EmptyString")
Private Sub EmptyString_ThrowsIfNotString()
    On Error Resume Next
    Guard.EmptyString True
    Guard.AssertExpectedError Assert, ErrNo.TypeMismatchErr
End Sub

'@TestMethod("Guard.EmptyString")
Private Sub EmptyString_ThrowsIfEmptyString()
    On Error Resume Next
    Guard.EmptyString vbNullString
    Guard.AssertExpectedError Assert, ErrNo.EmptyStringErr
End Sub

'@TestMethod("Guard.ObjectNotSet")
Private Sub ObjectNotSet_Pass()
    On Error Resume Next
    Guard.NullReference Guard
    Guard.AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.ObjectNotSet")
Private Sub ObjectNotSet_ThrowsIfNotObject()
    On Error Resume Next
    Guard.NullReference Empty
    Guard.AssertExpectedError Assert, ErrNo.ObjectRequiredErr
End Sub

'@TestMethod("Guard.ObjectNotSet")
Private Sub ObjectNotSet_ThrowsIfNothing()
    On Error Resume Next
    Guard.NullReference Nothing
    Guard.AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub

'@TestMethod("Guard.ObjectSet")
Private Sub ObjectSet_Pass()
    On Error Resume Next
    Guard.NonNullReference Nothing
    Guard.AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.ObjectSet")
Private Sub ObjectSet_ThrowsIfNotObject()
    On Error Resume Next
    Guard.NonNullReference Empty
    Guard.AssertExpectedError Assert, ErrNo.ObjectRequiredErr
End Sub

'@TestMethod("Guard.ObjectSet")
Private Sub ObjectSet_ThrowsIfNotNothing()
    On Error Resume Next
    Guard.NonNullReference Guard
    Guard.AssertExpectedError Assert, ErrNo.ObjectSetErr
End Sub

'@TestMethod("Guard.NonDefaultInstance")
Private Sub NonDefaultInstance_Pass()
    On Error Resume Next
    Guard.NonDefaultInstance Guard
    Guard.AssertExpectedError Assert, ErrNo.PassedNoErr
End Sub

'@TestMethod("Guard.NonDefaultInstance")
Private Sub NonDefaultInstance_ThrowsIfNothing()
    On Error Resume Next
    Guard.NonDefaultInstance Nothing
    Guard.AssertExpectedError Assert, ErrNo.ObjectNotSetErr
End Sub

'@TestMethod("Guard.DefaultInstance")
Private Sub DefaultInstance_ThrowsIfDefaultInstance()
    On Error Resume Next
    Guard.DefaultInstance Guard
    Guard.AssertExpectedError Assert, ErrNo.DefaultInstanceErr
End Sub


'''' Disabled test
'''' @@TestMethod("Guard.Self")
''Private Sub Self_CheckAvailability()
''    On Error GoTo TestFail
''
''Arrange:
''    Dim instanceVar As Object
''    Set instanceVar = Guard.Create
''Act:
''    Dim selfVar As Object
''    Set selfVar = instanceVar.Self
''Assert:
''    Assert.AreEqual TypeName(instanceVar), TypeName(selfVar), "Error: type mismatch: " & TypeName(selfVar) & " type."
''    Assert.AreSame instanceVar, selfVar, "Error: bad Self pointer"
''
''CleanExit:
''    Exit Sub
''TestFail:
''    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
''End Sub


'''' Disabled test
'''' @@TestMethod("Guard.Class")
''Private Sub Class_CheckAvailability()
''    On Error GoTo TestFail
''
''Arrange:
''    Dim classVar As Object
''    Set classVar = Guard
''Act:
''    Dim classVarReturned As Object
''    Set classVarReturned = classVar.Create.Class
''Assert:
''    Assert.AreEqual TypeName(classVar), TypeName(classVarReturned), "Error: type mismatch: " & TypeName(classVarReturned) & " type."
''    Assert.AreSame classVar, classVarReturned, "Error: bad Class pointer"
''
''CleanExit:
''    Exit Sub
''TestFail:
''    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
''End Sub
