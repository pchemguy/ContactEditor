Attribute VB_Name = "CommonRoutinesTests"
'@Folder "Common.Shared"
'@TestModule
'@IgnoreModule LineLabelNotUsed
'@IgnoreModule UnhandledOnErrorResumeNext: Test routines validating expected errors do not need to resume error handling
'@IgnoreModule FunctionReturnValueDiscarded: Test routines validating expected errors may not need the returned value
'@IgnoreModule AssignmentNotUsed, VariableNotUsed: Ignore dummy assignments in tests
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


'@TestMethod("ParramArray")
Private Sub ztcUnfoldParamArray_ThrowsIfScalarArgument()
    On Error Resume Next
    UnfoldParamArray 1
    Guard.AssertExpectedError Assert, ErrNo.ExpectedArrayErr
End Sub


'@TestMethod("ParramArray")
Private Sub ztcUnfoldParamArray_ThrowsIfObjectArgument()
    On Error Resume Next
    UnfoldParamArray Application
    Guard.AssertExpectedError Assert, ErrNo.ExpectedArrayErr
End Sub


'@TestMethod("ParramArray")
Private Sub ztcUnfoldParamArray_VerifiesUnfoldsWrappedArray()
    On Error GoTo TestFail

Arrange:
    Dim ParamArrayParamArrayArg(0 To 0) As Variant
    ParamArrayParamArrayArg(0) = Array(1024&, 2048&)
Act:
    Dim Actual As Variant
    Actual = UnfoldParamArray(ParamArrayParamArrayArg)
Assert:
    Assert.AreEqual 1024, Actual(0), "Unfolding error"
    Assert.AreEqual 2048, Actual(1), "Unfolding error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ParramArray")
Private Sub ztcUnfoldParamArray_VerifiesNoUnfoldingEmptyArray()
    On Error GoTo TestFail

Arrange:
Act:
    Dim Actual As Variant
    Actual = UnfoldParamArray(Array())
Assert:
    Assert.AreEqual -1, UBound(Actual), "Returned argument error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ParramArray")
Private Sub ztcUnfoldParamArray_VerifiesNoUnfoldingMismatch1stLB()
    On Error GoTo TestFail

Arrange:
    Dim ArgumentOuter(1 To 1) As Variant
    Dim ArgumentInner(0 To 1) As Long
    ArgumentInner(0) = 1024&
    ArgumentInner(1) = 2048&
    ArgumentOuter(1) = ArgumentInner
Act:
    Dim Actual As Variant
    Actual = UnfoldParamArray(ArgumentOuter)
Assert:
    Assert.AreEqual 1024, Actual(1)(0), "Returned argument error"
    Assert.AreEqual 2048, Actual(1)(1), "Returned argument error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ParramArray")
Private Sub ztcUnfoldParamArray_VerifiesNoUnfoldingMismatch1stUB()
    On Error GoTo TestFail

Arrange:
    Dim ArgumentOuter(0 To 1) As Variant
    Dim ArgumentInner(0 To 1) As Long
    ArgumentInner(0) = 1024&
    ArgumentInner(1) = 2048&
    ArgumentOuter(1) = ArgumentInner
Act:
    Dim Actual As Variant
    Actual = UnfoldParamArray(ArgumentOuter)
Assert:
    Assert.AreEqual 1024, Actual(1)(0), "Returned argument error"
    Assert.AreEqual 2048, Actual(1)(1), "Returned argument error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ParramArray")
Private Sub ztcUnfoldParamArray_VerifiesNoUnfoldingMismatch1stDim()
    On Error GoTo TestFail

Arrange:
    Dim Argument(0 To 0, 0 To 1) As Variant
    Argument(0, 0) = 1024&
    Argument(0, 1) = 2048&
Act:
    Dim Actual As Variant
    Actual = UnfoldParamArray(Argument)
Assert:
    Assert.AreEqual 1024, Actual(0, 0), "Returned argument error"
    Assert.AreEqual 2048, Actual(0, 1), "Returned argument error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ParramArray")
Private Sub ztcUnfoldParamArray_VerifiesNoUnfoldingMismatch2ndLB()
    On Error GoTo TestFail

Arrange:
    Dim ArgumentOuter(0 To 0) As Variant
    Dim ArgumentInner(1 To 2) As Long
    ArgumentInner(1) = 1024&
    ArgumentInner(2) = 2048&
    ArgumentOuter(0) = ArgumentInner
Act:
    Dim Actual As Variant
    Actual = UnfoldParamArray(ArgumentOuter)
Assert:
    Assert.AreEqual 1024, Actual(0)(1), "Returned argument error"
    Assert.AreEqual 2048, Actual(0)(2), "Returned argument error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ParramArray")
Private Sub ztcUnfoldParamArray_VerifiesNoUnfoldingMismatch2ndDim()
    On Error GoTo TestFail

Arrange:
    Dim ArgumentOuter(0 To 0) As Variant
    Dim ArgumentInner(0 To 0, 0 To 1) As Long
    ArgumentInner(0, 0) = 1024&
    ArgumentInner(0, 1) = 2048&
    ArgumentOuter(0) = ArgumentInner
Act:
    Dim Actual As Variant
    Actual = UnfoldParamArray(ArgumentOuter)
Assert:
    Assert.AreEqual 1024, Actual(0)(0, 0), "Returned argument error"
    Assert.AreEqual 2048, Actual(0)(0, 1), "Returned argument error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("ParramArray")
Private Sub ztcUnfoldParamArray_VerifiesNoUnfoldingMismatch2ndType()
    On Error GoTo TestFail

Arrange:
Act:
    Dim Actual As Variant
    Actual = UnfoldParamArray(Array(1024&, 2048&))
Assert:
    Assert.AreEqual 1024, Actual(0), "Returned argument error"
    Assert.AreEqual 2048, Actual(1), "Returned argument error"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("PathCheck")
Private Sub ztcVerifyOrGetDefaultPath_ValidatesFullPathName()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = Environ$("ComSpec")
Act:
    Dim Actual As String
    Actual = VerifyOrGetDefaultPath(Environ$("ComSpec"), vbNullString)
Assert:
    Assert.AreEqual Expected, Actual, "CheckPath failed with full valid path"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("PathCheck")
Private Sub ztcVerifyOrGetDefaultPath_ValidatesValidPathName()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = ThisWorkbook.Path & Application.PathSeparator & ThisWorkbook.Name
Act:
    Dim Actual As String
    Actual = VerifyOrGetDefaultPath(Expected, Array("db", "sqlite"))
Assert:
    Assert.AreEqual Expected, Actual, "CheckPath failed with valid pathname"
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("PathCheck")
Private Sub ztcVerifyOrGetDefaultPath_ValidatesEmptyFilePathName()
    On Error GoTo TestFail

Arrange:
    Dim PATHuSEP As String
    PATHuSEP = Application.PathSeparator
    Dim PROJuNAME As String
    PROJuNAME = ThisWorkbook.VBProject.Name
    Dim Expected As String
    Expected = ThisWorkbook.Path & _
               PATHuSEP & PROJuNAME & "." & "db"
Act:
    Dim Actual As String
    Actual = VerifyOrGetDefaultPath(vbNullString, Array("db", "sqlite"))
Assert:
    Assert.AreEqual Expected, Actual, "CheckPath failed with empty file pathname." _
                                    & "Expected: < " & Expected & " > "

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("PathCheck")
Private Sub ztcVerifyOrGetDefaultPath_ValidatesFileName()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = ThisWorkbook.Path & Application.PathSeparator & ThisWorkbook.Name
Act:
    Dim Actual As String
    Actual = VerifyOrGetDefaultPath(ThisWorkbook.Name, vbNullString)
Assert:
    Assert.AreEqual Expected, Actual, "CheckPath failed with filename"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("PathCheck")
Private Sub ztcVerifyOrGetDefaultPath_ValidatesRelativePath()
    On Error GoTo TestFail

Arrange:
    Dim DotPathSep As String
    DotPathSep = "." & Application.PathSeparator
    Dim Expected As String
    Expected = ThisWorkbook.Path & Application.PathSeparator & DotPathSep & ThisWorkbook.Name
Act:
    Dim Actual As String
    Actual = VerifyOrGetDefaultPath(DotPathSep & ThisWorkbook.Name, vbNullString)
Assert:
    Assert.AreEqual Expected, Actual, "CheckPath failed with relative path"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("PathCheck")
Private Sub ztcVerifyOrGetDefaultPath_ThrowsIfFileNotFound()
    On Error Resume Next
    Dim FilePathName As String
    FilePathName = VerifyOrGetDefaultPath(vbNullString, vbNullString)
    Guard.AssertExpectedError Assert, ErrNo.FileNotFoundErr
End Sub


'@TestMethod("PathCheck")
Private Sub ztcVerifyOrGetDefaultPath_ThrowsIfAbsolutePathSupplied()
    On Error Resume Next
    Dim FilePathName As String
    FilePathName = VerifyOrGetDefaultPath(Environ$("SystemRoot"), vbNullString)
    Guard.AssertExpectedError Assert, ErrNo.FileNotFoundErr
End Sub


'@TestMethod("PathCheck")
Private Sub ztcVerifyOrGetDefaultPath_ThrowsIfRootedPathSupplied()
    On Error Resume Next
    Dim FilePathName As String
    FilePathName = VerifyOrGetDefaultPath("\ABC\DEF", vbNullString)
    Guard.AssertExpectedError Assert, ErrNo.FileNotFoundErr
End Sub


'@TestMethod("IsFalsy")
Private Sub IsFalsy_VerifiesFalsiness()
    On Error GoTo TestFail
    
Arrange:
    Dim TestVar As Variant
    TestVar = Empty
    Dim TestObj As Object
    Set TestObj = Nothing
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
    Assert.IsTrue IsFalsy(vbNullString), "Empty string literal should be falsy"
    Assert.IsFalse IsFalsy("Some text"), "Non-empty should be truthy"
    Assert.IsTrue IsFalsy(TestVar), "Empty variant should be falsy"
    Assert.IsTrue IsFalsy(0&), "0 should be falsy"
    Assert.IsTrue IsFalsy(0#), "0.0 should be falsy"
    Assert.IsTrue IsFalsy(TestObj), "Not set object should be falsy"
    Assert.IsFalse IsFalsy(TestColl), "Set object should be truthy"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub

