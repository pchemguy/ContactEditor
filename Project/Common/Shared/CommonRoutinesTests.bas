Attribute VB_Name = "CommonRoutinesTests"
'@Folder "Common.Shared"
'@TestModule
'@IgnoreModule LineLabelNotUsed
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


'This method runs after every test in the module.
'@TestCleanup
Private Sub TestCleanup()
    Err.Clear
End Sub


'===================================================='
'==================== TEST CASES ===================='
'===================================================='


'@TestMethod("PathCheck")
Private Sub ztcVerifyOrGetDefaultPath_ValidatesValidPath()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = ThisWorkbook.Path & Application.PathSeparator & ThisWorkbook.Name
Act:
    Dim Actual As String
    Actual = VerifyOrGetDefaultPath(Expected, Array("db", "sqlite"))
Assert:
    Assert.AreEqual Expected, Actual, "CheckPath failed with valid path"
CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub


'@TestMethod("PathCheck")
Private Sub ztcVerifyOrGetDefaultPath_ValidatesDefaultPath()
    On Error GoTo TestFail

Arrange:
    Dim Expected As String
    Expected = ThisWorkbook.Path & Application.PathSeparator & ThisWorkbook.Name
    Expected = Left$(Expected, InStr(Len(Expected) - 5, Expected, ".")) & "db"
Act:
    Dim Actual As String
    Actual = VerifyOrGetDefaultPath(vbNullString, Array("db", "sqlite"))
Assert:
    Assert.AreEqual Expected, Actual, "CheckPath failed with default path"

CleanExit:
    Exit Sub
TestFail:
    Assert.Fail "Error: " & Err.Number & " - " & Err.Description
End Sub
