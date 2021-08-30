Attribute VB_Name = "ParameterProviderTests"
'@Folder "SecureADODB.DbParameterProvider.Tests"
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


Private Function GetSUT() As IParameterProvider
    Set GetSUT = AdoParameterProvider.Create(GetDefaultMappings)
End Function


Private Function GetDefaultMappings() As ITypeMap
    Set GetDefaultMappings = AdoTypeMappings.Default
End Function


'@TestMethod("Factory Guard")
Private Sub Create_ThrowsGivenNullMappings()
    
    On Error GoTo CleanFail
    Dim sut As IParameterProvider
    Set sut = AdoParameterProvider.Create(Nothing)
    On Error GoTo 0

CleanFail:
    If Err.Number = ErrNo.ObjectNotSetErr Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub


'@TestMethod("Guard Clauses")
Private Sub TypeMappings_ThrowsGivenNullMappings()
    
    On Error GoTo CleanFail
    Dim sut As AdoParameterProvider
    Set sut = AdoParameterProvider.Create(Nothing)
    On Error GoTo 0
    
CleanFail:
    If Err.Number = ErrNo.ObjectNotSetErr Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub


'@TestMethod("ParameterProvider")
Private Sub FromValue_MapsParameterSizeToStringLength()
    Dim sut As IParameterProvider
    Set sut = GetSUT
    
    Const Value = "ABC XYZ"
    
    Dim p As ADODB.Parameter
    Set p = sut.FromValue(Value)
    
    Assert.AreEqual Len(Value), p.Size
End Sub


'@TestMethod("ParameterProvider")
Private Sub FromValue_MapsParameterTypeAsPerMapping()
    Const Expected = DataTypeEnum.adNumeric
    Const Value = 42

    Dim typeMap As ITypeMap
    Set typeMap = AdoTypeMappings.Default()
    If typeMap.Mapping(TypeName(Value)) = Expected Then Assert.Inconclusive "'expected' data type should not be the default mapping for the specified 'value'."
    typeMap.Mapping(TypeName(Value)) = Expected

    Dim sut As IParameterProvider
    Set sut = AdoParameterProvider.Create(typeMap)
    
    Dim p As ADODB.Parameter
    Set p = sut.FromValue(Value)
    
    Assert.AreEqual Expected, p.Type
End Sub


'@TestMethod("ParameterProvider")
Private Sub FromValue_CreatesInputParameters()
    Const Expected = ADODB.adParamInput
    Const Value = 42
    
    Dim sut As IParameterProvider
    Set sut = GetSUT
    
    Dim p As ADODB.Parameter
    Set p = sut.FromValue(Value)
    
    Assert.AreEqual Expected, p.Direction
End Sub


'@TestMethod("ParameterProvider")
Private Sub FromValues_YieldsAsManyParametersAsSuppliedArgs()
    Dim sut As IParameterProvider
    Set sut = GetSUT
    
    Dim args(1 To 4) As Variant '1-based to match collection indexing
    args(1) = True
    args(2) = 42
    args(3) = 34567
    args(4) = "some string"
    
    Dim values As VBA.Collection
    Set values = sut.FromValues(args)
    
    Assert.AreEqual UBound(args), values.Count
End Sub

