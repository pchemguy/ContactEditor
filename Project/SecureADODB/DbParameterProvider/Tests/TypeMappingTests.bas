Attribute VB_Name = "TypeMappingTests"
'@Folder "SecureADODB.DbParameterProvider.Tests"
'@TestModule
'@IgnoreModule
Option Explicit
Option Private Module

Private Const InvalidTypeName As String = "this isn't a valid type name"

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


'@TestMethod("Factory Guard")
Private Sub Default_ThrowsIfNotInvokedFromDefaultInstance()
    On Error GoTo TestFail
    With New AdoTypeMappings
        On Error GoTo CleanFail
        Dim sut As AdoTypeMappings
        Set sut = .Default
        On Error GoTo 0
    End With
CleanFail:
    If Err.Number = ErrNo.NonDefaultInstanceErr Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub


Private Sub DefaultMapping_MapsType(ByVal Name As String)
    Dim sut As ITypeMap
    Set sut = AdoTypeMappings.Default
    Assert.IsTrue sut.IsMapped(Name)
End Sub


'@TestMethod("Type Mappings")
Private Sub Mapping_ThrowsIfUndefined()
    On Error GoTo TestFail
    With AdoTypeMappings.Default
        On Error GoTo CleanFail
        Dim Value As ADODB.DataTypeEnum
        Value = .Mapping(InvalidTypeName)
        On Error GoTo 0
    End With
CleanFail:
    If Err.Number = ErrNo.CustomErr Then Exit Sub
TestFail:
    Assert.Fail "Expected error was not raised."
End Sub

'@TestMethod("Type Mappings")
Private Sub IsMapped_FalseIfUndefined()
    Dim sut As ITypeMap
    Set sut = AdoTypeMappings.Default
    Assert.IsFalse sut.IsMapped(InvalidTypeName)
End Sub


'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsBoolean()
    Dim Value As Boolean
    DefaultMapping_MapsType TypeName(Value)
End Sub


'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsByte()
    Dim Value As Byte
    DefaultMapping_MapsType TypeName(Value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsCurrency()
    Dim Value As Currency
    DefaultMapping_MapsType TypeName(Value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsDate()
    Dim Value As Date
    DefaultMapping_MapsType TypeName(Value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsDouble()
    Dim Value As Double
    DefaultMapping_MapsType TypeName(Value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsInteger()
    Dim Value As Integer
    DefaultMapping_MapsType TypeName(Value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsLong()
    Dim Value As Long
    DefaultMapping_MapsType TypeName(Value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsSingle()
    Dim Value As Single
    DefaultMapping_MapsType TypeName(Value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsString()
    Dim Value As String
    DefaultMapping_MapsType TypeName(Value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsEmpty()
    Dim Value As Variant
    DefaultMapping_MapsType TypeName(Value)
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMapping_MapsNull()
    Dim Value As Variant
    Value = Null
    DefaultMapping_MapsType TypeName(Value)
End Sub

Private Function GetDefaultMappingFor(ByVal Name As String) As ADODB.DataTypeEnum
    On Error GoTo CleanFail
    Dim sut As ITypeMap
    Set sut = AdoTypeMappings.Default
    GetDefaultMappingFor = sut.Mapping(Name)
    Exit Function
CleanFail:
    Assert.Inconclusive "Default mapping is undefined for '" & Name & "'."
End Function

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForBoolean_MapsTo_adBoolean()
    Const Expected = adBoolean
    Dim Value As Boolean
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForByte_MapsTo_adInteger()
    Const Expected = adInteger
    Dim Value As Byte
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForCurrency_MapsTo_adCurrency()
    Const Expected = adCurrency
    Dim Value As Currency
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForDate_MapsTo_adDate()
    Const Expected = adDate
    Dim Value As Date
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForDouble_MapsTo_adDouble()
    Const Expected = adDouble
    Dim Value As Double
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForInteger_MapsTo_adInteger()
    Const Expected = adInteger
    Dim Value As Integer
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForLong_MapsTo_adInteger()
    Const Expected = adInteger
    Dim Value As Long
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForNull_MapsTo_DefaultNullMapping()
    Dim Expected As ADODB.DataTypeEnum
    Expected = AdoTypeMappings.DefaultNullMapping
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Null))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForEmpty_MapsTo_DefaultNullMapping()
    Dim Expected As ADODB.DataTypeEnum
    Expected = AdoTypeMappings.DefaultNullMapping
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Empty))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForSingle_MapsTo_adSingle()
    Const Expected = adSingle
    Dim Value As Single
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Value))
End Sub

'@TestMethod("Default Type Mappings")
Private Sub DefaultMappingForString_MapsTo_adVarWChar()
    Const Expected = adVarWChar
    Dim Value As String
    Assert.AreEqual Expected, GetDefaultMappingFor(TypeName(Value))
End Sub

