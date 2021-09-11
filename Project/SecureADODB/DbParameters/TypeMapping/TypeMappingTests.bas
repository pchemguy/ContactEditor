Attribute VB_Name = "TypeMappingTests"
'@Folder "SecureADODB.DbParameters.TypeMapping"
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
Private Sub DefaultMapping_MapsTypes()
    Dim sut As ITypeMap
    Set sut = AdoTypeMappings.Default
    
    Dim BooleanValue As Boolean
    Assert.IsTrue sut.IsMapped(TypeName(BooleanValue)), "Boolean is not mapped"
    Assert.AreEqual adBoolean, sut.Mapping(TypeName(BooleanValue)), "Expected Boolean->adBoolean"
    
    Dim ByteValue As Byte
    Assert.IsTrue sut.IsMapped(TypeName(ByteValue)), "Byte is not mapped"
    Assert.AreEqual adInteger, sut.Mapping(TypeName(ByteValue)), "Expected Byte->adInteger"
    
    Dim CurrencyValue As Currency
    Assert.IsTrue sut.IsMapped(TypeName(CurrencyValue)), "Currency is not mapped"
    Assert.AreEqual adCurrency, sut.Mapping(TypeName(CurrencyValue)), "Expected Currency->adCurrency"
    
    Dim DateValue As Date
    Assert.IsTrue sut.IsMapped(TypeName(DateValue)), "Date is not mapped"
    Assert.AreEqual adDate, sut.Mapping(TypeName(DateValue)), "Expected Date->adDate"
    
    Dim DoubleValue As Double
    Assert.IsTrue sut.IsMapped(TypeName(DoubleValue)), "Double is not mapped"
    Assert.AreEqual adDouble, sut.Mapping(TypeName(DoubleValue)), "Expected Double->adDouble"
    
    Dim IntegerValue As Integer
    Assert.IsTrue sut.IsMapped(TypeName(IntegerValue)), "Integer is not mapped"
    Assert.AreEqual adInteger, sut.Mapping(TypeName(IntegerValue)), "Expected Integer->adInteger"
    
    Dim LongValue As Long
    Assert.IsTrue sut.IsMapped(TypeName(LongValue)), "Long is not mapped"
    Assert.AreEqual adInteger, sut.Mapping(TypeName(LongValue)), "Expected Long->adInteger"
    
    Dim SingleValue As Single
    Assert.IsTrue sut.IsMapped(TypeName(SingleValue)), "Single is not mapped"
    Assert.AreEqual adSingle, sut.Mapping(TypeName(SingleValue)), "Expected Single->adSingle"
    
    Dim StringValue As String
    Assert.IsTrue sut.IsMapped(TypeName(StringValue)), "String is not mapped"
    Assert.AreEqual adVarWChar, sut.Mapping(TypeName(StringValue)), "Expected String->adVarWChar"
    
    Assert.IsTrue sut.IsMapped(TypeName(Empty)), "Empty is not mapped"
    Assert.AreEqual adVarChar, sut.Mapping(TypeName(Empty)), "Expected Empty->adVarChar"
    
    Dim NullValue As Variant
    NullValue = Null
    Assert.IsTrue sut.IsMapped(TypeName(NullValue)), "Null is not mapped"
    Assert.AreEqual adVarChar, sut.Mapping(TypeName(NullValue)), "Expected Null->adVarChar"
End Sub


'@TestMethod("Default Type Mappings")
Private Sub CSVMapping_MapsTypes()
    Dim sut As ITypeMap
    Set sut = AdoTypeMappings.CSV
    
    Dim StringValue As String
    Assert.IsTrue sut.IsMapped(TypeName(StringValue)), "String is not mapped"
    Assert.AreEqual adVarChar, sut.Mapping(TypeName(StringValue)), "Expected String->adVarChar for CSV"
End Sub
