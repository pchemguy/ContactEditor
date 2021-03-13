VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Provides SQL queries and templates"
'@Folder "ContactEditor.Storage.Table.SQLlib"
'@ModuleDescription "Provides SQL queries and templates"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

Private Type TSQL
    TableName As String
End Type
Private this As TSQL


Public Property Get TableName() As String
    TableName = this.TableName
End Property


Public Function Create(ByVal TableName As String) As SQLlib
    Dim SQLTool As SQLlib
    Set SQLTool = New SQLlib
    SQLTool.Init TableName
    Set Create = SQLTool
End Function


Public Sub Init(ByVal TableName As String)
    Guard.EmptyString TableName
    this.TableName = TableName
End Sub


Public Function SelectAll(Optional ByVal FieldArray As Variant = Empty) As String
    Dim Fields As String
    If (VarType(FieldArray) And VBA.VbVarType.vbArray) = VBA.VbVarType.vbArray Then
        Fields = Join(FieldArray, ", ")
    ElseIf Len(FieldArray) > 0 Then
        Fields = FieldArray
    Else
        Fields = "*"
    End If
    SelectAll = "SELECT " & Fields & " FROM """ & this.TableName & """"
End Function


Public Function SelectOne() As String
    SelectOne = "SELECT * FROM """ & this.TableName & """ LIMIT 1"
End Function


Public Function AsText(ByVal FieldName As String) As String
    Guard.EmptyString FieldName
    AsText = "CAST(" & FieldName & " AS TEXT) AS " & FieldName
End Function


Public Function SelectIdAsText(ByVal FieldNames As Variant) As String
    Dim TypeCheck As Boolean
    TypeCheck = ((VarType(FieldNames) And VBA.VbVarType.vbArray) = VBA.VbVarType.vbArray)
    Guard.Expression TypeCheck, "SQLlib", "FieldNames is not array"
    
    Dim Fields As Variant: Fields = FieldNames
    Dim IdIndex As Long: IdIndex = LBound(FieldNames, 1)
    Fields(IdIndex) = AsText(Fields(IdIndex))
    SelectIdAsText = SelectAll(Fields)
End Function


Public Function SelectAllAsText(ByVal FieldNames As Variant, ByVal FieldTypes As Variant) As String
    '''' Make sure that both FieldNames and FieldTypes are arrays and have the same base and size
    Dim TypeCheck As Boolean
    TypeCheck = ((VarType(FieldNames) And vbArray) = vbArray)
    Guard.Expression TypeCheck, "SQLlib", "FieldNames is not array"
    TypeCheck = ((VarType(FieldTypes) And vbArray) = vbArray)
    Guard.Expression TypeCheck, "SQLlib", "FieldTypes is not array"
    TypeCheck = (VarType(FieldNames(LBound(FieldNames))) = vbString)
    Guard.Expression TypeCheck, "SQLlib", "FieldNames must be strings"
    Guard.Expression LBound(FieldNames) = LBound(FieldTypes), "SQLlib", "FieldNames and FieldTypes arrays base mismatch"
    Guard.Expression UBound(FieldNames) = UBound(FieldTypes), "SQLlib", "FieldNames and FieldTypes arrays size mismatch"
    
    Dim Fields() As Variant: ReDim Fields(LBound(FieldNames) To UBound(FieldNames))
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        Select Case FieldTypes(FieldIndex)
            Case ADODB.DataTypeEnum.adVarWChar
                Fields(FieldIndex) = FieldNames(FieldIndex)
            Case Else
                Fields(FieldIndex) = AsText(FieldNames(FieldIndex))
        End Select
    Next FieldIndex
    SelectAllAsText = SelectAll(Fields)
End Function


Public Function UpdateSingleRecord(ByVal FieldNames As Variant) As String
    Dim FieldCount As Long
    Dim PKName As String: PKName = FieldNames(LBound(FieldNames)) ' Assumes that the primary key is the first field
    Dim SetFieldList As String
    SetFieldList = Join(FieldNames, ", ") ' Field list for SET clause
    SetFieldList = Right$(SetFieldList, Len(SetFieldList) - Len(PKName) - 2) ' Remove primary key field, which goes into the WHERE clause
    FieldCount = UBound(FieldNames, 1) - LBound(FieldNames, 1) + 1
    
    ' Since there is no standarad routine for repeating a multi-character patterns,
    ' prepare a sequence of placeholder characters with String and use Replace to
    ' generate the final pattern.
    Dim ValuePlaceholderList As String
    ValuePlaceholderList = Replace(String(FieldCount - 2, "@"), "@", "?, ") & "?"
    
    UpdateSingleRecord = "UPDATE """ & this.TableName & """ " & _
                         "SET (" & SetFieldList & ") = (" & ValuePlaceholderList & ") " & _
                         "WHERE " & PKName & " = ?"
End Function

