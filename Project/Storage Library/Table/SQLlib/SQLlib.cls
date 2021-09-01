VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLlib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Provides SQL queries and templates"
'@Folder "Storage Library.Table.SQLlib"
'@ModuleDescription "Provides SQL queries and templates"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit

Public Enum FieldFormat
    CastAsIs
    CastIdAsText
    CastAllAsText
End Enum

Private Type TSQL
    TableName As String
    LIMIT As Long
End Type
Private this As TSQL


Public Property Get TableName() As String
    TableName = this.TableName
End Property


Public Function Create(ByVal TableName As String) As SQLlib
    Dim Instance As SQLlib
    Set Instance = New SQLlib
    Instance.Init TableName
    Set Create = Instance
End Function


Friend Sub Init(ByVal TableName As String)
    Guard.EmptyString TableName
    this.TableName = TableName
End Sub


Public Sub SetLimit(Optional ByVal LimitValue As Long = 0)
    this.LIMIT = LimitValue
End Sub


Public Function SelectAll(Optional ByVal FieldArray As Variant = Empty, _
                          Optional ByVal EscapeFields As Boolean = True) As String
                          
    Dim EscapeOpen As String: EscapeOpen = IIf(EscapeFields, "[", vbNullString)
    Dim EscapeClose As String: EscapeClose = IIf(EscapeFields, "]", vbNullString)
    
    Dim Fields As String
    If (VarType(FieldArray) And vbArray) = vbArray Then
        Fields = EscapeOpen & Join(FieldArray, EscapeClose & ", " & EscapeOpen) & EscapeClose
    ElseIf Len(FieldArray) > 0 Then
        Fields = EscapeOpen & FieldArray & EscapeClose
    Else
        Fields = "*"
    End If
    Dim LimitClause As String
    If this.LIMIT > 0 Then LimitClause = " LIMIT " & CStr(this.LIMIT)
    SelectAll = "SELECT " & Fields & " FROM [" & this.TableName & "]" & LimitClause
End Function


Public Function AsText(ByVal FieldName As String) As String
    Guard.EmptyString FieldName
    AsText = "CAST([" & FieldName & "] AS TEXT) AS [" & FieldName & "]"
End Function


'@Ignore ProcedureNotUsed
Public Function AliasedField(ByVal FieldName As String) As String
    Guard.EmptyString FieldName
    AliasedField = "[" & FieldName & "] AS [" & FieldName & "]"
End Function


Public Function SelectIdAsText(ByVal FieldNames As Variant) As String
    Dim TypeCheck As Boolean
    TypeCheck = ((VarType(FieldNames) And vbArray) = vbArray)
    Guard.Expression TypeCheck, "SQLlib", "FieldNames is not array"
    
    Dim Fields() As String: ReDim Fields(LBound(FieldNames) To UBound(FieldNames))
    Dim FieldIndex As Long: FieldIndex = LBound(FieldNames)
    Fields(FieldIndex) = AsText(FieldNames(FieldIndex))
    For FieldIndex = LBound(FieldNames) + 1 To UBound(FieldNames)
        Fields(FieldIndex) = "[" & FieldNames(FieldIndex) & "]"
    Next FieldIndex
    SelectIdAsText = SelectAll(Fields, False)
End Function


Public Function SelectAllAsText(ByVal FieldNames As Variant, ByVal FieldTypes As Variant) As String
    '''' Make sure that both FieldNames and FieldTypes are arrays and have the same base and size
    Guard.IncompatibleArrays FieldNames, FieldTypes
    Guard.EmptyString FieldNames(LBound(FieldNames))
    
    Dim Fields() As Variant: ReDim Fields(LBound(FieldNames) To UBound(FieldNames))
    Dim FieldIndex As Long
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        Select Case FieldTypes(FieldIndex)
            Case ADODB.DataTypeEnum.adVarWChar
                Fields(FieldIndex) = "[" & FieldNames(FieldIndex) & "]"
            Case Else
                Fields(FieldIndex) = AsText(FieldNames(FieldIndex))
        End Select
    Next FieldIndex
    SelectAllAsText = SelectAll(Fields, False)
End Function


Public Function UpdateSingleRecord(ByVal FieldNames As Variant) As String
    Dim FieldCount As Long
    Dim PKName As String: PKName = FieldNames(LBound(FieldNames)) ' Assumes that the primary key is the first field
    Dim SetFieldList As String
    SetFieldList = "[" & Join(FieldNames, "], [") & "]" ' Field list for SET clause
    SetFieldList = Right$(SetFieldList, Len(SetFieldList) - Len(PKName) - 4) ' Remove primary key field, which goes into the WHERE clause
    FieldCount = UBound(FieldNames, 1) - LBound(FieldNames, 1) + 1
    
    ' Since there is no standarad routine for repeating a multi-character patterns,
    ' prepare a sequence of placeholder characters with String and use Replace to
    ' generate the final pattern.
    Dim ValuePlaceholderList As String
    ValuePlaceholderList = Replace(String(FieldCount - 2, "@"), "@", "?, ") & "?"
    
    UpdateSingleRecord = "UPDATE [" & this.TableName & "] " & _
                         "SET (" & SetFieldList & ") = (" & ValuePlaceholderList & ") " & _
                         "WHERE [" & PKName & "] = ?"
End Function

