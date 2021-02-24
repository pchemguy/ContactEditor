VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTableModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Represents a data table."
'@Folder "Storage.Table.Model"
'@ModuleDescription "Represents a data table."
'@IgnoreModule ProcedureNotUsed
'@Exposed
Option Explicit


Private Type TDataTableModel
    FieldMap As Scripting.Dictionary
    FieldNames As Variant
    Records As Variant
End Type
Private this As TDataTableModel

Private Sub Class_Initialize()
    Set this.FieldMap = New Scripting.Dictionary
    this.FieldMap.CompareMode = TextCompare
End Sub

Private Sub Class_Terminate()
    Set this.FieldMap = Nothing
End Sub

Public Property Get FieldMap() As Scripting.Dictionary
    Set FieldMap = this.FieldMap
End Property

Public Property Get FieldNames() As Variant
    FieldNames = this.FieldNames
End Property

Public Property Let FieldNames(ByVal FieldNames As Variant)
    this.FieldNames = FieldNames
End Property

Public Property Get Records() As Variant
    Records = this.Records
End Property

Public Property Let Records(ByVal Records As Variant)
    this.Records = Records
End Property
