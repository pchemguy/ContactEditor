VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Abstracts Excel worksheet"
'@Folder "Storage.Map.Backend"
'@ModuleDescription "Abstracts Excel worksheet"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess
'@Exposed
Option Explicit

Implements IDataMap

Private Type TDataMap
    Model As DataMapModel
    MapSheet As Excel.Worksheet
End Type
Private this As TDataMap


'@Description "Returns a new IDataMap object."
Public Function Create(ByVal Model As DataMapModel) As IDataMap
Attribute Create.VB_Description = "Returns a new IDataMap object."
    Dim result As SheetMap
    Set result = New SheetMap
    result.Init Model
    Set Create = result
End Function


Public Sub Init(ByVal Model As DataMapModel)
    Set this.Model = Model
End Sub


Private Sub Class_Terminate()
    Set this.Model = Nothing
End Sub


Private Sub GetProperties(ByVal FieldNames As Variant)
    this.Model.Properties.RemoveAll
    
    Dim FieldName As Variant
    For Each FieldName In FieldNames
        this.Model.Properties.Item(FieldName) = this.MapSheet.Range(FieldName).Value
    Next FieldName
End Sub


Private Sub IDataMap_GetData(ByVal ConnectionString As String, ByVal FieldNames As Variant)
    Dim BookSheet As Variant: BookSheet = Split(ConnectionString, "!")
    Set this.MapSheet = Excel.Workbooks(BookSheet(0)).Worksheets(BookSheet(1))
    
    GetProperties FieldNames
End Sub
