VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordEditorModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Forms.Record Editor")
Option Explicit


Private Type TRecordEditorModel
    CodeQuantityModel As DataMapModel
    CodeNamesModel As DataTableModel
End Type
Private this As TRecordEditorModel


Private Sub Class_Initialize()
    Set this.CodeQuantityModel = New DataMapModel
    Set this.CodeNamesModel = New DataTableModel
End Sub

Private Sub Class_Terminate()
    Set this.CodeQuantityModel = Nothing
    Set this.CodeNamesModel = Nothing
End Sub

Public Property Get CodeQuantityModel() As DataMapModel
    Set CodeQuantityModel = this.CodeQuantityModel
End Property

Public Property Get CodeNamesModel() As DataTableModel
    Set CodeNamesModel = this.CodeNamesModel
End Property

Public Function GetProperty(ByVal PropertyName As String) As Variant
    GetProperty = this.CodeQuantityModel.Properties(PropertyName)
End Function
