VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContactEditorModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@IgnoreModule ProcedureNotUsed
'@Folder "Forms.Contact Editor"
'@Exposed
Option Explicit


Private Type TContactEditorModel
    FormModel As DataRecordModel
    DataModel As DataTableModel
End Type
Private this As TContactEditorModel


Private Sub Class_Initialize()
    Set this.FormModel = New DataRecordModel
    Set this.DataModel = New DataTableModel
End Sub

Private Sub Class_Terminate()
    Set this.FormModel = Nothing
    Set this.DataModel = Nothing
End Sub

Public Property Get FormModel() As DataRecordModel
    Set FormModel = this.FormModel
End Property

Public Property Get DataModel() As DataTableModel
    Set DataModel = this.DataModel
End Property
