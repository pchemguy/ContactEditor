VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ContactEditorForm 
   Caption         =   "Contact Editor"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8910
   OleObjectBlob   =   "ContactEditorForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ContactEditorForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "ContactEditor.Forms.Contact Editor"
Option Explicit

'''' To avoid issues, populate ComboBox.List with array of strings,
'''' cast if necessary (ComboBox.List column elements used for
'''' ComboBox.Value must have the same type as ComboBox.Value,
'''' otherwise expect runtime errors and glitches.

Implements IDialogView

Public Event FormLoaded()
Public Event LoadRecord(ByVal RecordId As String)
Public Event ApplyChanges()
Public Event FormConfirmed()
Public Event FormCancelled(ByRef Cancel As Boolean)

Private Type TView
    Model As ContactEditorModel
    IsCancelled As Boolean
End Type
Private this As TView


Private Function OnCancel() As Boolean
    Dim cancelCancellation As Boolean: cancelCancellation = False
    RaiseEvent FormCancelled(cancelCancellation)
    If Not cancelCancellation Then Me.Hide
    OnCancel = cancelCancellation
End Function


Private Sub id_Change()
    If this.Model.SuppressEvents Then Exit Sub
    RaiseEvent LoadRecord(id.Value)
End Sub


Private Sub OkButton_Click()
    Me.Hide
    RaiseEvent FormConfirmed
End Sub


Private Sub CancelButton_Click()
    '@Ignore FunctionReturnValueDiscarded
    OnCancel
End Sub


Private Sub ApplyButton_Click()
    RaiseEvent ApplyChanges
End Sub


Private Sub UpdateDisabledRadio_Click()
    this.Model.PersistenceMode = DataPersistenceMode.DataPersistenceDisabled
End Sub


Private Sub UpdateOnApplyRadio_Click()
    this.Model.PersistenceMode = DataPersistenceMode.DataPersistenceOnApply
End Sub


Private Sub UpdateOnExitRadio_Click()
    this.Model.PersistenceMode = DataPersistenceMode.DataPersistenceOnExit
End Sub


Private Sub IDialogView_ShowDialog(ByVal viewModel As Object)
    Set this.Model = viewModel
    Me.Show vbModeless
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = Not OnCancel
    End If
End Sub


Private Sub UserForm_Activate()
    InitializeId
    InitializeAge
    InitializeGender
    InitializeTableUpdating
    RaiseEvent FormLoaded
End Sub


Private Sub InitializeGender()
    Dim listValues() As Variant
    listValues = Array("male", "female")
    Gender.Clear
    Gender.List = listValues
End Sub


Private Sub InitializeAge()
    Dim listValues(18 To 80) As Variant
    Dim AgeValue As Long
    For AgeValue = 18 To 80
        listValues(AgeValue) = CStr(AgeValue)
    Next AgeValue
    
    Age.Clear
    Age.List = listValues
End Sub


Private Sub InitializeId()
    id.Clear
    id.List = this.Model.RecordTableManager.Ids
End Sub


Private Sub InitializeTableUpdating()
    UpdateDisabledRadio.Value = True
End Sub


Private Sub FirstName_Change()
    If this.Model.SuppressEvents Then Exit Sub
    this.Model.RecordTableManager.RecordModel.SetField "FirstName", FirstName.Value
End Sub


Private Sub LastName_Change()
    If this.Model.SuppressEvents Then Exit Sub
    this.Model.RecordTableManager.RecordModel.SetField "LastName", LastName.Value
End Sub


Private Sub Age_Change()
    If this.Model.SuppressEvents Then Exit Sub
    this.Model.RecordTableManager.RecordModel.SetField "Age", Age.Value
End Sub


Private Sub Gender_Change()
    If this.Model.SuppressEvents Then Exit Sub
    this.Model.RecordTableManager.RecordModel.SetField "Gender", Gender.Value
End Sub


Private Sub Email_Change()
    If this.Model.SuppressEvents Then Exit Sub
    this.Model.RecordTableManager.RecordModel.SetField "Email", Email.Value
End Sub


Private Sub Country_Change()
    If this.Model.SuppressEvents Then Exit Sub
    this.Model.RecordTableManager.RecordModel.SetField "Country", Country.Value
End Sub


Private Sub Domain_Change()
    If this.Model.SuppressEvents Then Exit Sub
    this.Model.RecordTableManager.RecordModel.SetField "Domain", Domain.Value
End Sub
