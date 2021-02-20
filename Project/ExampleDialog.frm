VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExampleDialog 
   Caption         =   "Example Dialog"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   OleObjectBlob   =   "ExampleDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExampleDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Forms.Example"
Option Explicit

Public Event ApplyChanges(ByVal viewModel As ExampleModel)

Private Type TView
    IsCancelled As Boolean
    Model As ExampleModel
End Type
Private this As TView

Implements IDialogView

Private Sub AcceptButton_Click()
    Me.Hide
End Sub

Private Sub ApplyButton_Click()
    RaiseEvent ApplyChanges(this.Model)
End Sub

Private Sub CancelButton_Click()
    OnCancel
End Sub

Private Sub FieldABox_Change()
    this.Model.FieldA = FieldABox.value
End Sub

Private Sub FieldBBox_Change()
    this.Model.FieldB = FieldBBox.value
End Sub

Private Sub OnCancel()
    this.IsCancelled = True
    Me.Hide
End Sub

Private Function IDialogView_ShowDialog(ByVal viewModel As Object) As Boolean
    Set this.Model = viewModel
    Me.Show vbModal
    IDialogView_ShowDialog = Not this.IsCancelled
End Function

Private Sub UserForm_Activate()
    InitializeFieldA
    FieldABox.value = this.Model.FieldA
    FieldBBox.value = this.Model.FieldB
End Sub

Private Sub InitializeFieldA()
    Dim allValues As Collection
    Set allValues = this.Model.PossibleValues
    
    Dim listValues() As Variant
    ReDim listValues(0 To allValues.Count - 1, 0 To 1)
    
    Dim current As SheetBModel
    Dim i As Long
    For Each current In this.Model.PossibleValues
        listValues(i, 0) = current.Code
        listValues(i, 1) = current.Name
        i = i + 1
    Next
    FieldABox.Clear
    FieldABox.List = listValues
    FieldABox.ColumnCount = 2
    FieldABox.ColumnWidths = "0,70"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub
