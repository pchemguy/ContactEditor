VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RecordEditorDialog 
   Caption         =   "Example Dialog"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6165
   OleObjectBlob   =   "RecordEditorDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RecordEditorDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Forms.Record Editor"
Option Explicit

Public Event ApplyChanges(ByVal viewModel As ExampleModel)

Private Type TView
    IsCancelled As Boolean
    model As ExampleModel
End Type
Private this As TView

Implements IDialogView

Private Sub AcceptButton_Click()
    Me.Hide
End Sub

Private Sub ApplyButton_Click()
    RaiseEvent ApplyChanges(this.model)
End Sub

Private Sub CancelButton_Click()
    OnCancel
End Sub

Private Sub CodeNameBox_Change()
    this.model.Code = CodeNameBox.Value
End Sub

Private Sub QuantityBox_Change()
    this.model.Quantity = QuantityBox.Value
End Sub

Private Sub OnCancel()
    this.IsCancelled = True
    Me.Hide
End Sub

Private Function IDialogView_ShowDialog(ByVal viewModel As Object) As Boolean
    Set this.model = viewModel
    Me.Show vbModal
    IDialogView_ShowDialog = Not this.IsCancelled
End Function

Private Sub UserForm_Activate()
    InitializeFieldA
    CodeNameBox.Value = this.model.Code
    QuantityBox.Value = this.model.Quantity
End Sub

Private Sub InitializeFieldA()
    Dim allValues As Collection
    Set allValues = this.model.PossibleValues
    
    Dim listValues() As Variant
    ReDim listValues(0 To allValues.Count - 1, 0 To 1)
    
    Dim current As SheetBModel
    Dim i As Long
    For Each current In this.model.PossibleValues
        listValues(i, 0) = current.Code
        listValues(i, 1) = current.Name
        i = i + 1
    Next
    CodeNameBox.Clear
    CodeNameBox.List = listValues
    CodeNameBox.ColumnCount = 2
    CodeNameBox.ColumnWidths = "0,30"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub
