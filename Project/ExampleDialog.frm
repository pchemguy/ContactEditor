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

Private Sub Field1Box_Change()
    this.Model.Field1 = Field1Box.value
End Sub

Private Sub Field2Box_Change()
    this.Model.Field2 = Field2Box.value
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
    InitializeField1
    Field1Box.value = this.Model.Field1
    Field2Box.value = this.Model.Field2
End Sub

Private Sub InitializeField1()
    Dim allValues As Collection
    Set allValues = this.Model.PossibleValues
    
    Dim listValues() As Variant
    ReDim listValues(0 To allValues.Count - 1, 0 To 1)
    
    Dim current As SomeModel, i As Long
    For Each current In this.Model.PossibleValues
        listValues(i, 0) = current.Code
        listValues(i, 1) = current.Name
        i = i + 1
    Next
    Field1Box.Clear
    Field1Box.List = listValues
    Field1Box.ColumnCount = 2
    Field1Box.ColumnWidths = "0,70"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = VbQueryClose.vbFormControlMenu Then
        Cancel = True
        OnCancel
    End If
End Sub
