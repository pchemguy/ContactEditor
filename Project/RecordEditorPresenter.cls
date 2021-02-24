VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RecordEditorPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Forms.Record Editor"
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents view As RecordEditorDialog
Attribute view.VB_VarHelpID = -1
Private Type TPresenter
    model As RecordEditorModel
    CodeQuantitySM As IDataMap
    CodeNamesSM As IDataTable
End Type
Private this As TPresenter

Private Property Get Dialog() As IDialogView
    Set Dialog = view
End Property

Public Sub Show()
    Set this.model = New RecordEditorModel
    Set view = New RecordEditorDialog
    
    Dim viewModel As ExampleModel
    Set viewModel = New ExampleModel
    Set viewModel.PossibleValues = wbProxy.SheetB.Data
    viewModel.Code = wbProxy.SheetA.Code
    viewModel.Quantity = wbProxy.SheetA.Value
    
    If Dialog.ShowDialog(viewModel) Then ApplyChanges viewModel
    Set view = Nothing
    
End Sub

Private Sub ApplyChanges(ByVal viewModel As ExampleModel)
    wbProxy.SheetA.Code = viewModel.Code
    wbProxy.SheetA.Value = viewModel.Quantity
End Sub

Private Sub view_ApplyChanges(ByVal viewModel As ExampleModel)
    ApplyChanges viewModel
End Sub
