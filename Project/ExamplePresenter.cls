VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExamplePresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private wbProxy As WorkbookProxy
Private WithEvents view As ExampleDialog
Attribute view.VB_VarHelpID = -1

Private Property Get Dialog() As IDialogView
    Set Dialog = view
End Property

Public Sub Show(ByVal proxy As WorkbookProxy)
    
    Set wbProxy = proxy
    Set view = New ExampleDialog
    
    Dim viewModel As ExampleModel
    Set viewModel = New ExampleModel
    Set viewModel.PossibleValues = wbProxy.Sheet2.Data
    viewModel.Field1 = wbProxy.Sheet1.Field1
    viewModel.Field2 = wbProxy.Sheet1.Field2
    
    If Dialog.ShowDialog(viewModel) Then ApplyChanges viewModel
    Set view = Nothing
    
End Sub

Private Sub ApplyChanges(ByVal viewModel As ExampleModel)
    wbProxy.Sheet1.Field1 = viewModel.Field1
    wbProxy.Sheet1.Field2 = viewModel.Field2
End Sub

Private Sub view_ApplyChanges(ByVal viewModel As ExampleModel)
    ApplyChanges viewModel
End Sub
