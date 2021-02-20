VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExamplePresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Forms.Example"
Option Explicit
Private wbProxy As WorkbookProxy
'@MemberAttribute VB_VarHelpID, -1
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
    Set viewModel.PossibleValues = wbProxy.SheetB.Data
    viewModel.FieldA = wbProxy.SheetA.FieldA
    viewModel.FieldB = wbProxy.SheetA.FieldB
    
    If Dialog.ShowDialog(viewModel) Then ApplyChanges viewModel
    Set view = Nothing
    
End Sub

Private Sub ApplyChanges(ByVal viewModel As ExampleModel)
    wbProxy.SheetA.FieldA = viewModel.FieldA
    wbProxy.SheetA.FieldB = viewModel.FieldB
End Sub

Private Sub view_ApplyChanges(ByVal viewModel As ExampleModel)
    ApplyChanges viewModel
End Sub
