VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContactEditorPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ContactEditor.Forms.Contact Editor"
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents view As ContactEditorForm
Attribute view.VB_VarHelpID = -1

Private Type TPresenter
    viewModel As ContactEditorModel
    Dialog As IDialogView
End Type
Private this As TPresenter


Public Sub Show()
    Set view = New ContactEditorForm
    Set this.Dialog = view
    InitializeModel
    
    '''' Loads data from the backends into the Model
    this.viewModel.RecordTableManager.LoadDataIntoModel
    
    this.Dialog.ShowDialog this.viewModel
End Sub


Private Sub ApplyChanges()
    this.viewModel.RecordTableManager.SaveRecordDataToRecordStorage
    Select Case this.viewModel.PersistenceMode
        Case DataPersistenceMode.DataPersistenceOnApply
            this.viewModel.RecordTableManager.UpdateRecordToTable
            this.viewModel.RecordTableManager.SaveDataFromModel
        Case DataPersistenceMode.DataPersistenceOnExit
            this.viewModel.RecordTableManager.UpdateRecordToTable
        Case DataPersistenceMode.DataPersistenceDisabled
            Exit Sub
    End Select
End Sub


Private Sub view_ApplyChanges()
    ApplyChanges
End Sub


Private Sub view_FormLoaded()
    LoadFormFromModel
End Sub


Private Sub LoadFormFromModel()
    this.viewModel.SuppressEvents = True
        
    Dim FieldName As Variant
    Dim FieldIndex As Long
    Dim FieldNames As Variant: FieldNames = this.viewModel.RecordTableManager.FieldNames
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        FieldName = FieldNames(FieldIndex)
        view.Controls(FieldName).Value = CStr(this.viewModel.RecordTableManager.RecordModel.GetField(FieldName))
    Next FieldIndex

    this.viewModel.SuppressEvents = False
End Sub


Private Sub view_LoadRecord(ByVal RecordId As String)
    If this.viewModel.RecordTableManager.RecordModel.IsDirty Then
        Dim SaveChanges As Boolean
        SaveChanges = MsgBox("Apply unsaved changes?", vbYesNo + vbExclamation + vbDefaultButton2)
        If SaveChanges Then ApplyChanges
    End If
    this.viewModel.RecordTableManager.LoadRecordFromTable RecordId
    LoadFormFromModel
End Sub


Private Sub view_FormCancelled(ByRef Cancel As Boolean)
    'setting Cancel to True will leave the form open
    Cancel = MsgBox("Cancel this operation?", vbYesNo + vbExclamation) = vbNo
    If Not Cancel Then
        ' modeless form was cancelled and is now hidden.
        ' ...
        Set view = Nothing
    End If
End Sub


Private Sub view_FormConfirmed()
    'form was okayed and is now hidden.
    '...
    If this.viewModel.PersistenceMode <> DataPersistenceDisabled Then
        this.viewModel.RecordTableManager.UpdateRecordToTable
        this.viewModel.RecordTableManager.SaveDataFromModel
    Else
        this.viewModel.RecordTableManager.SaveRecordDataToRecordStorage
    End If
    Set view = Nothing
End Sub


'@Description "Instantiates model and binds it to the desired backends."
Private Sub InitializeModel()
Attribute InitializeModel.VB_Description = "Instantiates model and binds it to the desired backends."
    Set this.viewModel = New ContactEditorModel
    
    Dim ClassName As String
    Dim TableName As String
    Dim ConnectionString As String
    
    '''' Binds TableModel to its backend
    ClassName = "Worksheet"
    TableName = "Contacts"
    ConnectionString = ThisWorkbook.Name & "!" & Contacts.Name
    this.viewModel.RecordTableManager.InitTable ClassName, ConnectionString, TableName
    
    '''' Binds RecordModel to its backend
    ClassName = "Worksheet"
    TableName = vbNullString
    ConnectionString = ThisWorkbook.Name & "!" & ContactBrowser.Name
    this.viewModel.RecordTableManager.InitRecord ClassName, ConnectionString, TableName
End Sub
