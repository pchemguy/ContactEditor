VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContactEditorPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "ContactEditor"
Option Explicit

'@MemberAttribute VB_VarHelpID, -1
Private WithEvents view As ContactEditorForm
Attribute view.VB_VarHelpID = -1

Private Type TPresenter
    Model As ContactEditorModel
    Dialog As IDialogView
End Type
Private this As TPresenter


Public Sub Show(ByVal TableBackEnd As String)
    Set view = New ContactEditorForm
    Set this.Dialog = view
    InitializeModel TableBackEnd
    
    '''' Loads data from the backends into the Model
    this.Model.RecordTableManager.LoadDataIntoModel
    
    this.Dialog.ShowDialog this.Model
End Sub


Private Sub view_FieldChanged(ByVal FieldName As String, ByVal NewValue As Variant)
    this.Model.RecordTableManager.RecordModel.SetField FieldName, NewValue
End Sub


Private Sub view_ApplyChanges()
    ApplyChanges
End Sub


Private Sub ApplyChanges()
    this.Model.RecordTableManager.SaveRecordDataToRecordStorage
    Select Case this.Model.PersistenceMode
        Case DataPersistenceMode.DataPersistenceOnApply
            this.Model.RecordTableManager.UpdateRecordToTable
            this.Model.RecordTableManager.SaveDataFromModel
        Case DataPersistenceMode.DataPersistenceOnExit
            this.Model.RecordTableManager.UpdateRecordToTable
        Case DataPersistenceMode.DataPersistenceDisabled
            Exit Sub
    End Select
End Sub


Private Sub view_FormLoaded()
    LoadFormFromModel
End Sub


Private Sub LoadFormFromModel()
    this.Model.SuppressEvents = True
        
    Dim FieldName As Variant
    Dim FieldIndex As Long
    Dim FieldNames As Variant: FieldNames = this.Model.RecordTableManager.FieldNames
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        FieldName = FieldNames(FieldIndex)
        view.Controls(FieldName).Value = CStr(this.Model.RecordTableManager.RecordModel.GetField(FieldName))
    Next FieldIndex

    this.Model.SuppressEvents = False
End Sub


Private Sub view_LoadRecord(ByVal RecordId As String)
    If this.Model.RecordTableManager.RecordModel.IsDirty Then
        Dim SaveChanges As Boolean
        SaveChanges = MsgBox("Apply unsaved changes?", vbYesNo + vbExclamation + vbDefaultButton2)
        If SaveChanges Then ApplyChanges
    End If
    this.Model.RecordTableManager.LoadRecordFromTable RecordId
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
    If this.Model.PersistenceMode <> DataPersistenceDisabled Then
        this.Model.RecordTableManager.UpdateRecordToTable
        this.Model.RecordTableManager.SaveDataFromModel
    Else
        this.Model.RecordTableManager.SaveRecordDataToRecordStorage
    End If
    Set view = Nothing
End Sub


'''' Note that in present implementation, the database type used with
'''' ADODB/SecureADODB backends is hardcoded as SQLite.
'@Description "Instantiates model and binds it to the desired backends."
Private Sub InitializeModel(ByVal TableBackEnd As String)
Attribute InitializeModel.VB_Description = "Instantiates model and binds it to the desired backends."
    Set this.Model = New ContactEditorModel
    
    Dim ClassName As String
    Dim TableName As String
    Dim ConnectionString As String
    
    '''' Binds TableModel to its backend
    Select Case TableBackEnd
        Case "ADODB"
            ClassName = "ADODB"
            TableName = "Contacts"
            ConnectionString = "sqlite:"
        Case "SecureADODB"
            ClassName = "SecureADODB"
            TableName = "Contacts"
            ConnectionString = "sqlite:"
        Case "Worksheet"
            ClassName = "Worksheet"
            TableName = Contacts.Name
            ConnectionString = ThisWorkbook.Name
        Case "CSV"
            ClassName = "CSV"
            TableName = "Contacts.xsv!sep=,"
            ConnectionString = ThisWorkbook.Path
    End Select
    this.Model.RecordTableManager.InitTable ClassName, ConnectionString, TableName
    
    '''' Binds RecordModel to its backend
    ClassName = "Worksheet"
    ConnectionString = ThisWorkbook.Name
    TableName = ContactBrowser.Name
    this.Model.RecordTableManager.InitRecord ClassName, ConnectionString, TableName
End Sub
