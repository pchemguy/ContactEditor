Attribute VB_Name = "ContactEditorRunner"
'@Folder "ContactEditor.Forms.Contact Editor"
Option Explicit

'@Ignore MoveFieldCloserToUsage
Private presenter As ContactEditorPresenter


'@EntryPoint
Public Sub RunContactEditor()
    Set presenter = New ContactEditorPresenter
    Dim DataTableBackEnd As String
    
    '''' Available DataTable backends:
    ''''    "ADODB"
    ''''    "Worksheet"
    ''''    "CSV"
    ''''
    '''' Configuration - ContactEditorPresenter.InitializeModel
    ''''
    DataTableBackEnd = "ADODB"
    presenter.Show DataTableBackEnd
End Sub
