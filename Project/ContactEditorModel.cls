VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ContactEditorModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder "Forms.Contact Editor"
'@IgnoreModule ProcedureNotUsed
'@Exposed
Option Explicit

Public Enum DataPersistenceMode
    DataPersistenceDisabled
    DataPersistenceOnApply
    DataPersistenceOnExit
End Enum

Private Type TContactEditorModel
    RecordTableManager As DataCompositeManager
    PersistenceMode As DataPersistenceMode
    SuppressEvents As Boolean
End Type
Private this As TContactEditorModel


Private Sub Class_Initialize()
    Set this.RecordTableManager = New DataCompositeManager
    this.SuppressEvents = False
End Sub

Private Sub Class_Terminate()
    Set this.RecordTableManager = Nothing
End Sub

Public Property Get RecordTableManager() As DataCompositeManager
    Set RecordTableManager = this.RecordTableManager
End Property

Public Property Get PersistenceMode() As DataPersistenceMode
    PersistenceMode = this.PersistenceMode
End Property

Public Property Let PersistenceMode(ByVal Mode As DataPersistenceMode)
    this.PersistenceMode = Mode
End Property

Public Property Get SuppressEvents() As Boolean
    SuppressEvents = this.SuppressEvents
End Property

Public Property Let SuppressEvents(ByVal Mode As Boolean)
    this.SuppressEvents = Mode
End Property
