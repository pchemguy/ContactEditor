VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataTableFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@IgnoreModule ProcedureNotUsed
'@Folder "Storage Library.Table.Factory"
'@PredeclaredId
Option Explicit
Option Compare Text

Implements IDataTableFactory

Private Type TDataTableFactory
    ClassName As String
End Type
Private this As TDataTableFactory


'@Ignore ProcedureNotUsed
'@Description("Returns class reference")
Public Property Get Class() As DataTableFactory
Attribute Class.VB_Description = "Returns class reference"
    Set Class = DataTableFactory
End Property


Public Function Create(ByVal ClassName As String) As IDataTableFactory
    Guard.NonDefaultInstance Me
    
    Dim Instance As DataTableFactory
    Set Instance = New DataTableFactory
    Instance.Init ClassName
    Set Create = Instance
End Function


Friend Sub Init(ByVal ClassName As String)
    this.ClassName = ClassName
End Sub


'''' Creates an instance of IDataTableStorage
''''
'''' Args:
''''   ClassName (string):
''''     Name of a DataTableStorage class. Supported backends:
''''       - "ADODB"
''''       - "SecureADODB"
''''       - "Worksheet"
''''       - "CSV"
''''
''''   Model (DataTableModel):
''''     DataTableModel instance
''''
''''   ConnectionString (string):
''''     Information passed to the backend necessary to connect to peristent storage.
''''     Currently, only file-based storages have been implemented, and this argument
''''     is a file name (possibly with path) that can be resolved to an absolute path
''''     pointing to the database file (for ADODB type backends, see the backend).
''''     See backend constructors for resolution code.
''''
''''   TableName (string):
''''     Typically the name of the table. It may also include additional information.
''''
Public Function CreateInstance(ByVal ClassName As String, _
                               ByVal Model As DataTableModel, _
                               ByVal ConnectionString As String, _
                               ByVal TableName As String) As IDataTableStorage
    Dim DataTableClass As Object
    Select Case ClassName
        Case "ADODB"
            Set DataTableClass = DataTableADODB
        Case "SecureADODB"
            Set DataTableClass = DataTableSecureADODB
        Case "Worksheet"
            Set DataTableClass = DataTableWSheet
        Case "CSV"
            Set DataTableClass = DataTableCSV
    End Select
    
    If Not DataTableClass Is Nothing Then
        Set CreateInstance = DataTableClass.Create(Model, ConnectionString, TableName)
    Else
        Dim errorDetails As TError
        With errorDetails
            .Number = ErrNo.NotImplementedErr
            .Name = "NotImplementedErr"
            .Source = "IDataTableFactory"
            .Description = "Unsupported backend: " & ClassName
            .Message = .Description
        End With
        RaiseError errorDetails
    End If
End Function


Private Function IDataTableFactory_CreateInstance(ByVal Model As DataTableModel, _
                                                  ByVal ConnectionString As String, _
                                                  ByVal TableName As String) As IDataTableStorage
    Set IDataTableFactory_CreateInstance = CreateInstance(this.ClassName, Model, ConnectionString, TableName)
End Function
