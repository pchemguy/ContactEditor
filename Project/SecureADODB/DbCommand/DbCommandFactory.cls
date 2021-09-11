VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbCommandFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An implementation of an abstract factory that creates DbCommand objects."
'@Folder "SecureADODB.DbCommand"
'@ModuleDescription("An implementation of an abstract factory that creates DbCommand objects.")
'@PredeclaredId

Option Explicit

Implements IDbCommandFactory

Private Type TFactory
    commandClass As Object
    Base As IDbCommandBase
End Type
Private this As TFactory


'@Description("Default abstract factory")
Public Function Create(ByVal baseCommand As IDbCommandBase, _
              Optional ByVal commandClassName As String = "DbCommand") As IDbCommandFactory
Attribute Create.VB_Description = "Default abstract factory"
    Dim Instance As DbCommandFactory
    Set Instance = New DbCommandFactory
    Instance.Init baseCommand, commandClassName
    Set Create = Instance
End Function


'@Description("Default constructor")
Public Sub Init(ByVal baseCommand As IDbCommandBase, Optional ByVal commandClassName As String = "DbCommand")
Attribute Init.VB_Description = "Default constructor"
    Set this.Base = baseCommand
    Select Case commandClassName
        Case "DbCommand"
            Set this.commandClass = DbCommand
        Case Else
            Dim errorDetails As TError
            With errorDetails
                .Number = ErrNo.NoObject
                .Name = "NoObject"
                .Source = "DbCommandFactory"
                .Message = "Unknown DbCommand class: " & commandClassName
                .Description = .Message
            End With
            RaiseError errorDetails
    End Select
End Sub


'@Description("Default interface factory")
Private Function IDbCommandFactory_CreateInstance(ByVal db As IDbConnection) As IDbCommand
Attribute IDbCommandFactory_CreateInstance.VB_Description = "Default interface factory"
    Set IDbCommandFactory_CreateInstance = this.commandClass.Create(db, this.Base)
End Function
