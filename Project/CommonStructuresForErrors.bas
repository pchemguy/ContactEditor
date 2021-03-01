Attribute VB_Name = "CommonStructuresForErrors"
'@Folder "Common.Guard"
Option Explicit


Private Const adErrInvalidParameterType As Long = &HE3D&
Public Enum ErrNo
    PassedNoErr = 0&
    SubscriptOutOfRange = 9&
    TypeMismatchErr = 13&
    FileNotFoundErr = 53&
    ObjectNotSetErr = 91&
    ObjectRequiredErr = 424&
    InvalidObjectUseErr = 425&
    MemberNotExistErr = 438&
    ActionNotSupportedErr = 445&
    NoObject = 31004&
    
    CustomErr = VBA.vbObjectError + 1000&
    NotImplementedErr = VBA.vbObjectError + 1001&
    DefaultInstanceErr = VBA.vbObjectError + 1011&
    NonDefaultInstanceErr = VBA.vbObjectError + 1012&
    EmptyStringErr = VBA.vbObjectError + 1013&
    SingletonErr = VBA.vbObjectError + 1014&
    UnknownClassErr = VBA.vbObjectError + 1015&
    ObjectSetErr = VBA.vbObjectError + 1091&
    AdoFeatureNotAvailableErr = ADODB.ErrorValueEnum.adErrFeatureNotAvailable
    AdoInTransactionErr = ADODB.ErrorValueEnum.adErrInTransaction
    AdoNotInTransactionErr = ADODB.ErrorValueEnum.adErrInvalidTransaction
    AdoConnectionStringErr = ADODB.ErrorValueEnum.adErrProviderNotFound
    AdoInvalidParameterTypeErr = VBA.vbObjectError + adErrInvalidParameterType
End Enum


Public Type TError
    number As ErrNo
    Name As String
    source As String
    message As String
    description As String
    trapped As Boolean
End Type


'@Ignore ProcedureNotUsed
'@Description("Re-raises the current error, if there is one.")
Public Sub RethrowOnError()
Attribute RethrowOnError.VB_Description = "Re-raises the current error, if there is one."
    With VBA.Err
        If .number <> 0 Then
            Debug.Print "Error " & .number, .description
            .Raise .number
        End If
    End With
End Sub


'@Description("Formats and raises a run-time error.")
Public Sub RaiseError(ByRef errorDetails As TError)
Attribute RaiseError.VB_Description = "Formats and raises a run-time error."
    With errorDetails
        Dim message As Variant
        message = Array("Error:", _
            "name: " & .Name, _
            "number: " & .number, _
            "message: " & .message, _
            "description: " & .description, _
            "source: " & .source)
        Debug.Print Join(message, vbNewLine & vbTab)
        VBA.Err.Raise .number, .source, .message
    End With
End Sub


'@Description("Tests if argument is falsy: 0, False, vbNullString, Empty, Null, Nothing")
Public Function IsFalsy(ByVal arg As Variant) As Boolean
Attribute IsFalsy.VB_Description = "Tests if argument is falsy: 0, False, vbNullString, Empty, Null, Nothing"
    Select Case VarType(arg)
        Case vbEmpty, vbNull
            IsFalsy = True
        Case vbInteger, vbLong, vbSingle, vbDouble
            IsFalsy = Not CBool(arg)
        Case vbString
            IsFalsy = (arg = vbNullString)
        Case vbObject
            IsFalsy = (arg Is Nothing)
        Case vbBoolean
            IsFalsy = Not arg
        Case Else
            IsFalsy = False
    End Select
End Function
