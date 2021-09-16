VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbRecordsetStructures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "SecureADODB.DbRecordset.Meta"
'@PredeclaredId
'@IgnoreModule IndexedDefaultMemberAccess
Option Explicit


Public Property Get FieldAttributeEnumMap() As Scripting.Dictionary
    Dim FieldAttributeEnumMapLoc As Scripting.Dictionary
    Set FieldAttributeEnumMapLoc = New Scripting.Dictionary
    With FieldAttributeEnumMapLoc
        .CompareMode = TextCompare
        .Item(adFldUnspecified) = "adFldUnspecified"
        .Item(adFldCacheDeferred) = "adFldCacheDeferred"
        .Item(adFldFixed) = "adFldFixed"
        .Item(adFldIsChapter) = "adFldIsChapter"
        .Item(adFldIsCollection) = "adFldIsCollection"
        .Item(adFldIsDefaultStream) = "adFldIsDefaultStream"
        .Item(adFldIsNullable) = "adFldIsNullable"
        .Item(adFldIsRowURL) = "adFldIsRowURL"
        .Item(adFldLong) = "adFldLong"
        .Item(adFldMayBeNull) = "adFldMayBeNull"
        .Item(adFldMayDefer) = "adFldMayDefer"
        .Item(adFldNegativeScale) = "adFldNegativeScale"
        .Item(adFldRowID) = "adFldRowID"
        .Item(adFldRowVersion) = "adFldRowVersion"
        .Item(adFldUnknownUpdatable) = "adFldUnknownUpdatable"
        .Item(adFldUpdatable) = "adFldUpdatable"
    End With
    Set FieldAttributeEnumMap = FieldAttributeEnumMapLoc
End Property


Public Property Get RecordStatusEnumMap() As Scripting.Dictionary
    Dim RecordStatusEnumMapLoc As Scripting.Dictionary
    Set RecordStatusEnumMapLoc = New Scripting.Dictionary
    With RecordStatusEnumMapLoc
        .CompareMode = TextCompare
        .Item(adRecOK) = "adRecOK"
        .Item(adRecCanceled) = "adRecCanceled"
        .Item(adRecCantRelease) = "adRecCantRelease"
        .Item(adRecConcurrencyViolation) = "adRecConcurrencyViolation"
        .Item(adRecDBDeleted) = "adRecDBDeleted"
        .Item(adRecDeleted) = "adRecDeleted"
        .Item(adRecIntegrityViolation) = "adRecIntegrityViolation"
        .Item(adRecInvalid) = "adRecInvalid"
        .Item(adRecMaxChangesExceeded) = "adRecMaxChangesExceeded"
        .Item(adRecModified) = "adRecModified"
        .Item(adRecMultipleChanges) = "adRecMultipleChanges"
        .Item(adRecNew) = "adRecNew"
        .Item(adRecObjectOpen) = "adRecObjectOpen"
        .Item(adRecOutOfMemory) = "adRecOutOfMemory"
        .Item(adRecPendingChanges) = "adRecPendingChanges"
        .Item(adRecPermissionDenied) = "adRecPermissionDenied"
        .Item(adRecSchemaViolation) = "adRecSchemaViolation"
        .Item(adRecUnmodified) = "adRecUnmodified"
    End With
    Set RecordStatusEnumMap = RecordStatusEnumMapLoc
End Property

    
Public Property Get CursorOptionEnumMap() As Scripting.Dictionary
    Dim CursorOptionEnumMapLoc As Scripting.Dictionary
    Set CursorOptionEnumMapLoc = New Scripting.Dictionary
    With CursorOptionEnumMapLoc
        .CompareMode = TextCompare
        .Item(adAddNew) = "adAddNew"
        .Item(adApproxPosition) = "adApproxPosition"
        .Item(adBookmark) = "adBookmark"
        .Item(adDelete) = "adDelete"
        .Item(adFind) = "adFind"
        .Item(adHoldRecords) = "adHoldRecords"
        .Item(adIndex) = "adIndex"
        .Item(adMovePrevious) = "adMovePrevious"
        .Item(adNotify) = "adNotify"
        .Item(adResync) = "adResync"
        .Item(adSeek) = "adSeek"
        .Item(adUpdate) = "adUpdate"
        .Item(adUpdateBatch) = "adUpdateBatch"
    End With
    Set CursorOptionEnumMap = CursorOptionEnumMapLoc
End Property


Public Property Get RecordsetAttributes() As Variant
    Dim RecordsetAttributesLoc(0 To 13, 1 To 3) As Variant
    RecordsetAttributesLoc(0, 1) = "Name"
    RecordsetAttributesLoc(0, 2) = "Value"
    RecordsetAttributesLoc(0, 3) = "Type"
    
    RecordsetAttributesLoc(1, 1) = "CacheSize"
    RecordsetAttributesLoc(2, 1) = "CursorLocation"
    RecordsetAttributesLoc(3, 1) = "CursorType"
    RecordsetAttributesLoc(4, 1) = "LockType"
    RecordsetAttributesLoc(5, 1) = "MarshalOptions"
    RecordsetAttributesLoc(6, 1) = "MaxRecords"
    RecordsetAttributesLoc(7, 1) = "PageCount"
    RecordsetAttributesLoc(8, 1) = "PageSize"
    RecordsetAttributesLoc(9, 1) = "RecordCount"
    RecordsetAttributesLoc(10, 1) = "Source"
    RecordsetAttributesLoc(11, 1) = "State"
    RecordsetAttributesLoc(12, 1) = "Status"
    RecordsetAttributesLoc(13, 1) = "StayInSync"
    
    RecordsetAttributesLoc(1, 3) = "Long"
    RecordsetAttributesLoc(2, 3) = "CursorLocationEnum"
    RecordsetAttributesLoc(3, 3) = "CursorTypeEnum"
    RecordsetAttributesLoc(4, 3) = "LockTypeEnum"
    RecordsetAttributesLoc(5, 3) = "MarshalOptionsEnum"
    RecordsetAttributesLoc(6, 3) = "Long"
    RecordsetAttributesLoc(7, 3) = "Long"
    RecordsetAttributesLoc(8, 3) = "Long"
    RecordsetAttributesLoc(9, 3) = "Long"
    RecordsetAttributesLoc(10, 3) = "String"
    RecordsetAttributesLoc(11, 3) = "ObjectStateEnum"
    RecordsetAttributesLoc(12, 3) = "RecordStatusEnum"
    RecordsetAttributesLoc(13, 3) = "Boolean"
    RecordsetAttributes = RecordsetAttributesLoc
End Property


Public Function GetRecordStatusText(ByVal RecordStatus As Long) As String
    Dim RecordStatusEnumMapLoc As Scripting.Dictionary
    Set RecordStatusEnumMapLoc = RecordStatusEnumMap
    Dim RecordStatusEnumArray As Variant
    RecordStatusEnumArray = RecordStatusEnumMapLoc.Keys()
    
    Dim Status As RecordStatusEnum
    Status = RecordStatusEnumArray(0)
    If RecordStatus = Status Then
        GetRecordStatusText = RecordStatusEnumMap(Status)
        Exit Function
    End If
    
    Dim StatusText As String
    StatusText = vbNullString
    Dim StatusIndex As Long
    For StatusIndex = 1 To UBound(RecordStatusEnumArray)
        Status = RecordStatusEnumArray(StatusIndex)
        If (RecordStatus And Status) = Status Then
            StatusText = StatusText & RecordStatusEnumMapLoc(Status) & vbLf
        End If
    Next StatusIndex
    
    GetRecordStatusText = Left(StatusText, Len(StatusText) - 1)
End Function


Public Function GetObjectStateText(ByVal ObjectState As Long) As String
    Dim ObjectStateEnumArrMap As Variant
    ObjectStateEnumArrMap = Array( _
        "adStateClosed", "adStateOpen", "adStateConnecting", vbNullString, _
        "adStateExecuting", vbNullString, vbNullString, vbNullString, _
        "adStateFetching" _
    )
    
    If ObjectState = 0 Then
        GetObjectStateText = ObjectStateEnumArrMap(0)
        Exit Function
    End If

    Dim StateText As String
    StateText = vbNullString
    Dim State As ObjectStateEnum
    State = 1
    Do While State <= UBound(ObjectStateEnumArrMap)
        If (ObjectState And State) = State Then
            StateText = StateText & ObjectStateEnumArrMap(State) & vbLf
        End If
        State = State * 2
    Loop
    
    GetObjectStateText = Left(StateText, Len(StateText) - 1)
End Function


Public Function GetProperties(ByVal RawProperties As Properties) As Variant
    Dim TypeMapping As AdoTypeMappings
    Set TypeMapping = AdoTypeMappings.Default()
    Dim PropertyCount As Long
    PropertyCount = RawProperties.Count
    
    Dim Props As Variant: Props = Empty '''' RD workaround
    ReDim Props(0 To PropertyCount, 1 To 3)
    Props(0, 1) = "Name"
    Props(0, 2) = "Type"
    Props(0, 3) = "Value"
    Dim PropertyIndex As Long
    Dim Prop As Property
    For PropertyIndex = 1 To PropertyCount
        Set Prop = RawProperties(PropertyIndex - 1)
        Props(PropertyIndex, 1) = Prop.Name
        Props(PropertyIndex, 2) = TypeMapping.DataTypeEnumAsText(Prop.Type)
        Props(PropertyIndex, 3) = Prop.Value
    Next PropertyIndex
    GetProperties = Props
End Function


Public Property Get FieldAttributeNames() As Variant
    FieldAttributeNames = Array( _
        "Name", _
        "Type", _
        "Value", _
        "OriginalValue", _
        "UnderlyingValue", _
        "DefinedSize", _
        "Status", _
        "Attributes" _
    )
End Property


Public Function GetFieldAttributesText(ByVal FieldAttributes As Long) As String
    Dim FieldAttributeEnumMapLoc As Scripting.Dictionary
    Set FieldAttributeEnumMapLoc = FieldAttributeEnumMap
    Dim FieldAttributeEnumArray As Variant
    FieldAttributeEnumArray = FieldAttributeEnumMapLoc.Keys()
    
    If FieldAttributes = -1 Then
        GetFieldAttributesText = FieldAttributeEnumMapLoc(-1)
        Exit Function
    End If
    
    Dim FieldAttribute As FieldAttributeEnum
    Dim AttributeText As String
    AttributeText = vbNullString
    Dim AttributeIndex As Long
    For AttributeIndex = 0 To UBound(FieldAttributeEnumArray)
        FieldAttribute = FieldAttributeEnumArray(AttributeIndex)
        If (FieldAttributes And FieldAttribute) = FieldAttribute Then
            AttributeText = AttributeText & FieldAttributeEnumMapLoc(FieldAttribute) & vbLf
        End If
    Next AttributeIndex
    
    GetFieldAttributesText = Left(AttributeText, Len(AttributeText) - 1)
End Function


Public Function GetFieldStatusText(ByVal FieldStatus As Long) As String
    Dim FieldStatusEnumHighArrMap As Variant
    FieldStatusEnumHighArrMap = Array( _
        "adFieldPendingInsert", _
        "adFieldPendingDelete", _
        "adFieldPendingChange", _
        "adFieldPendingUnknown", _
        "adFieldPendingUnknownDelete" _
    )
    Dim FieldStatusEnumLowArrMap As Variant
    FieldStatusEnumLowArrMap = Array( _
        "adFieldOK", vbNullString, "adFieldCantConvertValue", "adFieldIsNull", _
        "adFieldTruncated", "adFieldSignMismatch", "adFieldDataOverflow", _
        "adFieldCantCreate", "adFieldUnavailable", "adFieldPermissionDenied", _
        "adFieldIntegrityViolation", "adFieldSchemaViolation", "adFieldBadStatus", _
        "adFieldDefault", vbNullString, "adFieldIgnore", "adFieldDoesNotExist", _
        "adFieldInvalidURL", "adFieldResourceLocked", "adFieldResourceExists", _
        "adFieldCannotComplete", "adFieldVolumeNotFound", "adFieldOutOfSpace", _
        "adFieldCannotDeleteSource", "adFieldReadOnly", _
        "adFieldResourceOutOfScope", "adFieldAlreadyExists" _
    )

    Dim FieldStatusText As String
    FieldStatusText = vbNullString
    If (FieldStatus And &HFFFF0000) <> 0 Then
        Dim StatusFlag As FieldStatusEnum
        StatusFlag = adFieldPendingInsert
        Dim StatusIndex As Long
        For StatusIndex = 0 To UBound(FieldStatusEnumHighArrMap)
            If (FieldStatus And StatusFlag) <> 0 Then
                FieldStatusText = FieldStatusText & _
                                  FieldStatusEnumHighArrMap(StatusIndex) & vbLf
            End If
            StatusFlag = StatusFlag * 2
        Next StatusIndex
    End If
    FieldStatusText = FieldStatusText & _
                      FieldStatusEnumLowArrMap(FieldStatus And &HFFFF&)

    GetFieldStatusText = FieldStatusText
End Function


'@Ignore ProcedureNotUsed
Public Function GetFilterGroupText(ByVal FilterGroup As FilterGroupEnum) As String
    GetFilterGroupText = Array( _
        "adFilterNone", _
        "adFilterPendingRecords", _
        "adFilterAffectedRecords", _
        "adFilterFetchedRecords", _
        Empty, _
        "adFilterConflictingRecords" _
    )(FilterGroup)
End Function
