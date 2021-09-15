VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbRecordsetStructures"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "SecureADODB.DbRecordset"
'@PredeclaredId
Option Explicit


Public Property Get FieldAttributeEnumMap() As Scripting.Dictionary
    Dim FieldAttributeEnumMapLoc As Scripting.Dictionary
    Set FieldAttributeEnumMapLoc = New Scripting.Dictionary
    With FieldAttributeEnumMapLoc
        .CompareMode = TextCompare
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
        .Item(adFldUnspecified) = "adFldUnspecified"
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


Friend Function GetRecordStatusText(ByVal RecordStatus As Long) As String
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
