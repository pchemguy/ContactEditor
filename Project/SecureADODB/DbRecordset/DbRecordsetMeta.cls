VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbRecordsetMeta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Retrieves extended informarion about the recordest object and its fields"
'@Folder "SecureADODB.DbRecordset"
'@ModuleDescription "Retrieves extended informarion about the recordest object and its fields"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
'@IgnoreModule ValueRequired: False positive
Option Explicit


Private Type TDbRecordsetMeta
    AdoRecordset As ADODB.Recordset
    FieldAttributeEnumMap As Scripting.Dictionary
    FieldAttributeEnumArray As Variant
    CursorLocationEnumMap(1 To 3) As String
    CursorTypeEnumMap(0 To 4) As String
    LockTypeEnumMap(0 To 5) As String
    MarshalOptionsEnumMap(0 To 1) As String
    ObjectStateEnumMap(0 To 8) As String
    RecordStatusEnumMap As Scripting.Dictionary
    RecordStatusEnumArray As Variant
    RstCoreAttr As Variant
    CursorOptionEnumMap As Scripting.Dictionary
    CursorOptionEnumArray As Variant
    DataTypeEnumMap As AdoTypeMappings
End Type
Private this As TDbRecordsetMeta


'@DefaultMember
Public Function Create(ByVal AdoRst As ADODB.Recordset) As DbRecordsetMeta
Attribute Create.VB_UserMemId = 0
    Dim Instance As DbRecordsetMeta
    Set Instance = New DbRecordsetMeta
    Instance.Init AdoRst
    Set Create = Instance
End Function


Friend Sub Init(ByVal AdoRst As ADODB.Recordset)
    Set this.AdoRecordset = AdoRst
    
    Set this.FieldAttributeEnumMap = New Scripting.Dictionary
    With this.FieldAttributeEnumMap
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
    this.FieldAttributeEnumArray = Array( _
        adFldCacheDeferred, _
        adFldFixed, _
        adFldIsChapter, _
        adFldIsCollection, _
        adFldIsDefaultStream, _
        adFldIsNullable, _
        adFldIsRowURL, _
        adFldLong, _
        adFldMayBeNull, _
        adFldMayDefer, _
        adFldNegativeScale, _
        adFldRowID, _
        adFldRowVersion, _
        adFldUnknownUpdatable, _
        adFldUnspecified, _
        adFldUpdatable _
    )
    
    Set this.RecordStatusEnumMap = New Scripting.Dictionary
    With this.RecordStatusEnumMap
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
    this.RecordStatusEnumArray = Array( _
        adRecOK, _
        adRecCanceled, _
        adRecCantRelease, _
        adRecConcurrencyViolation, _
        adRecDBDeleted, _
        adRecDeleted, _
        adRecIntegrityViolation, _
        adRecInvalid, _
        adRecMaxChangesExceeded, _
        adRecModified, _
        adRecMultipleChanges, _
        adRecNew, _
        adRecObjectOpen, _
        adRecOutOfMemory, _
        adRecPendingChanges, _
        adRecPermissionDenied, _
        adRecSchemaViolation, _
        adRecUnmodified _
    )
        
    Set this.CursorOptionEnumMap = New Scripting.Dictionary
    With this.CursorOptionEnumMap
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
    this.CursorOptionEnumArray = Array( _
        adAddNew, _
        adApproxPosition, _
        adBookmark, _
        adDelete, _
        adFind, _
        adHoldRecords, _
        adIndex, _
        adMovePrevious, _
        adNotify, _
        adResync, _
        adSeek, _
        adUpdate, _
        adUpdateBatch _
    )
    
    With this
        .CursorLocationEnumMap(adUseNone) = "adUseNone"
        .CursorLocationEnumMap(adUseServer) = "adUseServer"
        .CursorLocationEnumMap(adUseClient) = "adUseClient"
        
        .CursorTypeEnumMap(adOpenUnspecified + 1) = "adOpenUnspecified"
        .CursorTypeEnumMap(adOpenForwardOnly + 1) = "adOpenForwardOnly"
        .CursorTypeEnumMap(adOpenKeyset + 1) = "adOpenKeyset"
        .CursorTypeEnumMap(adOpenDynamic + 1) = "adOpenDynamic"
        .CursorTypeEnumMap(adOpenStatic + 1) = "adOpenStatic"
        
        .LockTypeEnumMap(adLockUnspecified + 1) = "adLockUnspecified"
        .LockTypeEnumMap(adLockReadOnly + 1) = "adLockReadOnly"
        .LockTypeEnumMap(adLockPessimistic + 1) = "adLockPessimistic"
        .LockTypeEnumMap(adLockOptimistic + 1) = "adLockOptimistic"
        .LockTypeEnumMap(adLockBatchOptimistic + 1) = "adLockBatchOptimistic"
        
        .MarshalOptionsEnumMap(adMarshalAll) = "adMarshalAll"
        .MarshalOptionsEnumMap(adMarshalModifiedOnly) = "adMarshalModifiedOnly"
        
        .ObjectStateEnumMap(adStateClosed) = "adStateClosed"
        .ObjectStateEnumMap(adStateOpen) = "adStateOpen"
        .ObjectStateEnumMap(adStateConnecting) = "adStateConnecting"
        .ObjectStateEnumMap(adStateExecuting) = "adStateExecuting"
        .ObjectStateEnumMap(adStateFetching) = "adStateFetching"
    End With
    
    Dim RstCoreAttr(0 To 13, 1 To 3) As Variant
    RstCoreAttr(0, 1) = "Name"
    RstCoreAttr(0, 2) = "Value"
    RstCoreAttr(0, 3) = "Type"
    RstCoreAttr(1, 1) = "CacheSize"
    RstCoreAttr(2, 1) = "CursorLocation"
    RstCoreAttr(3, 1) = "CursorType"
    RstCoreAttr(4, 1) = "LockType"
    RstCoreAttr(5, 1) = "MarshalOptions"
    RstCoreAttr(6, 1) = "MaxRecords"
    RstCoreAttr(7, 1) = "PageCount"
    RstCoreAttr(8, 1) = "PageSize"
    RstCoreAttr(9, 1) = "RecordCount"
    RstCoreAttr(10, 1) = "Source"
    RstCoreAttr(11, 1) = "State"
    RstCoreAttr(12, 1) = "Status"
    RstCoreAttr(13, 1) = "StayInSync"
    RstCoreAttr(1, 3) = "Long"
    RstCoreAttr(2, 3) = "CursorLocationEnum"
    RstCoreAttr(3, 3) = "CursorTypeEnum"
    RstCoreAttr(4, 3) = "LockTypeEnum"
    RstCoreAttr(5, 3) = "MarshalOptionsEnum"
    RstCoreAttr(6, 3) = "Long"
    RstCoreAttr(7, 3) = "Long"
    RstCoreAttr(8, 3) = "Long"
    RstCoreAttr(9, 3) = "Long"
    RstCoreAttr(10, 3) = "String"
    RstCoreAttr(11, 3) = "ObjectStateEnum"
    RstCoreAttr(12, 3) = "RecordStatusEnum"
    RstCoreAttr(13, 3) = "Boolean"
    
    this.RstCoreAttr = RstCoreAttr
    
    Set this.DataTypeEnumMap = AdoTypeMappings.Default()
End Sub


Public Function GetRecordsetCoreAttrbutes( _
            Optional ByVal TopLeftCell As Excel.Range = Nothing) As Variant
    With this.AdoRecordset
        this.RstCoreAttr(1, 2) = .CacheSize
        this.RstCoreAttr(2, 2) = this.CursorLocationEnumMap(.CursorLocation)
        this.RstCoreAttr(3, 2) = this.CursorTypeEnumMap(.CursorType + 1)
        this.RstCoreAttr(4, 2) = this.LockTypeEnumMap(.LockType + 1)
        this.RstCoreAttr(5, 2) = this.MarshalOptionsEnumMap(.MarshalOptions)
        this.RstCoreAttr(6, 2) = .MaxRecords
        this.RstCoreAttr(7, 2) = .PageCount
        this.RstCoreAttr(8, 2) = .PageSize
        this.RstCoreAttr(9, 2) = .RecordCount
        this.RstCoreAttr(10, 2) = .Source
        this.RstCoreAttr(11, 2) = GetObjectStateText(.State)
        this.RstCoreAttr(12, 2) = GetRecordStatusText(.Status)
        this.RstCoreAttr(13, 2) = .StayInSync
    End With
        
    If Not TopLeftCell Is Nothing Then
        Dim OutRange As Excel.Range
        With this
            Set OutRange = TopLeftCell.Resize( _
                UBound(.RstCoreAttr, 1) - LBound(.RstCoreAttr, 1) + 1, _
                UBound(.RstCoreAttr, 2) - LBound(.RstCoreAttr, 2) + 1 _
            )
        End With
        With OutRange
            .Clear
            .Range("B1").Value = "OOOOOOOOOOOOOOOO"
            .Columns(2).AutoFit
            .Value = this.RstCoreAttr
            .Rows(1).HorizontalAlignment = xlCenter
            .Rows(1).Font.Bold = True
            .Columns(1).AutoFit
            .Columns(3).AutoFit
        End With
    End If
    
    GetRecordsetCoreAttrbutes = this.RstCoreAttr
End Function


Friend Function GetRecordStatusText(ByVal RecordStatus As Long) As String
    Dim RecordStatusEnumArray As Variant
    RecordStatusEnumArray = this.RecordStatusEnumArray
    Dim RecordStatusEnumMap As Scripting.Dictionary
    Set RecordStatusEnumMap = this.RecordStatusEnumMap
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
            StatusText = StatusText & RecordStatusEnumMap(Status) & vbLf
        End If
    Next StatusIndex
    
    GetRecordStatusText = Left(StatusText, Len(StatusText) - 1)
End Function


Friend Function GetObjectStateText(ByVal ObjectState As Long) As String
    Dim ObjectStateEnumMap As Variant
    ObjectStateEnumMap = this.ObjectStateEnumMap
    If ObjectState = 0 Then
        GetObjectStateText = ObjectStateEnumMap(0)
        Exit Function
    End If

    Dim StateText As String
    StateText = vbNullString
    Dim State As ObjectStateEnum
    State = 1
    Do While State <= UBound(ObjectStateEnumMap)
        If (ObjectState And State) = State Then
            StateText = StateText & ObjectStateEnumMap(State) & vbLf
        End If
        State = State * 2
    Loop
    
    GetObjectStateText = Left(StateText, Len(StateText) - 1)
End Function


Public Function GetRecordsetProperties( _
            Optional ByVal TopLeftCell As Excel.Range = Nothing) As Variant
    Dim PropertyCount As Long
    PropertyCount = this.AdoRecordset.Properties.Count
    Dim Properties As Variant: Properties = Empty '''' RD workaround
    ReDim Properties(0 To PropertyCount, 1 To 3)
    Properties(0, 1) = "Name"
    Properties(0, 2) = "Type"
    Properties(0, 3) = "Value"
    Dim PropertyIndex As Long
    With this.AdoRecordset.Properties
        For PropertyIndex = 0 To PropertyCount - 1
            Properties(PropertyIndex + 1, 1) = .Item(PropertyIndex).Name
            Properties(PropertyIndex + 1, 2) = this.DataTypeEnumMap.DataTypeEnumAsText(.Item(PropertyIndex).Type)
            Properties(PropertyIndex + 1, 3) = .Item(PropertyIndex).Value
        Next PropertyIndex
    End With
    
    If Not TopLeftCell Is Nothing Then
        Dim OutRange As Excel.Range
        With this
            Set OutRange = TopLeftCell.Resize( _
                UBound(Properties, 1) - LBound(Properties, 1) + 1, _
                UBound(Properties, 2) - LBound(Properties, 2) + 1 _
            )
        End With
        With OutRange
            .Value = Properties
            .Rows(1).HorizontalAlignment = xlCenter
            .Rows(1).Font.Bold = True
            .Columns.AutoFit
        End With
    End If
    
    GetRecordsetProperties = Properties
End Function
