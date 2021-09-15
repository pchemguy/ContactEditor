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
'@Folder "SecureADODB.DbRecordset.Meta"
'@ModuleDescription "Retrieves extended informarion about the recordest object and its fields"
'@PredeclaredId
'@Exposed
'@IgnoreModule IndexedDefaultMemberAccess
'@IgnoreModule ValueRequired: False positive
Option Explicit


Private Type TDbRecordsetMeta
    AdoRecordset As ADODB.Recordset
    CursorLocationEnumArrMap(1 To 3) As String
    CursorTypeEnumArrMap(0 To 4) As String
    LockTypeEnumArrMap(0 To 5) As String
    MarshalOptionsEnumArrMap(0 To 1) As String
    DataTypeEnumMapping As AdoTypeMappings
End Type
Private this As TDbRecordsetMeta


Public Function Create(ByVal AdoRst As ADODB.Recordset) As DbRecordsetMeta
    Dim Instance As DbRecordsetMeta
    Set Instance = New DbRecordsetMeta
    Instance.Init AdoRst
    Set Create = Instance
End Function


Friend Sub Init(ByVal AdoRst As ADODB.Recordset)
    Set this.AdoRecordset = AdoRst
        
    With this
        .CursorLocationEnumArrMap(adUseNone) = "adUseNone"
        .CursorLocationEnumArrMap(adUseServer) = "adUseServer"
        .CursorLocationEnumArrMap(adUseClient) = "adUseClient"
        
        .CursorTypeEnumArrMap(adOpenUnspecified + 1) = "adOpenUnspecified"
        .CursorTypeEnumArrMap(adOpenForwardOnly + 1) = "adOpenForwardOnly"
        .CursorTypeEnumArrMap(adOpenKeyset + 1) = "adOpenKeyset"
        .CursorTypeEnumArrMap(adOpenDynamic + 1) = "adOpenDynamic"
        .CursorTypeEnumArrMap(adOpenStatic + 1) = "adOpenStatic"
        
        .LockTypeEnumArrMap(adLockUnspecified + 1) = "adLockUnspecified"
        .LockTypeEnumArrMap(adLockReadOnly + 1) = "adLockReadOnly"
        .LockTypeEnumArrMap(adLockPessimistic + 1) = "adLockPessimistic"
        .LockTypeEnumArrMap(adLockOptimistic + 1) = "adLockOptimistic"
        .LockTypeEnumArrMap(adLockBatchOptimistic + 1) = "adLockBatchOptimistic"
        
        .MarshalOptionsEnumArrMap(adMarshalAll) = "adMarshalAll"
        .MarshalOptionsEnumArrMap(adMarshalModifiedOnly) = "adMarshalModifiedOnly"
    End With
        
    Set this.DataTypeEnumMapping = AdoTypeMappings.Default()
End Sub


Public Function GetRecordsetAttrbutes( _
            Optional ByVal TopLeftCell As Excel.Range = Nothing) As Variant
    Dim RstAttr As Variant
    RstAttr = DbRecordsetStructures.RecordsetAttributes
    With this.AdoRecordset
        RstAttr(1, 2) = .CacheSize
        RstAttr(2, 2) = this.CursorLocationEnumArrMap(.CursorLocation)
        RstAttr(3, 2) = this.CursorTypeEnumArrMap(.CursorType + 1)
        RstAttr(4, 2) = this.LockTypeEnumArrMap(.LockType + 1)
        RstAttr(5, 2) = this.MarshalOptionsEnumArrMap(.MarshalOptions)
        RstAttr(6, 2) = .MaxRecords
        RstAttr(7, 2) = .PageCount
        RstAttr(8, 2) = .PageSize
        RstAttr(9, 2) = .RecordCount
        RstAttr(10, 2) = .Source
        RstAttr(11, 2) = DbRecordsetStructures.GetObjectStateText(.State)
        RstAttr(12, 2) = DbRecordsetStructures.GetRecordStatusText(.Status)
        RstAttr(13, 2) = .StayInSync
    End With
        
    If Not TopLeftCell Is Nothing Then
        Dim OutRange As Excel.Range
        Set OutRange = TopLeftCell.Resize( _
            UBound(RstAttr, 1) - LBound(RstAttr, 1) + 1, _
            UBound(RstAttr, 2) - LBound(RstAttr, 2) + 1 _
        )
        With OutRange
            .Clear
            .Range("B1").Value = "OOOOOOOOOOOOOOOO"
            .Columns(2).AutoFit
            .Value = RstAttr
            .Rows(1).HorizontalAlignment = xlCenter
            .Rows(1).Font.Bold = True
            .Columns(1).AutoFit
            .Columns(3).AutoFit
        End With
    End If
    
    GetRecordsetAttrbutes = RstAttr
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
            Properties(PropertyIndex + 1, 2) = this.DataTypeEnumMapping.DataTypeEnumAsText(.Item(PropertyIndex).Type)
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


Public Function GetCursorOptions( _
            Optional ByVal TopLeftCell As Excel.Range = Nothing) As Variant
    Dim OptionsMap As Scripting.Dictionary
    Set OptionsMap = DbRecordsetStructures.CursorOptionEnumMap
    Dim OptionIDs As Variant
    OptionIDs = OptionsMap.Keys
    Dim OptionStatuses As Variant
    ReDim OptionStatuses(0 To OptionsMap.Count, 1 To 2)
    OptionStatuses(0, 1) = "Option"
    OptionStatuses(0, 2) = "Status"
    
    Dim OptionID As CursorOptionEnum
    Dim OptionIndex As Long
    For OptionIndex = 1 To UBound(OptionIDs) + 1
        OptionID = OptionIDs(OptionIndex - 1)
        OptionStatuses(OptionIndex, 1) = OptionsMap(OptionID)
        OptionStatuses(OptionIndex, 2) = this.AdoRecordset.Supports(OptionID)
    Next OptionIndex
    
    If Not TopLeftCell Is Nothing Then
        Dim OutRange As Excel.Range
        With this
            Set OutRange = TopLeftCell.Resize( _
                UBound(OptionStatuses, 1) - LBound(OptionStatuses, 1) + 1, _
                UBound(OptionStatuses, 2) - LBound(OptionStatuses, 2) + 1 _
            )
        End With
        With OutRange
            .Value = OptionStatuses
            .Rows(1).HorizontalAlignment = xlCenter
            .Rows(1).Font.Bold = True
            .Columns.AutoFit
        End With
    End If
    
    GetCursorOptions = OptionStatuses
End Function

