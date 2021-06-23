---
layout: default
title: Code extensibility
nav_order: 7
parent: Class design patterns
permalink: /class-design/code-extensibility
---

#### Code extensibility

Let us extend the discussion from the previous section and devise a plan for extending *DataTableADODB* class, which is a part of the Storage library devloped for this project. Consider [Fig. 1](#FigDataTable). The goal is to extend functionality provided by the *DataTableADODB* class, while preserving backward compatibility. First, *DataTableADODB* can be extended via new methods or via etxtensions to existing methods without breaking compatibility. This new functionality needs to be exposed on the interface. Changing *IDataTableStorage*, however, will break compatibility, so the new functionality must be exposed via a separate new interface *IDataTableStorageV2*, which should also expose all functionality provided by *IDataTableStorage*. Now, each *DataTableXXX* class may implement  *IDataTableStorageV2* in addition to *IDataTableStorage* (if, e.g., *DataTableCSV* class does not provide functionality exposed via *IDataTableStorageV2*, it does not have to implement it). 

<a name="FigDataTable"></a>
<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/develop/Assets/Diagrams/Class%20Diagram%20-%20Table.svg" alt="FigDataTableModel" width="100%" />
<p align="center"><b>Figure 1. DataTable class diagram</b></p>

*IDataTableStorage* is related to two other classes, *DataTableFactory* and *DataTableManager* ([Fig. 1](#FigDataTable)). Based on the discussion above, several approaches are available for extending *DataTableFactory*. For example, a second abstract factory can be defined:

<p align="right"><b>DataTableFactory.cls</b></p>

```vb
Implements IDataTableStorage
Implements IDataTableStorageV2

'''' Default factory for IDataTableStorage
Public Function Create(ByVal Model As DataTableModel, _
                       ByVal ConnectionString As String, _
                       ByVal TableName As String) As IDataTableStorage
    Dim Instance As DataTableADODB
    Set Instance = New DataTableADODB
    Instance.Init Model, ConnectionString, TableName
    Set Create = Instance
End Function

'''' Default factory for IDataTableStorageV2
Public Function CreateV2(ByVal Model As DataTableModel, _
                         ByVal ConnectionString As String, _
                         ByVal TableName As String) As IDataTableStorageV2
    Dim Instance As DataTableADODB
    Set Instance = New DataTableADODB
    Instance.Init Model, ConnectionString, TableName
    Set Create = Instance
End Function
```

and the *CreateInstance* method should be appropriately implemented on individual interfaces.

*DataTableManager* incorporates *IDataTableStorage* as a private field. *DataTableManagerV2* incorporating *IDataTableStorageV2* should be introduced.  While it may not be immediatly necessary,  *IDataTableManagerV2* should also be introduced (which might be identical initially). In this case, *DataTableManager* is left unchanged, and *DataTableManagerV2* would implement *IDataTableManagerV2* only.
