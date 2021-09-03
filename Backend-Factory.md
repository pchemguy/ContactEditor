---
layout: default
title: Factory
nav_order: 1
parent: Backends
permalink: /backends/factory
---

Backends are a part of the [Storage Libary][] (StoreLib). The *IDataRecordStorage* and *IDataTableStorage* classes formalize interfaces of backends from the *DataRecord* and *DataTable* families, respectively. Backends can be instantiated either via their factories of via abstract factories, *DataRecordFactory*&nbsp;/&nbsp;*IDataRecordFactory* and *DataTableFactory*&nbsp;/&nbsp;*IDataTableFactory*. Both abstract factory interface classes expose a single method, *CreateInstance*, with similar signatures:

**IDataRecordFactory**

```vb
Public Function CreateInstance(ByVal Model As DataRecordModel, _
                               ByVal ConnectionString As String, _
                               ByVal TableName As String) As IDataRecordStorage
End Function
```

**IDataTableFactory**

```vb
Public Function CreateInstance(ByVal Model As DataTableModel, _
                               ByVal ConnectionString As String, _
                               ByVal TableName As String) As IDataTableStorage
End Function
```

The first argument should be self-explanatory and is the same for all backends. The meaning and format of the other two arguments vary depending on the backend.

*DataRecordWSheet* and *DataTableWSheet* abstract an Excel Worksheet (in the current implementation, only open Excel workbooks are supported). The backends interpret *ConnectionString* as workbook name and *TableName* as worksheet name: `Application.Workbooks(ConnectionString).Worksheets(TableName)`

*DataTableCSV* abstracts delimiter-separated values text files. The backend interprets *ConnectionString* as a file path and *TableName* as file name, possibly including the field separator in the suffix.

*DataTableADODB* and *DataTableSecureADODB* backends provide wrappers around the ADODB library. By default, they interpret *ConnectionString* as an ADODB connection string and *TableName* as the name of a database table. Additionally, specific database types may have connection string constructors. In the current implementation, only the SQLiteODBC driver has a connection string helper. Backend constructor calls this helper if *ConnectionString* has a predetermined prefix. The helper strips the prefix and interprets the remaining part of the *ConnectionString* as a database filename, possibly with a path prefix. Helpers for other database types may be defined similarly.



<!-- Refeences -->

[Storage Libary]: https://pchemguy.github.io/ContactEditor/storage-library
