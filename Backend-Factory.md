---
layout: default
title: Factory
nav_order: 1
parent: Backends
permalink: /backends/factory
---

#### DataRecordFactory


```vb
Public Function CreateInstance(ByVal Model As DataRecordModel, _
                               ByVal ConnectionString As String, _
                               ByVal TableName As String) As IDataRecordStorage
End Function
```

#### DataTableFactory
