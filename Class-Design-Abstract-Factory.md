---
layout: default
title: Abstract factory
nav_order: 3
parent: Design patterns
permalink: /class-design/abstract-factory
---

#### Abstract factory

An abstract factory class has two default factory methods. The *Create* method follows the same convention as for a regular class. It is available on the default predeclared instance of the abstract factory only and generates factory instances:

<p align="right"><b>DataTableFactory.cls</b></p>

```vb
Implements IDataTableFactory

'''' A boilerplate template for the default factory
'''' This method is called on the default predeclared class instance
Public Function Create(ByVal ClassName As String) As IDataTableFactory
    Dim Instance As DataTableFactory
    Set Instance = New DataTableFactory
    Instance.Init ClassName
    Set Create = Instance
End Function
```

The other factory is *CreateInstance*. It must be available on non-default factory instances, but it can also be available on the default instance. This factory generates instances of the target class, e.g.:

```vb
Public Function CreateInstance(ByVal ClassName As String, _
                               ByVal Model As DataTableModel, _
                               ByVal ConnectionString As String, _
                               ByVal TableName As String) As IDataTableStorage
    Select Case ClassName
        Case "ADODB"
            Set CreateInstance = DataTableADODB.Create(Model, ConnectionString, TableName)
        Case "Worksheet"
            Set CreateInstance = DataTableWSheet.Create(Model, ConnectionString, TableName)
        Case "CSV"
            Set CreateInstance = DataTableCSV.Create(Model, ConnectionString, TableName)
        Case Else
            Dim errorDetails As TError
            With errorDetails
                .Number = ErrNo.NotImplementedErr
                .Name = "NotImplementedErr"
                .Source = "IDataTableFactory"
                .Description = "Unsupported backend: " & ClassName
                .Message = .Description
            End With
            RaiseError errorDetails
    End Select
End Function

Private Function IDataTableFactory_CreateInstance( _
                    ByVal Model As DataTableModel, _
                    ConnectionString As String, _
                    ByVal TableName As String) As IDataTableStorage
    Set IDataTableFactory_CreateInstance = CreateInstance( _
        this.ClassName, Model, ConnectionString, TableName)
End Function
```
