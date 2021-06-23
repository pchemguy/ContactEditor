---
layout: default
title: Factory-Constructor pattern
nav_order: 1
parent: Class design conventions
permalink: /class-design/factory-constructor
---

#### Factory-Constructor pattern - parametrized class instantiation

A pair of a factory and a custom constructor performs parametrized class instantiation. The default factory *Create* and the default constructor *Init* are defined on the class's default interface only. Both methods have the same parameter signature but different return values. The Factory method should be a function returning a class instance, and the Constructor method should be a sub with no return value. The factory method called on the default (predeclared) class's instance (enabled via the "Predeclared" attribute) generates a new class instance (via the New operator) and then, to perform initialization, calls instance's constructor with all received arguments. For example, here is a snippet from a class, which is a part of the "Storage" library:

```vb

'''' _____ DataTableADODB.cls _____ ''''

Implements IDataTableStorage

'''' Encapsulated private fields
Private Type TDataTableADODB
    ' Define private fields here
End Type
Private this As TDataTableADODB

'''' A boilerplate template for the default factory
'''' This method is called on the default predeclared class instance
Public Function Create(ByVal Model As DataTableModel, _
                       ByVal ConnectionString As String, _
                       ByVal TableName As String) As IDataTableStorage
    Dim Instance As DataTableADODB
    Set Instance = New DataTableADODB
    Instance.Init Model, ConnectionString, TableName
    Set Create = Instance
End Function

'''' Constructor
'''' This method is called on the default interface of the newly generated class instance
Public Sub Init(ByVal Model As DataTableModel, _
                ByVal ConnectionString As String, _
                ByVal TableName As String)
    ' Check input parameters and initialize private data fields here.
End Sub

```

To simulate rudimentary introspection, *Class* and *Self* getters can also be defined. The *Class* getter returns the class's default instance. If a class instance presents a non-default interface, *Self* should return the same interface as well.

```vb

'''' Self attribute defined on the default interface
Public Property Get Self() As IDataTableStorage
    Set Self = Me
End Property

'''' Class attribute defined on the default interface
Public Property Get Class() As DataTableADODB
    Set Class = DataTableADODB
End Property

```
