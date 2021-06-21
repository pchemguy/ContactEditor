---
layout: default
title: Class design convention
nav_order: 3
permalink: /class-design
---

#### Factory-Constructor pattern - parametrized class instantiation

A pair of a factory and a custom constructor performs parametrized class instantiation. The default factory *Create* and the default constructor *Init* are defined on the class's default interface only. Both methods have the same parameter signature but different return values. The Factory method should be a function returning a class instance, and the Constructor method should be a sub with no return value. The factory method called on the default (predeclared) class's instance (enabled via the "Predeclared" attribute) generates a new class instance (via the New operator) and then, to perform initialization, calls instance's constructor with all received arguments. For example, here is a snippet from a class, which is a part of the "Storage" library:

```vba
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

```vba
'''' Self attribute defined on the default interface
Public Property Get Self() As IDataTableStorage
    Set Self = Me
End Property

'''' Class attribute defined on the default interface
Public Property Get Class() As DataTableADODB
    Set Class = DataTableADODB
End Property
```

#### Abstract factory

An abstract factory class has two default factory methods. The *Create* method follows the same convention as for a regular class. It is available on the default predeclared instance of the abstract factory only and generates factory instances:

```vba
'''' _____ DataTableFactory.cls _____ ''''

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

```vba
'''' _____ DataTableFactory.cls _____ ''''

Implements IDataTableFactory

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

#### Factory return type

Typically, a class's factory should return an instance of that class. *Abstract factory pattern* may be used to encapsulate a group of classes implementing the same interface. *CreateInstance* method of an abstract factory should call a concrete factory of the selected class returning an instance of that class. This pattern is illustrated by the code in the previous section and in [Fig. 1](#FigDataTableModel) with *DataTableXXX* family, *IDataTableStorage*, and *DataTableFactory*.

<a name="FigDataTableModel"></a>
<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/develop/Assets/Diagrams/Class%20Diagram%20-%20Table.svg" alt="FigDataTableModel" width="100%" />
<p align="center"><b>Figure 1. DataTableModel class diagram</b></p>

If a class implements an interface, it is customarily to define the factory's return type as an instance of such an interface:

```vba
'''' _____ DataTableADODB.cls _____ ''''

Implements IDataTableStorage

Public Function Create(ByVal Model As DataTableModel, _
					   ByVal ConnectionString As String, _
					   ByVal TableName As String) As IDataTableStorage ' <=== style #1
    Dim Instance As DataTableADODB
    Set Instance = New DataTableADODB
    Instance.Init Model, ConnectionString, TableName
    Set Create = Instance
End Function
```

However, this convention is not mandatory. Compare it with the following (the only difference is in the declared return type indicated by an arrow on the right):

```vba
'''' _____ DataTableADODB.cls _____ ''''

Implements IDataTableStorage

Public Function Create(ByVal Model As DataTableModel, _
					   ByVal ConnectionString As String, _
					   ByVal TableName As String) As DataTableADODB ' <=== style #2
    Dim Instance As DataTableADODB
    Set Instance = New DataTableADODB
    Instance.Init Model, ConnectionString, TableName
    Set Create = Instance
End Function
```

and yet another option:

```vba
'''' _____ DataTableADODB.cls _____ ''''

Implements IDataTableStorage

Public Function Create(ByVal Model As DataTableModel, _
					   ByVal ConnectionString As String, _
					   ByVal TableName As String) As Object ' <=== style #3
    Dim Instance As DataTableADODB
    Set Instance = New DataTableADODB
    Instance.Init Model, ConnectionString, TableName
    Set Create = Instance
End Function
```

Now the calling code instantiating IDataTableStorage would look like this:

```vba
	Dim DataTable as IDataTableStorage
	Set DataTable = DataTableADODB.Create(Model, ConnectionString, TableName)
```

yielding essentially identical behavior for all declaration styles. The subtle and technically irrelevant in this context detail is at which point *DataTableADODB* is "cast" as *IDataTableStorage* (the interface is changed within the factory in the first case and during assignment of the returned instance reference in the calling code for styles #2 and #3).

<a name="SectionDirectChaining"></a>

There is one circumstance, however, in which styles #1 and #2 affect the calling code. Style #1, which returns a specific non-default interface, permits chaining the interface methods directly on the factory (ClassA must have the predeclared attribute set to True):

```vba
'''' _____ ClassA.cls _____ ''''

Implements IClassA

Public Function Create() As IClassA
    Dim Instance As ClassA
    Set Instance = New ClassA
    Instance.Init 
    Set Create = Instance
End Function

Private Sub IClassA_SomeMethod()
	' Do something
End Sub

'''' _____ ClassAUser.bas _____ ''''
	ClassA.Create().SomeMethod()
```

Direct chaining is not possible with style #2, returning the default interface. The returned result must be assigned to an appropriately declared local variable to switch the interface. Another approach would be to add methods returning specific interfaces to the default interface:

```vba
'''' _____ ClassA.cls _____ ''''

Implements IClassA
Implements IClassAV2

Public Function Create() As ClassA
    Dim Instance As ClassA
    Set Instance = New ClassA
    Instance.Init 
    Set Create = Instance
End Function

Public Function IV1() As IClassA
    Set IV1 = Me
End Function

Public Function IV2() As IClassAV2
    Set IV2 = Me
End Function

Private Sub IClassA_SomeMethod()
	' Do something
End Sub

'''' _____ ClassAUser.bas _____ ''''
	ClassA.Create().IV1().SomeMethod()
	
	' or
	Dim Instance as IClassA
	Set Instance = ClassA.Create()
	Instance.SomeMethod()
```

Let us suppose we want to extend the functionality of *DataTableADODB* and expose it via a new *IDataTableStorageV2* interface while keeping the old interface unaffected for backward compatibility of the "Storage" library (assuming that the factory signature does not need to be changed):

```vba
'''' _____ DataTableADODB.cls _____ ''''

Implements IDataTableStorage
Implements IDataTableStorageV2

Public Function Create(ByVal Model As DataTableModel, _
					   ByVal ConnectionString As String, _
					   ByVal TableName As String) As DataTableADODB
    Dim Instance As DataTableADODB
    Set Instance = New DataTableADODB
    Instance.Init Model, ConnectionString, TableName
    Set Create = Instance
End Function
```
and the calling code:

```vba
	'''' Code using the old interface
	Dim DataTable as IDataTableStorage
	Set DataTable = DataTableADODB.Create(Model, ConnectionString, TableName)

	'''' Code using the new interface
	Dim DataTableEx as IDataTableStorageV2
	Set DataTableEx = DataTableADODB.Create(Model, ConnectionString, TableName)
```

In other words, the same factory can be used to generate both the old and new interfaces. Style #3, which is the most general object declaration, would work here as well. However, this approach imposes limitations on static code analysis, compile-time checks, and IntelliSense. Style #1 would also work based on preliminary testing, though RubberDuck complains about incompatible type assignment; nevertheless, style #2 is preferable over style #1 here.

The design of the abstract factory pattern shown in the previous section is not compatible with style #2, as its *CreateInstance* method, defined on the default interface, must return a common interface (see the snippet in the "Abstract factory" [section](#Abstract%20factory)). But, if we assume a slightly stricter convention, removing *CreateInstance* from the factory's default interface, style #2 will be usable:

```vba
'''' _____ DataTableFactory.cls _____ ''''

Implements IDataTableFactory
Implements IDataTableFactoryV2

'''' Compatible with both IDataTableFactory and IDataTableFactoryV2
Public Function Create(ByVal ClassName As String) As DataTableFactory
    Dim Instance As DataTableFactory
    Set Instance = New DataTableFactory
    Instance.Init ClassName
    Set Create = Instance
End Function

Private Function IDataTableFactory_CreateInstance( _
					ByVal Model As DataTableModel, _
			        ConnectionString As String, _
					ByVal TableName As String) As IDataTableStorage
	Dim Instance as IDataTableStorage
    Select Case ClassName
        Case "ADODB"
            Set Instance = DataTableADODB.Create(Model, ConnectionString, TableName)
        Case "Worksheet"
            Set Instance = DataTableWSheet.Create(Model, ConnectionString, TableName)
        Case "CSV"
            Set Instance = DataTableCSV.Create(Model, ConnectionString, TableName)
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
    Set IDataTableFactory_CreateInstance = Instance
End Function

Private Function IDataTableFactoryV2_CreateInstance( _
					ByVal Model As DataTableModel, _
			        ConnectionString As String, _
					ByVal TableName As String) As IDataTableStorageV2
	Dim Instance as IDataTableStorageV2
    Select Case ClassName
        Case "ADODB"
            Set Instance = DataTableADODB.Create(Model, ConnectionString, TableName)
        Case "Worksheet"
            Set Instance = DataTableWSheet.Create(Model, ConnectionString, TableName)
        Case "CSV"
            Set Instance = DataTableCSV.Create(Model, ConnectionString, TableName)
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
    Set IDataTableFactoryV2_CreateInstance = Instance
End Function
```

To enable direct chaining, methods IV1 and IV2 can be defined on the default instance as discussed [above](#SectionDirectChaining).  Alternatively, separate factories returning specific interfaces may be defined (concrete factories should only be available on the default interface, so such extension does not affect any other interfaces).

To make the code more robust, we should also add *InterfaceImplemented* "class method" (a method to be called on the predeclared instance) to all *DataTableXXX* classes. This method should take a string identifying a particular interface and return True/False based on whether such an interface is implemented. The calling code, such as abstract factory's specific *CreateInstance* can raise an error if the required interface is unavailable.

#### Extensibility considerations

Consider [Fig. 1](#FigDataTableModel). The goal is to extend functionality provided by the *DataTableADODB* class, while preserving backward compatibility. First, *DataTableADODB* can be extended via new methods or via etxtensions to existing methods without breaking compatibility. This new functionality needs to be exposed on the interface. Changing *IDataTableStorage*, however, will break compatibility, so the new functionality must be exposed via a separate new interface *IDataTableStorageV2*, which should also expose all functionality provided by *IDataTableStorage*. Now, each *DataTableXXX* class may implement  *IDataTableStorageV2* in addition to *IDataTableStorage* (if, e.g., *DataTableCSV* class does not provide functionality exposed via *IDataTableStorageV2*, it does not have to implement it). 

*IDataTableStorage* is related to two other classes, *DataTableFactory* and *DataTableManager* ([Fig. 1](#FigDataTableModel)). Based on the discussion above, several approaches are available for extending *DataTableFactory*. For example, a second abstract factory can be defined:

```vba
'''' _____ DataTableFactory.cls _____ ''''

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
