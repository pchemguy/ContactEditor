---
layout: default
title: Factory return type
nav_order: 6
parent: Design patterns
permalink: /class-design/factory-return-type
---

### Factory return type with single foreign interface

#### Concrete factory

Suppose we have a class implementing another class's interface and employing the "factory" pattern, such as the new [UserWSheet.cls][UserWSheetI.cls]. In such a case, it is customarily to define the factory return value as the non-default interface (though it can still return the default one):

```vb
Public Function Create(ByVal Model As UserModel, ByVal WSheetName As String) As IUserStorage
End Function
```

vs.

```vb
Public Function Create(ByVal Model As UserModel, ByVal WSheetName As String) As UserWSheet
End Function
```

For a class implementing one foreign interface, there is a subtle difference between these two options. Usually, the calling code assigns the value returned by the factory to a local variable. This variable's type, in turn, determines the interface it will present (see *SomeStorage* assignment in [DemoInterfaceSwitching.bas][]). Here, the two declarations would yield identical behavior.

#### Concrete factory

There is one use case, however, for which the two options affect the code. Consider an abstract factory class [UserStorageFactory.cls](#UserStorageFactory.cls) coded against the [IUserStorageFactory.cls](#IUserStorageFactory.cls) interface:

<a name="IUserStorageFactory.cls"></a>
<p align="right"><b>IUserStorageFactory.cls</b></p>

```vb
'''' This method takes the source name (e.g., a Worksheet name or a file name)
'''' and passes it to the appropriate IUserStorage class factory, such as
'''' UserWSheet or UserCSV. The previous call to UserStorageFactory.Create
'''' selects the desired class.
Public Function CreateInstance(ByVal Model As UserModel, ByVal SourceName As String) As IUserStorage
End Function
```

___

<a name="UserStorageFactory.cls"></a>
<p align="right"><b>UserStorageFactory.cls</b></p>

```vb
'''' N.B.: This class must be predeclared

Private Type TUserStorageFactory
	UserStorageType As String
End Type
Private this As TUserStorageFactory

'''' This method is called on the predeclared instance.
'''' It takes the type name of the IUserStorage source (e.g., "Worksheet" or
'''' "CSV") and returns an instance of the UserStorageFactory class to be
'''' used as a factory for IUserStorage instances.
Public Function Create(ByVal UserStorageType As String) As IUserStorageFactory
  Dim Instance As UserStorageFactory
  Set Instance = New UserStorageFactory
  Instance.Init UserStorageType
  Set Create = Instance
End Function
  
Friend Sub Init(ByVal UserStorageType As String)
  Set this.UserStorageType = UserStorageType
End Sub

Private Function IUserStorageFactory_CreateInstance( _
        ByVal Model As UserModel, ByVal SourceName As String) As IUserStorage
  Select Case this.UserStorageType
    Case "Worksheet"
      Set IUserStorageFactory_CreateInstance = UserWSheet.Create(Model, SourceName)
    Case "CSV"
      Set IUserStorageFactory_CreateInstance = UserCSV.Create(Model, SourceName)
    Case "SQLite"
      Set IUserStorageFactory_CreateInstance = UserSQLite.Create(Model, SourceName)
    Case Else
      Err.Raise ErrNo.NotImplementedErr, "UserStorageFactory", "Unsupported source"
  End Select
End Function
```

This abstract factory can be used like this:

<a name="DemoAbstractFactory.bas"></a>
<p align="right"><b>DemoAbstractFactory.bas</b></p>

```vb
Public Sub Main()
  Dim SourceType As String
  SourceType = "Worksheet"
  Dim SourceName As String
  SourceName = "UserData"
  Dim Model As UserModel
  Set Model = New UserModel
  
  Dim Storage As IUserStorage
  Set Storage = UserStorageFactory.Create(SourceType).CreateInstance(Model, SourceName)

  '''' Use Storage variable here
End Sub
```

Note how the *Storage* variable is assigned in [DemoAbstractFactory.bas](#DemoAbstractFactory.bas). The first call to the Create factory on the right-hand side yields an instance of an abstract factory, which we only need to use once and do not need to save. Because Create factory return type is declared as IUserStorageFactory, the returned reference has CreateInstance immediately accessible on it (compare to *SomeStorage* in [DemoInterfaceSwitching.bas][]). That is why we can chain calls to Create and CreateInstance here.

At the same time, if we switch the return type of Create in [UserStorageFactory.cls](#UserStorageFactory.cls) from IUserStorageFactory to UserStorageFactory, this will be the case of *Storage* in [DemoInterfaceSwitching.bas][]. *CreateInstance* will not be avaialble on the reference returned by Create, and chaining will no longer be possible. Instead, we would need to switch the interfaces explicitly as with *GenericStorage* in [DemoInterfaceSwitching.bas][].



[UserWSheetI.cls]: https://pchemguy.github.io/ContactEditor/class-design/intro-to-interfaces#UserWSheetI.cls
[DemoInterfaceSwitching.bas]: https://pchemguy.github.io/ContactEditor/class-design/interface-switching#DemoInterfaceSwitching.bas
