---
layout: default
title: Interface switching
nav_order: 4
parent: Class design patterns
permalink: /class-design/interface-switching
---

### Interface switching

An object instance of a class implementing multiple interfaces can present any of the class's interfaces, but only one interface at any given time. Interface selection occurs during the assignment operation based on the declared type of the assigned variable. For this reason, it is not possible to switch the interface of a particular variable directly. Instead, a new variable must be declared specifying the desired interface as its type. Then, the new variable is assigned an object reference from the existing one with a different interface, for example:

<a name="DemoInterfaceSwitching.bas"></a>
<p align="right"><b>DemoInterfaceSwitching.bas</b></p>

```vb

Public Sub Main()
  Dim Source As String
  Source = "UserData"
  Dim Model As UserModel
  Set Model = New UserModel
  
  '''' "Storage" variable is declared as UserWSheet, so it presents the default UserWSheet
  '''' interface, that is we can, for example, use its factory Create, but not the utility
  '''' methods, which are part of the IUserStorage interface, which is now hidden.
  Dim Storage As UserWSheet
  Set Storage = UserWSheet.Create(Model, Source)
  Storage.LoadData '''' <-- Compile time error
  
  '''' "GenericStorage" variable is declared as IUserStorage, so an attempt to call the
  '''' factory would result in an error, but LoadData/SaveData methods are now available.
  Dim GenericStorage as IUserStorage
  Set GenericStorage = Storage
  GenericStorage.LoadData '''' <-- Should work fine
  
  '''' Direct assignment. Again, an attempt to call the factory would result in an error,
  '''' but LoadData/SaveData methods are available.
  Dim SomeStorage as IUserStorage
  Set SomeStorage = UserWSheet.Create(Model, Source)  
  SomeStorage.LoadData '''' <-- Should work fine
End Sub

```

In [DemoInterfaceSwitching.bas](#DemoInterfaceSwitching.bas), we created one object instance of the UserWSheet class with two references (*Storage* and *GenericStorage*) pointing to this object and presenting two different interfaces implemented by the UserWSheet class. *SomeStorage* points to a second UserWSheet class instance presenting the IUserStorage interface.
