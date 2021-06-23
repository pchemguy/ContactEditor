### Overview

In VBA, *interface* is a class feature that is comprised of all public declarations of that class, including method signatures (name, parameters, and return type), property signatures, and fields. Given an instance of a class, one can use any of its public methods and property getters/setters. The code within those methods and getters/setters comprises an implementation of the interface. In other words, each class implicitly implements its interface. However, classes are also allowed to implement interfaces of other classes, and this feature is the basis for [polymorphic][Polymorphism] programming in VBA.

Consider an example constisting of three files shown below, two class modules and one regular module. In this example, user's name and login saved in the file needs to be loaded, presented to the user for editing (outside the scope of this example), and saved back to the file. The [UserModel.cls](#UserModel.cls) class is responsible for holding the user data, the [UserWSheet.cls](#UserWSheet.cls) class loads/saves the data to/from an Excel Worsheet, and the [EditUserData.bas](#EditUserData.bas) module executes the workflow.

<a name="UserModel.cls"></a>
<p align="center"><b>UserModel.cls</b></p>

```vb

Private Type TUserModel
	FirstName As String
	LastName As String
	Login As String
End Type
Private this As TUserModel

```

<details><summary>Getters and Setters</summary>

```vb

Public Property Let FirstName(ByVal Value As String)
  FirstName = Value
End Property

Public Property Get FirstName() As String
  FirstName = this.FirstName
End Property

Public Property Let LastName(ByVal Value As String)
  LastName = Value
End Property

Public Property Get LastName() As String
  LastName = this.LastName
End Property

Public Property Let Login(ByVal Value As String)
  Login = Value
End Property

Public Property Get Login() As String
  Login = this.Login
End Property

```

</details>

___
<a name="UserWSheet.cls"></a>
<p align="center"><b>UserWSheet.cls</b></p>

```vb

'''' N.B.: This class must be predeclared

Private Type TUserWSheet
  Model As UserModel
  WSheetName As String
End Type
Private this As TUserWSheet

Public Sub LoadData()
  With this.Model
    .FirstName = ThisWorkbook.Worksheets(this.WSheet).Range("A1").Value
    .LastName  = ThisWorkbook.Worksheets(this.WSheet).Range("B1").Value
    .Login     = ThisWorkbook.Worksheets(this.WSheet).Range("C1").Value
  End With
End Sub

Public Sub SaveData()
  With this.Model
    ThisWorkbook.Worksheets(this.WSheet).Range("A1").Value = .FirstName
    ThisWorkbook.Worksheets(this.WSheet).Range("B1").Value = .LastName
    ThisWorkbook.Worksheets(this.WSheet).Range("C1").Value = .Login
  End With
End Sub

```

<details><summary>Factory/Constructor</summary>

```vb

Public Function Create(ByVal Model As UserModel, ByVal WSheetName As String) As UserWSheet
  Dim Instance As UserWSheet
  Set Instance = New UserWSheet
  Instance.Init Model, WSheetName
  Set Create = Instance
End Function
  
Public Sub Init(ByVal Model As UserModel, ByVal WSheetName As String)
  Set this.Model = Model
  this.WSheetName = WSheetName
End Sub

```
  
</details>

<details><summary>Getters and Setters</summary>

```vb

Public Property Get Model() As UserModel
  Set Model = this.Model
End Property

Public Property Set Model(ByVal Instance As UserModel)
  Set this.Model = Instance
End Property

```

</details>

___
<a name="EditUserData.bas"></a>
<p align="center"><b>EditUserData.bas</b></p>

```vb

Public Sub Main()
  Dim Source As String
  Source = "UserData"
  Dim Model As UserModel
  Set Model = New UserModel
  Dim Storage As UserWSheet
  Set Storage = UserWSheet.Create(Model, Source)
  
  Storage.LoadData
  
  '''' Some functionality goes here, for example,
  '''' Use a modal UserForm to show the data to the user
  '''' Then save the changes
  
  Storage.SaveData
End Sub

```

The [UserWSheet.cls](#UserWSheet.cls) class has several sections, including field and property definitions, a factory/constructor pair, and utility methods (LoadData/SaveDate). If we consider the role of this class from the calling code ([EditUserData.bas](#EditUserData.bas)), however, the utility methods is all that matters. Moreover, the calling code does not really care what is the nature of the data source or what code is inside the utility methods; the only important thing is that the Model property of the Storage object must contain the data after the .LoadData method is called, and the data must be saved from the Model attribute to persistent storage after the .SaveData is called. In fact, the source of the data could also be, for example, a plain-text file or an SQL database. Further, the actual source type may be defined by the user or saved program settings. 

Ideally, we would want to have one class module "UserXXX.cls" (similar to UserWSheet.cls) for each source type, set up the *Storage* variable based on the selected source, and be able to use the following code without any modifications regardless of the selected source type. We only require that the *Storage* object must have the *Model* attribute and two utility methods LoadData/SaveDate characterized by the specified outcome. This requirement can be formalized by defining an additional class, say, [IUserStorage.cls](#IUserStorage.cls), which includes "declarations" for all our required members:

<a name="IUserStorage.cls"></a>
<p align="center"><b>IUserStorage.cls</b></p>

```vb

Public Property Get Model() As UserModel
End Property

Public Property Set Model(ByVal Instance As UserModel)
End Property

Public Sub LoadData()
End

Public Sub SaveData()
End

```

This is module is treated by VBA like any other class module. Alone, it is not very useful as it does not have any functionality; in a sense, it declares its interface like any other class module, but it does not provide any useful implementation. However, other classes, such as [UserWSheet.cls](#UserWSheet.cls), can implement it:

<a name="UserWSheetI.cls"></a>
<p align="center"><b>UserWSheet.cls</b></p>

```vb

'''' N.B.: This class must be predeclared

Implements IUserStorage

Private Type TUserWSheet
  Model As UserModel
  WSheetName As String
End Type
Private this As TUserWSheet

```

<details><summary>IUserStorage interface implementation</summary>

```vb

Private Sub IUserStorage_LoadData()
  With this.Model
    .FirstName = ThisWorkbook.Worksheets(this.WSheet).Range("A1").Value
    .LastName  = ThisWorkbook.Worksheets(this.WSheet).Range("B1").Value
    .Login     = ThisWorkbook.Worksheets(this.WSheet).Range("C1").Value
  End With
End Sub

Private Sub IUserStorage_SaveData()
  With this.Model
    ThisWorkbook.Worksheets(this.WSheet).Range("A1").Value = .FirstName
    ThisWorkbook.Worksheets(this.WSheet).Range("B1").Value = .LastName
    ThisWorkbook.Worksheets(this.WSheet).Range("C1").Value = .Login
  End With
End Sub

Private Property Get IUserStorage_Model() As UserModel
  Set IUserStorage_Model = this.Model
End Property

Private Property Set IUserStorage_Model(ByVal Instance As UserModel)
  Set this.Model = Instance
End Property

```

</details>

<details><summary>Factory/Constructor</summary>

```vb

Public Function Create(ByVal Model As UserModel, ByVal WSheetName As String) As UserWSheet
  Dim Instance As UserWSheet
  Set Instance = New UserWSheet
  Instance.Init Model, WSheetName
  Set Create = Instance
End Function
  
Public Sub Init(ByVal Model As UserModel, ByVal WSheetName As String)
  Set this.Model = Model
  this.WSheetName = WSheetName
End Sub

```

</details>

The new [UserWSheet.cls](#UserWSheetI.cls) implements two interfaces, that is two sets of methods/attributes: the class's own default interface, *UserWSheet*, comprised of its public methods/attributes (in this case this set includes the factory and constructor methods) and the *IUserStorage*. Similarly, classes may implement more then one additional interface.

### Interface switching

The next important concept is interface switching. When a class implements other interfaces besides its own, an instance of such a class can present any of the implemented interfaces, but only one interface at any given time. Switching occurs during the assignment operation, and a particular interface is selected based on the declared type of the variable being assigned, for example:

<a name="DemoInterfaceSwitching.bas"></a>
<p align="center"><b>DemoInterfaceSwitching.bas</b></p>

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

Importantly, in [DemoInterfaceSwitching.bas](#DemoInterfaceSwitching.bas) we created one object instance of the UserWSheet class, but two references pointing to this object, *Storage* and *GenericStorage*, presenting two different interfaces implemented by the UserWSheet class.

### Factory return type

If a class, implementing an additional interface, uses the "factory" pattern, such as the new [UserWSheet.cls](#UserWSheetI.cls) class, it is customarily to define the factory return value as non-default (though it can still return the default) interface:

```vb

Public Function Create(ByVal Model As UserModel, ByVal WSheetName As String) As IUserStorage
End Function

```

vs.

```vb

Public Function Create(ByVal Model As UserModel, ByVal WSheetName As String) As UserWSheet
End Function

```

When only one additional interface is implemented, there is a subtle difference between these two options. If the calling code assigns the value returned by the factory to a variable, then the selected interface is determined by the declaration of the assigned variable (*SomeStorage* assignment in [DemoInterfaceSwitching.bas](#DemoInterfaceSwitching.bas) would yield identical result in both cases).

There is one use case, however, for which the two options are not the same. Consider an abstract factory class [UserStorageFactory.cls](#UserStorageFactory.cls) coded against the [IUserStorageFactory.cls](#IUserStorageFactory.cls) interface:

<a name="IUserStorageFactory.cls"></a>
<p align="center"><b>IUserStorageFactory.cls</b></p>

```vb

'''' This method takes the name of the source, such as Worksheet name or file name
'''' and passes it to the appropriate IUserStorage class factory, such as UserWSheet
'''' or UserCSV (the target class is selected by UserStorageFactory.Create).
Public Function CreateInstance(ByVal Model As UserModel, ByVal SourceName As String) As IUserStorage
End Function

```

___

<a name="UserStorageFactory.cls"></a>
<p align="center"><b>UserStorageFactory.cls</b></p>

```vb

'''' N.B.: This class must be predeclared

Private Type TUserStorageFactory
	UserStorageType As String
End Type
Private this As TUserStorageFactory

'''' This method is called on the predeclared instance.
'''' It takes the type name of the IUserStorage source, such as "Worksheet" or "CSV"
'''' and returns an instance of the UserStorageFactory class that should be used as
'''' a factory for IUserStorage instances.
Public Function Create(ByVal UserStorageType As String) As IUserStorageFactory
  Dim Instance As UserStorageFactory
  Set Instance = New UserStorageFactory
  Instance.Init UserStorageType
  Set Create = Instance
End Function
  
Public Sub Init(ByVal UserStorageType As String)
  Set this.UserStorageType = UserStorageType
End Sub

Private Function IUserStorageFactory_CreateInstance(ByVal Model As UserModel, _
                                                    ByVal SourceName As String) As IUserStorage
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
<p align="center"><b>DemoAbstractFactory.bas</b></p>

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

Note how the *Storage* variable is assigned in [DemoAbstractFactory.bas](#DemoAbstractFactory.bas). What happens there is that the first call to Create factory yields an instance of an abstract factory, which we only need to use once and do not need to save. Because Create factory return type is declared as IUserStorageFactory, the returned reference has CreateInstance immediately accessible on it (compare to *SomeStorage* in [DemoInterfaceSwitching.bas](#DemoInterfaceSwitching.bas)). This is why we can chain calls to Create and CreateInstance here.

At the same time, if we switch the return type of Create in [UserStorageFactory.cls](#UserStorageFactory.cls) from IUserStorageFactory to UserStorageFactory, this will be the case of *Storage* in [DemoInterfaceSwitching.bas](#DemoInterfaceSwitching.bas). *CreateInstance* will not be avaialble on the reference returned by Create and chaining will no longer be possible. Instead, we would need to switch the interfaces explicitly as with *GenericStorage* in [DemoInterfaceSwitching.bas](#DemoInterfaceSwitching.bas).



[Polymorphism]: https://en.wikipedia.org/wiki/Polymorphism_(computer_science)
