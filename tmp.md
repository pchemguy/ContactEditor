
### Overview

In VBA, *interface* is a class feature comprised of all public declarations of that class, including method signatures (name, parameters, and return type), property signatures, and fields. Given an instance of a class, one can use all its public methods and property getters/setters. The code within those methods and getters/setters comprises an implementation of the interface. In other words, each class implicitly implements its interface. Additionally, a class may implement another class's interface, and this feature is the basis for [polymorphic][Polymorphism] programming in VBA.

Consider an example consisting of three files shown below, one regular and two class modules. In this example, the user's name and login information saved in the file needs to be loaded, presented to the user for editing (outside the scope of this example), and saved back to the file. The [UserModel.cls](#UserModel.cls) class is responsible for holding the user data, the [UserWSheet.cls](#UserWSheet.cls) class loads/saves the data to/from an Excel Worsheet, and the [EditUserData.bas](#EditUserData.bas) module executes the workflow.

<a name="UserModel.cls"></a>
<p align="right"><b>UserModel.cls</b></p>

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
<p align="right"><b>UserWSheet.cls</b></p>

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

Public Function Create(ByVal Model As UserModel, _
                       ByVal WSheetName As String) As UserWSheet
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
<p align="right"><b>EditUserData.bas</b></p>

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

The [UserWSheet.cls](#UserWSheet.cls) class has several sections, including field and property definitions, a factory/constructor pair, and utility methods (LoadData/SaveDate). If we consider the role of this class from the calling code ([EditUserData.bas](#EditUserData.bas)), however, the utility methods are all that matters. The calling code does not care what the nature of the data source is or what code is inside the utility methods. The only important matter is that the Model property of the Storage object must contain the data after calling the LoadData method, and the data must be saved from the Model attribute to persistent storage after calling the SaveData. Moreover, the source of the data could also be, for example, a plain-text file or an SQL database. Further, the actual source type may be defined by the user or program settings. 

Ideally, we would want to have one class module "UserXXX.cls" (similar to UserWSheet.cls) for each source type, set up the *Storage* variable based on the selected source, and be able to use the following code without any modifications regardless of the selected source type. We only require that the *Storage* object has the *Model* attribute and two utility methods LoadData/SaveDate characterized by the specified outcome. This requirement can be formalized by defining an additional class, say, [IUserStorage.cls](#IUserStorage.cls), which includes "declarations" for all our required members:

<a name="IUserStorage.cls"></a>
<p align="right"><b>IUserStorage.cls</b></p>

```vb

Public Property Get Model() As UserModel
End Property

Public Property Set Model(ByVal Instance As UserModel)
End Property

Public Sub LoadData()
End Sub

Public Sub SaveData()
End Sub

```

VBA treats this module like any other class module. Alone, it is not very useful as it does not have any functionality; in a sense, it declares its interface like any other class module, but it does not provide any useful implementation. However, other classes, such as [UserWSheet.cls](#UserWSheetI.cls), may implement it:

<a name="UserWSheetI.cls"></a>
<p align="right"><b>UserWSheet.cls</b></p>

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

Public Function Create(ByVal Model As UserModel, _
                       ByVal WSheetName As String) As UserWSheet
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

The new [UserWSheet.cls](#UserWSheetI.cls) implements two interfaces (two sets of methods/attributes): the class's default interface, *UserWSheet*, includes its public methods/attributes (in this case, the factory and constructor methods) and the *IUserStorage*. Similarly, classes may implement more than one additional interface.

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

### Factory return type with single foreign interface

Suppose we have a class implementing another class's interface and employing the "factory" pattern, such as the new [UserWSheet.cls](#UserWSheetI.cls). In such a case, it is customarily to define the factory return value as the non-default interface (though it can still return the default one):

```vb

Public Function Create(ByVal Model As UserModel, ByVal WSheetName As String) As IUserStorage
End Function

```

vs.

```vb

Public Function Create(ByVal Model As UserModel, ByVal WSheetName As String) As UserWSheet
End Function

```

For a class implementing one additional interface, there is a subtle difference between these two options. Usually, the calling code assigns the value returned by the factory to a local variable. This variable's type, in turn, determines the interface it will present (see *SomeStorage* assignment in [DemoInterfaceSwitching.bas](#DemoInterfaceSwitching.bas)). Here, the two declarations would yield identical behavior.

There is one use case, however, for which the two options are not the same. Consider an abstract factory class [UserStorageFactory.cls](#UserStorageFactory.cls) coded against the [IUserStorageFactory.cls](#IUserStorageFactory.cls) interface:

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
  
Public Sub Init(ByVal UserStorageType As String)
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

Note how the *Storage* variable is assigned in [DemoAbstractFactory.bas](#DemoAbstractFactory.bas). What happens there is that the first call to Create factory yields an instance of an abstract factory, which we only need to use once and do not need to save. Because Create factory return type is declared as IUserStorageFactory, the returned reference has CreateInstance immediately accessible on it (compare to *SomeStorage* in [DemoInterfaceSwitching.bas](#DemoInterfaceSwitching.bas)). That is why we can chain calls to Create and CreateInstance here.

At the same time, if we switch the return type of Create in [UserStorageFactory.cls](#UserStorageFactory.cls) from IUserStorageFactory to UserStorageFactory, this will be the case of *Storage* in [DemoInterfaceSwitching.bas](#DemoInterfaceSwitching.bas). *CreateInstance* will not be avaialble on the reference returned by Create, and chaining will no longer be possible. Instead, we would need to switch the interfaces explicitly as with *GenericStorage* in [DemoInterfaceSwitching.bas](#DemoInterfaceSwitching.bas).

### Factory return type with multiple foreign interfaces

Consider again the [UserWSheet.cls](#UserWSheetI.cls) class. Assume it is a part of, say, "UserLogin" library, and some applications already use it. Now we wish to extend the functionality of the UserWSheet, while preserving backward compatibility of the library. Perhaps, we now have multiple users saved in a table and we want to load a specific user. Assuming that our UserModel class is still satisfactory and does not need to be updated, we need to add new methods for loading/saving data to the UserWSheet, LoadDataRow/SaveDataRow:

```vb

Public Sub LoadDataRow(ByVal RecordIndex as Long)
  With this.Model
    .FirstName = ThisWorkbook.Worksheets(this.WSheet).Range("A" & CStr(RecordIndex)).Value
    .LastName  = ThisWorkbook.Worksheets(this.WSheet).Range("B" & CStr(RecordIndex)).Value
    .Login     = ThisWorkbook.Worksheets(this.WSheet).Range("C" & CStr(RecordIndex)).Value
  End With
End Sub

Public Sub SaveDataRow(ByVal RecordIndex as Long)
  With this.Model
    ThisWorkbook.Worksheets(this.WSheet).Range("A" & CStr(RecordIndex)).Value = .FirstName
    ThisWorkbook.Worksheets(this.WSheet).Range("B" & CStr(RecordIndex)).Value = .LastName
    ThisWorkbook.Worksheets(this.WSheet).Range("C" & CStr(RecordIndex)).Value = .Login
  End With
End Sub

```

Changing signatures of existing methods or interfaces would break backward compatibility, so these changes are prohibited. The new functionality still needs to be exposed via an interface. Since we cannot changes the existing one, we create a new one:


<a name="IUserStorageV2.cls"></a>
<p align="right"><b>IUserStorageV2.cls</b></p>

```vb

Public Property Get Model() As UserModel
End Property

Public Property Set Model(ByVal Instance As UserModel)
End Property

Public Sub LoadData()
End Sub

Public Sub SaveData()
End Sub

Public Sub LoadDataRow(ByVal RecordIndex as Long)
End Sub

Public Sub SaveDataRow(ByVal RecordIndex as Long)
End Sub

```

The new interface [IUserStorageV2.cls](#IUserStorageV2.cls) also incorporates all the functionality of the old interface that will be used on the new interface. Let us extend the UserWSheet class:

<a name="UserWSheet.cls"></a>
<p align="right"><b>UserWSheetEx2.cls</b></p>

```vb

'''' N.B.: This class must be predeclared

Implements IUserStorage
Implements IUserStorageV2

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

<details><summary>IUserStorageV2 interface implementation</summary>

```vb

Private Sub IUserStorageV2_LoadDataRow(ByVal RecordIndex as Long)
  With this.Model
    .FirstName = ThisWorkbook.Worksheets(this.WSheet).Range("A" & CStr(RecordIndex)).Value
    .LastName  = ThisWorkbook.Worksheets(this.WSheet).Range("B" & CStr(RecordIndex)).Value
    .Login     = ThisWorkbook.Worksheets(this.WSheet).Range("C" & CStr(RecordIndex)).Value
  End With
End Sub

Private Sub IUserStorageV2_SaveDataRow(ByVal RecordIndex as Long)
  With this.Model
    ThisWorkbook.Worksheets(this.WSheet).Range("A" & CStr(RecordIndex)).Value = .FirstName
    ThisWorkbook.Worksheets(this.WSheet).Range("B" & CStr(RecordIndex)).Value = .LastName
    ThisWorkbook.Worksheets(this.WSheet).Range("C" & CStr(RecordIndex)).Value = .Login
  End With
End Sub

Private Sub IUserStorageV2_LoadData()
  With this.Model
    .FirstName = ThisWorkbook.Worksheets(this.WSheet).Range("A1").Value
    .LastName  = ThisWorkbook.Worksheets(this.WSheet).Range("B1").Value
    .Login     = ThisWorkbook.Worksheets(this.WSheet).Range("C1").Value
  End With
End Sub

Private Sub IUserStorageV2_SaveData()
  With this.Model
    ThisWorkbook.Worksheets(this.WSheet).Range("A1").Value = .FirstName
    ThisWorkbook.Worksheets(this.WSheet).Range("B1").Value = .LastName
    ThisWorkbook.Worksheets(this.WSheet).Range("C1").Value = .Login
  End With
End Sub

Private Property Get IUserStorageV2_Model() As UserModel
  Set IUserStorageV2_Model = this.Model
End Property

Private Property Set IUserStorageV2_Model(ByVal Instance As UserModel)
  Set this.Model = Instance
End Property

```

</details>

<details><summary>Factory/Constructor</summary>

```vb

Public Function Create(ByVal Model As UserModel, _
                       ByVal WSheetName As String) As UserWSheet
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

Note that this extended class definition is backward compatible with old software that uses the IUserStorage interface. Now, what should we do about the factory, if we want to be able to generate instances of both new and old interfaces. There are two possible approaches here. The simplest scenario is if the old factory/constructor are compatible with the new extended IUserStorageV2 instances except for the type. Apparently different foreign interfaces of a class can be switched directly the sam way via assignment. That is if an IUserStorageV2 variable is assigned a reference from an IUserStorage, the interface is switched as before. However, RubberDuck VBA extension considers such an assignment as illegal. While this might be a limitation of the RDVBA, a safer approach is to declare the factory as returning the default interface. In this case, the calling code declares a local variable as either IUserStorage (old code) or IUserStorageV2 (new code) and in both cases the default interfaces returned by the factory is automatically switched to the desired foreign interface. Further, to accomodate the extended functionality, factory's argument list can be extended with optional arguments and new private fields can be added to the UserWSheet class and initialized without affecting the old code.

If this approach cannot be accommodated, a new separate factory CreateV2 returning the new interface will need to be defined.

These two approaches can, in principle, accommodate any number of foreign interfaces, while ensuring that the class is backward compatible with all previous software.

[Polymorphism]: https://en.wikipedia.org/wiki/Polymorphism_(computer_science)
