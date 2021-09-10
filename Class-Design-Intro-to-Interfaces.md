---
layout: default
title: Intro to interfaces in VBA
nav_order: 4
parent: Design patterns
permalink: /class-design/intro-to-interfaces
---

### Introduction to interfaces in VBA

Interfaces are a somewhat advanced topic in VBA, and this tutorial is not for beginners. A good starting point might be the [Flyable and Swimable][] tutorial, which uses a simple model example and builds an easy-to-follow working code, or another simple example from [Better Solutions][]. [OOP VBA Examples][] tutorial is also worth mentioning; it discusses a few more advanced topics related to the use of interfaces. The main reason for preparing this tutorial was to form the foundation for discussing more advanced aspects in the following sections. The more elaborate model example used here focuses on a practical application of interfaces relevant to this project. While I only provide code fragments in the text, this tutorial is a part of a complete working demo application.

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

The new [UserWSheet.cls](#UserWSheetI.cls) implements two interfaces (two sets of methods/attributes): the class's default interface, *UserWSheet*, includes its public methods/attributes (in this case, the factory and constructor methods) and a foreign interface *IUserStorage*. Similarly, classes may implement more than one foreign interface.


[Flyable and Swimable]: https://riptutorial.com/vba/topic/8784/interfaces
[Better Solutions]: https://bettersolutions.com/vba/class-modules/implements.htm
[OOP VBA Examples]: https://riptutorial.com/vba/topic/5357/object-oriented-vba
