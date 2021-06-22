
In VBA, *interface* is a class feature that is comprised of all public declarations of that class, including method signatures (name, parameters, and return type), property signatures, and fields. Given an instance of a class, one can use any of its public methods and property getters/setters. The code within those methods and getters/setters comprises an implementation of the interface. In other words, each class implicitly implements its interface. However, classes are also allowed to implement interfaces of other classes, and this feature is the basis for [polymorphic][Polymorphism] programming in VBA.

Consider an example constisting of three files, two class modules and one regular module. In this example, user's name and login saved in the file needs to be loaded, presented to the user for editing (outside the scope of this example), and saved back to the file. The "UserModel.cls" class is responsible for holding the user data, the "UserWSheet.cls" class loads/saves the data to/from an Excel Worsheet, and the "EditUserData.bas" module executes the workflow.

___
_____ UserModel.cls _____

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
_____ UserWSheet.cls _____

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

<details><summary>Factory</summary>

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
  Model = this.Model
End Property

```

</details>

___
_____ EditUserData.bas _____

```vb

Public Sub Main()
  Dim Model As UserModel
  Set Model = New UserModel
  Dim Source As String
  Source = "UserData"
  Dim Storage As UserWSheet
  Set Storage = UserWSheet.Create(Model, Source)
  
  Storage.LoadData
  
  '''' Use a *modal UserForm* to show the data to the user
  '''' The save changes
  
  Storage.SaveData
End Sub

```

___


[Polymorphism]: https://en.wikipedia.org/wiki/Polymorphism_(computer_science)
