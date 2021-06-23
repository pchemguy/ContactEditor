---
layout: default
title: Factory and multiple interfaces
nav_order: 6
parent: Class design conventions
permalink: /class-design/factory-multiple-interfaces
---

### Factory return type with multiple foreign interfaces - code extensibility

Consider again the [UserWSheet.cls][UserWSheetI.cls] class. Suppose this class is a part of some "UserLogin" library, and some applications already use it. Now, we wish to extend the functionality of the UserWSheet, while preserving the backward compatibility of the library. Perhaps, we now have multiple users saved in a table and want to load a specific user. Assuming that our UserModel class is still satisfactory and does not need to be updated, we need to add new methods for loading/saving data to the UserWSheet, LoadDataRow/SaveDataRow:

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

Changing signatures of existing methods or interfaces would break backward compatibility, so these changes are prohibited. The new functionality still needs to be exposed via an interface, and since we cannot change the existing one, we create a new one:

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

The new interface [IUserStorageV2.cls](#IUserStorageV2.cls) also incorporates the old interface functionality that will be used on the new interface. Let us extend the UserWSheet class:

<a name="UserWSheetEx2.cls"></a>
<p align="right"><b>UserWSheet.cls</b></p>

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

This definition is backward compatible with old software that uses the IUserStorage interface. Now, what should we do about the factory if we want to generate instances of both new and old interfaces? There are two possible approaches here.  The simplest scenario is if the current factory is compatible with the new IUserStorageV2 instances except for the type. It appears that different foreign interfaces of a class can be switched directly the same way via assignment. For example, if an IUserStorageV2 variable is assigned a reference from an IUserStorage, the interface is switched as before, though the RubberDuck VBA extension flags such an assignment as illegal.

A better approach is to declare the factory's return type as the default interface. Interface switching will occur in the calling code when the factory's returned reference is assigned to a local variable (see *SomeStorage* assignment in [DemoInterfaceSwitching.bas][]). Similarly to this demo, the new code would declare the *StorageEx* variable as IUserStorageV2, switching the default interface to IUserStorageV2. Further, the factory's argument list can be extended with optional arguments, and new private fields can be added to the UserWSheet class and initialized without affecting the old code.

If this approach does not provide sufficient flexibility, a new separate factory CreateV2 returning the new or default interface can be defined.

These two approaches can, in principle, accommodate any number of foreign interfaces and code upgrades while preserving full backward compatibility.


[Polymorphism]: https://en.wikipedia.org/wiki/Polymorphism_(computer_science)
[UserWSheetI.cls]: https://pchemguy.github.io/ContactEditor/class-design/intro-to-interfaces#UserWSheetI.cls
[DemoInterfaceSwitching.bas]: https://pchemguy.github.io/ContactEditor/class-design/interface-switching#DemoInterfaceSwitching.bas
