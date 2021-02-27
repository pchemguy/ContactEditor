There is no worksheet. – Rubberduck News

There is no worksheet.
======================

Posted on [December 8, 2017March 28, 2018](https://rubberduckvba.wordpress.com/2017/12/08/there-is-no-worksheet/) by [Rubberduck VBA](https://rubberduckvba.wordpress.com/author/rubberduckvba/)

Your VBA project is embedded in an Excel workbook. It references the VBA standard library; it references the library that exposes the host application’s (i.e. in this case, Excel’s) object model; it includes global-scope objects of types that are declared in these libraries – like `Sheet1` (an `Excel.Worksheet` instance) and `ThisWorkbook` (an `Excel.Workbook` instance). These free, global-scope objects are right here to take and run with.

You’re free to use them, *wisely*.

True, they’re global – they *can* be accessed from anywhere in the code.
They *can*… and that doesn’t mean they *should*.

And if you’re willing to do whatever it takes to *abstract away* the host application’s object model in your “business logic”, then you can isolate your logic from the worksheet boilerplate and write pretty much the same code you’d be writing in, say, VB.NET… or any other object-oriented language for that matter.

* * * * *

Abstracting Worksheets
======================

There is no worksheet. There is *data*. Data going in, data going out: data is all it is. When the data is coming from a database, many programmers immediately say “*I know! Let’s implement a repository pattern!*“, or otherwise come up with various ways to *abstract away* the data-handling boilerplate. If you think of worksheets as *data*, then it’s not any different, really.

So we shall treat worksheets as such: data. What do we need to do with this data?

Some tried to make worksheets `Implements` interfaces, and [ran into issues](https://stackoverflow.com/q/43358109/1188513) ([here too](https://stackoverflow.com/q/18500674/1188513), and [oh another](https://stackoverflow.com/q/18699566/1188513)). I completely agree with [this post](http://productivebytes.blogspot.ca/2012/10/using-implements-behind-excel-worksheet.html), which basically boils down to **don’t**.

### Whatever you do, don’t make worksheets implement an interface.

*Wrap* them instead. Make a *proxy *class implement the interface you need, and then make sure everything that needs anything on a worksheet, accesses it through an interface, say `IWorkbookData.FooSheet`, where `FooSheet` is a property that returns a `FooSheetProxy` instance, exposed as an `IFooSheet`.

![Diagram](https://rubberduckvba.files.wordpress.com/2017/12/capture.png)

The only thing that ever needs to use a `FooSheet` directly, is a `FooSheetProxy`.

I don’t know about you, but I don’t like diagrams. So how about some real-world code instead?

Say you have an *order form* workbook, that you distribute to your sales reps and that come back filled with customer order details. Now imagine you need a macro that reads the form contents and imports it into your ERP system.

You could write a macro that reads the customer account number in cell O8, the order date in cell J6, the delivery and cancel dates in cells J8 and J10, and loops through rows 33 to 73 to pull the model codes from column F, and the per-size quantities in columns V through AQ… and it would work.

…Until it doesn’t anymore, because an order came back without a customer account number because it’s a new customer and the data validation wouldn’t let them enter an account that didn’t exist at the time you issued the order form. Or you had to shift all the sized units columns to the right because someone asked if it was possible to enter arbitrary notes at line item level. Or a new category needed to be added and now you have two size scales atop your sized units columns, and you can’t just grab the size codes from row 31 anymore. In other words, it works, until someone else uses (or sees) it and the requirements change.

Sounds familiar?

If you’ve written a *script-like* god-procedure that starts at the top and finishes with a `MsgBox "Completed!"` call at the bottom, (*because that’s all VBA is good for, right?*), then you’re going to modify your code, increase the cyclomatic complexity with new cases and conditions, and rinse and repeat. Been there, done that.

Not anymore.

Name things.
------------

**Abstraction** is key. Your code doesn’t *really* care about what’s in cell O8. Your code needs to know about a *customer account number*. So you name that range `Header_AccountNumber`, proceed to name all the things, and before you know it you’re looking at `Header_OrderDate`, `Header_DeliveryDate` and `Header_CancelDate`, and then `Details_DetailArea` and `Details_SizedUnits` named ranges, you’ve ajusted your code to use them instead of hard-coding cell references, and that’s already a tremendous improvement: now the code isn’t going to break every time something needs to move around.

But you’re still looking at a god-like procedure that does everything, and the only way to test it is to run it: the more complex things are, the less possible it is to cover everything and guarantee that the code behaves as intended in every single corner case. So you download Rubberduck and think “I’m going to write a bunch of unit tests”, and there you are, writing “unit tests” that interact with the real worksheet, firing worksheet events at every change, calculating totals and whatnot: they’re automated tests, but they’re not *unit *tests. You simply *can’t* write unit tests against a god-like macro procedure that knows everything and does everything. Abstraction level needs to go further up.

The [Order Form] worksheet has a code-behind module. Here’s mine:

```vb
'@Folder("OrderForm.OrderInfo")
Option Explicit
```
Yup. That’s absolutely all of it. *Your princess is in another castle*. A “proxy” class implements an interface that the rest of the code uses. The interface exposes everything we’re going to need – something like this:

```vb
'@Folder("OrderForm.OrderInfo")
Option Explicit

Public Property Get AccountNumber() As String
End Property

Public Property Get OrderNumber() As String
End Property

'...

Public Function CreateOrderHeaderEntity() As ISageEntity
End Function

Public Function CreateOrderDetailEntities() As VBA.Collection
End Function

Public Sub Lockdown(Optional ByVal locked As Boolean = True)
End Sub
```

Then there’s a *proxy* class that implements this interface; the `AccountNumber` property implementation might look like this:

```vb
Private Property Get IOrderSheet_AccountNumber() As String
    Dim value As Variant
    value = OrderSheet.Range("Header_Account").value
    If Not IsError(value) Then IOrderSheet_AccountNumber = value
End Property
```

And then the `CreateOrderHeaderEntity` creates and returns an object that your “import into ERP system” macro will consume, using the named ranges defined on `OrderSheet`. Now instead of depending directly on `OrderSheet`, your macro depends on this `OrderSheetProxy` class, and you can even refactor the macro into its own class and make it work against an `IOrderSheet` instead.

What gives? Well, now that you have code that works off an `IOrderSheet` interface, you can write some `OrderSheetTestProxy` implementation that doesn’t even know or care about the actual `OrderSheet` worksheet, and just like that, you can write unit tests that don’t use any worksheet at all, and *still* be able to automatically test the entire set of functionliaties!

* * * * *

Of course this isn’t the full picture, but it gives an idea. A recent *order form* project of mine currently contains 86 class modules, 3 standard modules, 11 user forms, and 25 worksheets (total worksheet code-behind procedures: 0) – not counting anything test-related – and using this pattern (combined with [MVP](https://rubberduckvba.wordpress.com/2017/10/25/userform1-show/)), the code is extremely clear and surprisingly simple; most macros look more or less like this:

```vb
Public Sub AddCustomerAccount()
    Dim proxy As IWorkbookData
    Set proxy = New WorkbookProxy
    If Not proxy.SettingsSheet.EnableAddCustomerAccounts Then
        MsgBox MSG_FeatureDisabled, vbExclamation
        Exit Sub
    End If


    With New AccountsPresenter
        .Present proxy
    End With
End Sub
```

Without abstraction, this project would be a huge unscalable copy-pasta mess, impossible to extend or maintain, let alone debug.

See, *there is no worksheet*!

### Published by Rubberduck VBA
---------------------------
