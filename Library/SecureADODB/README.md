This project proposes several modifications to the [SecureADODB] library, as well as explores several alternative design options.

### Class diagram and connections to ADODB

The class diagram [below](#FigClassDiagram) shows the core SecureADODB classes (this fork, blue) and the mapping to the core ADODB classes (green).

<a name="FigClassDiagram"></a>

<img src="https://raw.githubusercontent.com/pchemguy/SecureADODB-Fork/master/UML%20Class%20Diagrams/SecureADODB%20-%20ADODB%20Class%20Mapping.svg" alt="Overview" width="100%" />

<p align="center"><b>SecureADODB Fork class diagram (interfaces not shown)</b></p>

*DbRecordset* class has been added in this fork, while the DbManager class shown at the bottom is functionally similar to the UnitOfWork class from the base project.

*DbManager* class streamlines the interaction with the database by facilitating typical workflows. This class glues together individual base classes. It takes connection parameters or connection string and then instantiates other classes and injects dependencies as necessary.

*DbConnection* and *DbRecordset* classes receive and handle events raised by the corresponding ADODB classes (which was part of the motivation for creating the *DbRecordset* class). The core events (associated with connection, execution, and transaction) are accessible via the “Connection” class. Accessing asynchronous fetching events, however, requires the “Recordset” class.

### Usage examples

Please see *ExamplesDbManager.bas* (examples using this fork of the SecureADODB library) and *ExamplesPlainADODB.bas* (plain ADODB examples) in "VBAProject\SecureADODB\DbManager" for usage examples. SecureADODB examples produce output via Debug.Print and via the QueryTable feature. Plain ADODB module produces very little output, and its results can be examined via the debugger or by modifying the code.

### Core differences from RDVBA SecureADODB

1). A coupling loop between *DbConnection* and *DbCommand* has been removed (issue [IDbConnection_CreateCommand interface][Issue 14]).  

  <a name="FigSecureADODBloop"></a>

  <img src="https://raw.githubusercontent.com/pchemguy/SecureADODB-Fork/master/UML%20Class%20Diagrams/SecureADODB_CreateCommand%20Loop.svg" alt="SecureADODB loop" width="100%" />

  <p align="center"><b>SecureADODB dependency loop</b></p>

2). *AutoDbCommand* and *DefaultDbCommand* have been replaced with *DbCommand* and *DefaultDbCommandFactory* replaced with *DbCommandFactory*. *DbCommand* always takes an existing *DbConnection* class as a dependency, and is only responsible for ExecuteNoQuery functionality ([NoQuery flag] commit), while queries returning a Recordset or a scalar are executed via the *DbRecordset* class.  
3). *DbManager* takes a flag, turning transactions on/off. Additionally, the BeginTransaction method now has a transaction error handler. If this handler traps an error, it sets a flag on the DbConnection object disabling further transaction handling.  
4). A new Guard class replaces the Errors module with some refactoring and additional functionality. A  "Scripting.Dictionary" backed logger prototype has also been implemented.  
5). Design patterns:  

  - *Factory-Constructor pattern*. Following the convention of the base project, the default concrete factory is the "Create" method defined on default class instances. Initialization, on the other hand, is not performed by a set of public setters but rather via a corresponding constructor ([Factory-Constructor pattern][] issue). Please see [Contact Editor tutorial][Factory-Constructor - Contact Editor] for additional discussion about the returned value.
  - *Abstract Factory and CreateInstance convention*. Аbstract factory's Create method generates factory instances. Factory instance's *CreateInstance* method, in turn, generates instances of the target class ([CreateInstance convention] issue).  
  - *Duplicate Guard clauses*. Factories hold only the non-default instance guard, which might be redundant when the factory produces non-default interface objects lacking the factory method. The factory passes all initial values to the new instance constructor responsible for validation guards/checks.  

6). *DbRecordset* class handles queries returning disconnected or online Recordsets, as well as scalars. A fully initialized “ADODB.Command” sets most of the *DbRecordset*’s properties (via injected *DbCommand*). Several options (such as return type and cursor type/location) are supplied to the *DbRecordset* factory directly.  
7). A new module, DbManagerITests, runs a set of tests against mock CSV and SQLite databases. This way, actual SecureADODB classes (as opposed to stubs) are tested. DbManagerITests tests also serve as use templates.  

[SecureADODB]: https://github.com/rubberduck-vba/examples/tree/master/SecureADODB
[Class Diagram]: https://raw.githubusercontent.com/pchemguy/SecureADODB-Fork/master/UML%20Class%20Diagrams/SecureADODB%20-%20ADODB%20Class%20Mapping.svg
[Issue 14]: https://github.com/pchemguy/RDVBA-examples/issues/14
[SecureADODB loop]: https://raw.githubusercontent.com/pchemguy/SecureADODB-Fork/master/UML%20Class%20Diagrams/SecureADODB_CreateCommand%20Loop.svg
[NoQuery flag]: https://github.com/pchemguy/RDVBA-examples/commit/ffc12ffb361ecc5a2338a321d84e8a756b48e109
[Factory-Constructor pattern]: https://github.com/pchemguy/RDVBA-examples/issues/11
[Factory-Constructor - Contact Editor]: https://pchemguy.github.io/ContactEditor/class-design
[CreateInstance convention]: https://github.com/pchemguy/RDVBA-examples/issues/10