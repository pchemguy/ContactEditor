---
layout: default
title: Overview
nav_order: 1
parent: Storage Library
permalink: /storage-library/overview
---

### Functional structure

[Fig. 1](#FigDataFlow) illustrates a simplified data flow process in a data management application. Storage Library (StoreLib), shown in the center of the diagram, mediates data transfer between the user (UI) and storage (Database). The library contains two core class families, *DataRecord* and *DataTable*. *DataRecord* class family directly interacts with the UI, handles the data chunk currently presented to the user (e.g., a particular record), and can save a copy of this active chunk to a local file. *DataTable* interacts directly with the data storage and holds multiple chunks of data (e.g., a set of records).

<a name="FigDataFlow"></a>  
<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/master/Assets/Diagrams/Data Flow Diagram.svg" alt="FigDataFlow" width="100%" />  
<p align="center"><b>Figure 1. Basic data flow diagram</b></p>  

Data management processes, including loading, editing, and saving the data via the StoreLib, are illustrated on the flow [chart](#FigDataManagementFlowChat) below.

<details><summary>Data management flow chart</summary>
<a name="FigDataManagementFlowChat"></a>  
<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/master/Assets/Diagrams/Data Manager Flow Chart.svg" alt="FigDataManagementFlowChat" width="100%" />  
<p align="center"><b>Figure 2. Data management flow chart</b></p>  
</details>

A dedicated class, *DataCompositeManager*, bridges the two class families. It handles data transfers between *DataRecord* and *DataTable* without introducing unnecessary coupling.

### Data model

*DataRecordModel* holds a single data record (table row). This model class has one backend, *DataRecordWSheet*, which implements the *IDataRecordStorage* interface, saves the data to Excel Worksheet, and populates the model at application startup. The *DataRecordFactory* abstract factory for record backends implements the *IDataRecordFactory* ([Fig. 3](#FigDataRecord)).

<a name="FigDataRecord"></a>  
<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/master/Assets/Diagrams/Class Diagram - Record.svg" alt="FigDataRecordModel" width="100%" />  
<p align="center"><b>Figure 3. DataRecord class diagram</b></p>  

*DataTableModel* contains a whole data table or a subset of rows, abstracting persistent storage. This model class has four backends, including *DataTableWSheet*, *DataTableCSV*, and *DataTableADODB*&nbsp;/&nbsp;*DataTableSecureADODB*. They handle a Worksheet, a delimited text file, and a relational database, respectively. The *IDataTableStorage* class formalizes their interface and the *DataTableFactory*/*IDataTableStorage* abstract factory instantiates them ([Fig. 4](#FigDataTable)).

<a name="FigDataTable"></a>  
<img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/master/Assets/Diagrams/Class Diagram - Table.svg" alt="FigDataTableModel" width="100%" />  
<p align="center"><b>Figure 4. DataTable class diagram</b></p>  

Data Managers comprise the last part of the library. Each such class incorporates by composition a model class and an appropriate backend class, yielding "backend-managed" models. Applications using this library should typically instantiate and keep a reference to one of these classes. The library defines three manager classes, including two individual and one composite.

*DataRecordManager* ([Fig. 3](#FigDataRecord)) and *DataTableManager* ([Fig. 4](#FigDataTable)) implement *IDataRecordManager* and *IDataTableManager* interfaces and manage their respective model classes. Additionally, since *DataRecordModel* and *DataTableModel* may work cooperatively (one holding a record from the recordset held in the other), the data may need to be transferred between the two model classes. Thus, a third composite manager is necessary.

*DataCompositeManager* is used where *DataRecordModel* and *DataTableModel* work together, and it handles data transfers between the model classes. A composite manager can encapsulate either two backend-managed classes or two model and two backend classes directly. *DataCompositeManager* uses the latter approach ([Fig. 5](#FigCompositeManager)).

<a name="FigCompositeManager"></a>  
<img src="https://github.com/pchemguy/ContactEditor/blob/master/Assets/Diagrams/Class Diagram.svg?raw=true" alt="Overview" width="100%" />  
<p align="center"><b>Figure 5. Composite manager</b></p>  

Lets us overlap the structural application schematic from the previous section and the class hierarchy, as shown in [Fig. 6](#FigFunctionalClassMapping). Note that the figure has only one DataTable backend class, one Data Manager class, while abstract factory classes are not shown. The arrows indicate the flow of data, and the classes placed near the middle of these arrows facilitate/implement such transfers (*ContactEditorPresenter* is responsible for "GUI&nbsp;&#x21D4;&nbsp;*DataRecordModel*" transfer).

<a name="FigFunctionalClassMapping"></a>  
<img src="https://github.com/pchemguy/ContactEditor/blob/master/Assets/Diagrams/Overview Class Map.svg?raw=true" alt="Functional class mapping" width="100%" />  
<p align="center"><b>Figure 6. Functional class mapping.</b></p>  

The *ContactEditorModel* (MVP) and *DataRecordModel*/*DataTableModel* (Storage Library) classes act as a container for the data. As illustrated in the figure, the former encapsulates the latter via indirect composition through the DataCompositeManager class. The four classes thus bridge the MVP and Storage Library components.