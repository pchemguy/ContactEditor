---
layout: default
title: Home
nav_order: 1
permalink: /
---

### Overview

"Contact Editor" demos the Model-View-Presenter (MVP) pattern backed by persistent storage (MVP-DB) in VBA. While studying OOP design patterns, I developed this mock application as a course project to serve as an MVP-DB template/prototype for my VBA experiments. It partially implements the "user <---> database" interaction workflow schematically shown in the [figure](#FigUserDbLoop). Typically, the user sends a query to the database, receives a response table, browses record data via a user form, and, possibly, updates the database.

<a name="FigUserDbLoop"></a>

<img src="https://github.com/pchemguy/ContactEditor/blob/develop/Assets/Diagrams/Overview.jpg?raw=true" alt="Overview" width="100%" />

<p align="center"><b>User - database interaction workflow.</b></p>

### Development environment and repo structure

Contact Editor demo is a VBA app and is part of an Excel Workbook [Contact Editor.xls][Contact Editor] available from the root of this repo. Additionally, all code modules and user forms are available from the [Project][Project] folder (which acts as a container and corresponds to the VBA project root within the .xls file). [Rubber Duck VBA][Rubber Duck VBA] add-in greatly facilitated the development process, and the project structure is exported/imported using the [RDVBA Project Utils][RDVBA Project Utils] VBA module. Primarily, I use Excel 2002 for development and also run tests on Excel 2016.

### Running the demo with different backends

*ContactEditorRunner.RunContactEditor* is the main entry for the demo.  
*ContactEditorPresenter.InitializeModel* performs basic backend configuration.  
The *DataTableBackEnd* variable found in the entry point sub determines the type of the main storage backend. It can take one of the following values:

- "Worksheet" for an Excel Worksheet (demo database [file][ContactEditor.xsv]),
- "CSV" for a text delimited file (demo database [file][Contact Editor]), and
- "ADODB" for a relational database (demo SQLite database [file][ContactEditor.db]).

While the "ADODB" backend can connect to both "CSV" and "Worksheet" databases, the dedicated backends should be more efficient.

### Documentation and further information

Documentation and technical details are available from the project [WiKi][WiKi].

### Acknowledgments

A special thanks goes to Mathieu Guindon, a co-founder of the [Rubber Duck VBA][Rubber Duck VBA] project and his [RDVBA blog][RDVBA blog]. RDVBA [blog post][RDVBA No Worksheet] describing a possible approach to abstracting a Worksheet-based persistent storage and a [demo file][RDVBA No Worksheet Demo] helped me jump-start with storage integration. I also followed the [blog post][RDVBA UserForm1.Show] regarding the best practices for UserForm handling and the [SO answer][RDVBA Modeless Form] (and the last comment to that answer) regarding the modeless user forms.


[Rubber Duck VBA]: https://rubberduckvba.com
[RDVBA blog]: https://rubberduckvba.wordpress.com
[RDVBA No Worksheet]: https://rubberduckvba.wordpress.com/2017/12/08/there-is-no-worksheet
[RDVBA No Worksheet Demo]: https://rubberduckvba.wordpress.com/2017/12/08/there-is-no-worksheet/#div-comment-286
[RDVBA UserForm1.Show]: https://rubberduckvba.wordpress.com/2017/10/25/userform1-show
[RDVBA Modeless Form]: https://stackoverflow.com/questions/47357708/vba-destroy-a-modeless-userform-instance-properly#answer-47358692
[RDVBA Project Utils]: https://github.com/pchemguy/RDVBA-Project-Utils
[Overview]: https://github.com/pchemguy/ContactEditor/blob/develop/Assets/Diagrams/Overview.jpg?raw=true
[Project]: https://github.com/pchemguy/ContactEditor/tree/master/Project
[Contact Editor]: https://github.com/pchemguy/ContactEditor/blob/master/ContactEditor.xls
[ContactEditor.db]:  https://github.com/pchemguy/ContactEditor/blob/master/ContactEditor.db
[ContactEditor.xsv]: https://github.com/pchemguy/ContactEditor/blob/master/ContactEditor.xsv
[WiKi]: https://github.com/pchemguy/ContactEditor/wiki
