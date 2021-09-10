---
layout: default
title: Project structure
nav_order: 1
parent: Design patterns
permalink: /class-design/project-structure
---

By now, the ContactEditor app contains several libraries/subpackages. In the RDVBA Code Explorer, I place subpackages either in the project root (e.g., SecureADODB or Storage Library) or under the virtual *Common* folder (e.g., Guard and Logger).

On disk, I have the *Project* folder next to my host file (Excel Workbook), which contains all code modules and folder hierarchy matching the virtual folder hierarchy. *Assets\Diagrams* folder contains documentation diagrams in several formats. Usually, I create originals in [yEd][yEd]-desktop, save them in the native format, and export them as .eps files. Then I open .eps files in Adobe Illustrator, remove the white background, and save them as .ai, .svg, and .jpg/.png. The *Library* folder contains additional files used by the libraries organized in subfolders named after libraries.

[yEd]: https://www.yworks.com/products/yed