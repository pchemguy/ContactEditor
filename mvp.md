---
layout: default
title: Model-view-presenter - GUI
nav_order: 2
permalink: /mvp
---

<a name="FigMainForm"></a>  
<div align="center"><img src="https://raw.githubusercontent.com/pchemguy/ContactEditor/develop/Assets/Diagrams/ContactEditorForm.jpg" alt="Main form" width="80%" /></div>
<p align="center"><b>Main form</b></p>  

The Model-View-Presenter (MVP) part of the application consists of three classes, *ContactEditorModel* (M), *ContactEditorForm* (V), and *ContactEditorPresenter* (P). A regular module, *ContactEditorRunner*, holds the main entry point.

*ContactEditorForm* acts as a table browser, displaying one record at a time. Each control is associated with a particular database field. By convention, both control elements and associated fields should have the same names, so only identifiers that can satisfy this condition are allowed. When the application starts, backends can perform introspection and populate the *FieldNames* array. (While introspection is natural for the DataTable backends, the *DataRecodWSheet* backend can also collect field names based on a convention described in the code.) Then *ContactEditorPresenter* populates the form from *DataRecordModel*:

```vb
Private Sub LoadFormFromModel()
    this.Model.SuppressEvents = True
        
    Dim FieldName As Variant
    Dim FieldIndex As Long
    Dim FieldNames As Variant
    FieldNames = this.Model.RecordTableManager.FieldNames
    For FieldIndex = LBound(FieldNames) To UBound(FieldNames)
        FieldName = FieldNames(FieldIndex)
        view.Controls(FieldName).Value = CStr(this.Model.RecordTableManager.RecordModel.GetField(FieldName))
    Next FieldIndex

    this.Model.SuppressEvents = False
End Sub
```

Note the For loop, which goes through the array of collected field names. Since *DataRecordModel* wraps a dictionary object, field names can be used as the keys to retrieve the associated values. Individual control objects can be accessed using their names via the *Controls* collection. Because by convention, the names of controls and fields match, the code that populates the form does so without hardcoding any actual names.

Individual form controls raise "Change" events updating the *DataRecordModel*. The section with "Change" event handlers is the only place requiring hardcoded control/field names. When the user presses the "Apply" or "Ok" buttons, the *DataRecordModel* backend saves changes made by the user to the DataRecordStorage via an instance of the *DataCompositeManager* class described later.

On the right, the form presents a radio control determining persistence mode. When the user presses the "Apply" or "Ok" buttons and, if Table Updating is not Disabled, *ContactEditorPresenter* copies the changes from *DataRecordModel* to the corresponding record in *DataTableModel* via *DataCompositeManager*. *DataTableModel* backend persists individual edits in the "On apply" mode and does a single batch update in the "On exit" via the click handler of the "Ok" button (*ContactEditorPresenter* initiates the updating process via a call to *DataCompositeManager*).

*ContactEditorForm* is *modeless* necessitating custom events for passing control to *ContactEditorPresenter*. Thus, the presenter holds a WithEvents form reference as a module-level attribute, and *ContactEditorRunner* holds a presenter reference as a module-level variable. The presenter uses a default interface reference to *ContactEditorForm* to capture the form's events and the form's IDialogView interface to show it.

*ContactEditorForm* uses the *SuppressEvents* boolean flag of *ContactEditorModel* to cancel event processing when the form fields are programmatically updated. (Perhaps, the flag should be moved to the form.)