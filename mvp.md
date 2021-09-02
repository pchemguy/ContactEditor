---
layout: default
title: Model-view-presenter - GUI
nav_order: 2
permalink: /mvp
---

The *ContactEditorForm* acts as a table browser, displaying one record at a time. The name of each form control matches the associated field name. *ContactEditorPresenter* populates the form from *DataRecordModel* by matching control and field names, minimizing the need for hardcoding the field names. Individual form controls raise "Change" events updating the *DataRecordModel*. The section with "Change" event handlers is the only place requiring hardcoded control/field names. When the user presses the "Apply" or "Ok" buttons, the *DataRecordModel* backend saves changes made by the user.

The form also presents a radio control determining persistence mode. When the user presses the "Apply" or "Ok" buttons and if table updating is not disabled, *ContactEditorPresenter* copy changes from *DataRecordModel* to the corresponding record in *DataTableModel*. *DataTableModel* backend persists individual changes in the "On apply" mode and does a single batch update in the "On exit" via the click handler of the "Ok" button.

*ContactEditorForm* is *modeless* necessitating custom events for passing control to *ContactEditorPresenter*. Thus, the presenter holds a WithEvents form reference as a module-level attribute, and *ContactEditorRunner* holds a presenter reference as a module-level variable. The presenter uses a default interface reference to *ContactEditorForm* to capture the form's events and the form's IDialogView interface to show it.

*ContactEditorForm* uses the *SuppressEvents* boolean flag of *ContactEditorModel* to cancel event processing when the form fields are programmatically updated. (Perhaps, the flag should be moved to the form.)