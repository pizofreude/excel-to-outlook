# excel-to-outlook

This is an Excel tool to create mass e-mail with custom fields from an XLSM file.

## Description

Using an Excel tool, you may create bulk emails with configurable fields.

When using an Outlook template, the tool scans the "words" in the columns and replaces them with the values in the table before displaying all the created messages for review.

The utility also inserts the "To:" email box, "CC:", Subject, and, if a valid path is provided, adds a file as an attachment. For each line in the table, the tool creates a distinct email. Blank "customfields" will not be taken into account.

Important: Do not alter the arrangement of the rows and columns. You can add columns if you need to retrieve more fields. The columns must have distinctive names.

## Documentation

1. To make each email unique, open Outlook and start a new message.

2. Replace any words or phrases with the table's column headers. Please make sure the words you want to replace exactly match the wording in the column header (an extra space for example could lead to errors).
Outlook message should be saved as an Outlook Template (*.oft).

3. Select the outlook template by clicking the spreadsheet's "Select Template" button.

4. Select "Generate Emails" from the menu.

5. Examine created emails and manually send each of them.

6. Check that all the columns are named appropriately, that no columns have empty rows or columns, and that the headers are on row 10 if an error occurred.
