Attribute VB_Name = "m2_get_wb_full_path"
Option Explicit

Sub GetExcelFilePath()
    'Create and set dialog box as variable
    Dim dialogBox As FileDialog
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    
    'Do not allow multiple files to be selected
    dialogBox.AllowMultiSelect = False
    
    'Set the title of of the DialogBox
    dialogBox.Title = "Select a file"
    
    'Set the default folder to open
    dialogBox.InitialFileName = ActiveWorkbook.Path
    
    'Clear the dialog box filters
    dialogBox.Filters.Clear
    
    'Apply file filters - use ; to separate filters for the same name
    dialogBox.Filters.Add "Excel workbooks", "*.xlsx;*.xls;*.xlsm"
    
    'Show the dialog box and output full file name
    If dialogBox.Show = -1 Then
        'ActiveSheet.Range("filePath").Value = dialogBox.SelectedItems(1)
        ActiveCell.Value = dialogBox.SelectedItems(1)
    End If
End Sub
