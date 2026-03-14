
Public Sub CopyPaste_PrintArea_NewWorkbook_WithAutosave()
'_________________________ Copy the Print selected area to a new workbook_____________________________
'//////////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'/////// This script selects the pre-setup print area on the current worksheet(from page break view) \\\\\\\\\\\\\\\\\\\
'////// and copys it to a new workbook created by the script. It is copied to the same columns as the \\\\\\\\\\\\\\\\\
'////// source print area, but the start row can be adjusted using a variable in this script (default=2) \\\\\\\\\\\\\\\\
'///// It also gives users the option to select a file save name an location, with checks on those.    \\\\\\\\\\\\\\\\\\
'///// This autosave feature can be turned off using a variable in this script, although there is a version \\\\\\\\\\\\\
'//// of this script which does not have autosave built in, which might be easier                        \\\\\\\\\\\\\\\\\
'//////////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'******************************* General Properties and Quirks **********************************************************
    'Can CopyPaste into new file?: Yes
    '    - Must be a module
    'Can be used on Apple?: Yes, so long as SkipAutoSave is turned off (set to 1)
    '    -if there is a version of this script without autosave, use that instead
'******************************* Required Calls *************************************************************************
'This is the list of other subroutines you need for this script to function, as well as there
'location in this file. VBA will not allow the sub to run without them, so they are a must.
    'Located in Module ExternalFileManipulation: FileSaveAs
    
'********************************** Adjustable Features/Variables *******************************************************
    '1. Row Height of Paste location _
    '        Description: When you are pasting the print area to a new workbook, you can adjust the start row
    '        you are pasting it to.
    '        Varible to adjust: StartNum
    '        Possible Options: Any discrete number
    '        Location in Script: Just Before Part 1
    '2. Toggle Autosave on or off. _
    '    Description: This script has the ability to get a filename a save location from the user so the workbook
    '                 they paste into is already named and saved. This can be turned off, so the new workbook will
    '                 be 'book1' and the user will have to use the regular method of file saving to save it.
    '    Variable to Adjust: SkipAutoSave
    '    Possible Options: 1 or 0 (not 1)
    '    Option Outcomes: 1 will skip the autosave feature. Otherwise autosave will not be skipped
    '    Location in Script: Just Before Part 1
'((((((((((((((((((((((((((((((((((((((((((((((((()))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))


Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim Wrkbook_New As Workbook
Name_Wrksht_Active = ActiveSheet.Name
Name_Wrkbook_Host = ThisWorkbook.Name
StartNum = 2 'This is the row number you start your pasting to
SkipAutoSave = 0 'Set to 1 to skip save procedure. Data will be copied into unsaved workbook, user saves manually
'------------------------------ Part 1: Getting Info for setting up new file ------------------------------------
On Error GoTo Skip_Auto_Save 'If there is an error with the file name setup, users can create the file but save the old fashion way
Err_Location = "setting the new file name and path."

Begin_NewFile_Procedure:
Title_1 = "Excel Files": Extension_1 = "*.xlsx;*.xlsm;*.xls"
Call FileSaveAs(FileName_2, FilePath_2, Title_1, Extension_1) 'Call to setup file name and path
If FilePath_2 = "" Or FileName_2 = "" Then
    GoTo Skip_Auto_Save
End If

FullPath = FilePath & "\" & FileName


'----------------------------------- Part 2: Setting up New File ---------------------------------------------
'################################### Section 2.1: Saving New File #############################################
If 1 = 2 Then
Skip_Auto_Save:
    SkipAutoSave = 1 'Allows for file to be created but not saved so that user can save manually
End If
On Error GoTo -1
On Error GoTo 0
Set Wrkbook_New = Workbooks.Add

'################################### Section 2.2: Paste Values #################################################
 '### Get Print Area Range
Adr_Copy = Split(Workbooks(Name_Wrkbook_Host).Worksheets(Name_Wrksht_Active).PageSetup.PrintArea, ",") 'Divides up non-contiguous ranges

For i = LBound(Adr_Copy) To UBound(Adr_Copy) 'Loops through non-contiguous ranges to copy separatly. Excel cannot copy non-contiguous ranges apparently...
    'Create Paste Range At Top of worksheet
    RowNum_PasteStart = Split(Split(Adr_Copy(i), ":")(0), "$")(2)
    RowNum_PasteEnd = Split(Split(Adr_Copy(i), ":")(1), "$")(2)
    MoveGap = RowNum_PasteStart - StartNum 'How Many numbers it must move up
    Adr_Paste = Replace(Adr_Copy(i), RowNum_PasteStart, StartNum)
    Adr_Paste = Replace(Adr_Paste, RowNum_PasteEnd, RowNum_PasteEnd - MoveGap)
    'Setting Ranges and copypasting
    Set Range_Copy = Workbooks(Name_Wrkbook_Host).Worksheets(Name_Wrksht_Active).Range(Adr_Copy(i))
    Set Range_Paste = Wrkbook_New.Worksheets(1).Range(Adr_Paste)
    Range_Copy.Copy
    Range_Paste.PasteSpecial xlPasteAll
    Range_Paste.Copy
    Range_Paste.PasteSpecial xlPasteValues
Next i
Application.CutCopyMode = False
'#Error Section
On Error Resume Next
Set NewSheet = Wrkbook_New.Worksheets(1)
NewSheet.Name = Name_Wrksht_Active
'################################### Section 2.3: Saving Workbook ##############################################
If Skip_Auto_Save <> 1 Then
    Application.AlertBeforeOverwriting = False
    Wrkbook_New.SaveAs FullPath
    Application.AlertBeforeOverwriting = True
End If
'////////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

If 1 = 2 Then
Exit_Error_Generic:
    On Error GoTo -1
    On Error GoTo 0
    MsgBox "Unfortunately there was an error when " & Err_Location & " If this issue persists, please contact the developer", vbOKOnly + vbInformation, "Error Occured"
End If
Exit_Normal:
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub
