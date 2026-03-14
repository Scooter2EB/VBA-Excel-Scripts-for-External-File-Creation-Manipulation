Attribute VB_Name = "ExternalFileManipulation"
'---------------------------------------------------------------------------------------------------------------
'_______________________________________ Module Name: ExternalFileManipulation __________________________________
'-----------------------------------------------------------------------------------------------------------------
'///////////////// This Module contains several subs designed to do simple manipulations with external files  \\\\\
'//////////////// and folders, such as opening the windows dailog and getting file or folder names or pasting  \\\\\
'/////////////// the print area to a new workbook                                                               \\\\\
'///////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ List Of Subroutines ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ _
'NOTE: Each sub contains better documentation on its purpose and uses. Only Copypaste_PrintArea_NewWorkbook_WithoutAutosave
'       works on non-windows machines
'While some of the subs appear to do things that built-in error catchers already flag, those built-in ones often do
'  not tell the end user why the error occured. By pre-checking user input, you can give them more information as to their mistake
'1. FileFinder
'       Pulls up the windows dialog window allowing users to select a file. Sends file name and path back to calling sub.
'2. FolderFinder
'       Pulls up the windows dialog window allowing users to select a folder. Sends the folder name back to the calling sub
'3. FileSaveAs
'       Pulls up the windows dialog window and allows the user to save the current file in a new location and/or with a new name.
'4. CopyPaste_PrintArea_NewWorkbook_WithoutAutoSave
'       Copies the print area of the current worksheet to a new workbook.
'5. CopyPaste_PrintArea_NewWorkbook_WithoutAutoSave
'       Copies the print area of the current worksheet to a new workbook AND gives the user the option to select a name and
'       location to save the new workbook as.


Public Sub FileFinder(FileName_2, FilePath_2, Optional Extension_1, Optional Title_1)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'____________________________________ User Selection of File From Computer ____________________________________
'////// This script will open up a window based on the current files location and allow the user to select a SINGLE FILE
'////// It then saves the file name and file path as separate variables before passing them back to the calling sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'************************************* General Properties and Quirks ***********************************************
    'Can direct paste script into sub: Yes
    'Can use on Apple: No, unless you use the Mac version of FileDialog
    'Built in Error Handling: Goes to end of sub. Passes FileName_2 and FilePath_2 as Empty, so check for that when
        'when building an error checker
    'User cancel handling: Same result as Error Handling. Done for referntial simplicity. Can easily be modified to
        'differentiate if necessary
'*******************************************************************************************************************
'(((((((((((((((((((((((((((((((((( Required Variables to Pass Through ))))))))))))))))))))))))))))))))))))))))))
    'Variable Name: FileName_2
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: Not Required
        'Output Value: User selected file name w/extension
        'Description: This variable will store the name of the file selected by the user, including the extension,
                      'and pass it through to the calling sub
    'Variable Name: FilePath_2
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: Not Required
        'Output Value: File path for file selected. DOES NOT INCLUDE SLASH AT THE END OF THE PATH
        'Description: This variable will store the entire file path of the file selected by the user,
                      'and pass it through to the calling sub
    'Variable Name: Extension_1
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: Text list of desired visible file extensions, semi colon separated
            'Ex: "*.xlsx; *.xlsm; *.xls; *xlsb"
            'Note 1: If left blank, will default to all files.
        'Output: Input
        'Description: Specifies the file extension that are displayed and can be selected in the file selection window.
                      'Useful for narrowing down choices for users. Defaults to excel files. See Title_1 description for more
    'Variable Name: Title_1
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: Description of the type of files that can be selected
            'Ex: "Excel Files"
            'Note 1: If left blank, will default to All
            'Note 2: This does not effect the types of visible files, it is only a handy reference for users
        'Output: ""
        'Description: (This will be hard to understand). Whenever a window opens up in Windows allowing you to select
                    ' a file, at the bottom of the window is a line where the selected file name will display.
                    'Beside this line is a little box/drop-down menu allowing you to select the types of files you
                    'can see/select. Typically, it will show a name, followed by several extensions which represent
                    'the extensions you can select. This sets the name shown while the previous variable (Extension_1)
                    'sets the list of extensions in the box AND what you can actually see.
                        'For example, when not specified, the default value for Extension_1 is "*.xlsx; *.xlsm; *.xls; *xlsb"
                        'while the default value for Title_1 is "Excel Files". When the file selector opens up,
                        ' the little box/drop-down menu will show "Excel Files (*.xlsx; *.xlsm; *.xls)". So setting
                        'this variable (Title_1) just gives a description of the possible files in english rather than extension language.
                        'You will understand once you play around with it. Maybe.
'((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((()))))))))))))))))))))))))))))))))))))))))))))))))))))))

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ Script to insert in your sub for calling and after call error checking $$$$$$$$$$$$$$$$
'Extension_1={Either semi-colan separated list of extensions or "" for default (excel files)}
'Title_1={Either 1-3 word Description of the types of files or "" for default (excel files)}
'Call FileFinder(FileName_2, FilePath_2, Extension_1, Title_1)
'If FileName_2=Empty or FilePath_2=Empty then
    'Dev Note: I have left a generic error handler which exits the script. You replace it with your own if desired
    ' MsgBox "No File Chosen. Now exiting this subrountine"
    'Application.EnableEvents=True
    'Application.ScreenUpdating=True
    'Exit Sub
'End if
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'##################################################### Beginnig of Sub ###############################################


On Error GoTo Exit_1
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Settable In Sub Options @@@@@@@@@@@@@@@@@@@@@@@@@@@
'### Variable To Turn on/off own file selection selection. Default is off (0)
Self_Select = 0

'########################### Setting Up Title_1 and Extension_1 #####################################################
'''''''' This checks to see if the user has set values for Title_1 and Extension_1. If not, sets default
If Title_1 = Empty Then 'Dev Note: Cannot Handle if Title_1 is non-String for file selector
    Title_1 = "All"
End If
If Extension_1 = Empty Then 'Dev Note: Cannot Handle if Title_1 is non-String for file selector
    Extension_1 = "*.*" 'all files
End If
'############################ Setting up and Executing the File Selector ###########################
'### Setting up File Dialog Application
Section_Execution:
Dim Dialog As Office.FileDialog
Set Dialog = Nothing
Set Dialog = Application.FileDialog(msoFileDialogFilePicker)
'##### Doing Operations with Dialog Selector
With Dialog
    .InitialFileName = ThisWorkbook.path 'Sets where the file selection box opens
    .AllowMultiSelect = False 'Can only Select one file at a time. Cannot change due to lack of ability to record multi files
    .ButtonName = "Select File"
    .Filters.Clear
    .Filters.Add Title_1, Extension_1, 1 'Adding in filters
    .Title = "Choose A File" 'Window Title. Can be Changed
    .InitialView = msoFileDialogViewList 'Don't remember. Leave unless you look it up and know what you are doing
    .Show 'shows the file selector
    '######### Setting the Variables ###########
    On Error GoTo No_Selection
    If 1 = 2 Then
No_Selection:
        'Where it goes if there is a error. If no file is selected, setting FullPath_List throws and error, leads here
        On Error GoTo -1
        On Error GoTo Exit_1
        FileName_2 = Empty
        FilePath_2 = Empty
    Else
        FullPath_List = Split(.SelectedItems(1), Var_Slash) 'creates vector of file path and name seperated by /
        FileName_2 = FullPath_List(UBound(FullPath_List)) 'pulls the file name from the list
        FilePath_2 = Replace(.SelectedItems(1), Var_Slash & FileName_2, "") 'replaces the file name in the full path with a ""; tl/dr sets file path
        If Self_Select = 0 Then  ' Little script to handle when users select the file this script is operating from
            If FileName_2 = ThisWorkbook.Name And FilePath_2 = ThisWorkbook.path Then
                Self_Select_Msg = MsgBox("The file you have selected is this file very file (" & FileName_2 & ")!" & vbNewLine & _
                                        "This is not currently allowed. Please click 'Ok' to select a different file, or click 'Cancel' to exit", _
                                        vbOKCancel + vbExclamation, "File Self Selection Error")
                If Self_Select_Msg = vbOK Then
                    GoTo Section_Execution
                Else
                    GoTo No_Selection
                End If
            End If
        End If
        
    End If
End With

Exit_1:
End Sub

Public Sub FolderFinder(FilePath_2)
'V 1.0
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'____________________________________ User Selection of Folder From Computer ____________________________________
'////// This script will open up a window based on the current files location and allow the user to select a SINGLE FOLDER
'////// It then saves the folder path as a separate variable before passing them back to the calling sub
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'************************************* General Properties and Quirks ***********************************************
    'Can direct paste script into sub: Yes
    'Can Use on non-windows:No
    'Built in Error Handling: Goes to end of sub. Passes FilePath_2 as Empty, so check for that when
        'when building an error checker
    'User cancel handling: Same result as Error Handling. Done for referntial simplicity. Can easily be modified to
        'differentiate if necessary
'*******************************************************************************************************************
'(((((((((((((((((((((((((((((((((( Required Variables to Pass Through ))))))))))))))))))))))))))))))))))))))))))
    'Variable Name: FilePath_2
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: Not Required
        'Output Value: File path for file selected. DOES NOT INCLUDE SLASH AT THE END OF THE PATH
        'Description: This variable will store the entire file path of the file selected by the user,
                      'and pass it through to the calling sub

'((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((()))))))))))))))))))))))))))))))))))))))))))))))))))))))

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ Script to insert in your sub for calling and after call error checking $$$$$$$$$$$$$$$$
'Call FileFinder(FilePath_2)
'If FilePath_2=Empty then
    'Dev Note: I have left a generic error handler which exits the script. You replace it with your own if desired
    ' MsgBox "No Folder Chosen. Now exiting this subrountine"
    'Application.EnableEvents=True
    'Application.ScreenUpdating=True
    'Exit Sub
'End if
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'##################################################### Beginnig of Sub ###############################################


On Error GoTo Exit_1
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Settable In Sub Options @@@@@@@@@@@@@@@@@@@@@@@@@@@
'### Variable To Turn on/off self selection. Default is off (0)
Self_Select = 0
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'##### Determining if computer is Windows or Apple (well, not windows actually but c'mon, linux will never see this)
'      for backslash setting
    'Windows uses \ in file paths, while Apple uses /. This checks the operating system and sets the slashes accordingly
If Not Left(Application.OperatingSystem, 1) Like "W" Then
    Var_Slash = "/"
Else
    Var_Slash = "\"
End If

'########################### Setting Up Title_1 and Extension_1 #####################################################
'''''''' This checks to see if the user has set values for Title_1 and Extension_1. If not, sets default
If Title_1 = Empty Then 'Dev Note: Cannot Handle if Title_1 is non-String for file selector
    Title_1 = "Excel Files"
End If
'############################ Setting up and Executing the File Selector ###########################
'### Setting up File Dialog Application
Section_Execution:
Dim Dialog As Office.FileDialog
Set Dialog = Nothing
Set Dialog = Application.FileDialog(msoFileDialogFolderPicker)
'##### Doing Operations with Dialog Selector
With Dialog
    .InitialFileName = ThisWorkbook.path 'Sets the window to open at the location of the current file
    .ButtonName = "Select Folder"
    .Title = "Choose A Folder" 'Window Title. Can be Changed
    .InitialView = msoFileDialogViewList 'Don't remember. Leave unless you look it up and know what you are doing
    .Show 'shows the file selector
    '######### Setting the Variables ###########
    On Error GoTo No_Selection
    If 1 = 2 Then
No_Selection:
        'Where it goes if there is a error. If no file is selected, setting FullPath_List throws and error, leads here
        On Error GoTo -1
        On Error GoTo Exit_1
        FilePath_2 = Empty
    Else
        FilePath_2 = .SelectedItems(1)
    End If
End With

Exit_1:
End Sub

Private Sub FileSaveAs(FileName_2, FilePath_2, Title_1, Extension_1)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'____________________________________ Save Current File as New one ____________________________________
'////// This script will open up a window based on the current files location and allow the user to navigate anywhere and _
////// save their file in the new location with a new name. This file then becomes the new one
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'************************************* General Properties and Quirks ***********************************************
    'Can direct paste script into sub: Yes
    'Can use on Non-Windows:No
    'Built in Error Handling: Goes to end of sub. Passes FileName_2 and FilePath_2 as Empty, so check for that when
        'when building an error checker.
    'User cancel handling: FileName_2 will equal "User Cancel".
'*******************************************************************************************************************
'(((((((((((((((((((((((((((((((((( Required Variables to Pass Through ))))))))))))))))))))))))))))))))))))))))))
    'Variable Name: FileName_2
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: Not Required
        'Output Value: User selected file name w/extension
        'Description: This variable will store the name of the file selected by the user, including the extension,
                      'and pass it through to the calling sub
    'Variable Name: FilePath_2
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: Not Required
        'Output Value: File path for file selected. DOES NOT INCLUDE SLASH AT THE END OF THE PATH
        'Description: This variable will store the entire file path of the file selected by the user,
                      'and pass it through to the calling sub
    
    'Variable Name: Title_1
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: Description of the type of files that can be selected
            'Ex: "Excel Files"
            'Note 1: If left blank, will default to the example. Useful if opening other excel files
            'Note 2: For all Files, type "All"
            'Note 3: This does not effect the types of visible files, it is only a handy reference for users
        'Output: ""
        'Description: (This will be hard to understand). Whenever a window opens up in Windows allowing you to select
                    ' a file, at the bottom of the window is a line where the selected file name will display.
                    'Beside this line is a little box/drop-down menu allowing you to select the types of files you
                    'can see/select. Typically, it will show a name, followed by several extensions which represent
                    'the extensions you can select. This sets the name shown while the previous variable (Extension_1)
                    'sets the list of extensions in the box AND what you can actually see.
                        'For example, when not specified, the default value for Extension_1 is "*.xlsx; *.xlsm; *.xls"
                        'while the default value for Title_1 is "Excel Files". When the file selector opens up,
                        ' the little box/drop-down menu will show "Excel Files (*.xlsx; *.xlsm; *.xls)". So setting
                        'this variable (Title_1) just gives a description of the possible files in english rather than extension language.
                        'You will understand once you play around with it. Maybe.
    'Variable Name: Extension_1
        'Type: String
        'Min Declaration Level: SUb
        'Input Value: A colon separated list of file extensions you want users to be able to choose from. Each must Start with ".*" _
            Ex: "*.xlsb;*.xlsx"
            'Note 1: If left blank, will default to the current files type
        'Output: ""
        'Description: Works along side Title 1. Gives the list of available file types to view
'((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((()))))))))))))))))))))))))))))))))))))))))))))))))))))))

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ Script to insert in your sub for calling and after call error checking $$$$$$$$$$$$$$$$

'Title_1={Either 1-3 word Description of the types of files or "" for default (excel files)}
'Call FileSaveAs(FileName_2, FilePath_2, Extension_1, Title_1)
'If FileName_2="User Cancel" then
    'Put in something here to exit probably
'ElseIf FileName_2=Empty or FilePath_2=Empty then
    'Dev Note: I have left a generic error handler which exits the script. You replace it with your own if desired
    ' MsgBox "No File Chosen. Now exiting this subrountine"
    'Application.EnableEvents=True
    'Application.ScreenUpdating=True
    'Exit Sub
'End if
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'##################################################### Beginnig of Sub ###############################################

Title_2 = "Save File As" 'Title that appears at the top of the dialog window
On Error GoTo No_Selection
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@ Settable In Sub Options @@@@@@@@@@@@@@@@@@@@@@@@@@@
'### Variable To Turn on/off self selection. Default is off (0)
Self_Select = 0
'@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
'##### Determining if computer is Windows or Apple (well, not windows actually but c'mon, linux will never see this)
'      for backslash setting
    'Windows uses \ in file paths, while Apple uses /. This checks the operating system and sets the slashes accordingly
If Not Left(Application.OperatingSystem, 1) Like "W" Then
    Var_Slash = "/"
Else
    Var_Slash = "\"
End If
'#################### Getting The Current File Name and Extension #####################################
Name_Wrkbook_Active = ActiveSheet.Parent.Name

'########################### Setting Up Title_1 and Extension_1 #####################################################
'''''''' This checks to see if the user has set values for Title_1 and Extension_1. If not, sets default
If Title_1 = Empty Then 'Dev Note: Cannot Handle if Title_1 is non-String for file selector
    Title_1 = "Excel Files"
End If
If Extension_1 = Empty Then 'Dev Note: Cannot Handle if Title_1 is non-String for file selector
    Extension_1 = "*." & Split(Name_Wrkbook_Active, ".")(UBound(Split(Name_Wrkbook_Active, ".")))
ElseIf UBound(Split(Extension_1, "*")) < UBound(Split(Extension_1, ".")) _
    Or UBound(Split(Extension_1, ".")) - 1 <> UBound(Split(Extension_1, ";")) Then 'error checking the extensions
        Debug.Print "You forgot to setup the file extensions properly. Its '*.ex1;*.ex2' etc."
        FileName_2 = Empty
        FilePath_2 = Empty
        GoTo Exit_1
End If

        
'############################ Setting up and Executing the File Selector ###########################
'### Setting up File Dialog Application
If FilePath_2 = "" Then
    'FilePath_2 = Replace(ThisWorkbook.Path, Split(ThisWorkbook.Path, Var_Slash)(UBound(Split(ThisWorkbook.Path, Var_Slash))), "")
    FilePath_2 = ThisWorkbook.path & Var_Slash
ElseIf Right(FilePath_2, 1) <> Var_Slash Then
    FilePath_2 = FilePath_2 & Var_Slash
End If
If FileName_2 <> "" Then
    If Right(FilePath_2, 1) = Var_Slash Then
        FilePath_2 = FilePath_2 & FileName_2
    Else
        FilePath_2 = FilePath_2 & Var_Slash & FileName_2
    End If
End If

Section_Execution:
Dialog = Application.GetSaveAsFilename(InitialFileName:=FilePath_2, _
                                          FileFilter:=Title_1 & "(" & Extension_1 & ")," & Extension_1, Title:=Title_2)

    
If Dialog = False Then
User_Cancel:
    FileName_2 = "User Cancel"
    FilePath_2 = Empty
    GoTo Exit_1
No_Selection:
    'Where it goes if there is a error. If no file is selected, setting FullPath_List throws and error, leads here
    On Error GoTo -1
    On Error GoTo Exit_1
    FileName_2 = Empty
    FilePath_2 = Empty
    GoTo Exit_1
Else
    FullPath_List = Split(Dialog, Var_Slash) 'creates vector of file path and name seperated by /
    FileName_2 = FullPath_List(UBound(FullPath_List)) 'pulls the file name from the list
    FilePath_2 = Replace(Dialog, FileName_2, "") 'replaces the file name in the full path with a ""; tl/dr sets file path
    If Self_Select = 0 Then  ' Little script to handle when users select the file this script is operating from
        If FileName_2 = ThisWorkbook.Name And FilePath_2 = ThisWorkbook.path & Var_Slash Then
        Self_Select_Msg = MsgBox("The file you have selected is this very file (" & FileName_2 & ")!" & vbNewLine & _
                                "This is not currently allowed. Please click 'Ok' to select a different file, or click 'Cancel' to exit", _
                                vbOKCancel + vbExclamation, "File Self Selection Error")
            If Self_Select_Msg = vbOK Then
                GoTo Section_Execution
            Else
                GoTo User_Cancel
            End If
        End If
    End If
End If


ThisWorkbook.SaveAs FilePath_2 & FileName_2
Exit_1:
End Sub

Public Sub CopyPaste_PrintArea_NewWorkbook_WithOutAutosave()
'_________________________ Copy the Print selected area to a new workbook_____________________________
'//////////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ _
/////// This script selects the pre-setup print area on the current worksheet(from page break view) \\\\\\\\\\\\\\\\\\\ _
////// and copys it to a new workbook created by the script. It is copied to the same columns as the \\\\\\\\\\\\\\\\\ _
////// source print area, but the start row can be adjusted using a variable in this script (default=2) \\\\\\\\\\\\\\\\ _
//////////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'******************************* General Properties and Quirks **********************************************************
    'Can CopyPaste into new file?: Yes _
        - Must be a module
    'Can be used on Apple?: Yes
'********************************** Adjustable Features/Variables *******************************************************
    '1. Row Height of Paste location _
            Description: When you are pasting the print area to a new workbook, you can adjust the start row _
            you are pasting it to. _
            Varible to adjust: StartNum _
            Possible Options: Any discrete number _
            Location in Script: Just Before Part 1
'((((((((((((((((((((((((((((((((((((((((((((((((()))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))


Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim Wrkbook_New As Workbook
Name_Wrksht_Active = ActiveSheet.Name
Name_Wrkbook_Host = ThisWorkbook.Name
StartNum = 2 'This is the row number you start your pasting to

'----------------------------------- Part 1: Setting up New File ---------------------------------------------
'################################### Section 1.1: Creating New File #############################################
Set Wrkbook_New = Workbooks.Add
'################################### Section 1.2: Paste Values #################################################
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
'////////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'#Error Section
On Error Resume Next
Set NewSheet = Wrkbook_New.Worksheets(1)
NewSheet.Name = Name_Wrksht_Active
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

