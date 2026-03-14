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
