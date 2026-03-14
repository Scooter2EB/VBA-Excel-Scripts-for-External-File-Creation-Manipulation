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
