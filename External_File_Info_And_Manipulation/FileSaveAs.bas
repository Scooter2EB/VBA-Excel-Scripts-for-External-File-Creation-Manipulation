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
