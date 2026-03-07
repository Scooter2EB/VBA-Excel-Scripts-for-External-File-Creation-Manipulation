Sub CheckIf_FileProvided_AlreadyExists(FileName, FilePath, FileExtension, Optional Var_Slash, Optional Display_Input1, Optional FileExists, Optional FileOverwrite)

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'__________________________ Check if file already exists in specified location ____________________________________
'////// This script takes the file name provided by the user and checks if it already exists in the specified location.
'////// If it does, gives users the option to change the file name before sending it back. _
'////// Will check the full path legnth if change.
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'************************************* General Properties and Quirks ***********************************************
    'Can direct paste script into sub: Yes
    'Can Direct paste script into worksheet sub: Yes
    'Required inputs: Hard 2, soft 4
    'Can use on Apple: Yes
    'Built in Error Handling: Quietly quits sub. Sends back FileExists=2
    'User cancel handling: Only occurs if file does exist. Sends back FileExists=1 and FileOverwrite=2
    'Calls to other subs: None, but Checker_Length_FullPath is embedded and could be replaced
    'EXTRA NOTES _
        '1. 'ASSUMES THE FILE PATH PROVIDED EXISTS. Checking for this is a different script. _
        2. If Filename has an extension and File Extension is non-empty, the one in the file name will be used.
'*******************************************************************************************************************
'(((((((((((((((((((((((((((((((((( Variables Passed Through and/or Returned ))))))))))))))))))))))))))))))))))))))))))
    'Variable Name: FileName
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: The name of the file. For more, see description
        'Output Value: Either the same name or a new file name depending on whether they change it
        'Description: This varibale will consist of the name (no path) of the file to be checked for existence. _
                        If also passing the FileExtension variable through to this sub, DON'T include the extension _
                        in the file name. If File Extension is empty, you MUST include the file extension
    'Variable Name: FilePath
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: The path for the file being checked. Include drive letters
        'Output Value: Same file path.
        'Description: Used to create the full file path for the file being checked
    'Variable Name: FileExtension
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: (Soft Requirement) The file extension of the file you are checking (e.g. .xlsm)
        'Output Value: Same as input if no issue. Can be newly determine extension. If =1, user quit at new ext stage
        'Description: Holds the file exentions of the file being checked. Used in constructing the full file name _
                        and path. Can be left empty if the file name already includes
    'Variable Name: Var_Slash
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: (Soft Requirement) The direction of the slash for file paths. Its \ unless on Apple. Leave blank for default of \
        'Output Value: Same as Input
        'Description: This just lets the sub know which slash to use for file paths. Will usually be \ for windows computers. _
                        If this variable is already declared for local or public use, you can ignore
    'Variable Name: Display_Input1
        'Type: Byte
        'Min Declaration Level: Sub
        'Input Value: (Soft Requirement) 1 or 0. If not defined, defaults to 0
        'Output Value: Same as Input
        'Description: This tells this sub whether to display the inputbox for a new file name in the event the file already exists. _
                      Setting it to 1 skips this option for the user, allowing for custom messages in the calling sub
    'Variable Name: FileExists
        'Type: Byte
        'Min Declaration Level: Sub
        'Input Value: None
        'Output Value: 0,1 or 2
        'Description: Indicator for the calling sub to let it know whether the file did exist. 0 means no _
                    , 1 means yes, and 2 means there was an error thrown and it could not be determined
    'Variable Name: FileOverwrite
        'Type: Byte
        'Min Declaration Level: Sub
        'Input Value: None
        'Output Value: 0,1,2 or 3
        'Description: Indicator for the calling sub to let it know that the user intends to overwrite the existing file _
                      These indicators work semi-dependantly with the FileExists indicator. For example, if this indicator says _
                      it wants to overwrite (0), that does not mean there is a file to overwrite, only that the calling sub _
                      does not have to worry about saving the existing _
                      0. Can Overwrite _
                      1. Means Cannot Overwrite _
                      2. Means the user tried to cancel the whole procedure _
                      3. Means the file exists but the user was not asked about whether they want to overwrite
       
       'Interpetation Note: FileExists(FE) and FileOverwrite(FO) _
            FE=0,FO=0: This means the file did not already exist so there is no issue saving it _
            FE=1,FO=0: This means the file does exist but you can overwrite it _
            FE=1,FO=1: This means the file does exist and you cannot overwrite _
            FE=0,FO=1: Should not occur.
                                        
                            
                                     
'((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((((()))))))))))))))))))))))))))))))))))))))))))))))))))))))

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ Script to insert in your sub for calling and after call error checking $$$$$$$$$$$$$$$$

'FileName={Name of the file. May or may not have file extension}
'Filepath={File path of the file}
'FileExtension={the extension of the file, including the period. (ex. .xlsx)}
'Call CheckIf_FileProvided_AlreadyExists(FileName, FilePath, FileExtension, Var_Slash,Display_Input1, FileExists, FileOverwrite)
'On Error Goto Exit_Error_Generic 'IF this gives you an error, delete it or make it go to where you send errors
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'##################################################### Beginnig of Sub ###############################################
'////////////////////////////////////////// Section 1: Setup \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
On Error GoTo Exit_Error_Generic
If FileName = "" Or FilePath = "" Then 'Incase there is no FileName or FilePath
    MsgBox "File Name or File Path not provided!"
    FileExists = 2
    FileOverwrite = 2
    Exit Sub
End If
If Var_Slash = "" Then
    Var_Slash = "\" 'for windows
    If Not Left(Application.OperatingSystem, 1) Like "W" Then: Var_Slash = "/"
End If
'################################# Part 1.1 File Extension Checking ##################################################
CurrentExt = 1 'Indicates whether the current file name has an extension tacked on
Vector_FileName = Split(FileName, ".") 'Splits the file name by each period in the name
'## Determine if filename provided has a file extension on it.
If UBound(Vector_FileName) = 0 Then 'When there are no periods in the file name (so no file extension)
    CurrentExt = 0
ElseIf Len(Vector_FileName(UBound(Vector_FileName))) < 3 _
Or Len(Vector_FileName(UBound(Vector_FileName))) > 5 Then 'If number of characters after the last period are not between 3-5
    CurrentExt = 0
End If
    '## When no file extension, checking FileExtension Variable and/or asking for new extension
If CurrentExt = 0 Then
    If FileExtension = "" Then 'when FileExtension is blank
New_FileExtension:
        Input_Ext = InputBox("It appears as though the file you wish to save, " & FileName & ", does not have a file extension " & _
                             " specified. Please specify the file extension you would like to use (ex: .xlsx, .docx), " & _
                             "in the box below and click 'Ok', or click 'Cancel' to stop the procedure", "No File Extension")
        If Input_Ext = "" Then 'When they indicate they want to leave
            FileExtension = 1
            Exit Sub
        ElseIf Len(Input_Ext) > 5 Then 'If file extension is too long
            MsgBox "It appears as though the Extension you inputted, " & Input_Ext & " is not a recognized file extension." & _
                   "Please try again."
            GoTo New_FileExtension
        End If
        FileExtension = Input_Ext
    End If
    If Left(FileExtension, 1) <> "." Then 'Making sure there is a period
        FileExtension = "." & FileExtension
    End If
    FileName = FileName & FileExtension 'Attaching the New File Extension
End If
Vector_FileName = Split(FileName, ".")
FileExtension = "." & Vector_FileName(UBound(Vector_FileName))
'/////////////////////////////////////// Part 2: Checking if the file Exists \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
If Right(FilePath, 1) <> Var_Slash Then
    FilePath = FilePath & Var_Slash
End If
FullPath = FilePath & FileName
If Dir(FullPath) <> "" Then 'Checking if File exists at current path
    '###################### Section 2.1: If File Exists at current path ##########################
File_Exists_Start:
    FileExists = 1
    If Display_Input1 = 0 Then 'Optional built in menu to choose to overwrite or rename. Can be skipped if external menu chosen
        Input_FileExists = InputBox("It appears as though the file you are looking create/save, " & FileName & _
                             " already exists in the location you have specified (" & FilePath & ")." & vbNewLine & _
                             " Would you like to: " & vbNewLine & vbNewLine & "1. Overwrite this file" & vbNewLine & vbNewLine & _
                             "2.Change the file name" & vbNewLine & vbNewLine & "3. Stop this procedure", "File already exists", Replace(FileName, Vector_FileName(UBound(Vector_FileName)), ""))
        If Input_FileExists = 1 Then 'Choosing to overwrite the existing file
            File_Overwrite = 0
        ElseIf Input_FileExists = 2 Then 'Choosing to Rename the file
Input_NewName:
            Input_FileExists = InputBox("Please enter the new name of your file (excluding file extension)" & vbNewLine & _
                                        "Note: You cannot use these characters: /\:*?<>|", "New File Name", File)
            If Input_FileExists <> "" Then
                FileName = Input_FileExists
                FullPath = FilePath & FileName & FileExtension
                '##### Check if new file exists ####
                If Dir(FullPath) <> "" Then
                    GoTo File_Exists_Start:
                End If
                '#### Checking for special characters in name (Requires Call) ###########
                Check_FileName = FileName
                Has_Special = 0
                Call Checker_FileName_SpecialCharacters(Check_FileName, Has_Special, Display_Msg1)
                FileName = Check_FileName
                If Has_Special = 2 Then 'If user cancels at renaming stage
                    GoTo Cancel_Operations
                End If
                '###### Check length of new full path #######
                If Len(FullPath) > 255 Then 'Checking if Fullpath is less than 255 characters
                    MsgBox "Unfortuntaly, your the full file path you have specified, " & FullPath & " exceeds windows 255 character" & _
                        "limit on full paths by " & Len(FullPath) - 255 & " Please shorten the file name and try again"
                    GoTo Input_NewName
                End If
                FileExists = 0
            Else
                GoTo Cancel_Operations 'User choosing to cancel all operations
            End If
        Else
            GoTo Cancel_Operations 'User choosing to cancel all operations
        End If
    Else
        FileOverwrite = 3 'File exists but user not asked what to do
    End If
Else
    FileExists = 0 'File does not exist
    FileOverwrite = 0
End If
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ Error Catching and Related Section \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
If 1 = 2 Then
Cancel_Operations:
    File_Overwrite = 2
ElseIf 1 = 2 Then
Exit_Error_Generic:
    MsgBox "There was an error when attempting to check if the new file being created already exists. Exiting Procedure"
    File_Overwrite = 2 'Maybe Different Number
    On Error GoTo -1
    On Error GoTo 0
End If
End Sub
