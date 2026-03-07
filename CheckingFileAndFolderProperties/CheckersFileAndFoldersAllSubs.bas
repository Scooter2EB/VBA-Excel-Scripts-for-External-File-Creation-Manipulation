Attribute VB_Name = "CheckersFileManipulation"
'---------------------------------------------------------------------------------------------------------------
'_______________________________________ Module Name: CheckersFileManipulation __________________________________
'-----------------------------------------------------------------------------------------------------------------
'///////////////// This Module contains several subs designed to do checks on things you might need if you  \\\\\
'//////////////// were opening, saving , or creating a new file or folder.                                   \\\\\
'///////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ List Of Subroutines ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ _
'NOTE: Each sub contains better documentation on its purpose and uses. Many should work on non-windows machines
'While some of the subs appear to do things that built-in error catchers already flag, those built-in ones often do
'  not tell the end user why the error occured. By pre-checking user input, you can give them more information as to their mistake
'1. Checker_Length_FullPath
'       Checks the given file name and path to see if it exceeds windows 255 character limit.
'2. CheckIf_FileProvided_AlreadyExists
'       Checks if the given file exists at the given path. Also lets you know whether you could save that file there or if its blocked
'3. Check_If_File_Open_CurrentInstanceOnly
'       This goes through the current instance of excel (i.e. the one the script is running from) and checks if a given file name is open
'       (If you don't know what an instance of excel is, go to the next sub and read or google it)
'4. Check_If_File_Open_AllInstances
'       This goes through all instatnces of excel and checks if the given file name is open. (Unsure if MAC friendly)
'5. Check_Folder_Make_Folder_Basic
'       This checks if a given folder exists and has the option to create the folder if it does not
'6. Checker_FileName_SpecialCharacters

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ Module Declarations $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'All of these declared variables are used by Check_If_File_Open_AllInstances. That sub is a modified version of the script I got from
'    https://stackoverflow.com/questions/30363748/having-multiple-excel-instances-launched-how-can-i-get-the-application-object-f
'They have to be declared here for that sub to work. If you want to know why, check the link
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" ( _
    ByVal hwnd As LongPtr, ByVal dwId As Long, riid As Any, ppvObject As Object) As Long
Private Declare PtrSafe Function FindWindowExA Lib "user32" ( _
    ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
    ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
Private Sub Checker_Length_FullPath(FullPath, Display_Msg, Len_Excess)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'____________________________________ Checking the Full File path length ____________________________________
'////// This script checks the given file name and path to see if it exceeds windows limit of 255.
'////// If so, can display message and sends back variable indicating excess length
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'************************************* General Properties and Quirks ***********************************************
    'Can direct paste script into regular/sheet sub: Yes/Yes
    'Can Toggle message: Yes
    'Required Inputs:1 hard, 1 soft
    'Can use on Apple: Yes
    'Built in Error Handling: None
    'User cancel handling: None
'*******************************************************************************************************************
'(((((((((((((((((((((((((((((((((( Required Variables to Pass Through ))))))))))))))))))))))))))))))))))))))))))
    'Variable Name: FullPath
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: Full file name and path (although any string, including a blank will not be an issue)
        'Output Value: Unchanged
        'Description: Used to check the length of the file name and path
    'Variable Name: Display_Msg
        'Type: Byte
        'Min Declaration Level: Sub
        'Input Value: 1 or 0. Can also leave undefined and will default to 0
        'Output Value: Same as input
        'Description: This variable lets the script know whether or not to display the message that you have exceed _
                        the filepath limit for windows. 0 means display, 1 means do not display
    'Variable Name: Len_Excess
        'Type: Byte
        'Min Declaration Level: Sub
        'Input Value: None
        'Output Value: 1 or 0.
        'Description: Records whether the full file path exceeds the limit for windows. A 1 means it did. A 0 means it did not. _
            Used by calling script to make further decisions
If Len(FullPath) > 255 Then
    Len_Excess = 1
    If Display_Msg = 0 Then
        MsgBox "Unfortuntaly, your the full file path you have specified, " & FullPath & " exceeds windows 255 character" & _
           "limit on full paths by " & Len(FullPath) - 255 & ". You will have to manually change the file name or path to ensure any future actions " & _
           "relying on this file will work."
    End If
End If
End Sub
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
Sub Check_If_File_Open_CurrentInstanceOnly(Check_FullPath, Is_Wrkbook_Open)
'_____________________________ Checks if the file is currently open _________________________________
'/////////////// Checks the current instance of excel for any file matching the provided file information \\\\\\\\\
'//////////////  and returns a variable letting the calling sub know \\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'/////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Should work on Non-Windows PCs
'********************************* Variables ***************************************************************
    'Variable Name: Check_FullPath
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: The Full file path and name
        'Output Value: Same
        'Description: This varibale will consist of the name and path of the file to be checked
    'Variable Name: Is_Workbook_Open
        'Type: Byte
        'Input Value: NA
        'Output Value: True or False
        'Description: Lets you know if the file is open. True means open, False means its not.

'Got from https://www.exceldome.com/solutions/check-if-workbook-is-open/

On Error Resume Next
FileNo = FreeFile()
Open Check_FullPath For Input Lock Read As #FileNo
Close FileNo

ErrorNo = Err

On Error GoTo 0

Select Case ErrorNo
Case 0: Is_Wrkbook_Open = False
Case 70: Is_Wrkbook_Open = True
Case Else: Error ErrorNo
End Select

End Sub
Private Sub Check_If_File_Open_AllInstances(Check_FileName, Is_Wrkbook_Open, Wrkbook_Open)
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'____________________________________ Check and Retrieve workbook object if open ____________________________________
'////// This script loops through all excel instances to see if the given workbook is open. If so, it returns the workbook object _
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&& Definitions/Technical Terms &&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&&
'An Instance of excel: This refers to an version of the excel application being open, rather than a workbook being open. _
    If you open excel, open a workbook using excel, then from that workbook you open another workbook, they are both open _
    under the same instance of the excel application. If you now go to through the windows file explorer (just looking through _
    you harddrive) and open an excel file from there, it should be on a different instance of excel. You cannot copy+paste _
    as easily between them, they have their own versions of VBA (Hit alt+f11 and in project window you see all WB on that instance). _
    If you check the task manager (hit start then type Task manager and open), you will see two instances of Excel, and if you _
    expand them, one of them should have two workbooks open on them. _
    This is important to understand as VBA has the native ability to loop through all workbooks in one instance of excel, _
    but not through instances of excel. Hence why this script has complex declarations.

'************************************* General Properties and Quirks ***********************************************
    'Can direct paste script into sub: Kinda, if you work for it. Only one passthrough variable requires input, and it is just a file name and path _
                                        Requires modual/sheet level declarations. See error handling for more difficulties
    'Can use on Apple: Unlikely. I don't think the required ,modual/sheet declarations are apple friendly
    'Built in Error Handling: None. I recommend a special one in the calling sub which will jump to message asking the user _
                              if the file is open
'Modual/Sheet level DECLARATIONS:
'Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" ( _
'    ByVal hwnd As LongPtr, ByVal dwId As Long, riid As Any, ppvObject As Object) As Long
'Private Declare PtrSafe Function FindWindowExA Lib "user32" ( _
'    ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
'    ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr
'*******************************************************************************************************************
'(((((((((((((((((((((((((((((((((( Required Variables to Pass Through ))))))))))))))))))))))))))))))))))))))))))
    'Variable Name: Check_FileName
        'Type: String
        'Min Declaration Level: Sub
        'Input Value: The file name you wish to see open. Includes the file extension (same format as produced by ThisWorkbook.name
        'Output Value: Same as input
        'Description: Stores the name of the file you wish to check is open
    'Variable Name: Is_Wrkbook_Open
        'Type: byte
        'Min Declaration Level: N/A
        'Input Value: Not Required
        'Output Value: 0 if workbook not open. 1 if it is
        'Description: Tells calling sub whether workbook is open or not
    'Variable Name: Wrkbook_Open
        'Type: Variant/Object/Workbook
        'Min Declaration Level: N/A
        'Input Value: None
        'Output: If the workbook is open, the workbook object. If not open, Empty
        'Description: Stores the workbook object if the workbook is open. Only way to ensure you can get it if open in a different instance
'((((((((((((((((((((((((((((((((((((((((((((((((((((((()))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))

'######################################### Section 1: Create Collection of Open Excel Instances #########################

Is_Wrkbook_Open = 0
If Var_Slash = "" Then: Var_Slash = "\" 'Setting the Slash variable for files
'Came From https://stackoverflow.com/questions/30363748/having-multiple-excel-instances-launched-how-can-i-get-the-application-object-f
Dim Excl As Application
Dim guid&(0 To 3), acc As Object, hwnd, hwnd2, hwnd3
guid(0) = &H20400
guid(1) = &H0
guid(2) = &HC0
guid(3) = &H46000000

Set GetExcelInstances = New Collection 'Creating a collection of open excel instances
Do
    hwnd = FindWindowExA(0, hwnd, "XLMAIN", vbNullString)
    If hwnd = 0 Then: Exit Do
    hwnd2 = FindWindowExA(hwnd, 0, "XLDESK", vbNullString)
    hwnd3 = FindWindowExA(hwnd2, 0, "EXCEL7", vbNullString)
    If AccessibleObjectFromWindow(hwnd3, &HFFFFFFF0, guid(0), acc) = 0 Then
        GetExcelInstances.Add acc.Application
    End If
Loop
'############ End of section from Website
'####################################### Section 2: Loop through excel instances to check for workbook ##################

For Each Excl In GetExcelInstances 'loop through excel instances
    'Debug.Print "Handle: " & Excl.ActiveWorkbook.FullName 'loop through workbooks in xcel instances
    
    For Each Wrkbook In Excl.Workbooks
        If Wrkbook.Name = Check_FileName Then
            'Debug.Print Wrkbook.Name
            Is_Wrkbook_Open = 1
            Set Wrkbook_Open = Wrkbook
        End If
    Next
Next

End Sub

Sub Check_Folder_Make_Folder_Basic(FolderPath, Optional Toggle_MakeFolder As Byte = 0, Optional Toggle_ErrCatch As Byte = 0)
'_________________________ Check if a folder exists and attempt to make it (Basic) ______________________________________________
'//////////////// When given the folder path, this sub checks if it exists. If not, it can attempt \\\\\\\\\\\\\\\\\\\\\\\\\ _
/////////////// to make a new one or not depending on what you tell it to do (default is). It has \\\\\\\\\\\\\\\\\\\\\\\\\\ _
///////////// error catching and a message (default on), but you can turn that off and have your calling script handle it \\\ _
//////////// This is basic as it does NOT have the ability for the user to enter or select a new folder \\\\\\\\\\\\\\\\\\\\\ _
////////////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'*******************************************************************************************************************
'(((((((((((((((((((((((((((((((((( Variables Passed Through and/or Returned ))))))))))))))))))))))))))))))))))))))))))
    'Variable Name: FolderPath
        'Type: String
        'Min Declaration Level: Calling Sub
        'Input Value: The Full Folderpath. (ex. C:\Program Files)
        'Output Value: ?
        'Description: The full folder path. does not have to end with a slash but it can
    'Variable Name: Toggle_MakeFolder
        'Type: String
        'Min Declaration Level: Calling Sub
        'Input Value: 0 or 1
        'Output Value: Input Value
        'Description: 0 (default) attempts to make the folder if it cannot be found. 1 skips this. If this equals 1 AND the folder _
                      does not exist, it will set Toggle_ErrCatch to 5.
    'Variable Name: Toggle_ErrCatch
        'Type: String
        'Min Declaration Level: Calling Sub
        'Input Value: 0 or 1
        'Output Value: 2 if folder existed, 3 if new folder created,4 if no folder provided, _
                        5 for certain Non-exist Stituation (See Toggle_MakeFolder), input value if error (because if you turn off error catching, I cannot modify the value to let you know)
        'Description: 0 (Default) turns error catching on, so if folder is not found and cannot be created, script goes to a message _
                      for the user and sends this variable back as a 1. If 1, the calling subs error catching will handle the issue


If Len(FolderPath) = 0 Then 'If no folder path was provided
    If Toggle_MakeFolder = 0 Then: MsgBox "No folder path provided", vbOKOnly, "No Folder Path"
    Toggle_ErrCatch = 4
    Exit Sub
End If

If Toggle_ErrCatch = 0 Then 'Turn on error catching if requested
    On Error Resume Next
End If

FolderPathCheck = Dir(FolderPath, vbDirectory) 'check if folder exists
If FolderPathCheck = "" Then 'this will be a blank if the folder does not exist
    On Error GoTo No_Permission
    If Toggle_MakeFolder = 0 Then
        Err_Text = "Unfortunately, the custom Folder path '" & FolderPath & "' does not currently exist on this " & _
                   "computer AND excel does not have permission to create said folder."
        MkDir (FolderPath)
        Toggle_ErrCatch = 3: Exit Sub
    Else
        Err_Text = "Unfortunately, the custom Folder path '" & FolderPath & "' does not currently exist on this " & _
                  "computer."
        If Toggle_ErrCatch = 0 Then 'If error catching is turned on, shows error message. Otherwise skips it
            Toggle_ErrCatch = 5: GoTo No_Permission
        Else
            Toggle_ErrCatch = 5: Exit Sub
        End If
    End If
Else 'if the folder exists
    Toggle_ErrCatch = 2
    Exit Sub
End If

If 1 = 2 Then
No_Permission:
    On Error GoTo 0
    NoSave = MsgBox(Err_Text, vbOKOnly + vbInformation, "No Backup Folder Available")
End If

End Sub
 
Private Sub Checker_FileName_SpecialCharacters(Check_FileName, Has_Special, Display_Msg1)
'_________________________ Check if a file name has special characters ______________________________________________
'//////////////// Checks if the given file name contains illegal characters and sends back a flag to \\\\\\\\\\\\\\\\\\\\\\
'/////////////// the calling sub. Has a toggleable error message with ability for user to submit a new name \\\\\\\\\\\\\\\\
'////////////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'*******************************************************************************************************************
'(((((((((((((((((((((((((((((((((( Variables Passed Through and/or Returned ))))))))))))))))))))))))))))))))))))))))))
    'Variable Name: Check_FileName
        'Type: String
        'Min Declaration Level: Calling Sub
        'Input Value: The proposed file name. Does not need extension
        'Output Value: Input Value
        'Description: Any string the user wants to make into a file name
    'Variable Name: Has_Special
        'Type: Byte
        'Min Declaration Level: NA
        'Input Value: NA
        'Output Value: 0, or 2
        'Description: 0 means no special characters. 1 Means there is. 2 Means the user had the option to change the file name but cancelled.
    'Variable Name: Display_Msg1
        'Type: Byte
        'Min Declaration Level: Calling Sub
        'Input Value: 0 or 1
        'Output Value: Input Value
        'Description: Lets script know whether to display its own message and give the user a chance to correct the name.
                      '0=message off; 1=On

SpecialCharacters = "/\:*?<>|"
Special_Chr_Start:
Has_Special = 0
If Check_FileName <> "" Then
    For x = 1 To 8
        If InStr(1, Check_FileName, Mid(SpecialCharacters, x, 1), vbTextCompare) <> 0 Then
            GoTo Special_Chr
        End If
    Next x
Else
    GoTo Special_Chr
End If
        
If 1 = 2 Then
Special_Chr:
    Has_Special = 1
    If Display_Msg1 = 1 Then
        Check_FileName = InputBox("The file name you have given, " & Check_FileName & " is either blank or has a special " & _
                       "character (" & SpecialCharacters & "). Please type a new name in the box below and click Ok", "Naming Error")
        If Check_FileName = "" Then
            Has_Special = 2
            Exit Sub
        End If
        GoTo Special_Chr_Start
    End If
End If

End Sub

