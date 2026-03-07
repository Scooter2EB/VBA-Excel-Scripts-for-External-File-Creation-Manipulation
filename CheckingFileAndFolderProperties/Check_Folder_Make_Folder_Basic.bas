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
