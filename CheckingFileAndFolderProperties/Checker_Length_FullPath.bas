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
