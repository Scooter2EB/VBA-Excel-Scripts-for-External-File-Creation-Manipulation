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
