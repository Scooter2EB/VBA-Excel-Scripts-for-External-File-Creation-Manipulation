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
