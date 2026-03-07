Attribute VB_Name = "BackingUpFiles"
'--------------------------------------------------------------------------------------------------------------- _
______________________________________________ Module Name: Backing Up Files __________________________________ _
----------------------------------------------------------------------------------------------------------------- _
'//////////////////////////// This array contains scripts designed to make backup copies of your file \\\\\\\\\\\\ _
///////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ List Of Subroutines ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^ _
 1. SaveWorkbookBackup _
        This saves a backup copy of your active workbook in a folder labelled 'Backups' which it creates in the _
        files current location. It also adds the current date and a counter to the backup file name

Dim BackupCount As Byte 'Used to count the number of backups saved per session
Sub SaveWorkbookBackUp()
'/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'___________________________ Saving a Backup Copy of the Workbook ________________________
'////// This sub saves a backup copy of the open file in a directory called 'Backups' within said file's \\\\\ _
/////// existing directory. Each backup copy is named: Filename Backup <Save Date> <(Count of backups that day)> \\\ _
/////// If the Backups directory does not exist, it will be created. If it cannot be created, it the backup \\\\\ _
/////// will be saved in the file's current directory.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

'$$$$$$ Required Declarations (Module Level): Dim BackUpCount As Byte

BackupCount = BackupCount + 1 'So you can save more than one backup copy per session

Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual 'This turns off calculations if you have alot of them. They are turned back _
                                                on at the end of the script. You can remove this if desired
'%%%%%%%%%%%%%%%%%%%%%%%%% Checking for operating system %%%%%%%%%%%%%%%%%%%%%%%%%%%%
Var_Slash = "\" 'for windows
If Not Left(Application.OperatingSystem, 1) Like "W" Then: Var_Slash = "/"

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

On Error GoTo NoSave

'%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Checking If you want to save the current workbook %%%%%%%%%%%%%%%
ShouldSave = MsgBox("Would you like to save the current file BEFORE creating a backup?", vbYesNo + vbQuestion, "Save File?")
If ShouldSave = vbYes Then
    ThisWorkbook.Save
End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%% Checking and Creating File Path %%%%%%%%%%%%%%%%%%%%
File_Name = Split(ThisWorkbook.Name, ".")
File_Ext = "." & File_Name(UBound(File_Name)) 'use the split list of the file name by '.' to ensure you get the file extension (they might have multiple '.' in the file name
File_Name = Left(ThisWorkbook.Name, Len(ThisWorkbook.Name) - Len(File_Ext))
File_Path = ThisWorkbook.path & Var_Slash & "Backups" & Var_Slash
File_PathCheck = Dir(File_Path, vbDirectory)
If File_PathCheck = "" Then
    On Error GoTo No_Permission
    MkDir (File_Path)
End If
If 1 = 2 Then
No_Permission:
    On Error GoTo -1: On Error GoTo NoSave
    NoSave = MsgBox("Unfortunately, the custom file path '" & File_Path & "' does not currently exist on this " & _
                  "computer AND excel does not have permission to create said folder. Until one of these " & _
                  "two conditions are changed, all backups will be saved in the same folder as this file.", vbOKOnly + vbInformation, "No Backup Folder Available")
                  
    File_Path = ThisWorkbook.path & Var_Slash
End If
'%%%%%%%%%%%%%%%%%%%%%%%%%%%% Creating a Filename and checking the path length %%%%%%%%%%%%%%%%%%%%%%%%%%%%

BackupName = File_Path & File_Name & " Backup " & Format(CDate(Now()), "dd-mmm-yyyy") & " (" & BackupCount & ")" & File_Ext
If Len(BackupName) >= 255 Then
    TooLong = MsgBox("Issue: The Current file path and file name for you back up, which is: " & vbNewLine & vbNewLine & _
                   BackupName & vbNewLine & vbNewLine & "exceeds windows built in length limit of 255 characters." & vbNewLine & _
                   "The backup will now be saved in the folder of the current directory instead", vbOKOnly + vbInformation, "File Path and Name Too Long")
    
    File_Path = ThisWorkbook.path & Var_Slach
    BackupName = File_Path & File_Name & " Backup " & Format(CDate(Now()), "dd-mmm-yyyy") & " (" & BackupCount & ")" & File_Ext
End If
'%%%%%%%%%%%%%%%%%%%%%%%%% Saving Backup %%%%%%%%%%%%%%%%%%%%%%%%%%%%
ThisWorkbook.SaveCopyAs FileName:=BackupName

If 1 = 2 Then
NoSave:
    NoBackUp = MsgBox("For Some Reason, generating a backupcopy failed. It might have something to do with " & _
                    "folder permissions.", vbOKOnly + vbInformation, "Error: Could Not Create Backup")
    BackupCount = BackupCount - 1
Else
    BackUp_msg = MsgBox("Congratulations! You have created a backup file! I am so proud of you! Your backup copy " & _
                  " is saved at " & BackupName, vbOKOnly + vbInformation, "You ain't neva gonna lose yo work now")
End If
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub


