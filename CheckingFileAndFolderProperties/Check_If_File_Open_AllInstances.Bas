'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$ Module Declarations $$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'This sub is a modified version of the script I got from
'    https://stackoverflow.com/questions/30363748/having-multiple-excel-instances-launched-how-can-i-get-the-application-object-f
'They have to be declared here for that sub to work. If you want to know why, check the link
Private Declare PtrSafe Function AccessibleObjectFromWindow Lib "oleacc" ( _
    ByVal hwnd As LongPtr, ByVal dwId As Long, riid As Any, ppvObject As Object) As Long
Private Declare PtrSafe Function FindWindowExA Lib "user32" ( _
    ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
    ByVal lpszClass As String, ByVal lpszWindow As String) As LongPtr

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
