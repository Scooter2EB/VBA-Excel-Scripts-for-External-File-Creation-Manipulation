Public Sub CopyPaste_PrintArea_NewWorkbook_WithOutAutosave()
'_________________________ Copy the Print selected area to a new workbook_____________________________
'//////////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\ _
/////// This script selects the pre-setup print area on the current worksheet(from page break view) \\\\\\\\\\\\\\\\\\\ _
////// and copys it to a new workbook created by the script. It is copied to the same columns as the \\\\\\\\\\\\\\\\\ _
////// source print area, but the start row can be adjusted using a variable in this script (default=2) \\\\\\\\\\\\\\\\ _
//////////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'******************************* General Properties and Quirks **********************************************************
    'Can CopyPaste into new file?: Yes _
        - Must be a module
    'Can be used on Apple?: Yes
'********************************** Adjustable Features/Variables *******************************************************
    '1. Row Height of Paste location _
            Description: When you are pasting the print area to a new workbook, you can adjust the start row _
            you are pasting it to. _
            Varible to adjust: StartNum _
            Possible Options: Any discrete number _
            Location in Script: Just Before Part 1
'((((((((((((((((((((((((((((((((((((((((((((((((()))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))))


Application.EnableEvents = False
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual

Dim Wrkbook_New As Workbook
Name_Wrksht_Active = ActiveSheet.Name
Name_Wrkbook_Host = ThisWorkbook.Name
StartNum = 2 'This is the row number you start your pasting to

'----------------------------------- Part 1: Setting up New File ---------------------------------------------
'################################### Section 1.1: Creating New File #############################################
Set Wrkbook_New = Workbooks.Add
'################################### Section 1.2: Paste Values #################################################
 '### Get Print Area Range
Adr_Copy = Split(Workbooks(Name_Wrkbook_Host).Worksheets(Name_Wrksht_Active).PageSetup.PrintArea, ",") 'Divides up non-contiguous ranges

For i = LBound(Adr_Copy) To UBound(Adr_Copy) 'Loops through non-contiguous ranges to copy separatly. Excel cannot copy non-contiguous ranges apparently...
    'Create Paste Range At Top of worksheet
    RowNum_PasteStart = Split(Split(Adr_Copy(i), ":")(0), "$")(2)
    RowNum_PasteEnd = Split(Split(Adr_Copy(i), ":")(1), "$")(2)
    MoveGap = RowNum_PasteStart - StartNum 'How Many numbers it must move up
    Adr_Paste = Replace(Adr_Copy(i), RowNum_PasteStart, StartNum)
    Adr_Paste = Replace(Adr_Paste, RowNum_PasteEnd, RowNum_PasteEnd - MoveGap)
    'Setting Ranges and copypasting
    Set Range_Copy = Workbooks(Name_Wrkbook_Host).Worksheets(Name_Wrksht_Active).Range(Adr_Copy(i))
    Set Range_Paste = Wrkbook_New.Worksheets(1).Range(Adr_Paste)
    Range_Copy.Copy
    Range_Paste.PasteSpecial xlPasteAll
    Range_Paste.Copy
    Range_Paste.PasteSpecial xlPasteValues
Next i
Application.CutCopyMode = False
'////////////////////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'#Error Section
On Error Resume Next
Set NewSheet = Wrkbook_New.Worksheets(1)
NewSheet.Name = Name_Wrksht_Active
If 1 = 2 Then
Exit_Error_Generic:
    On Error GoTo -1
    On Error GoTo 0
    MsgBox "Unfortunately there was an error when " & Err_Location & " If this issue persists, please contact the developer", vbOKOnly + vbInformation, "Error Occured"
End If
Exit_Normal:
Application.EnableEvents = True
Application.ScreenUpdating = True
Application.Calculation = xlCalculationAutomatic
End Sub
