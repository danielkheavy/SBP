Attribute VB_Name = "Module10"
Option Explicit

Public Function ExportRecordSetToExcel(ByVal v_rstData As ADODB.Recordset, _
                                       ByVal v_strFileName As String, _
                                       Optional ByVal v_strReportName As String = vbNullString, _
                                       Optional ByVal v_strWSheetName As String = vbNullString) As Boolean

    '*******************************************************************
    '   Name        :   ExportRecordsetToExcel
    '   Purpose     :   Exports the supplied recordset to Excel
    '   Parameters  :   v_rstData       : Recordset to export
    '                   v_strFileName   : Excel filename
    '                   v_strReportName : Report filename (Optional)
    '                   v_strWSheetName : Worksheet name  (Optional)
    '   Returns     :   TRUE if recordset exported, FALSE if not
    '   Author      :   Jeff Valdon
    '   Date        :   Jan 2006
    '*******************************************************************

    'Declare local variables
    Dim blnReturn       As Boolean

    Dim intAnswer       As VbMsgBoxResult

    Dim objExcel        As Excel.Application

    Dim intSheets       As Integer

    Dim objWBook        As Excel.Workbook

    Dim objWSheet       As Excel.Worksheet

    Dim intFieldCount   As Integer

    Dim intRow          As Integer

    Dim intCol          As Integer

    Dim intIndex        As Integer

    Dim objRange        As Excel.Range

    'Declare local constants
    Const FUNCTION_NAME As String = "ExportRecordsetToExcel"

    On Error GoTo ErrorHandler

    'Set default return value
    blnReturn = False
   
    'Does the file alreay exist?
    If Dir$(v_strFileName) > vbNullString Then
    
        'File exists
        intAnswer = MsgBox("Archivo con el nombre " & v_strFileName & " Ya EXiste" & vbcrlf & "Desea Reescribir?", vbYesNo Or vbDefaultButton1 Or vbQuestion, App.Title)
    
        If intAnswer = vbNo Then
        
            'User wishes to cancel the output
            MsgBox "Excel Grabacion cancelado", vbInformation, App.Title
            
            'Exit function
            ExportRecordSetToExcel = blnReturn
            Exit Function
            
        Else
        
            'Delete the file
            Kill v_strFileName
        
        End If
        
    End If
    
    'Create the Excel object
    Set objExcel = New Excel.Application
    
    'Set Excel options
    With objExcel
        intSheets = .SheetsInNewWorkbook
        .SheetsInNewWorkbook = 1
        .Visible = False

    End With
    
    'Create the workbook
    Set objWBook = objExcel.Workbooks.Add
    
    'Reference the first worksheet
    Set objWSheet = objWBook.Worksheets(1)
    
    'Get the field count
    intFieldCount = v_rstData.Fields.count
    
    'Initialise the row and column counters
    intRow = 1
    intCol = 1
    
    'Create the column headers
    For intIndex = 0 To intFieldCount - 1
        objWSheet.Cells(intRow, intIndex + 1) = v_rstData.Fields(intIndex).Name
    Next
    
    'Select the column headers
    Set objRange = objWSheet.Range(objWSheet.Cells(intRow, intCol), objWSheet.Cells(intRow, intFieldCount))
    
    'Format the column headers
    With objRange.Cells
        .Font.bold = True
        .Interior.ColorIndex = 15
        .Interior.Pattern = xlSolid

    End With
    
    'Increment the row counter
    intRow = intRow + 1
    
    'Add the recordset data
    objWSheet.Cells(intRow, intCol).CopyFromRecordset v_rstData
    
    'Tidy up the worksheet
    With objWSheet
        .columns.AutoFit
        .PageSetup.CenterFooter = "Page &P of &N"
        .PageSetup.CenterHorizontally = True
        .PageSetup.Orientation = xlLandscape

    End With
    
    'Add borders to the data
    With objWSheet.UsedRange.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic

    End With
    
    'Has a report name been supplied?
    If v_strReportName <> vbNullString Then
        
        'Name the report
        objWSheet.PageSetup.CenterHeader = v_strReportName
        
    End If
                
    'Has a worksheet name been supplied?
    If v_strWSheetName <> vbNullString Then
        
        'Name the worksheet
        objWSheet.Name = v_strWSheetName
    
    End If
    
    'Save the workbook
    objWBook.SaveAs v_strFileName
        
    'Reset the Excel default
    objExcel.SheetsInNewWorkbook = intSheets
    
    'All OK
    blnReturn = True

CleanExit:

    On Error Resume Next

    'Destroy the Range
    If Not objRange Is Nothing Then
        Set objRange = Nothing

    End If
    
    'Destroy the Worksheet
    If Not objWSheet Is Nothing Then
        Set objWSheet = Nothing

    End If
    
    'Destroy the Workbook
    If Not objWBook Is Nothing Then
        objWBook.Close
        Set objWBook = Nothing

    End If
    
    'Destroy the Excel object
    If Not objExcel Is Nothing Then
        objExcel.Quit
        Set objExcel = Nothing

    End If

    ExportRecordSetToExcel = blnReturn

    Exit Function

ErrorHandler:

    'Display the error
    MsgBox "Error No " & Err.Number & vbcrlf & Err.Description & vbcrlf & "Has occured in " & FUNCTION_NAME & vbcrlf & "Please contact Technical Support", vbCritical, App.Title
    
    Resume CleanExit
    
End Function

