Attribute VB_Name = "m_main"
Option Explicit

Const Debug_Mode                As Boolean = True

Const Ws_TM3StockReportSheetName                As String = "Stock report"
Const Ws_TM3StockReportAddMaterialsSheetName    As String = "Add materials"
Const stockReportMaterialTypeColumn             As String = "A"
Const stockReportMaterialCodeColumn             As String = "C"
Const stockReportMaterialBatchColumn            As String = "D"
Const stockReportCommentsColumn                 As String = "S"
Const stockReportDeletedMaterialColumn          As String = "U"
Const columnWithSafetyStockEntered              As String = "I"
Const columnStock                               As String = "J"
Const columnFreeStockWarehouse                  As String = "K"

Const rowToStart                                As Long = 6

Public Wb_Current               As Workbook
Public Ws_MaterialsInput        As Worksheet
Public Ws_MaterialsInput_Name   As String
Public Ws_Output                As Worksheet
Public Ws_Output_Name           As String
Public WS_helpSheet             As Worksheet
Public Ws_settingsSheet         As Worksheet

Public Ws_TM3StockReportSheet   As Worksheet

Public conn                     As ADODB.Connection



Public DbServerAddress          As String
Public DbName                   As String
Public AdUserName               As String

Sub readme()


''' ******************************************************************
''' Author:  Paramonov M.
''' Version: 0.1
''' Date:    17.01.2021
''' Updates:
'''      +
''' ******************************************************************
    

End Sub

Sub init()

    If Not Debug_Mode Then On Error GoTo ErrHandler

    ' Praise the Omnissiah!

    Set Wb_Current = ThisWorkbook
    Ws_MaterialsInput_Name = "Materials Input"
    
    Set WS_helpSheet = Wb_Current.Sheets("helpSheet")
    Set Ws_settingsSheet = Wb_Current.Sheets("settingsSheet")
    
    DbServerAddress = "erw-dev.spd.ru\INS01"
    DbName = "SPD_MRP"
    AdUserName = GetADUsername
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error in sub init: " & Err.Description)
    
End Sub

Sub LoadTM3StockReport()

    If Not Debug_Mode Then On Error GoTo ErrHandler

    Call init

    Call ImprovePerformance(True)
    
    Application.EnableEvents = False
    

    Dim rangeToInsertData               As String
    Dim sqlLoadTM3Report                As String
    Dim chosenMaterialType              As String
    Dim sqlInitMaterialTypesCombobox    As String
    Dim headerRange                     As String
    Dim columnWithComments              As String
    Dim columnWithSafetyStockEntered    As String
    Dim columnWithDeliveryDates         As String
    Dim columnFreeStockWarehouse        As String
    Dim lr                              As Long
    Dim rowsCounter                     As Long
    Dim firstRow                        As Long
    
    headerRange = "A5:U5"
    columnWithComments = stockReportCommentsColumn
    columnWithSafetyStockEntered = "I"
    columnWithDeliveryDates = "Q"
    columnFreeStockWarehouse = "K"
    firstRow = 6
    
    rangeToInsertData = "A5"
    sqlInitMaterialTypesCombobox = "SELECT [Type] FROM dbo.TM3_WSStockReport_MasterMaterialTypes ORDER BY [Department], [Type];"
    sqlLoadTM3Report = "EXEC [TM3].[DBSUB_WSStockReport_LoadReport]"
    
    If ThisWorkbook.Sheets(Ws_TM3StockReportSheetName).ComboBox1.value <> "" Then
        chosenMaterialType = ThisWorkbook.Sheets(Ws_TM3StockReportSheetName).ComboBox1.value
        sqlLoadTM3Report = sqlLoadTM3Report & " " & StringToMSSQLFormat(chosenMaterialType)
    End If
    
    Set conn = CreateConnection(DbServerAddress, DbName)
    
    Set Ws_TM3StockReportSheet = ThisWorkbook.Sheets(Ws_TM3StockReportSheetName)
    
    ' Turn off protection.
    If Ws_TM3StockReportSheet.ProtectContents = True Then Ws_TM3StockReportSheet.Unprotect
    
    ' Init combobox.
    Call InitComboBoxFromSqlQuery(ThisWorkbook.Sheets(Ws_TM3StockReportSheetName).ComboBox1, sqlInitMaterialTypesCombobox, conn)
    
    ' Clear sheet before load data.
    Ws_TM3StockReportSheet.Rows("5:10000").EntireRow.Delete
    
    ' Load data on sheet.
    Call RunSQLSelect(Ws_TM3StockReportSheetName, sqlLoadTM3Report, DbServerAddress, DbName, ThisWorkbook, rangeToInsertData)
    
    ' Get last row.
    lr = GetBorders("LR", ThisWorkbook.Sheets(Ws_TM3StockReportSheetName).Name, ThisWorkbook)
    
    ' Format result.
    With Ws_TM3StockReportSheet
    
        .Range(headerRange).EntireColumn.AutoFit
        .Range(headerRange).AutoFilter
        .Range(headerRange).Font.Bold = True
        .Columns(columnWithComments).ColumnWidth = 40
        .Columns(columnWithDeliveryDates).ColumnWidth = 40
    
        
    Call indicateSafetyStockEntered(ThisWorkbook.Sheets("settingsSheet").Range("safetyStockMode").value)

    End With
    
    Call FreezeHeader(5)
    
    If Ws_TM3StockReportSheet.ProtectContents = False Then
        Ws_TM3StockReportSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
            False, AllowFormattingCells:=True, AllowFormattingColumns:=True, _
            AllowFormattingRows:=True, AllowInsertingHyperlinks:=True, AllowSorting:= _
            True, AllowFiltering:=True, AllowUsingPivotTables:=True
    End If

    Call MsgBox("Finish!")

    Call ImprovePerformance(False)
    
    Call hideColumns
    
    Application.EnableEvents = True
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error in sub LoadTM3StockReport: " & Err.Description)

End Sub

Sub addDataValidationTypesToAddMaterialsSheet(Optional column As String = "B")

    If Not Debug_Mode Then On Error GoTo ErrHandler

    Call init
    
    Dim rowToStart                          As Long
    Dim rowToEnd                            As Long
    Dim rowsCounter                         As Long
    Dim materialTypedForDropDownList        As String
    Dim sqlGetMaterialTypedForDropDownList  As String
    Dim resultRs                            As ADODB.Recordset
    Dim conn                                As ADODB.Connection
    
    rowToStart = 6
    rowToEnd = 100
    sqlGetMaterialTypedForDropDownList = "EXEC dbo.DBSUB_TM3_WSStockReport_GetAllTypesInRow"
    
    Set conn = CreateConnection(DbServerAddress, DbName)
    Set resultRs = conn.Execute(sqlGetMaterialTypedForDropDownList)
    
    materialTypedForDropDownList = resultRs.Fields(0)
    
    With ThisWorkbook.Sheets(Ws_TM3StockReportAddMaterialsSheetName).Range(column & rowToStart & ":" & column & rowToEnd).Validation
    
        
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Formula1:=materialTypedForDropDownList
    
    End With

Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error in sub addDataValidationTypesToAddMaterialsSheet: " & Err.Description)

End Sub

Sub addMMType()

    If Not Debug_Mode Then On Error GoTo ErrHandler

    Call init
    
    Dim typeToAdd As String
    Dim sqlAddMaterialType As String
    sqlAddMaterialType = "EXEC dbo.DBSUB_TM3_WSStockReport_AddMaterialType "
    
    typeToAdd = InputBox("Enter a MM type.")
    
    If Trim(typeToAdd) = "" Then
        Exit Sub
    End If
    
    sqlAddMaterialType = sqlAddMaterialType & StringToMSSQLFormat(typeToAdd)
    
    Call RunSQLInsert(sqlAddMaterialType, DbServerAddress, DbName)
    
    
    MsgBox ("Type added.")

Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error in sub addMMType: " & Err.Description)

End Sub

Sub addMMsToMasterMaterialsList()

    If Not Debug_Mode Then On Error GoTo ErrHandler

    Call init

    Dim lr                              As Long
    Dim rowsCounter                     As Long
    Dim rowNumbetToStart                As Long
    Dim sqlAddMMsToMasterMaterialsList  As String
    
    rowNumbetToStart = 6
    

    With ThisWorkbook.Sheets(Ws_TM3StockReportAddMaterialsSheetName)
    
        lr = GetBorders("LR", .Name, ThisWorkbook)
        
        For rowsCounter = rowNumbetToStart To lr
        
            sqlAddMMsToMasterMaterialsList = "EXEC dbo.DBSUB_TM3_WSStockReport_AddMaterialToTMaster "
            sqlAddMMsToMasterMaterialsList = sqlAddMMsToMasterMaterialsList & .Range("A" & rowsCounter) & ", " & StringToMSSQLFormat(.Range("B" & rowsCounter))
            Call RunSQLInsert(sqlAddMMsToMasterMaterialsList, DbServerAddress, DbName)
        
        Next rowsCounter
    
    End With
    
    MsgBox ("Materials added to Master.")
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error in sub addMMsToMasterMaterialsList: " & Err.Description)

End Sub

Sub markUnmarkMaterialAsDeleted()

    If Not Debug_Mode Then On Error GoTo ErrHandler

    Call init

    Dim sqlMarkMaterialAsDeleted    As String
    Dim selectedRow                 As Variant
    Dim selectedMaterialCode        As String
    Dim sqlQuery                    As String
    Dim rs                          As ADODB.Recordset
    Dim conn                        As ADODB.Connection
    
    sqlMarkMaterialAsDeleted = "EXEC dbo.[DBSUB_TM3_WSStockReport_MarkMaterialAsDeleted] "
    selectedRow = selection.Row
    
    Set conn = m_common.CreateConnection(DbServerAddress, DbName, , 5)
    
    With ThisWorkbook.Sheets(Ws_TM3StockReportSheetName)
    
        .Unprotect
    
        ' Insert value in DB.
        selectedMaterialCode = .Range(stockReportMaterialCodeColumn & CStr(selectedRow)).value
        sqlMarkMaterialAsDeleted = sqlMarkMaterialAsDeleted & StringToMSSQLFormat(selectedMaterialCode)
    
        Call RunSQLInsert(sqlMarkMaterialAsDeleted, DbServerAddress, DbName)


        ' Update value on sheet.
        sqlQuery = "SELECT [Deleted] FROM [dbo].[TM3_WSStockReport_MasterMaterialsList] WHERE [Material code] = " & selectedMaterialCode
        Set rs = conn.Execute(sqlQuery)
        .Range(stockReportDeletedMaterialColumn & selectedRow).value = rs.Fields(0).value
        
        .Protect
    
    End With
    
    MsgBox ("Finished.")
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error in sub markUnmarkMaterialAsDeleted: " & Err.Description)

End Sub
Sub saveMaterialGroups()

    If Not Debug_Mode Then On Error GoTo ErrHandler

    Dim Wb                       As Workbook
    Dim ws                       As Worksheet
    Dim ws_Name                  As String:  ws_Name = Ws_TM3StockReportSheetName
    Dim ws_LR                    As Long
    Dim rowsCounter              As Long
    Dim rowNumberToStart         As Long:    rowNumberToStart = 6
    Dim sqlInsertQuery           As String
    Dim sqlInsertSSQuery         As String
    Dim sqlValues                As String
    Dim MMCodeColumn             As String: MMCodeColumn = "C"
    Dim MMGroupColumn            As String: MMGroupColumn = "B"
    
    Set Wb = ThisWorkbook
    Set ws = Wb.Sheets(ws_Name)
    ws_LR = GetBorders("LR", ws_Name, Wb)
    sqlInsertQuery = "EXEC dbo.[DBSUB_TM3_WSStockReport_UpdateUserMaterialGroup] "
    
    Call init
    
    With ws
            
        ' Insert data.
        For rowsCounter = rowNumberToStart To ws_LR
            ' Update Comments.
            If Trim(.Range(MMGroupColumn & rowsCounter).value) <> "" Then
                sqlValues = .Range(MMCodeColumn & rowsCounter).value & "," & _
                                    StringToMSSQLFormat(.Range(stockReportMaterialTypeColumn & rowsCounter).value, True) & "," & _
                                    StringToMSSQLFormat(.Range(MMGroupColumn & rowsCounter).value, True)
                Debug.Print sqlInsertQuery & sqlValues
                Call RunSQLInsert(sqlInsertQuery & sqlValues, DbServerAddress, DbName)
            End If

        Next rowsCounter
    
    End With
    
    MsgBox ("Data updated!")
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error in sub saveComments: " & Err.Description)

End Sub

Sub saveCommentsAndSafetyStocks()

    If Not Debug_Mode Then On Error GoTo ErrHandler

    Dim Wb                       As Workbook
    Dim ws                       As Worksheet
    Dim ws_Name                  As String:  ws_Name = Ws_TM3StockReportSheetName
    Dim ws_LR                    As Long
    Dim rowsCounter              As Long
    Dim rowNumberToStart         As Long:    rowNumberToStart = 6
    Dim sqlInsertQuery           As String
    Dim sqlInsertSSQuery         As String
    Dim sqlValues                As String
    Dim MMCodeColumn             As String: MMCodeColumn = "C"
    Dim CommentColumn            As String: CommentColumn = stockReportCommentsColumn
    Dim safetyStockEnteredColumn As String: safetyStockEnteredColumn = "I"
    Dim batchColumn              As String: batchColumn = "D"
    
    Set Wb = ThisWorkbook
    Set ws = Wb.Sheets(ws_Name)
    ws_LR = GetBorders("LR", ws_Name, Wb)
    sqlInsertQuery = "EXEC dbo.[DBSUB_TM3_WSStockReport_UpdateUserComments] "
    sqlInsertSSQuery = "EXEC dbo.[DBSUB_TM3_WSStockReport_UpdateMaterialData] "
    
    Call init
    
    With ws
        
        ' Check data is correct.
        For rowsCounter = rowNumberToStart To ws_LR
            If Trim(.Range(MMCodeColumn & rowsCounter).value) = "" Then
                MsgBox ("Data error: Material code can not be empty!")
                Exit Sub
            End If
        Next rowsCounter
        
    
        ' Insert data.
        For rowsCounter = rowNumberToStart To ws_LR
            ' Update Comments.
            If Trim(.Range(CommentColumn & rowsCounter).value) <> "" Then
                sqlValues = .Range(MMCodeColumn & rowsCounter).value & "," & _
                            StringToMSSQLFormat(.Range(stockReportMaterialTypeColumn & rowsCounter).value, True) & "," & _
                            StringToMSSQLFormat(.Range(CommentColumn & rowsCounter).value, True) & "," & _
                            StringToMSSQLFormat(.Range(batchColumn & rowsCounter).value, True)
                Debug.Print sqlInsertQuery & sqlValues
                Call RunSQLInsert(sqlInsertQuery & sqlValues, DbServerAddress, DbName)
            End If
            
            ' Update SS.
            If Trim(.Range(safetyStockEnteredColumn & rowsCounter).value) <> "" Then
                sqlValues = .Range(MMCodeColumn & rowsCounter).value & "," & NumberToMSSQLFormat(.Range(safetyStockEnteredColumn & rowsCounter).value)
                Call RunSQLInsert(sqlInsertSSQuery & sqlValues, DbServerAddress, DbName)
            End If
            
        Next rowsCounter
    
    End With
    
    MsgBox ("Comments are updated!")
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error in sub saveComments: " & Err.Description)

End Sub

Function gatherForInput(ws_Name As String, Optional column As String = "A", Optional Wb As Workbook) As String

    If Not Debug_Mode Then On Error GoTo ErrHandler
    
    Dim inputRangeAddress       As String
    Dim cll                     As Range
    Dim materialsInputString    As String
    
    Dim ws As Worksheet
    Set ws = Wb.Sheets(ws_Name)
    
    inputRangeAddress = column & "2:" & column & GetBorders("LR", ws_Name, Wb)
    
    If WorksheetFunction.CountA(Ws_MaterialsInput.Range(column & "2:" & column & "1048576")) > 0 Then
    
        For Each cll In ws.Range(inputRangeAddress)
            If cll.value <> "" Then
                materialsInputString = materialsInputString & Trim(CStr(cll.value)) & ","
            End If
        Next cll
    
        gatherForInput = StringToMSSQLFormat(Mid(materialsInputString, 1, Len(materialsInputString) - 1))
    
    Else
    
        gatherForInput = "null"
    
    End If
    
Done:
    Exit Function
ErrHandler:
    MsgBox ("Error in function gatherForInput: " & Err.Description)
    
End Function

Function CellInRange(RangeToCheck As String, WsNameToCheck As String, WbToCheck As Workbook)
    'Tests if a Cell is within a specific range.

Dim testRange As Range
Dim myRange As Range

'Set the range
Set testRange = WbToCheck.Sheets(WsNameToCheck).Range(RangeToCheck)

'Get the cell or range that the user selected
Set myRange = selection

If WbToCheck.Sheets(WsNameToCheck).Name = Application.ActiveSheet.Name Then
    'Check if the selection is inside the range.
    If Intersect(testRange, myRange) Is Nothing Then
        'Selection is NOT inside the range.
    
        CellInRange = False
    
    Else
        'Selection IS inside the range.
    
        CellInRange = True
    
    End If
End If

End Function

Sub UnderlineTheCases(rowToStart As Long, columnToCheck As String, underlineStart As String, UnderlineEnd As String, ws_Name As String, Optional Wb As Workbook)

    Set Wb = ThisWorkbook
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(ws_Name)
    
    

    With ws

        Dim counter     As Long
        Dim SavedValue  As String
        Dim lr          As Long
        
        lr = GetBorders("LR", ws_Name, Wb)
        
        SavedValue = ws.Range(columnToCheck & CStr(rowToStart))
        For counter = rowToStart To lr
        
            If CStr(ws.Range(columnToCheck & CStr(counter)).value) <> CStr(SavedValue) Then
                SavedValue = .Range(columnToCheck & CStr(counter)).value
                ws.Range(underlineStart & CStr(counter) & ":" & UnderlineEnd & CStr(counter)).Borders(xlEdgeTop).LineStyle = xlContinuous
            End If
            
        Next counter
    
    End With

End Sub

Sub FreezeHeader(Optional RowNumber As Integer = 2)

    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitRow = RowNumber
        .FreezePanes = True
    End With

End Sub

Sub test()

    Call FreezeHeader(5)

End Sub

Function CreateWorkbook(NewWorkbookName As String, TemplateSheetName As String, Optional sql As String = "", Optional rangeToInsert As String = "", Optional WorkbookExtension As String = ".xlsx", Optional closeWorkbook As Boolean = False) As Workbook

    If Not Debug_Mode Then On Error GoTo ErrHandler
    
    Call ImprovePerformance(True)
    
    Dim targetWb        As Workbook
    Dim currentWb       As Workbook
    
    Dim wsToWriteData   As String
    
    'Dim conn    As New ADODB.connection
    Dim rs      As New ADODB.Recordset
    
    
    Dim choiceFileDialog            As Integer
    Dim resultExtension             As Integer
    Dim shtIndex                    As Integer
    
    Dim chosenExtension             As String
    Dim pathToFile                  As String
    Dim fullPathToSave              As Variant
    
    Dim defaultSheetName            As String
    Dim HeaderLastRowTemplate       As Long
    Dim RFPDocLastRow               As Long
    Dim LongDescLastRow             As Long
    
    Dim LongDescRowsCounter         As Long
    Dim RFPDocRowsCounter           As Long
    Dim RFPDocMMColumnNumber        As Long: RFPDocMMColumnNumber = 15
    Dim RFPDocLongDescColumnNumber  As Long: RFPDocLongDescColumnNumber = 4
    Dim ValuesForQuery              As String
    
        
    Set currentWb = ThisWorkbook
                                                         
    fullPathToSave = Application.ActiveWorkbook.Path & "\" & NewWorkbookName & WorkbookExtension
    
    'Interrupt sub if user pressed Cancel or X in FileDialog.
    If fullPathToSave = False Then
        MsgBox "Path is not chosen."
        Exit Function
    End If

    chosenExtension = RxMatch(fullPathToSave, "\.[\w]+$")       'Get chosen excel workbook's extension.
    'fullPathToSave = RxReplace(fullPathToSave, "\.[\w]+$", "")  'Cut extension from full path to saving workbook.
    
    Select Case chosenExtension
        Case ".xlsx"
            'You want to save Excel 2007-2016 file
            resultExtension = xlWorkbookDefault
        Case ".xlsb"
            'You want ta save Excel 2007-2016 BINARY file
            resultExtension = xlExcel12
    End Select
    
    ' Save new Workbook to specified folder with specified in FilDialog name.
    Workbooks.Add
    Set targetWb = ActiveWorkbook
    
    If TemplateSheetName <> "" Then
            
        ThisWorkbook.Sheets(TemplateSheetName).Copy Before:=targetWb.Sheets(1)
        targetWb.Sheets(TemplateSheetName).Visible = xlSheetVisible

        ' Detect defualt sheet
        If SheetExists("Sheet1", targetWb) = True Then
        
            targetWb.Worksheets("Sheet1").Delete
            
        ElseIf SheetExists(WS_helpSheet.Range("defaultSheetNameRus").value, targetWb) = True Then
        
            targetWb.Worksheets(WS_helpSheet.Range("defaultSheetNameRus").value).Delete
            
        End If
        
        wsToWriteData = TemplateSheetName
    Else
    
        wsToWriteData = targetWb.Sheets(1).Name
    
    End If
    
    
    If sql <> "" Then
    
        Set conn = CreateConnection(DbServerAddress, DbName)
        Set rs = conn.Execute(sql)
    
        If rs.RecordCount = 1 And rs.Fields(0) = 2 Then


            MsgBox (rs.Fields(1))
            targetWb.Close
            Exit Function

        Else
        
            targetWb.Worksheets(wsToWriteData).Range(rangeToInsert).CopyFromRecordset rs
        
        End If
    
    End If
    
   Set CreateWorkbook = targetWb
    
    
    targetWb.SaveAs FileName:=fullPathToSave, FileFormat:=resultExtension
    
    ' Save RFP doc.
    targetWb.Save
    
    If closeWorkbook = True Then: targetWb.Close
    
    Call ImprovePerformance(False)
    
Done:
    Exit Function
ErrHandler:
    MsgBox ("Error in sub CreateRFPWorkbook: " & Err.Description)
        
End Function

Sub setColumnsToHide()

    Call init
    
    Dim columnsToHide As String
        
    columnsToHide = Trim(InputBox("Enter columns to hide. Example: A,B,G,AZ." & vbNewLine & "The current selection is: " & Ws_settingsSheet.Range("columnsToHide").value))
    
    If columnsToHide = "" Then
        Exit Sub
    ElseIf RxMatch(columnsToHide, "^[A-z\,]+$") = "" Then
        MsgBox ("Wrong input! You must enter column letters delimeted by comma ( , ). Try again.")
        Exit Sub
    End If
    
    ' Trim extra comma.
    If Right(columnsToHide, 1) = "," Then columnsToHide = Left(columnsToHide, Len(columnsToHide) - 1)
    
    Ws_settingsSheet.Range("columnsToHide").value = columnsToHide

End Sub


Sub hideColumns()

    Call init

    Dim Ws_TM3StockReportSheet  As Worksheet
    Dim columnsToHide           As String
    Dim columnLetter            As Variant
    Dim colsArray               As Variant
    
    Set Ws_TM3StockReportSheet = Wb_Current.Sheets(Ws_TM3StockReportSheetName)
    
    columnsToHide = Ws_settingsSheet.Range("columnsToHide").value
    colsArray = Split(columnsToHide, ",")
    
    'Ws_TM3StockReportSheet.Columns("A:AZ").Hidden = False
    
    If columnsToHide = "" Then
        MsgBox ("No columns chosen to hide.")
        Exit Sub
    End If

    
    ' If columns hidden.
    If Ws_TM3StockReportSheet.Columns(colsArray(0)).Hidden = True Then
    
        For Each columnLetter In colsArray
        
            If Trim(columnLetter) <> "" Then
                Ws_TM3StockReportSheet.Columns(columnLetter).EntireColumn.Hidden = False
            End If
            
        Next columnLetter
        
    ' If not hidden.
    ElseIf Ws_TM3StockReportSheet.Columns(colsArray(0)).Hidden = False Then
    
        For Each columnLetter In colsArray
        
            If Trim(columnLetter) <> "" Then
                Ws_TM3StockReportSheet.Columns(columnLetter).EntireColumn.Hidden = True
            End If
        
        Next columnLetter

    
    End If


End Sub

Sub deleteMMFromMaterialList()

    If Not Debug_Mode Then On Error GoTo ErrHandler

    Call init

    Dim lr                              As Long
    Dim rowsCounter                     As Long
    Dim rowNumbetToStart                As Long
    Dim sqlAddMMsToMasterMaterialsList  As String
    
    rowNumbetToStart = 6
    

    With ThisWorkbook.Sheets(Ws_TM3StockReportAddMaterialsSheetName)
    
        lr = GetBorders("LR", .Name, ThisWorkbook)
        
        For rowsCounter = rowNumbetToStart To lr
        
            sqlAddMMsToMasterMaterialsList = "EXEC dbo.DBSUB_TM3_WSStockReport_DeleteMaterialFromTMaster "
            sqlAddMMsToMasterMaterialsList = sqlAddMMsToMasterMaterialsList & .Range("A" & rowsCounter) & ", " & StringToMSSQLFormat(.Range("B" & rowsCounter))
            Call RunSQLInsert(sqlAddMMsToMasterMaterialsList, DbServerAddress, DbName)
        
        Next rowsCounter
    
    End With
    
    MsgBox ("Materials deleted from Master.")
    
Done:
    Exit Sub
ErrHandler:
    MsgBox ("Error in sub deleteMMFromMaterialList: " & Err.Description)

End Sub

Sub indicateSafetyStockCWI_OptionButton()

    Call indicateSafetyStockEntered("CWI")

End Sub

Sub indicateSafetyStockWE_OptionButton()

    Call indicateSafetyStockEntered("WE")

End Sub

Sub indicateSafetyStockEntered(mode As String)

    Call ImprovePerformance(True)
    
    Application.EnableEvents = False

    Call init
    
    Dim columnToCheck   As String
    Dim lr              As Long
        
    Dim rowsCounter     As Long
    
    lr = GetBorders("LR", ThisWorkbook.Sheets(Ws_TM3StockReportSheetName).Name, ThisWorkbook)
    
    Select Case mode
        Case "CWI"
            columnToCheck = columnFreeStockWarehouse
            Ws_settingsSheet.Range("safetyStockMode").value = "CWI"
        Case "WE"
            columnToCheck = columnStock
            Ws_settingsSheet.Range("safetyStockMode").value = "WE"
        Case Else
            Exit Sub
    End Select
    
    With Wb_Current.Sheets(Ws_TM3StockReportSheetName)
    
        For rowsCounter = rowToStart To lr
        
            .Range(columnWithSafetyStockEntered & rowsCounter).Interior.ColorIndex = 15
        
            If .Range(columnToCheck & rowsCounter).value < .Range(columnWithSafetyStockEntered & rowsCounter).value Then
                .Range(columnWithSafetyStockEntered & rowsCounter).Interior.ColorIndex = 3
            End If
        
        Next rowsCounter
    
    End With

    Application.EnableEvents = True

    Call ImprovePerformance(False)

End Sub
