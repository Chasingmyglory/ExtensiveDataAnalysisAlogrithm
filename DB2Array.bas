Attribute VB_Name = "DB2Array"
Sub b_Create_DB_Plan_array_New()
''**************************************************************************************************************************
''  createPlanArray_16.20
''  b_Create_DB_Plan_array_New : this code is for creating DB2 from the planning filr from the perspective of Array datastructure
''  Created : 05-09-2023
''  @Author : Shamim, Mohammad Raisul Hasan
''**************************************************************************************************************************


Dim SourceMS As Workbook

DateOrder = Application.International(xlDateOrder)
DateSeparator = Application.International(xlDateSeparator)

If DateOrder = 0 Then
    SystemDateFormat = "MM" & DateSeparator & "dd" & DateSeparator & "yyyy"
ElseIf DateOrder = 1 Then
    SystemDateFormat = "dd" & DateSeparator & "MM" & DateSeparator & "yyyy"
ElseIf DateOrder = 2 Then
    SystemDateFormat = "yyyy" & DateSeparator & "MM" & DateSeparator & "dd"
End If

Call update_values_RD

'''Define the specific information of this version
With MS_Version_Input
    .Height = 200
    .Width = 600
    '.StartPosition = 0 ' Manual positioning
    .Top = Application.Top + (Application.Height - .Height) / 2
    .Left = Application.Left + (Application.Width - .Width) / 2
End With

MS_Version_Input.Show

Current_Version = MS_Version_Input.TB_Current_Version_Name.Text
Current_VersionYear = MS_Version_Input.TB_Version_Year.Text
Current_VersionMonth = MS_Version_Input.TB_Version_Month.Text
Current_VersionType = MS_Version_Input.TB_Version_Type.Text
Current_VersionStatus = 1
Answer_UserForm = MS_Version_Input.Answer_TB.Text

MS_Version_Input.TB_Current_Version_Name.Text = ""
MS_Version_Input.TB_Version_Year.Text = "Select Year"
MS_Version_Input.TB_Version_Month.Text = "Select Month"
MS_Version_Input.TB_Version_Type.Text = "Select Version Type"
MS_Version_Input.Answer_TB.Text = ""

If Answer_UserForm = 2 Then
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
End If
    

Current_VersionDate = Date
Current_DataStatus = "PlanningFile_Data"
Current_Control = "Done"
Current_RowPrevVersion = "-"
Current_User = Application.UserName

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
MacroStartTime = Timer

'''Run an update of the latest Project Information - Always
Call a_Project_Info_Update

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

''''For testing - specific Planning File
'Set SourceMS = Workbooks("MS11-2020 PlanningFile_Final_NoFormats_20201112_DBv2_Test.xlsb")


''' Set Source and Target File
'''Select Planning File
Set DialogBox = Application.FileDialog(msoFileDialogOpen)  'Set Source and Source File and only allow the user to select one file
DialogBox.Title = "Open the Planning File"
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False

intChoice = Application.FileDialog(msoFileDialogOpen).Show 'make the file dialog visible to the user

If intChoice <> 0 Then 'determine what choice the user made

Source_MS = Application.FileDialog(msoFileDialogOpen).SelectedItems(1) 'get the file path selected by the user
Else
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
End If


Set SourceMS = Workbooks.Open(Source_MS)

Source_MS = SourceMS.Name

Set SourceMS = Workbooks(Source_MS)

Set Source_Sheet = SourceMS.Sheets("04_Planning Template")
Set MS_Planning_Sheet = ThisWorkbook.Sheets("MS_Planning")
Set MS_Project_Info_Sheet = ThisWorkbook.Sheets("MS_Project_Info")
Set Milestones_Sheet = ThisWorkbook.Sheets("MS_Milestones")
Set MS_Versions_Sheet = ThisWorkbook.Sheets("MS_Versions")
Set TransferConflicts_Sheet = ThisWorkbook.Sheets("Transfer_Conflicts")


ThisWorkbook.Activate

''' Set information from Source
    'On Error Resume Next ( Filtering is commented out as array takes all the data)
    'Source_Sheet.ShowAllData
    'On Error GoTo 0
''''

''''***********Start of the Columns

    ' Planning file's relevant data's first row
    Source_First_Row_Header = Source_Sheet.Range("A1").End(xlDown).row
    Source_First_Row = Source_Sheet.Range("A1").End(xlDown).row + 1
   
    
    'Set lastCell = Source_Sheet.Cells(Source_Sheet.Rows.Count, 1).Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    'If Not lastCell Is Nothing Then
    'Source_Last_Row = lastCell.row
    'End If
    
    'Doing the filter to get the best data
    Source_Sheet.AutoFilterMode = False

    ' Planning file relevant data's last row
    Source_Last_Row = Source_Sheet.Cells(Source_Sheet.Rows.Count, 1).End(xlUp).row

    ' Planning file's relevant data's first row
    Source_Last_Col = Source_Sheet.Cells((Source_First_Row - 1), Source_Sheet.Columns.Count).End(xlToLeft).Column
    
    Source_MSIndex_Col = Application.Match("Index", Source_Sheet.Range("21:21"), 0)
    Source_ItemCode_Col = Application.Match("Item Code", Source_Sheet.Range("21:21"), 0)
    
    ' Plan start column
    Source_PlanStart_Col = Application.Match("Week", Source_Sheet.Range("1:1"), 0)
    
    ' Plan end column
    Source_PlanEnd_Col = Application.Match("Week", Source_Sheet.Range("1:1"), 0) + Application.WorksheetFunction.CountIf(Source_Sheet.Range("1:1"), "Week") - 1
    
    ' Previous Years not in MS column
    Source_PrevYears_Col = Application.Match("Previous years not in MS", Source_Sheet.Range("21:21"), 0)
    Source_SupSplit_Col = Application.Match("Supplier split", Source_Sheet.Range("21:21"), 0)
    Source_ProdSemiCompl_Col = Application.Match("Production" & Chr(10) & "Semi/Complete", Source_Sheet.Range("21:21"), 0)
    Source_IFRS15_Col = Application.Match("IFRS15", Source_Sheet.Range("21:21"), 0)
    
    ' Factory Column
    Source_Factory_Col = Application.Match("Factory", Source_Sheet.Range("21:21"), 0)
    Source_Comment_Col = Application.Match("Comment", Source_Sheet.Range("21:21"), 0)
    Source_ChangedBy_Col = Application.Match("Changed by", Source_Sheet.Range("21:21"), 0)
    Source_ChangedDate_Col = Application.Match("Last change date", Source_Sheet.Range("21:21"), 0)
        
'''************************End of the column declaration

    
    'as sorting in array is time and memory intensive hence sorting is  done directly on excel
    'before loading data to the array
    With Source_Sheet.Sort
        .SortFields.Add Key:=Cells(Source_First_Row, Source_MSIndex_Col), Order:=xlAscending
        .SortFields.Add Key:=Cells(Source_First_Row, Source_ItemCode_Col), Order:=xlAscending
        
        .SetRange Range(Cells(Source_First_Row, 1), Cells(Source_Last_Row, Source_Last_Col))
        .Header = xlNo
        .Apply
        .SortFields.Clear
    End With
    
    Dim entireData As Variant
    entireData = Source_Sheet.Range(Source_Sheet.Cells(Source_First_Row, 1), Source_Sheet.Cells(Source_Last_Row, Source_Last_Col)).value

    ' Define the columns you want to extract
    Dim sourceColIndices As Collection
    Set sourceColIndices = New Collection
    
    ' Add the columns of interest to the colIndices
    sourceColIndices.Add Source_MSIndex_Col
    sourceColIndices.Add Source_PrevYears_Col
    sourceColIndices.Add Source_SupSplit_Col
    sourceColIndices.Add Source_ProdSemiCompl_Col
    sourceColIndices.Add Source_IFRS15_Col
    sourceColIndices.Add Source_ItemCode_Col
    sourceColIndices.Add Source_Factory_Col
    For i = Source_PlanStart_Col To Source_PlanEnd_Col
        sourceColIndices.Add i
    Next i
    sourceColIndices.Add Source_Comment_Col
    sourceColIndices.Add Source_ChangedBy_Col
    sourceColIndices.Add Source_ChangedDate_Col
    
    'Timer
    MacroStartTime = Timer
    ' Loading final Source Planning data
    Dim finalSourceData As Variant
    finalSourceData = ExtractColumns(entireData, sourceColIndices)
    
    Dim finalSourceComments() As String
    ReDim finalSourceComments(UBound(finalSourceData, 1), UBound(finalSourceData, 2))
    
    Dim p As Long, q As Long, excelCol As Long
    For p = 1 To UBound(finalSourceData, 1)
        For q = 1 To UBound(finalSourceData, 2)
            ' Translate c back to Excel column index
            excelCol = sourceColIndices(q)
            
            If Not Source_Sheet.Cells(Source_First_Row + p - 1, excelCol).Comment Is Nothing Then
                finalSourceComments(p, q) = Source_Sheet.Cells(Source_First_Row + p - 1, excelCol).Comment.Text
            Else
                finalSourceComments(p, q) = ""
            End If
        Next q
        
    Next p
    
    
    
    
    
    'Putting the filters back
    Source_Sheet.Range("A" & Source_First_Row_Header & ":" & Source_Sheet.Cells(Source_Last_Row, Source_Sheet.Columns.Count).End(xlToLeft).Address).AutoFilter

    
    MS_DB_First_Row = 2
    MS_DB_Last_Row = MS_Planning_Sheet.Cells(MS_Planning_Sheet.Rows.Count, 1).End(xlUp).row
    
     ' Now source planning date should be updated
    Dim planDateArray As Variant
    planDateArray = Source_Sheet.Range(Source_Sheet.Cells(19, Source_PlanStart_Col), Source_Sheet.Cells(20, Source_PlanEnd_Col)).value

   
    
    'We need MS_Project Info sheets info in another array so we first we are identifying the
    'columns
    MS_ProjInfo_First_Row = 2
    
    MS_ProjInfo_Last_Row = MS_Project_Info_Sheet.Cells(MS_Project_Info_Sheet.Rows.Count, 1).End(xlUp).row
    MS_ProjInfo_Last_Col = MS_Project_Info_Sheet.Cells((MS_ProjInfo_First_Row - 1), MS_Project_Info_Sheet.Columns.Count).End(xlToLeft).Column
    
    MS_ProjInfo_MS_Index_Col = Application.Match("MS_Index", MS_Project_Info_Sheet.Range("1:1"), 0)
    MS_ProjInfo_WTG_Col = Application.Match("WTG_Total", MS_Project_Info_Sheet.Range("1:1"), 0)
    MS_ProjInfo_VersionID_Col = Application.Match("Version_ID", MS_Project_Info_Sheet.Range("1:1"), 0)
    
    
    Dim msPInfoColIndices As Collection
    Set msPInfoColIndices = New Collection
    
    msPInfoColIndices.Add MS_ProjInfo_MS_Index_Col
    msPInfoColIndices.Add MS_ProjInfo_WTG_Col
    msPInfoColIndices.Add MS_ProjInfo_VersionID_Col
    
    
    ' Total MS_Project_Info_Data
    Dim entireMS_Project_Info_Data As Variant
    entireMS_Project_Info_Data = MS_Project_Info_Sheet.Range(MS_Project_Info_Sheet.Cells(MS_ProjInfo_First_Row, 1), MS_Project_Info_Sheet.Cells(MS_ProjInfo_Last_Row, MS_ProjInfo_Last_Col)).value
    
    ' Loading final MS_Project_Info sheet's data for our requirement
    Dim finalMSPInfoData As Variant
    finalMSPInfoData = ExtractColumns(entireMS_Project_Info_Data, msPInfoColIndices)
     
  
    
    ''MS_TransferConflict_Sheet
    ' the challenge here is, this array will be filled up one by one
    'but now after thought just realized that that will not be that hard
    'the hard thing would be to understand nested loops inside one another
    MS_TransConflicts_First_Row = 2
    MS_TransConflicts_Last_Row = TransferConflicts_Sheet.Cells(TransferConflicts_Sheet.Rows.Count, 1).End(xlUp).row
    MS_TransConflicts_VersionID_Col = Application.Match("Version_ID", TransferConflicts_Sheet.Range("1:1"), 0)
    MS_TransConflicts_MSIndex_Col = Application.Match("MS_Index", TransferConflicts_Sheet.Range("1:1"), 0)
    MS_TransConflicts_ItemCode_Col = Application.Match("Item_Code", TransferConflicts_Sheet.Range("1:1"), 0)
    MS_TransConflicts_RowPlanningFile_Col = Application.Match("Row_PlanningFile", TransferConflicts_Sheet.Range("1:1"), 0)
    MS_TransConflicts_ColPlanningFile_Col = Application.Match("Col_PlanningFile", TransferConflicts_Sheet.Range("1:1"), 0)
    MS_TransConflicts_Comment_Col = Application.Match("Comment", TransferConflicts_Sheet.Range("1:1"), 0)
    
    
    '' Milestone sheet put in another array
    Dim MilestoneSheetData As Variant
    milestoneSheetFirstRow = 2
    milestoneSheetLastRow = Milestones_Sheet.Cells(Milestones_Sheet.Rows.Count, 1).End(xlUp).row
    milestoneSheetFirstColumn = 1
    milestoneSheetLastColumn = Milestones_Sheet.Cells((milestoneSheetFirstRow - 1), Milestones_Sheet.Columns.Count).End(xlToLeft).Column
    
    
    MilestoneSheetData = Milestones_Sheet.Range(Milestones_Sheet.Cells(milestoneSheetFirstRow, 1), Milestones_Sheet.Cells(milestoneSheetLastRow, milestoneSheetLastColumn)).value
    
    
    
    ''''Remove all information of plans, since this is always creating the first version
    MS_Planning_Sheet.Rows(MS_DB_First_Row & ":" & (MS_DB_Last_Row + 1)).Delete
    MS_Versions_Sheet.Rows(2).Delete
    TransferConflicts_Sheet.Rows(MS_TransConflicts_First_Row & ":" & (MS_TransConflicts_Last_Row + 1)).Delete
    
    '''Add the information of the version in the Project Info Tab too
    With MS_Project_Info_Sheet
        .Range(.Cells(MS_ProjInfo_First_Row, MS_ProjInfo_VersionID_Col), .Cells(MS_ProjInfo_Last_Row, MS_ProjInfo_VersionID_Col)) = Current_Version
    End With
    
    
    
    MS_MSVersions_VersionID_Col = Application.Match("Version_ID", MS_Versions_Sheet.Range("1:1"), 0)
    MS_MSVersions_VersionDate_Col = Application.Match("Version_Date", MS_Versions_Sheet.Range("1:1"), 0)
    MS_MSVersions_VersionYear_Col = Application.Match("Version_Year", MS_Versions_Sheet.Range("1:1"), 0)
    MS_MSVersions_VersionMonth_Col = Application.Match("Version_Month", MS_Versions_Sheet.Range("1:1"), 0)
    MS_MSVersions_VersionType_Col = Application.Match("Version_Type", MS_Versions_Sheet.Range("1:1"), 0)
    MS_MSVersions_VersionStatus_Col = Application.Match("Version_Status", MS_Versions_Sheet.Range("1:1"), 0)
    MS_MSVersions_VersionUploadBy_Col = Application.Match("Version_Upload_By", MS_Versions_Sheet.Range("1:1"), 0)
    
    '''Add the information of the version in the MS_Version tab
    With MS_Versions_Sheet
        .Cells(2, MS_MSVersions_VersionID_Col) = Current_Version
        .Cells(2, MS_MSVersions_VersionDate_Col) = Current_VersionDate
        .Cells(2, MS_MSVersions_VersionYear_Col) = Current_VersionYear
        .Cells(2, MS_MSVersions_VersionMonth_Col) = Current_VersionMonth
        .Cells(2, MS_MSVersions_VersionType_Col) = Current_VersionType
        .Cells(2, MS_MSVersions_VersionStatus_Col) = 1
        .Cells(2, MS_MSVersions_VersionUploadBy_Col) = Current_User
    End With
    
    
      ' info for Version ID info for array
    MS_MSVersions_First_Row = 2
    MS_MSVersions_Last_Row = MS_Versions_Sheet.Cells(1, 1).End(xlUp).row
    MS_MSVersions_First_Column = 1
    MS_MSVersions_Last_Column = MS_Versions_Sheet.Cells((MS_MSVersions_First_Row - 1), MS_Versions_Sheet.Columns.Count).End(xlToLeft).Column
    
    Dim versionInfo As Variant
    versionInfo = MS_Versions_Sheet.Range(MS_Versions_Sheet.Cells(MS_MSVersions_First_Row, 1), MS_Versions_Sheet.Cells(MS_MSVersions_Last_Row, MS_MSVersions_Last_Column)).value
    
    
    
    '' the first loop to go through the rows of the planning file range
    firstRowIndexSouce = LBound(finalSourceData, 1)
    lastRowIndexSource = UBound(finalSourceData, 1)
    lastColindexSource = UBound(finalSourceData, 2)
    
    'for the transfer conflict sheet data I am using this array apporach
    Dim transferData As Variant
    Dim dataRow As Long
    ReDim transferData(1 To lastRowIndexSource, 1 To 6) 'Assuming you have 6 columns to store
    dataRow = 1
    
    Dim ms_Planning_Array As Variant
    Dim msPlanningDataArrayRow As Long
    ReDim ms_Planning_Array(1 To 1200000, 1 To 17)
    msPlanningDataArrayRow = 1
    
    
    
    SourceMS.Close savechanges:=False
    Dim counter As Long
    counter = 0
    Dim cR As Long
    cR = firstRowIndexSouce
    ' Check the dimensions of the finalSourceData array
    'Debug.Print "finalSourceData Rows (1st dimension): " & UBound(finalSourceData, 1)
    'Debug.Print "finalSourceData Columns (2nd dimension): " & UBound(finalSourceData, 2)

    ' Check the range of loop variables
    'Debug.Print "firstRowIndexSouce: " & firstRowIndexSouce
    'Debug.Print "lastRowIndexSource: " & lastRowIndexSource

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ROW Starts ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim lastFoundPosition As Long
    lastFoundPosition = LBound(ms_Planning_Array, 1)
    Dim lastFilledRow As Long ' last filled row in MS_Planning Sheet
    lastFilledRow = 0
    For cR = firstRowIndexSouce To lastRowIndexSource
            
            
           
            Current_MS_Index = finalSourceData(cR, 1)
        
            
            Current_ItemCode = finalSourceData(cR, 6)
            
            
            Current_ProdSemiCompl = finalSourceData(cR, 4)
            
            
            Current_IFRS15 = finalSourceData(cR, 5)
            
            
            Current_Factory = finalSourceData(cR, 7)
            
            ' Check if Current_Factory is empty
             If IsEmpty(Current_Factory) Or Current_Factory = "" Then
            ' Loop through MilestoneSheetData to find a match
                For j = LBound(MilestoneSheetData, 1) To UBound(MilestoneSheetData, 1)
                    If MilestoneSheetData(j, 1) = Current_ItemCode Then
                        Current_Factory = MilestoneSheetData(j, 5)
                        
                        Exit For ' Exit the loop once the match is found
                    
                    End If
                 Next j
            End If
            
            
            Current_MSComment = finalSourceData(cR, 216)
            Current_ChangedBy = finalSourceData(cR, 217)
            Current_LastChangeDate = finalSourceData(cR, 218)
             
            ' Calculate Sum_Items_Row
            Dim col As Integer
            Sum_Items_Row = 0
            
            ' Check if the value at finalSourceData(i, 2) is numeric before adding
            If IsNumeric(finalSourceData(cR, 2)) Then
                Sum_Items_Row = finalSourceData(cR, 2)
            End If
            
            For col = 8 To 215
                ' Check if the value at finalSourceData(i, col) is numeric before adding
                If IsNumeric(finalSourceData(cR, col)) Then
                    Sum_Items_Row = Sum_Items_Row + finalSourceData(cR, col)
                End If
            Next col
            
            ' Get the date
            For k = LBound(MilestoneSheetData, 1) To UBound(MilestoneSheetData, 1)
                If MilestoneSheetData(k, 1) = Current_ItemCode Then
                    UseDate = MilestoneSheetData(k, 4)
                    Exit For
                End If
            Next k
            
            
            If Sum_Items_Row = 0 Then
                    transferData(dataRow, 1) = Current_Version
                    transferData(dataRow, 2) = Current_MS_Index
                    transferData(dataRow, 3) = Current_ItemCode
                    transferData(dataRow, 4) = cR + 21    'current item row
                    transferData(dataRow, 5) = "-"
                    transferData(dataRow, 6) = "No planning in PlanningFile"
                    dataRow = dataRow + 1
            Else
                           Dim r As Long
                           r = 2
                           Source_PrevYears_Col = r
                           CurrentItem_Col = Source_PrevYears_Col
                           Source_Plan_Start_Col = 8
                               
                           If PrevRow_MS_Index = Current_MS_Index And PrevRow_ItemCode = Current_ItemCode Then
                                   Current_ItemNoID = Current_ItemNoID
                           Else
                                   Current_ItemNoID = 1
                           End If
                           Decimal_Prev_Item = 0
                            'Current_ItemNoID = 1
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++Column Starts +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                           For CurrentItem_Col = Source_PrevYears_Col To (lastColindexSource - 3)
                               
                               If CurrentItem_Col = Source_PrevYears_Col Or CurrentItem_Col >= Source_Plan_Start_Col Then
                                       Current_Item_Value = finalSourceData(cR, CurrentItem_Col)
                                       
                                       If Not IsNumeric(Current_Item_Value) Then
                                               transferData(dataRow, 1) = Current_Version
                                               transferData(dataRow, 2) = Current_MS_Index
                                               transferData(dataRow, 3) = Current_ItemCode
                                               transferData(dataRow, 4) = cR + 21    'current item row
                                               transferData(dataRow, 5) = "-"
                                               transferData(dataRow, 6) = "Non numeric item in plan"
                                               dataRow = dataRow + 1
                                       Else
                                               If Current_ItemCode = "05_1_0" Or Current_ItemCode = "05_1_1" Or Current_ItemCode = "05_1_2" Or Current_ItemCode = "05_1_3" Then
                                                       Current_Item_Value = Round(3 * Current_Item_Value, 0)
                                               End If
                                               
                                              
                                               
                                               
                                               Current_Item_Value = Current_Item_Value + Decimal_Prev_Item
                                              
                                               Round_Item_Value = CustomFloor(Current_Item_Value)
                                              
                                               Decimal_Prev_Item = Round(Current_Item_Value - Round_Item_Value, 15)
                                              
                                               Current_Item_Value = CLng(Round_Item_Value)
                                               
                                               If Current_Item_Value <= 0 Then
                                                       'go to next column
                                               Else
                                                       
                                                       
                                                       
                                                       Total_Lines_Num = Current_Item_Value
                                                       If IsEmpty(finalSourceComments(cR, CurrentItem_Col)) Or Trim(finalSourceData(cR, CurrentItem_Col)) = "" Then
                                                               Current_MSIndividualComment = ""
                                                       Else
                                                               Current_MSIndividualComment = finalSourceComments(cR, CurrentItem_Col)
                                                       End If
                                                       
                                                       
                                                       
                                                       
                                                       
                                                       ' Calculating the start and end rows
                                                       Current_Row_Start = lastFilledRow + 1
                                                       Current_Row_End = Current_Row_Start + Total_Lines_Num - 1
                                                       
                                                       Dim Date_Column_Maatched_With_Planning_Col As Long
                                                       Date_Column_Maatched_With_Planning_Col = CurrentItem_Col - 7
                                                       
                                                       If Source_PrevYears_Col = CurrentItem_Col Then                                                                                                                                                                                                     'CurrentItem_Col = m + 6
                                                               FirstDate_MS = planDateArray(1, 1)
                                                               Current_MSDate = DateSerial(Month:=12, Day:=15, Year:=Year(FirstDate_MS - 15))             'End If
                                                       Else                                                                                                                                        'Next m
                                                               If UseDate = 20 Then
                                                                       Current_MSDate = planDateArray(2, Date_Column_Maatched_With_Planning_Col) - 2
                                                               Else
                                                                       Current_MSDate = planDateArray(1, Date_Column_Maatched_With_Planning_Col)
                                                               End If 'UseDate = 20
                                                       End If 'CurrentItem_Col = Source_PrevYears_Col
                                                       
                                                     
                                                       For row = Current_Row_Start To Current_Row_End
                                                           ms_Planning_Array(row, 1) = Current_MS_Index
                                                           ms_Planning_Array(row, 2) = Current_ItemCode
                                                            ms_Planning_Array(row, 3) = Current_ItemNoID
                                                           ' Handling ItemNoID
                                                           'If row = Current_Row_Start Then
                                                                   'ms_Planning_Array(row, 3) = Current_ItemNoID
                                                           'Else
                                                                   'ms_Planning_Array(row, 3) = ms_Planning_Array(row - 1, 3) + 1
                                                           'End If
                                                   
                                                           ms_Planning_Array(row, 4) = 1
                                                           ms_Planning_Array(row, 5) = Current_ProdSemiCompl
                                                           ms_Planning_Array(row, 6) = Current_IFRS15
                                                           ms_Planning_Array(row, 7) = Current_Factory
                                                           ms_Planning_Array(row, 8) = Current_MSDate
                                                           ms_Planning_Array(row, 9) = Current_MSIndividualComment
                                                           ms_Planning_Array(row, 10) = Current_MSComment
                                                           ms_Planning_Array(row, 11) = Current_ChangedBy
                                                           ms_Planning_Array(row, 12) = Current_LastChangeDate
                                                           ms_Planning_Array(row, 13) = Current_Version
                                                           ms_Planning_Array(row, 14) = Current_VersionDate
                                                           ms_Planning_Array(row, 15) = Current_DataStatus
                                                           ms_Planning_Array(row, 16) = Current_Control
                                                           ms_Planning_Array(row, 17) = Current_RowPrevVersion
                                                            
                                                            Current_ItemNoID = Current_ItemNoID + 1
                                                       Next row
                                                       
                                                       PrevRow_MS_Index = Current_MS_Index
                                                       PrevRow_ItemCode = Current_ItemCode
                                                       
                                                        lastFilledRow = Current_Row_End
                                               End If 'Current_Item_Value <= 0
                                               
                                       End If ' Not IsNumeric(Current_Item_Value)
                               End If 'CurrentItem_Col = Source_PrevYears_Col Or CurrentItem_Col >= Source_Plan_Start_Col
                               
                               
                        Next CurrentItem_Col ' column iteration
                        
                        
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++Column Looping ends ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                            NextRow_MS_Index = finalSourceData(cR + 1, 1)
                            NextRow_ItemCode = finalSourceData(cR + 1, 6)
                            
                            If NextRow_MS_Index = Current_MS_Index And NextRow_ItemCode = Current_ItemCode Then
                                    'do nothing
                            Else
                                    Current_Item_Count = 0
                                    For ms_row = lastFoundPosition To lastFilledRow
                                        If ms_Planning_Array(ms_row, 1) = Current_MS_Index And _
                                            ms_Planning_Array(ms_row, 2) = Current_ItemCode And _
                                            ms_Planning_Array(ms_row, 5) <> "Semi" Then
                                            Current_Item_Count = Current_Item_Count + ms_Planning_Array(ms_row, 4)
                                            ' Update last found position
                                            lastFoundPosition = ms_row
                                            'Exit For
                                        End If ' if ms_Planning_Array(ms_row, 1) = Current_......
                                    Next ms_row
                                    
                                    Dim Current_ProjInfo_Row As Long
                                    
                                    ' Find the row that matches Current_MS_Index in the first column of finalMSPInfoData
                                    For Current_ProjInfo_Row = LBound(finalMSPInfoData, 1) To UBound(finalMSPInfoData, 1)
                                            If finalMSPInfoData(Current_ProjInfo_Row, 1) = Current_MS_Index Then
                                                Exit For
                                            End If
                                    Next Current_ProjInfo_Row
                                    
                                    ' Get the value from the second column of the matched row
                                    Current_Item_WTG_Total = finalMSPInfoData(Current_ProjInfo_Row, 2)
        
                                    ' Apply the same conditions to modify the value if necessary
                                    If Current_ItemCode = "05_1_0" Or Current_ItemCode = "05_1_1" Or Current_ItemCode = "05_1_2" Or Current_ItemCode = "05_1_3" Then
                                            Current_Item_WTG_Total = Round(3 * Current_Item_WTG_Total, 0)
                                    Else
                                    ' Use the chance to round up decimals
                                            Current_Item_WTG_Total = Round(Current_Item_WTG_Total, 0)
                                    End If
                                    'If Current_Item_Count = Current_Item_WTG_Total Then
                                            '''Do Nothing
                                    If Current_Item_Count > Current_Item_WTG_Total Then
                                            '''Too many items in the plan
                                            transferData(dataRow, 1) = Current_Version
                                            transferData(dataRow, 2) = Current_MS_Index
                                            transferData(dataRow, 3) = Current_ItemCode
                                            transferData(dataRow, 4) = cR + 21    'current item row
                                            transferData(dataRow, 5) = "-"
                                            transferData(dataRow, 6) = "Too many items in plan"
                                            dataRow = dataRow + 1
                            
                                    ElseIf Current_Item_Count < Current_Item_WTG_Total Then
                                            '''Too few items in the plan
                                            transferData(dataRow, 1) = Current_Version
                                            transferData(dataRow, 2) = Current_MS_Index
                                            transferData(dataRow, 3) = Current_ItemCode
                                            transferData(dataRow, 4) = cR + 21    'current item row
                                            transferData(dataRow, 5) = "-"
                                            transferData(dataRow, 6) = "Too few items in plan"
                                            dataRow = dataRow + 1
                            
                                    End If 'Current_Item_Count = Current_Item_WTG_Total Then
                            End If ' if NextRow_MS_Index = Current_MS_Index And.......
            
            End If 'Sum_Items_Row = 0
            'Debug.Print "cR" & cR
        
    Next cR ' row iteration
 'Debug.Print "CR Last " & cR
    
                With MS_Planning_Sheet
                    '''Count Items in Transfer_Conflict Sheet
                    MS_Planning_Sheet_Used_Rows = .UsedRange.Rows.Count
                    If MS_Planning_Sheet_Used_Rows = 1 Then
                        MS_Planning_Sheet_Next_Rows = 2
                    Else
                        MS_Planning_Sheet_Next_Rows = MS_Planning_Sheet.Cells(MS_Planning_Sheet.Rows.Count, 1).End(xlUp).row + 1

                    End If
                End With
        ' Write the data to the worksheet in one go
        
        MS_Planning_Sheet.Cells(MS_Planning_Sheet_Next_Rows, 1).Resize(lastFilledRow, UBound(ms_Planning_Array, 2)).value = ms_Planning_Array


        '''An empty row could be assumed as a conflict/mistake
                With TransferConflicts_Sheet
                    '''Count Items in Transfer_Conflict Sheet
                    TransferConflicts_Used_Rows = .UsedRange.Rows.Count
                    If TransferConflicts_Used_Rows = 1 Then
                        MS_TransConflicts_Next_Row = 2
                    Else
                        MS_TransConflicts_Next_Row = TransferConflicts_Sheet.Cells(TransferConflicts_Sheet.Rows.Count, 1).End(xlUp).row + 1  '.Range("A1").End(xlDown).row + 1
                    End If
                End With
        ' Write the data to the worksheet in one go
        TransferConflicts_Sheet.Cells(MS_TransConflicts_Next_Row, 1).Resize(dataRow - 1, UBound(transferData, 2)).value = transferData


Application.StatusBar = False
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
TimeGone = (Timer - MacroStartTime) / 86400
TimeGone = Format(TimeGone, "hh:mm:ss")
MsgBox "Finished the planning transfer in " & TimeGone
         
            
             
End Sub

