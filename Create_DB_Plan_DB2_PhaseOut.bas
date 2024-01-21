Attribute VB_Name = "Create_DB_Plan_DB2_PhaseOut"
'Public Const SqlServer = "KGHSQL61V"
Public Const DataBase_PRD = "GPO" ' =GPO;
Public Const DataBase_DEV = "GPO_DEV" ' =GPO_DEV;
Public Const SqlServer = "10.5.7.3"
Public Const Database = "GPO_DEV"
Public conn As ADODB.Connection
Public newArray As Variant
Public entireMS_Project_Info_Data_new As Variant
Public ms_Planning_Array As Variant
Public lastFilledRow As Long
Public MS_MSVersions_2D(1 To 1, 1 To 7) As Variant
Public MilestoneSheetData As Variant
Sub Create_DB_Plan_PhaseOutDB2()
''**************************************************************************************************************************
''  DB2 Phase Out
''  Create_DB_Plan_PhaseOutDB2 : this code is for creating DB2 structure in the array from the perspective of Array datastructure
''  Created : 17-10-2023
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
    .Height = 220
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

'''**********INPUT FILE IMPORT*****************************************************************
'************************************************Project info**********************************+
''' Set Source and Target File
'''Select Input File file
Set DialogBox = Application.FileDialog(msoFileDialogOpen)  'Set Target and Target File and only allow the user to select one file
DialogBox.Title = "Open the INPUT File"
Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False
intChoice = Application.FileDialog(msoFileDialogOpen).Show 'make the file dialog visible to the user
If intChoice <> 0 Then 'determine what choice the user made
InputFile = Application.FileDialog(msoFileDialogOpen).SelectedItems(1) 'get the file path selected by the user
Else
    Application.StatusBar = False
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Exit Sub
End If

Set InputFile = Workbooks.Open(InputFile)

Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False
Application.StatusBar = "Progress: 0% - Time passed: 00:00:00"

'''Set systems date format
DateOrder = Application.International(xlDateOrder)
DateSeparator = Application.International(xlDateSeparator)

If DateOrder = 0 Then
    SystemDateFormat = "MM" & DateSeparator & "dd" & DateSeparator & "yyyy"
ElseIf DateOrder = 1 Then
    SystemDateFormat = "dd" & DateSeparator & "MM" & DateSeparator & "yyyy"
ElseIf DateOrder = 2 Then
    SystemDateFormat = "yyyy" & DateSeparator & "MM" & DateSeparator & "dd"
End If

'Set InputFile = Workbooks("INPUT FILE_2_MS03-2020_FinalSalesInput_20200327.xlsb")

MacroStartTime = Timer
Set MS_Input_Sheet = InputFile.Sheets("MS_INPUT")

ThisWorkbook.Activate

''' Set information from Source
Source_First_Row = 9
Source_Last_Row = MS_Input_Sheet.Cells(Source_First_Row, 1).End(xlDown).row - 1
Source_MSIndex_Col = Application.Match("MS_Index", MS_Input_Sheet.Range("8:8"), 0)
SourceVoltageLevel_Col = Application.Match("Voltage Level", MS_Input_Sheet.Range("8:8"), 0)


entireMS_Project_Info_Data_new = MS_Input_Sheet.Range(MS_Input_Sheet.Cells(Source_First_Row, Source_MSIndex_Col), MS_Input_Sheet.Cells(Source_Last_Row, SourceVoltageLevel_Col)).value
        
' Find the number of rows in your original data
Dim numRows As Long
numRows = UBound(entireMS_Project_Info_Data_new, 1) - LBound(entireMS_Project_Info_Data_new, 1) + 1
' Find the number of columns in your original data
Dim numCols As Long
numCols = UBound(entireMS_Project_Info_Data_new, 2) - LBound(entireMS_Project_Info_Data_new, 2) + 1

' Create a new array with one additional column for the Version ID
ReDim newArray(1 To numRows, 1 To numCols)
' Fill the new array with data from the original array and the Current_Version
Dim i As Long
For i = 1 To numRows
    ' Set the Version ID column to the Current_Version
    newArray(i, 1) = Current_Version
    ' Set the other columns to the appropriate values from the original data
    newArray(i, 2) = entireMS_Project_Info_Data_new(i, 1)
    newArray(i, 3) = entireMS_Project_Info_Data_new(i, 9)
Next i

' now adding the version info to the Project Info sheet
Dim loopIndex As Long
Dim numRowsOfProjectInfoDataArray As Long

'Get the nmber of rows in the array
numRowsOfProjectInfoDataArray = UBound(entireMS_Project_Info_Data_new, 1)

'Resize the array to add a new column
ReDim Preserve entireMS_Project_Info_Data_new(1 To numRowsOfProjectInfoDataArray, 1 To UBound(entireMS_Project_Info_Data_new, 2) + 1)

'Resize the array to add a new column
For loopIndex = 1 To numRowsOfProjectInfoDataArray
    'Shift the existing columns to the right
    For jam = UBound(entireMS_Project_Info_Data_new, 2) - 1 To 1 Step -1
        entireMS_Project_Info_Data_new(loopIndex, jam + 1) = entireMS_Project_Info_Data_new(loopIndex, jam)
    Next jam
    
    'Set the value of the new column to the value of the "Current_Version" variable
    entireMS_Project_Info_Data_new(loopIndex, 1) = Current_Version
    
    
Next loopIndex

' Now newArray contains your desired data with the Version ID column filled with Current_Version

' Create a new worksheet
'Dim NewSheet As Worksheet
'Set NewSheet = Sheets.Add(After:=Sheets(Sheets.Count))

' Define the destination range on the new sheet
'Dim DestRange As Range
'Set DestRange = NewSheet.Cells(1, 1)

' Get the dimensions of the source data
'Dim numRowsPaste As Long
'Dim numColsPaste As Long
'numRowsPaste = UBound(newArray, 1) - LBound(newArray, 1) + 1
'numColsPaste = UBound(newArray, 2) - LBound(newArray, 2) + 1

' Define the destination range using the dimensions of the source data
'Set DestRange = DestRange.Resize(numRows, numCols)

' Paste the data onto the new sheet
'DestRange.value = newArray

ProgressMacro = 100 - (Source_Last_Row - Source_First_Row + 1 - CurrentItem_Row) * 100 / (Source_Last_Row - Source_First_Row + 1)
TimeGone = (Timer - MacroStartTime) / 86400
TimeGone = Format(TimeGone, "hh:mm:ss")
Application.StatusBar = "Progress: " & Round(ProgressMacro, 0) & "% - Time passed: " & TimeGone
InputFile.Close SaveChanges:=False

Application.StatusBar = False
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

MsgBox "Finished updating project information in " & TimeGone

'''''''''''''********************************************************************************

'****************************************'Planning File Update**********************************
Application.ScreenUpdating = False
Application.Calculation = xlCalculationManual
Application.EnableEvents = False

'Dim intChoice As Integer
Dim Source_MS As String
Dim Source_Sheet As Worksheet
Dim userResponse As VbMsgBoxResult
Dim wasWorkbookOpened As Boolean ' New variable to track if a new workbook open

' Ask the user if they want to use the current planning file
userResponse = MsgBox("Do you want to use this planning file?", vbYesNoCancel)


Select Case userResponse
    Case vbYes
        'Use the current planning file
        ThisWorkbook.Activate
        Set SourceMS = ThisWorkbook
    Case vbNo
        ' Show file dialog to choose a different planning file
        With Application.FileDialog(msoFileDialogOpen)
            .Title = "Select a Planning File"
            .AllowMultiSelect = False
            intChoice = .Show
            If intChoice <> 0 Then
                Source_MS = .SelectedItems(1)
                Set SourceMS = Workbooks.Open(Source_MS)
                wasWorkbookOpened = True
            Else
                ' If user cancels file dialog, exit sub
                Exit Sub
            End If
        End With
    Case vbCancel
        'if user cancels the message box, exit sub
        Exit Sub
End Select

Set Source_Sheet = SourceMS.Sheets("04_Planning Template")
    
        



'Set DialogBox = Application.FileDialog(msoFileDialogOpen)  'Set Source and Source File and only allow the user to select one file
'DialogBox.Title = "If you want to a different Planning File then select otherwise press cancel"
'Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False

'intChoice = Application.FileDialog(msoFileDialogOpen).Show

'If intChoice <> 0 Then
'    Source_MS = Application.FileDialog(msoFileDialogOpen).SelectedItems(1)
'    Set SourceMS = Workbooks.Open(Source_MS)
'    Source_MS = SourceMS.Name
'    Set SourceMS = Workbooks(Source_MS)
'Else
''    Application.StatusBar = False
''    Application.EnableEvents = True
''    Application.Calculation = xlCalculationAutomatic
''    Application.ScreenUpdating = True
''    Exit Sub
'    ThisWorkbook.Activate
'    Set SourceMS = ThisWorkbook
'End If

'Set SourceMS = Workbooks.Open(Source_MS)
'Source_MS = SourceMS.Name
'Set SourceMS = Workbooks(Source_MS)
'Set Source_Sheet = SourceMS.Sheets("04_Planning Template")

' Planning file's relevant data's first row
Source_First_Row_Header = Source_Sheet.Range("A1").End(xlDown).row
Source_First_Row_PL = Source_Sheet.Range("A1").End(xlDown).row + 1
Source_Sheet.AutoFilterMode = False
Source_Last_Row_Pl = Source_Sheet.Cells(Source_Sheet.Rows.Count, 1).End(xlUp).row
Source_Last_Col = Source_Sheet.Cells((Source_First_Row_PL - 1), Source_Sheet.Columns.Count).End(xlToLeft).column
Source_MSIndex_Col = Application.Match("Index", Source_Sheet.Range("21:21"), 0)
Source_ItemCode_Col = Application.Match("Item Code", Source_Sheet.Range("21:21"), 0)
Source_PlanStart_Col = Application.Match("Week", Source_Sheet.Range("1:1"), 0)
Source_PlanEnd_Col = Application.Match("Week", Source_Sheet.Range("1:1"), 0) + Application.WorksheetFunction.CountIf(Source_Sheet.Range("1:1"), "Week") - 1
Source_PrevYears_Col = Application.Match("Previous years not in MS", Source_Sheet.Range("21:21"), 0)
Source_SupSplit_Col = Application.Match("Supplier split", Source_Sheet.Range("21:21"), 0)
Source_ProdSemiCompl_Col = Application.Match("Production" & Chr(10) & "Semi/Complete", Source_Sheet.Range("21:21"), 0)
Source_IFRS15_Col = Application.Match("IFRS15", Source_Sheet.Range("21:21"), 0)
Source_Factory_Col = Application.Match("Factory", Source_Sheet.Range("21:21"), 0)
Source_Comment_Col = Application.Match("Comment", Source_Sheet.Range("21:21"), 0)
Source_ChangedBy_Col = Application.Match("Changed by", Source_Sheet.Range("21:21"), 0)
Source_ChangedDate_Col = Application.Match("Last change date", Source_Sheet.Range("21:21"), 0)

Source_Vessel_Col = Application.Match("Vessel", Source_Sheet.Range("21:21"), 0) '--Vessel Column introduced
Source_Released_Items_Col = Application.Match("Released Items", Source_Sheet.Range("21:21"), 0) '--Vessel Column introduced
Source_Free_Component_Column_Col_1 = Application.Match("Free Component Column 1", Source_Sheet.Range("21:21"), 0) '--Vessel Column introduced
Source_Free_Component_Column_Col_2 = Application.Match("Free Component Column 2", Source_Sheet.Range("21:21"), 0)  '--Vessel Column introduced
Source_Free_Component_Column_Col_3 = Application.Match("Free Component Column 3", Source_Sheet.Range("21:21"), 0)  '-Vessel Column introduced


With Source_Sheet.Sort
        .SortFields.Add Key:=Cells(Source_First_Row_PL, Source_MSIndex_Col), Order:=xlAscending
        .SortFields.Add Key:=Cells(Source_First_Row_PL, Source_ItemCode_Col), Order:=xlAscending
        
        .SetRange Range(Cells(Source_First_Row_PL, 1), Cells(Source_Last_Row_Pl, Source_Last_Col))
        .Header = xlNo
        .Apply
        .SortFields.Clear
End With

Dim entireData As Variant
entireData = Source_Sheet.Range(Source_Sheet.Cells(Source_First_Row_PL, 1), Source_Sheet.Cells(Source_Last_Row_Pl, Source_Last_Col)).value

' Define the columns you want to extract
Dim sourceColIndices As Collection
Set sourceColIndices = New Collection
   
'--Every new column which will be added will need to be extracted from here
' Add the columns of interest to the colIndices
sourceColIndices.Add Source_MSIndex_Col '--1 (Before 1)
sourceColIndices.Add Source_PrevYears_Col '--2 (Before 2)
sourceColIndices.Add Source_SupSplit_Col '--3 (Before 3)
sourceColIndices.Add Source_ProdSemiCompl_Col  '--4 (Before 4)
sourceColIndices.Add Source_IFRS15_Col '--5 (Before 5)
'--If Now want to add another vessel column, that should be extracted here
sourceColIndices.Add Source_Vessel_Col '--6 *** Now from here indices are changed)
sourceColIndices.Add Source_ItemCode_Col '--7 (Before 6)
sourceColIndices.Add Source_Factory_Col '--8 ( before 7)
sourceColIndices.Add Source_Released_Items_Col '--9
sourceColIndices.Add Source_Free_Component_Column_Col_1 '---10
sourceColIndices.Add Source_Free_Component_Column_Col_2 '--11
sourceColIndices.Add Source_Free_Component_Column_Col_3 '--12
For i = Source_PlanStart_Col To Source_PlanEnd_Col '---13
    sourceColIndices.Add i
Next i
sourceColIndices.Add Source_Comment_Col '-- Now would be 221(before 216)
sourceColIndices.Add Source_ChangedBy_Col '-- Now 222(before 217)
sourceColIndices.Add Source_ChangedDate_Col ' -- Now 223 (before 218)
    
    
'Timer
MacroStartTime = Timer
' Loading final Source Planning data
Dim FinalSourceData As Variant

FinalSourceData = ExtractColumns(entireData, sourceColIndices)

'' ***************The Commment Importing
Dim finalSourceComments() As String
ReDim finalSourceComments(UBound(FinalSourceData, 1), UBound(FinalSourceData, 2))
    
Dim p As Long, q As Long, excelCol As Long
For p = 1 To UBound(FinalSourceData, 1)
    For q = 1 To UBound(FinalSourceData, 2)
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
Source_Sheet.Range("A" & Source_First_Row_Header & ":" & Source_Sheet.Cells(Source_Last_Row_Pl, Source_Sheet.Columns.Count).End(xlToLeft).Address).AutoFilter

Dim planDateArray As Variant
planDateArray = Source_Sheet.Range(Source_Sheet.Cells(19, Source_PlanStart_Col), Source_Sheet.Cells(20, Source_PlanEnd_Col)).value


Dim MS_MSVersions(1 To 7) As Variant
Dim s As Integer
MS_MSVersions(1) = Current_Version
MS_MSVersions(2) = Current_VersionDate
MS_MSVersions(3) = Current_VersionYear
MS_MSVersions(4) = Current_VersionMonth
MS_MSVersions(5) = Current_VersionType
MS_MSVersions(6) = 1
MS_MSVersions(7) = Current_User

For s = 1 To 7
    MS_MSVersions_2D(1, s) = MS_MSVersions(s)
Next s

firstRowIndexSource = LBound(FinalSourceData, 1)
LastRowIndexSource = UBound(FinalSourceData, 1)
lastColindexSource = UBound(FinalSourceData, 2)


Dim msPlanningDataArrayRow As Long
ReDim ms_Planning_Array(1 To 1200000, 1 To 22) '----17 to 22
msPlanningDataArrayRow = 1


Dim counter As Long
counter = 0
Dim cR As Long
cR = firstRowIndexSource
Dim lastFoundPosition As Long
lastFoundPosition = LBound(ms_Planning_Array, 1)

lastFilledRow = 0
Call LoadMilestoneSheetData

'Row Processing ******************************************************************************************************************************************************
For cR = firstRowIndexSource To LastRowIndexSource
    Current_MS_Index = FinalSourceData(cR, 1)
    Current_ItemCode = FinalSourceData(cR, 7) '-- (cR, before 6)
    Current_ProdSemiCompl = FinalSourceData(cR, 4)
    Current_IFRS15 = FinalSourceData(cR, 5)
    Current_Factory = FinalSourceData(cR, 8) '--before 7
    Current_Vessel = FinalSourceData(cR, 6)
    Current_Released_Items = FinalSourceData(cR, 9)
    Current_Free_Component_Column_1 = FinalSourceData(cR, 10)
    Current_Free_Component_Column_2 = FinalSourceData(cR, 11)
    Current_Free_Component_Column_3 = FinalSourceData(cR, 12)
    
    If IsEmpty(Current_Factory) Or Current_Factory = "" Then
        For j = LBound(MilestoneSheetData, 1) To UBound(MilestoneSheetData, 1)
            If MilestoneSheetData(j, 1) = Current_ItemCode Then
                Current_Factory = MilestoneSheetData(j, 5)
                Exit For
            End If
        Next j
    End If
    
    Current_MSComment = FinalSourceData(cR, 221) '-- before 216
    Current_ChangedBy = FinalSourceData(cR, 222) ' -- before 217
    Current_LastChangeDate = FinalSourceData(cR, 223) '--  before 218
    
    Dim col As Integer
    Sum_Items_Row = 0
    
    If IsNumeric(FinalSourceData(cR, 2)) Then
        Sum_Items_Row = FinalSourceData(cR, 2)
    End If
    
    For col = 13 To 220 '-- eta chilo column 8 theke 215
        If IsNumeric(FinalSourceData(cR, col)) Then
            Sum_Items_Row = Sum_Items_Row + FinalSourceData(cR, col)
        End If
    Next col
    
    For k = LBound(MilestoneSheetData, 1) To UBound(MilestoneSheetData, 1)
        If MilestoneSheetData(k, 1) = Current_ItemCode Then
            UseDate = MilestoneSheetData(k, 4)
            Exit For
        End If
    Next k
    
    If Sum_Items_Row = 0 Then
        'old transferConflict
    Else
        Dim r As Long
        r = 2
        Source_PrevYears_Col = r
        CurrentItem_Col = Source_PrevYears_Col
        Source_Plan_Start_Col = 13 '-- eta chilo 8
        
        If PrevRow_MS_Index = Current_MS_Index And PrevRow_ItemCode = Current_ItemCode Then
            Current_ItemNoID = Current_ItemNoID
        Else
            Current_ItemNoID = 1
        End If
        Decimal_Prev_Item = 0
        '' Column Processing Starts*********************************************************************************************************************************
        For CurrentItem_Col = Source_PrevYears_Col To (lastColindexSource - 3) ' last Plan column is 3 column before last columns of planning file
            If CurrentItem_Col = Source_PrevYears_Col Or CurrentItem_Col >= Source_Plan_Start_Col Then
                Current_Item_Value = FinalSourceData(cR, CurrentItem_Col)
                If Not IsNumeric(Current_Item_Value) Then
                    'old transferConflict
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
                        If IsEmpty(finalSourceComments(cR, CurrentItem_Col)) Or Trim(FinalSourceData(cR, CurrentItem_Col)) = "" Then
                            Current_MSIndividualComment = ""
                        Else
                            Current_MSIndividualComment = finalSourceComments(cR, CurrentItem_Col)
                        End If
                        ' Calculating the start and end rows
                        Current_Row_Start = lastFilledRow + 1
                        Current_Row_End = Current_Row_Start + Total_Lines_Num - 1
                        Dim Date_Column_Maatched_With_Planning_Col As Long
                        Date_Column_Maatched_With_Planning_Col = CurrentItem_Col - 12 '--- previously it was 7
                        If Source_PrevYears_Col = CurrentItem_Col Then                                                                                                                                                                                                     'CurrentItem_Col = m + 6
                            FirstDate_MS = planDateArray(1, 1)
                            Current_MSDate = DateSerial(Month:=12, Day:=15, Year:=Year(FirstDate_MS - 15))             'End If
                        Else                                                                                                                                        'Next m
                            If UseDate = 20 Then
                                Current_MSDate = planDateArray(2, Date_Column_Maatched_With_Planning_Col) - 2
                            Else
                                Current_MSDate = planDateArray(1, Date_Column_Maatched_With_Planning_Col)
                            End If 'UseDate = 20
                        End If
                        For row = Current_Row_Start To Current_Row_End
                            ms_Planning_Array(row, 1) = Current_MS_Index
                            ms_Planning_Array(row, 2) = Current_ItemCode
                            ms_Planning_Array(row, 3) = Current_ItemNoID
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
                            ms_Planning_Array(row, 18) = Current_Vessel '--vessel
                            ms_Planning_Array(row, 19) = Current_Released_Items '--Released Items
                            ms_Planning_Array(row, 20) = Current_Free_Component_Column_1 '..Free Component Column 1
                            ms_Planning_Array(row, 21) = Current_Free_Component_Column_2 '--Free Component Coulumn 2
                            ms_Planning_Array(row, 22) = Current_Free_Component_Column_3 '--Free Component Column 3
                            Current_ItemNoID = Current_ItemNoID + 1
                        Next row
                        PrevRow_MS_Index = Current_MS_Index
                        PrevRow_ItemCode = Current_ItemCode
                        lastFilledRow = Current_Row_End
                    End If 'If Current_Item_Value <= 0 Then
                End If 'If Not IsNumeric(Current_Item_Value)
            End If 'If CurrentItem_Col = Souce_PrevYears_Col Or CurrentItem_Col >= Source_Plan_Start_Col
        Next CurrentItem_Col '*****************************************************************************************************************************************
        
        NextRow_MS_Index = FinalSourceData(cR + 1, 1)
        NextRow_ItemCode = FinalSourceData(cR + 1, 6)
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
            For Current_ProjInfo_Row = LBound(newArray, 1) To UBound(newArray, 1)
                If newArray(Current_ProjInfo_Row, 2) = Current_MS_Index Then
                    Exit For
                End If
            Next Current_ProjInfo_Row
            ' Get the value from the second column of the matched row
            'Debug.Print "Current_ProjInfo_Row " & Current_ProjInfo_Row
'            Current_Item_WTG_Total = newArray(Current_ProjInfo_Row, 3)
''            ' Apply the same conditions to modify the value if necessary
'            If Current_ItemCode = "05_1_0" Or Current_ItemCode = "05_1_1" Or Current_ItemCode = "05_1_2" Or Current_ItemCode = "05_1_3" Then
'                Current_Item_WTG_Total = Round(3 * Current_Item_WTG_Total, 0)
'            Else
''                ' Use the chance to round up decimals
'                Current_Item_WTG_Total = Round(Current_Item_WTG_Total, 0)
'            End If 'Current_ItemCode = "05_1_0" Or Current_ItemCode = "05_1_1" Or Current_ItemCode = "05_1_2" Or Current_ItemCode = "05_1_3" Then
'            If Current_Item_Count > Current_Item_WTG_Total Then
''                '''Too many items in the plan
'            ElseIf Current_Item_Count < Current_Item_WTG_Total Then
'                '''Too few items in the plan
'            End If 'Current_Item_Count > Current_Item_WTG_Total Then
        End If 'NextRow_MS_Index = Current_MS_Index And NextRow_ItemCode = Current_ItemCode Then
    End If '  If Sum_Items_Row = 0
Next cR '**************************************************************************************************************************************************************


If wasWorkbookOpened And SourceMS.Name <> ThisWorkbook.Name Then
    SourceMS.Close SaveChanges:=False  ' Change to True if you want to save changes
End If

Application.StatusBar = False
Application.EnableEvents = True
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
TimeGone = (Timer - MacroStartTime) / 86400
TimeGone = Format(TimeGone, "hh:mm:ss")
MsgBox "Finished the planning transfer in " & TimeGone
         
            
Dim answer As VbMsgBoxResult

answer = MsgBox("Do you want to proceed with the Database Upload?", vbYesNo + vbQuestion, "Database Upload")

If answer = vbYes Then
    'Insert your databse upload code here
    '''This Macro exports the DB Tables into the SQL Data Base

    MacroStartTime = Timer
    UploadToDB = DataBase_PRD '"GPO_DEV"
    
    '''MS_Project_Info table
    sConn = "Provider=SQLOLEDB;Data Source=" & SqlServer & ";Initial Catalog=" & UploadToDB & ";Integrated Security=SSPI"
    
    tblname = "MS_Project_Info"
    Call ExportArrayToSQL(entireMS_Project_Info_Data_new, sConn, tblname, MacroStartTime, UploadToDB)
    
    
    Dim MS_Planning_Mapping As Object
    Set MS_Planning_Mapping = CreateObject("Scripting.Dictionary")
    
    With MS_Planning_Mapping
        .Add 1, "MS_Index"
        .Add 2, "Item_Code"
        .Add 3, "Item_Number_ID"
        .Add 4, "Number_of_Items"
        .Add 5, "Production_Semi_Complete"
        .Add 6, "IFRS15"
        .Add 7, "Factory"
        .Add 8, "MS_Date"
        .Add 9, "MS_individual_comment"
        .Add 10, "MS_comment"
        .Add 11, "Changed_by"
        .Add 12, "Last_Change_Date"
        .Add 13, "Version_ID"
        'No mapping for "Version_Date" because it doesn't exist in SQL table
        .Add 15, "Data_Status"
        'Skipping "Control" as it might not have a matching column in SQL
        .Add 18, "Vessel"
        .Add 19, "Released_Items"
        .Add 20, "Free_Component_Column_1"
        .Add 21, "Free_Component_Column_2"
        .Add 22, "Free_Component_Column_3"
    End With
    
    '''MS_Planning table
    sConn = "Provider=SQLOLEDB;Data Source=" & SqlServer & ";Initial Catalog=" & UploadToDB & ";Integrated Security=SSPI"
    
    tblname = "MS_Planning"
    Call ExportArrayToSQL(ms_Planning_Array, sConn, tblname, MacroStartTime, UploadToDB, MS_Planning_Mapping, lastFilledRow) ' only here MS_Planning Mapping is used
    
    '''MS_Version table
    sConn = "Provider=SQLOLEDB;Data Source=" & SqlServer & ";Initial Catalog=" & UploadToDB & ";Integrated Security=SSPI"
    
    tblname = "MS_Versions"
    Call ExportArrayToSQL(MS_MSVersions_2D, sConn, tblname, MacroStartTime, UploadToDB)
    
    
    Application.StatusBar = "Uploaded to GPO Prod database, now copying to DEV and Test "
    Dim str_QUERY As String
    str_MS_VERISON = MS_MSVersions_2D(1, 1)
    str_QUERY = "EXEC [GPO].[dbo].[GPO_MS_copy_to_DB] '" & str_MS_VERISON & "'"
    If connectDB() Then
        Dim rs As ADODB.Recordset
        Set rs = conn.Execute(str_QUERY)
    Else
        'MsgBox "shit"
    End If
    
    
    
    TimeGone = (Timer - MacroStartTime) / 86400
    TimeGone = Format(TimeGone, "hh:mm:ss")
    
    MsgBox "Finished uploading information to Database in " & TimeGone
    Application.StatusBar = False

Else
    MsgBox "Database Upload Cancelled"
    Exit Sub
End If

End Sub
Function ExportArrayToSQL(ByVal msPlanningArray As Variant, _
    ByVal conString As String, ByVal table As String, _
    ByVal StartTime As Single, ByVal DB_Dest As String, _
    Optional ByVal colMapping As Object = Nothing, _
    Optional ByVal filledRowsCount As Long = -1, _
    Optional ByVal beforeSQL = "", Optional ByVal afterSQL As String) As Integer


'    On Error Resume Next

    ' Object type and CreateObject function are used instead of ADODB.Connection,
    ' ADODB.Command for late binding without reference to
    ' Microsoft ActiveX Data Objects 2.x Library
    ' ADO API Reference
    ' https://msdn.microsoft.com/en-us/library/ms678086(v=VS.85).aspx
    ' Dim con As ADODB.Connection
    Dim con As Object
    Set con = CreateObject("ADODB.Connection")

    con.ConnectionString = conString
    con.Open

    ' Dim cmd As ADODB.Command
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")

    ' BeginTrans, CommitTrans, and RollbackTrans Methods (ADO)
    ' http://msdn.microsoft.com/en-us/library/ms680895(v=vs.85).aspx

    Dim level As Long
    level = con.BeginTrans

    cmd.CommandType = 1             ' adCmdText
    If beforeSQL > "" Then
        cmd.CommandText = beforeSQL
        cmd.ActiveConnection = con
        cmd.Execute
    End If

    ' Dim rst As ADODB.Recordset
    Dim rst As Object
    Set rst = CreateObject("ADODB.Recordset")

    With rst
        Set .ActiveConnection = con
        .Source = "SELECT TOP 1 * FROM " & table
        .CursorLocation = 3         ' adUseClient
        .LockType = 4               ' adLockBatchOptimistic
        .CursorType = 0             ' adOpenForwardOnly
        .Open

        ' Column mappings

        Dim tableFields(100) As Integer
        Dim arrayFields(100) As Integer

        Dim exportFieldsCount As Integer
        exportFieldsCount = 0

        Dim col As Integer
        Dim index As Integer

        If colMapping Is Nothing Then
        ' Original linear approach
            For col = 0 To .Fields.Count - 1
                For i = 1 To UBound(msPlanningArray, 2)
                    If col = i - 1 Then
                        exportFieldsCount = exportFieldsCount + 1
                        tableFields(exportFieldsCount) = col
                        arrayFields(exportFieldsCount) = i
                        Exit For
                    End If
                Next i
            Next col
        Else
            ' The new mapping mechanism
            For i = 1 To UBound(msPlanningArray, 2)
                If colMapping.Exists(i) Then
                    columnInSQL = colMapping(i)
                    colIndexInSQL = -1
                    
                    For j = 0 To .Fields.Count - 1
                        If .Fields(j).Name = columnInSQL Then
                            colIndexInSQL = j
                            Exit For
                        End If
                    Next j
        
                    If colIndexInSQL > -1 Then
                        exportFieldsCount = exportFieldsCount + 1
                        tableFields(exportFieldsCount) = colIndexInSQL
                        arrayFields(exportFieldsCount) = i
                    End If
                End If
            Next i
        End If


        If exportFieldsCount = 0 Then
            ExportArrayToSQL = 1
            GoTo ConnectionEnd
        End If

        ' Fast read of Excel range values to an array
        ' for further fast work with the array

        'Dim arr As Variant
        'arr = sourceRange.value

        ' The range data transfer to the Recordset

        Dim row As Long
        Dim rowCount As Long
        If filledRowsCount > -1 Then
            rowCount = filledRowsCount
        Else
            rowCount = UBound(msPlanningArray, 1)
        End If

        Dim val As Variant
        xstep = Round(rowCount / (rowCount / 50), 0) + 1
        batchnumber = 1
        For xbatch = 1 To rowCount Step xstep
            TimeGone = (Timer - StartTime) / 86400
            TimeGone = Format(TimeGone, "hh:mm:ss")
            Application.StatusBar = "Uploading " & table & " in " & DB_Dest & " - uploading batch " & CStr(batchnumber) & " of " & (rowCount / 50) & " batches - " & TimeGone
            'Application.StatusBar = "Uploading " & table & " in " & DB_Dest & " - uploading batch " & CStr(batchnumber) & " of " & CStr(Round(rowCount / 200, 0)) & " batches - " & TimeGone
            DoEvents
            For row = xbatch To Application.WorksheetFunction.Min(xbatch + xstep - 1, rowCount)
                .AddNew 'I have done until this
                For col = 1 To exportFieldsCount
                        ' Insert debug statements here
                'Debug.Print "Row: " & row
                'Debug.Print "msPlanningArray First Dimension Upper Bound: " & UBound(msPlanningArray, 1)
                'Debug.Print "msPlanningArray First Dimension Lower Bound: " & LBound(msPlanningArray, 1)
        
                'Debug.Print "Col: " & col
                'Debug.Print "arrayFields Upper Bound: " & UBound(arrayFields)
                'Debug.Print "arrayFields Lower Bound: " & LBound(arrayFields)
        
                ' Since arrayFields(col) might be the problematic part, we will surround it with error handling to avoid interrupting the debug print
                On Error Resume Next
                'Debug.Print "arrayFields(col): " & arrayFields(col)
                On Error GoTo 0
        
                'Debug.Print "msPlanningArray Second Dimension Upper Bound: " & UBound(msPlanningArray, 2)
                'Debug.Print "msPlanningArray Second Dimension Lower Bound: " & LBound(msPlanningArray, 2)
        
                ' Original line causing error

                    val = msPlanningArray(row, arrayFields(col)) 'rangeFields(col) ised because rangeFields
                    If val = "" Then ' is representing the column order of the
                        .Fields(tableFields(col)).value = Null
                    Else ' source range
                        On Error GoTo ErrHandler
                            .Fields(tableFields(col)) = val
                        On Error GoTo 0
                    End If
                Next
            Next
    
             .UpdateBatch
            batchnumber = batchnumber + 1
        Next
    End With

    rst.Close
    Set rst = Nothing

    If afterSQL > "" Then
        cmd.CommandText = afterSQL
        cmd.ActiveConnection = con
        cmd.Execute
    End If

    ExportRangeToSQL = 0

ConnectionEnd:

    con.CommitTrans

    con.Close
    Set cmd = Nothing
    Set con = Nothing
    
Exit Function
ErrHandler:
MsgBox "Error inserting into column '" & rst.Fields(tableFields(col)).Name & _
"' (Index: " & tableFields(col) & ") with value '" & val & _
"' of type '" & TypeName(val) & "'. Error " & err.Number & _
": " & err.Description

End Function
Function connectDB() As Boolean
'Function to estalish and check the Database connection
connectDB = False

Set conn = New ADODB.Connection
If conn.State = 0 Then

'''Change DB name
    sConn = "Provider=SQLOLEDB;Data Source=" & SqlServer & ";Initial Catalog=" & Database & ";Integrated Security=SSPI"
    With conn
        .ConnectionString = sConn
        .ConnectionTimeout = 250
        .CommandTimeout = 2000
        
    End With
    conn.Open
End If
If conn.State = adStateOpen Then connectDB = True

End Function
Function GetPingResult(Host) As Boolean

'Function to ping the SqlServer and check if it is available, and if it is reached via intranet (IP begins with 10.)
Dim objPing As Object
Dim objStatus As Object
Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}"). _
       ExecQuery("Select * from Win32_PingStatus Where Address = '" & Host & "'")
For Each objStatus In objPing
    If objStatus.Statuscode = 0 Then
        IPrange = Left(objStatus.ProtocolAddress, InStr(1, objStatus.ProtocolAddress, ".") - 1)
        If IPrange = 10 Then
            GetPingResult = True
        Else
            GetPingResult = False
        End If
    Else
        GetPingResult = False
    End If
Next objStatus

Set objPing = Nothing

End Function
Function connect_Database() As Boolean
Dim xcount As Integer
xcount = 0

If ActiveWorkbook.Name <> ThisWorkbook.Name Then ThisWorkbook.Activate
StartTime = Timer

connect_Database = False
Do While xcount < 4
    xcount = xcount + 1
    If GetPingResult(SqlServer) = False Then
        If xcount > 3 Then
            Exit Function
        End If
    Else
        Exit Do
    End If
Loop

If ActiveWorkbook.Name <> ThisWorkbook.Name Then ThisWorkbook.Activate
xcount = 0

StartTime = Timer
Do While xcount < 4
    xcount = xcount + 1
    If connectDB() = False Then
        If xcount > 3 Then
            Exit Function
        End If
    Else
        Exit Do
    End If
Loop

connect_Database = True
End Function
Function sqlDate(IN_Date As Date) As String
    sqlDate = Right("0000" + Trim(Str(Year(IN_Date))), 4) & _
              Right("00" + Trim(Str(Month(IN_Date))), 2) & _
              Right("00" + Trim(Str(Day(IN_Date))), 2)
End Function


Sub update_values_RD()
Dim rs1 As ADODB.Recordset

On Error GoTo err
'Connect to DB, and if not possible quit
If ActiveWorkbook.Name <> ThisWorkbook.Name Then ThisWorkbook.Activate
If connect_Database() = False Then
    MsgBox "Connection to DB failed. Please check that you are in NX/AWP intranet. If so, please contact GPO"
    Exit Sub
End If

Application.Calculation = xlCalculationManual

Querybasis = "SELECT DISTINCT Version_ID FROM v_MS_Name_Full_Test_v1"

Set rs1 = conn.Execute(Querybasis)
If Not rs1.EOF Then
    Sheets("Help_Tab").ListObjects("MS_List").Range.Offset(1, 0).ClearContents
    Sheets("Help_Tab").ListObjects("MS_List").Range.Offset(1, 0).CopyFromRecordset rs1
    Sheets("Help_Tab").ListObjects("MS_List").Resize Sheets("Help_Tab").Range(Sheets("Help_Tab").ListObjects("MS_List").Range.Cells(1, 1).Address, Sheets("Help_Tab").ListObjects("MS_List").Range.End(xlDown).Offset(0, rs1.Fields.Count - 1).Address)
End If

If Not rs1 Is Nothing Then
    rs1.Close
    Set rs1 = Nothing
End If
If CBool(conn.State And adStateOpen) Then conn.Close

Exit Sub

err:
strtmp = strtmp & vbCrLf & "VB Error # " & Str(err.Number)
strtmp = strtmp & vbCrLf & "   Generated by " & err.Source
strtmp = strtmp & vbCrLf & "   Description  " & err.Description

' Enumerate Errors collection and display properties of
' each Error object.
If Not conn Is Nothing Then
    Set Errs1 = conn.Errors
    For Each errLoop In Errs1
         With errLoop
           strtmp = strtmp & vbCrLf & "Error #" & i & ":"
           strtmp = strtmp & vbCrLf & "   ADO Error   #" & .Number
           strtmp = strtmp & vbCrLf & "   Description  " & .Description
           strtmp = strtmp & vbCrLf & "   Source       " & .Source
           i = i + 1
        End With
    Next
    If CBool(conn.State And adStateOpen) Then conn.Close
End If

If ActiveWorkbook.Name <> ThisWorkbook.Name Then ThisWorkbook.Activate
Application.Calculation = xlCalculationAutomatic
MsgBox "There was an error. Please contact GPO"
End Sub

Sub LoadMilestoneSheetData()
    
  ReDim MilestoneSheetData(1 To 59, 1 To 5)

' Row 1
MilestoneSheetData(1, 1) = "0_0": MilestoneSheetData(1, 2) = "NTP": MilestoneSheetData(1, 3) = "-": MilestoneSheetData(1, 4) = 20: MilestoneSheetData(1, 5) = "-"

' Row 2
MilestoneSheetData(2, 1) = "01_1_0": MilestoneSheetData(2, 2) = "NAC": MilestoneSheetData(2, 3) = "Exworks": MilestoneSheetData(2, 4) = 20: MilestoneSheetData(2, 5) = "NAC Supplier"

' Row 3
MilestoneSheetData(3, 1) = "01_1_1": MilestoneSheetData(3, 2) = "NAC": MilestoneSheetData(3, 3) = "PickUp": MilestoneSheetData(3, 4) = 19: MilestoneSheetData(3, 5) = "NAC Supplier"

' Row 4
MilestoneSheetData(4, 1) = "01_1_2": MilestoneSheetData(4, 2) = "NAC": MilestoneSheetData(4, 3) = "FOB": MilestoneSheetData(4, 4) = 19: MilestoneSheetData(4, 5) = "NAC Supplier"

' Row 5
MilestoneSheetData(5, 1) = "01_1_3": MilestoneSheetData(5, 2) = "NAC": MilestoneSheetData(5, 3) = "DDP": MilestoneSheetData(5, 4) = 19: MilestoneSheetData(5, 5) = "-"
' Row 6
MilestoneSheetData(6, 1) = "01_1_4": MilestoneSheetData(6, 2) = "NAC": MilestoneSheetData(6, 3) = "FWA": MilestoneSheetData(6, 4) = 19: MilestoneSheetData(6, 5) = "FWASupplier"

' Row 7
MilestoneSheetData(7, 1) = "02_1_0": MilestoneSheetData(7, 2) = "HUB": MilestoneSheetData(7, 3) = "Exworks": MilestoneSheetData(7, 4) = 20: MilestoneSheetData(7, 5) = "HUB Supplier"

' Row 8
MilestoneSheetData(8, 1) = "02_1_1": MilestoneSheetData(8, 2) = "HUB": MilestoneSheetData(8, 3) = "PickUp": MilestoneSheetData(8, 4) = 19: MilestoneSheetData(8, 5) = "HUB Supplier"

' Row 9
MilestoneSheetData(9, 1) = "02_1_2": MilestoneSheetData(9, 2) = "HUB": MilestoneSheetData(9, 3) = "FOB": MilestoneSheetData(9, 4) = 19: MilestoneSheetData(9, 5) = "HUB Supplier"

' Row 10
MilestoneSheetData(10, 1) = "02_1_3": MilestoneSheetData(10, 2) = "HUB": MilestoneSheetData(10, 3) = "DDP": MilestoneSheetData(10, 4) = 19: MilestoneSheetData(10, 5) = "-"
' Row 11
MilestoneSheetData(11, 1) = "02_1_4": MilestoneSheetData(11, 2) = "HUB": MilestoneSheetData(11, 3) = "FWA": MilestoneSheetData(11, 4) = 19: MilestoneSheetData(11, 5) = "FWASupplier"

' Row 12
MilestoneSheetData(12, 1) = "03_1_0": MilestoneSheetData(12, 2) = "DTR": MilestoneSheetData(12, 3) = "Exworks": MilestoneSheetData(12, 4) = 20: MilestoneSheetData(12, 5) = "DTR Supplier"

' Row 13
MilestoneSheetData(13, 1) = "03_1_1": MilestoneSheetData(13, 2) = "DTR": MilestoneSheetData(13, 3) = "PickUp": MilestoneSheetData(13, 4) = 19: MilestoneSheetData(13, 5) = "DTR Supplier"

' Row 14
MilestoneSheetData(14, 1) = "03_1_2": MilestoneSheetData(14, 2) = "DTR": MilestoneSheetData(14, 3) = "FOB": MilestoneSheetData(14, 4) = 19: MilestoneSheetData(14, 5) = "DTR Supplier"

' Row 15
MilestoneSheetData(15, 1) = "03_1_3": MilestoneSheetData(15, 2) = "DTR": MilestoneSheetData(15, 3) = "DDP": MilestoneSheetData(15, 4) = 19: MilestoneSheetData(15, 5) = "-"

' Row 16
MilestoneSheetData(16, 1) = "03_1_4": MilestoneSheetData(16, 2) = "DTR": MilestoneSheetData(16, 3) = "FWA": MilestoneSheetData(16, 4) = 19: MilestoneSheetData(16, 5) = "FWASupplier"

' Row 17
MilestoneSheetData(17, 1) = "04_1_0": MilestoneSheetData(17, 2) = "Concrete Laying": MilestoneSheetData(17, 3) = "-": MilestoneSheetData(17, 4) = 20: MilestoneSheetData(17, 5) = "-"

' Row 18
MilestoneSheetData(18, 1) = "04_2_0": MilestoneSheetData(18, 2) = "Concrete Tower": MilestoneSheetData(18, 3) = "Exworks": MilestoneSheetData(18, 4) = 20: MilestoneSheetData(18, 5) = "Concrete Tower Supplier"

' Row 19
MilestoneSheetData(19, 1) = "04_2_1": MilestoneSheetData(19, 2) = "Concrete Tower": MilestoneSheetData(19, 3) = "PickUp": MilestoneSheetData(19, 4) = 19: MilestoneSheetData(19, 5) = "Concrete Tower Supplier"

' Row 20
MilestoneSheetData(20, 1) = "04_2_2": MilestoneSheetData(20, 2) = "Concrete Tower": MilestoneSheetData(20, 3) = "FOB": MilestoneSheetData(20, 4) = 19: MilestoneSheetData(20, 5) = "Concrete Tower Supplier"

' Row 21
MilestoneSheetData(21, 1) = "04_2_3": MilestoneSheetData(21, 2) = "Concrete Tower": MilestoneSheetData(21, 3) = "DDP": MilestoneSheetData(21, 4) = 19: MilestoneSheetData(21, 5) = "-"

' Row 22
MilestoneSheetData(22, 1) = "04_2_4": MilestoneSheetData(22, 2) = "Concrete Tower": MilestoneSheetData(22, 3) = "FWA": MilestoneSheetData(22, 4) = 19: MilestoneSheetData(22, 5) = "FWASupplier"

' Row 23
MilestoneSheetData(23, 1) = "04_3_0": MilestoneSheetData(23, 2) = "Steel Tower": MilestoneSheetData(23, 3) = "Exworks": MilestoneSheetData(23, 4) = 20: MilestoneSheetData(23, 5) = "Steel Tower Supplier"

' Row 24
MilestoneSheetData(24, 1) = "04_3_1": MilestoneSheetData(24, 2) = "Steel Tower": MilestoneSheetData(24, 3) = "PickUp": MilestoneSheetData(24, 4) = 19: MilestoneSheetData(24, 5) = "Steel Tower Supplier"

' Row 25
MilestoneSheetData(25, 1) = "04_3_2": MilestoneSheetData(25, 2) = "Steel Tower": MilestoneSheetData(25, 3) = "FOB": MilestoneSheetData(25, 4) = 19: MilestoneSheetData(25, 5) = "Steel Tower Supplier"

' Row 26
MilestoneSheetData(26, 1) = "04_3_3": MilestoneSheetData(26, 2) = "Steel Tower": MilestoneSheetData(26, 3) = "DDP": MilestoneSheetData(26, 4) = 19: MilestoneSheetData(26, 5) = "-"

' Row 27
MilestoneSheetData(27, 1) = "04_3_4": MilestoneSheetData(27, 2) = "Steel Tower": MilestoneSheetData(27, 3) = "FWA": MilestoneSheetData(27, 4) = 19: MilestoneSheetData(27, 5) = "FWASupplier"
' Row 28
MilestoneSheetData(28, 1) = "05_1_0": MilestoneSheetData(28, 2) = "Blade": MilestoneSheetData(28, 3) = "Exworks": MilestoneSheetData(28, 4) = 20: MilestoneSheetData(28, 5) = "Blade Supplier"

' Row 29
MilestoneSheetData(29, 1) = "05_1_1": MilestoneSheetData(29, 2) = "Blade": MilestoneSheetData(29, 3) = "PickUp": MilestoneSheetData(29, 4) = 19: MilestoneSheetData(29, 5) = "Blade Supplier"

' Row 30
MilestoneSheetData(30, 1) = "05_1_2": MilestoneSheetData(30, 2) = "Blade": MilestoneSheetData(30, 3) = "FOB": MilestoneSheetData(30, 4) = 19: MilestoneSheetData(30, 5) = "Blade Supplier"

' Row 31
MilestoneSheetData(31, 1) = "05_1_3": MilestoneSheetData(31, 2) = "Blade": MilestoneSheetData(31, 3) = "DDP": MilestoneSheetData(31, 4) = 19: MilestoneSheetData(31, 5) = "-"

' Row 32
MilestoneSheetData(32, 1) = "05_1_4": MilestoneSheetData(32, 2) = "Blade": MilestoneSheetData(32, 3) = "FWA": MilestoneSheetData(32, 4) = 19: MilestoneSheetData(32, 5) = "FWASupplier"

' Row 33
MilestoneSheetData(33, 1) = "06_1_0": MilestoneSheetData(33, 2) = "Foundations": MilestoneSheetData(33, 3) = "Exworks": MilestoneSheetData(33, 4) = 20: MilestoneSheetData(33, 5) = "-"

' Row 34
MilestoneSheetData(34, 1) = "06_1_1": MilestoneSheetData(34, 2) = "Foundations": MilestoneSheetData(34, 3) = "PickUp": MilestoneSheetData(34, 4) = 19: MilestoneSheetData(34, 5) = "-"

' Row 35
MilestoneSheetData(35, 1) = "06_1_2": MilestoneSheetData(35, 2) = "Foundations": MilestoneSheetData(35, 3) = "FOB": MilestoneSheetData(35, 4) = 19: MilestoneSheetData(35, 5) = "-"

' Row 36
MilestoneSheetData(36, 1) = "06_1_3": MilestoneSheetData(36, 2) = "Foundations": MilestoneSheetData(36, 3) = "DDP": MilestoneSheetData(36, 4) = 19: MilestoneSheetData(36, 5) = "-"

' Row 37
MilestoneSheetData(37, 1) = "07_1": MilestoneSheetData(37, 2) = "Installation Start": MilestoneSheetData(37, 3) = "Installation": MilestoneSheetData(37, 4) = 20: MilestoneSheetData(37, 5) = "-"
' Row 38
MilestoneSheetData(38, 1) = "08_0": MilestoneSheetData(38, 2) = "Foundations": MilestoneSheetData(38, 3) = "Installation": MilestoneSheetData(38, 4) = 20: MilestoneSheetData(38, 5) = "-"

' Row 39
MilestoneSheetData(39, 1) = "08_1": MilestoneSheetData(39, 2) = "Pre-assembly": MilestoneSheetData(39, 3) = "Installation": MilestoneSheetData(39, 4) = 20: MilestoneSheetData(39, 5) = "-"

' Row 40
MilestoneSheetData(40, 1) = "08_2": MilestoneSheetData(40, 2) = "Tower": MilestoneSheetData(40, 3) = "Installation": MilestoneSheetData(40, 4) = 20: MilestoneSheetData(40, 5) = "-"

' Row 41
MilestoneSheetData(41, 1) = "08_2_0": MilestoneSheetData(41, 2) = "Cubible assembly": MilestoneSheetData(41, 3) = "Installation": MilestoneSheetData(41, 4) = 20: MilestoneSheetData(41, 5) = "-"

' Row 42
MilestoneSheetData(42, 1) = "08_2_0_1": MilestoneSheetData(42, 2) = "Montaje Losa Superior": MilestoneSheetData(42, 3) = "Installation": MilestoneSheetData(42, 4) = 20: MilestoneSheetData(42, 5) = "-"

' Row 43
MilestoneSheetData(43, 1) = "08_2_0_2": MilestoneSheetData(43, 2) = "Tower LTM1500 (8 piezas)": MilestoneSheetData(43, 3) = "Installation": MilestoneSheetData(43, 4) = 20: MilestoneSheetData(43, 5) = "-"

' Row 44
MilestoneSheetData(44, 1) = "08_2_0_3": MilestoneSheetData(44, 2) = "Tower LTM1500 (cota 80)": MilestoneSheetData(44, 3) = "Installation": MilestoneSheetData(44, 4) = 20: MilestoneSheetData(44, 5) = "-"

' Row 45
MilestoneSheetData(45, 1) = "08_2_0_4": MilestoneSheetData(45, 2) = "Tower LR1300 (equivalente)": MilestoneSheetData(45, 3) = "Installation": MilestoneSheetData(45, 4) = 20: MilestoneSheetData(45, 5) = "-"

' Row 46
MilestoneSheetData(46, 1) = "08_3": MilestoneSheetData(46, 2) = "NAC": MilestoneSheetData(46, 3) = "Installation": MilestoneSheetData(46, 4) = 20: MilestoneSheetData(46, 5) = "-"

' Row 47
MilestoneSheetData(47, 1) = "08_4": MilestoneSheetData(47, 2) = "Blade": MilestoneSheetData(47, 3) = "Installation": MilestoneSheetData(47, 4) = 20: MilestoneSheetData(47, 5) = "-"
' Row 48
MilestoneSheetData(48, 1) = "09_1": MilestoneSheetData(48, 2) = "Quality Reviews": MilestoneSheetData(48, 3) = "Installation": MilestoneSheetData(48, 4) = 20: MilestoneSheetData(48, 5) = "-"

' Row 49
MilestoneSheetData(49, 1) = "10_1": MilestoneSheetData(49, 2) = "Commissioning": MilestoneSheetData(49, 3) = "Commissioning": MilestoneSheetData(49, 4) = 20: MilestoneSheetData(49, 5) = "-"

' Row 50
MilestoneSheetData(50, 1) = "10_2": MilestoneSheetData(50, 2) = "Commissioning End": MilestoneSheetData(50, 3) = "Commissioning": MilestoneSheetData(50, 4) = 20: MilestoneSheetData(50, 5) = "-"

' Row 51
MilestoneSheetData(51, 1) = "11_1_1": MilestoneSheetData(51, 2) = "COD": MilestoneSheetData(51, 3) = "-": MilestoneSheetData(51, 4) = 20: MilestoneSheetData(51, 5) = "-"

' Row 52
MilestoneSheetData(52, 1) = "11_1_2": MilestoneSheetData(52, 2) = "Final Invoice": MilestoneSheetData(52, 3) = "-": MilestoneSheetData(52, 4) = 20: MilestoneSheetData(52, 5) = "-"

' Row 53
MilestoneSheetData(53, 1) = "12_1_1": MilestoneSheetData(53, 2) = "Nacelle + blade + hub": MilestoneSheetData(53, 3) = "FOB": MilestoneSheetData(53, 4) = 19: MilestoneSheetData(53, 5) = "-"

' Row 54
MilestoneSheetData(54, 1) = "12_1_2": MilestoneSheetData(54, 2) = "Nacelle + blade + tower": MilestoneSheetData(54, 3) = "FOB": MilestoneSheetData(54, 4) = 19: MilestoneSheetData(54, 5) = "-"

' Row 55
MilestoneSheetData(55, 1) = "12_1_3": MilestoneSheetData(55, 2) = "Nacelle + blade": MilestoneSheetData(55, 3) = "FOB": MilestoneSheetData(55, 4) = 19: MilestoneSheetData(55, 5) = "-"

' Row 56
MilestoneSheetData(56, 1) = "12_2_1": MilestoneSheetData(56, 2) = "Nacelle + blade + hub": MilestoneSheetData(56, 3) = "DDP": MilestoneSheetData(56, 4) = 19: MilestoneSheetData(56, 5) = "-"

' Row 57
MilestoneSheetData(57, 1) = "12_2_2": MilestoneSheetData(57, 2) = "Nacelle + blade + tower": MilestoneSheetData(57, 3) = "DDP": MilestoneSheetData(57, 4) = 19: MilestoneSheetData(57, 5) = "-"

' Row 58
MilestoneSheetData(58, 1) = "12_2_3": MilestoneSheetData(58, 2) = "Nacelle + blade": MilestoneSheetData(58, 3) = "DDP": MilestoneSheetData(58, 4) = 19: MilestoneSheetData(58, 5) = "-"

' Row 59
MilestoneSheetData(59, 1) = "13_1": MilestoneSheetData(59, 2) = "IFRS-15": MilestoneSheetData(59, 3) = "-": MilestoneSheetData(59, 4) = 20: MilestoneSheetData(59, 5) = "-"


End Sub

Function CustomFloor(ByVal value As Double) As Double
    
    Dim strValue As String
    strValue = CStr(value)
    Dim pos As Integer
    pos = InStr(strValue, ".")
    If pos = 0 Then pos = InStr(strValue, ",")
    If pos > 0 Then
        strValue = Left(strValue, pos - 1)
    End If
    CustomFloor = CDbl(strValue)
    
    
        
End Function
Function ExtractColumns(dataArray As Variant, colIndices As Collection) As Variant
    Dim numRows As Long
    Dim numCols As Long
    Dim etractedData() As Variant
    Dim i As Long, j As Long
    
    numRows = UBound(dataArray, 1)
    numCols = colIndices.Count
    
    ' Resize the extractedData array based on the number of the rows and columns of interest
    ReDim extractedData(1 To numRows, 1 To numCols)
    
    'Extract only the columns of interest
    For i = 1 To numRows
        For j = 1 To numCols
            extractedData(i, j) = dataArray(i, colIndices(j))
        Next j
    Next i
    
    ExtractColumns = extractedData
End Function






