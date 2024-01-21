Attribute VB_Name = "MismatchingCounts"
Option Explicit
Sub MismatchingCounts()
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")

    ' Define your server and database
    Dim SqlServer As String
    Dim DataBase_PRD As String

    ' Set your server and database
    SqlServer = "10.5.7.3"  ' Replace with your server name
    DataBase_PRD = "GPO"  ' Replace with your database name

    ' Connection String
    conn.Open "Provider=SQLOLEDB;Data Source=" & SqlServer & ";Initial Catalog=" & DataBase_PRD & ";Integrated Security=SSPI"

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")
    
    'Prompting user for the Version_ID Input
    Dim versionID As String
    versionID = InputBox("Input the Version_ID value:")
    
    If versionID = "" Then
         Exit Sub
    End If

    ' SQL query
    Dim sqlQuery As String
    sqlQuery = "SELECT p.MS_Index, p.Version_ID, p.Item_Code, i.Project_Name, COUNT(*) AS ActualCount, " & _
               "CASE WHEN LEFT(p.Item_Code, 2) = '05' THEN 3 * i.WTG_Total ELSE i.WTG_Total END AS ExpectedCount " & _
               "FROM dbo.MS_Planning p INNER JOIN dbo.MS_Project_Info i ON p.MS_Index = i.MS_Index AND p.Version_ID = i.Version_ID " & _
               "WHERE p.Version_ID = '" & versionID & "' GROUP BY p.MS_Index, p.Version_ID, p.Item_Code, i.Project_Name, i.WTG_Total " & _
               "HAVING COUNT(*) != CASE WHEN LEFT(p.Item_Code, 2) = '05' THEN 3 * i.WTG_Total ELSE i.WTG_Total END"

    ' Executing the SQL query
    rs.Open sqlQuery, conn

    ' Creating a new worksheet
    Dim NewSheet As Worksheet
    Set NewSheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    NewSheet.Name = "Mismatch Tab"  ' Optional: Rename the sheet

    ' Print the column headers in A1 to F1
    NewSheet.Range("A1").value = "MS_Index"
    NewSheet.Range("B1").value = "Version_ID"
    NewSheet.Range("C1").value = "Item_Code"
    NewSheet.Range("D1").value = "Project_Name"
    NewSheet.Range("E1").value = "MS Count"
    NewSheet.Range("F1").value = "Expected Count"

    ' Pasting the recordset into the new worksheet
    NewSheet.Range("A2").CopyFromRecordset rs
    
    'Alignment
    NewSheet.Columns("A:F").AutoFit
    NewSheet.Range("A1:F1").HorizontalAlignment = xlCenter
    
    ' Closing the recordset and connection
    rs.Close
    conn.Close

    Set rs = Nothing
    Set conn = Nothing
End Sub


