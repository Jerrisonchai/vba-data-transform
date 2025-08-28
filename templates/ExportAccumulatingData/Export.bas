Attribute VB_Name = "Export"
Option Explicit
Sub ExceltoText()
Call capturetime
Call MyShape_Click
Dim startTime As Date, endTime As Date, timetaken As Date
startTime = Now()
Sheets("Summary 2023").Activate
'Declaring the variables
 Dim FileName, sLine, Deliminator As String, LastCol, LastRow, FileNumber As Integer, i As Long, j As Long
'Excel Location and File Name
 FileName = Sheets("Dashboard").Range("C20") & "\DailyData " & Format(Now(), "YYYY-MM-DD hh mm AMPM") & ".txt"
'Field Separator
 Deliminator = vbTab
'Identifying the Last Cell
 LastCol = Sheets("Summary 2023").Cells.SpecialCells(xlCellTypeLastCell).Column
 LastRow = Sheets("Summary 2023").Cells.SpecialCells(xlCellTypeLastCell).Row
 FileNumber = FreeFile
'Creating or Overwrighting a text file
 Open FileName For Output As FileNumber
'Reading the data from Excel using For Loop
 For i = 1 To LastRow
 For j = 1 To LastCol
'Removing Deliminator if it is wrighting the last column
 If j = LastCol Then
 sLine = sLine & Cells(i, j).Value
 Else
 sLine = sLine & Cells(i, j).Value & Deliminator
 End If
 Next j
'Wrighting data into text file
 Print #FileNumber, sLine
 sLine = ""
 Next i
'Closing the Text File
 Close #FileNumber
'Generating message to display
 MsgBox "Text file has been generated"
[Start_Time].Value = startTime: [UserName].Value = Environ("Username")
Call captureendtime
End Sub
Sub ExceltoCsv()
Call capturetime
Call MyShape_Click
Dim startTime As Date, endTime As Date, timetaken As Date
startTime = Now()
Sheets("Summary 2023").Activate
'Declaring the variables
 Dim FileName, sLine, Deliminator As String, LastCol, LastRow, FileNumber As Integer, i As Long, j As Long
'Excel Location and File Name
 FileName = Sheets("Dashboard").Range("C20") & "\DailyData " & Format(Now(), "YYYY-MM-DD hh mm AMPM") & ".csv"
'Field Separator
 Deliminator = vbTab
'Identifying the Last Cell
 LastCol = Sheets("Summary 2023").Cells.SpecialCells(xlCellTypeLastCell).Column
 LastRow = Sheets("Summary 2023").Cells.SpecialCells(xlCellTypeLastCell).Row
 FileNumber = FreeFile
'Creating or Overwrighting a text file
 Open FileName For Output As FileNumber
'Reading the data from Excel using For Loop
 For i = 1 To LastRow
 For j = 1 To LastCol
'Removing Deliminator if it is wrighting the last column
 If j = LastCol Then
 sLine = sLine & Cells(i, j).Value
 Else
 sLine = sLine & Cells(i, j).Value & Deliminator
 End If
 Next j
'Wrighting data into text file
 Print #FileNumber, sLine
 sLine = ""
 Next i
'Closing the Text File
 Close #FileNumber
'Generating message to display
 MsgBox "Text file has been generated"
Call captureendtime
[Start_Time].Value = startTime: [UserName].Value = Environ("Username")
End Sub

