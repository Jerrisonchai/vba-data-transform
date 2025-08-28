Attribute VB_Name = "ExportData"
Option Explicit
Sub CopySheetsToNewWorkbook()
Call capturetime
Call MyShape_Click
Dim startTime As Date, endTime As Date, timetaken As Date
startTime = Now()
    Dim sourceWorkbook As Workbook, newWorkbook As Workbook, sourceSheet1 As Worksheet, sourceSheet2 As Worksheet, newSheet1 As Worksheet, newSheet2 As Worksheet
    Set sourceWorkbook = ThisWorkbook
    Set sourceSheet1 = sourceWorkbook.Sheets("Monthly")
    Set sourceSheet2 = sourceWorkbook.Sheets("Summary 2023")
    OptimizedMode True
    sourceWorkbook.RefreshAll
    Set newWorkbook = Workbooks.Add
    On Error Resume Next
    Set newSheet1 = newWorkbook.Sheets(sourceSheet1.Name)
    Set newSheet2 = newWorkbook.Sheets(sourceSheet2.Name)
    On Error GoTo 0
    ' If the sheets don't exist in the new workbook, add them
    If newSheet1 Is Nothing Then
        sourceSheet1.Copy Before:=newWorkbook.Sheets(1)
        Set newSheet1 = newWorkbook.Sheets(1)
    End If
    If newSheet2 Is Nothing Then
        sourceSheet2.Copy Before:=newWorkbook.Sheets(1)
        Set newSheet2 = newWorkbook.Sheets(1)
    End If
    ' Copy formatting from sourceSheet1 to newSheet1
    sourceSheet1.Cells.Copy
    newSheet1.Cells.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False ' Clear clipboard
    ' Copy formatting from sourceSheet2 to newSheet2
    sourceSheet2.Cells.Copy
    newSheet2.Cells.PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False ' Clear clipboard
    ' Optional: Save the new workbook with a specific name and path
    newWorkbook.SaveAs sourceWorkbook.Sheets("Dashboard").Range("C20").Value & "\" & "Daily Sample List " & Format(Now(), "YYYY-MM-DD hh mm AMPM") & ".xlsx"
    sourceWorkbook.Activate
    OptimizedMode False
    ' Clean up
    Set newWorkbook = Nothing
    Set newWorkbook = Nothing
    Set sourceSheet1 = Nothing
    Set sourceSheet2 = Nothing
    Set newSheet1 = Nothing
    Set newSheet2 = Nothing
[Start_Time].Value = startTime: [UserName].Value = Environ("Username")
 Call captureendtime
End Sub
Sub SaveAs()    'Button - Save As
Call capturetime
Call MyShape_Click
Dim startTime As Date, endTime As Date, timetaken As Date
startTime = Now()
ThisWorkbook.Save
OptimizedMode True
Application.DisplayAlerts = False
ThisWorkbook.RefreshAll
Dim sFolder As String   'Step 1 - Choose Folder to save file
    sFolder = Sheets("Dashboard").Range("C20").Value & "\"
    Dim FolderPicker As FileDialog, mypath As String, ar, n As Integer
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
        With FolderPicker
            .Title = "Choose folder to Save"
            .InitialFileName = sFolder
            .AllowMultiSelect = False
            .ButtonName = "Confirm"
                If .Show = -1 Then
                    mypath = .SelectedItems(1)
                    Else
                        Exit Sub
                End If
        End With
    Sheets("Dashboard").Range("C20").Value = mypath
ThisWorkbook.SaveAs FileName:=ThisWorkbook.Sheets("Dashboard").Range("C20").Value & "\" & "Daily Sample List " & Format(Now(), "YYYY-MM-DD hh mm AMPM") & ".xlsx", _
    FileFormat:=51, _
    CreateBackup:=False     'Step 2 - Save as excel
    With ThisWorkbook
        ar = .LinkSources(1)    'Step 3 - Update all links in formula
        If Not IsEmpty(ar) Then
            For n = 1 To UBound(ar)
                .ChangeLink Name:=ar(n), _
                    NewName:=.Name, Type:=xlExcelLinks
            Next
        End If
    End With
'Step 4 - Delete unused sheets
    Sheets("Dashboard").Delete
    Sheets("LOG").Delete
Application.DisplayAlerts = True
OptimizedMode False
Set FolderPicker = Nothing
mypath = vbNullString: ar = Empty: n = Empty
ThisWorkbook.Save
MsgBox "Report has been exported"
End Sub

