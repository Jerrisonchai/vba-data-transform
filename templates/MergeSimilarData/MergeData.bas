Attribute VB_Name = "MergeData"
Option Explicit
Sub CopyData2()
Call capturetime
Call MyShape_Click
Dim startTime As Date, endTime As Date, timetaken As Date
startTime = Now()
    Dim sFolder As String, FolderPicker As FileDialog, usedrow As Long, mypath As String, sFile As String, wshT As Worksheet, wshO As Worksheet, t As Long, wbkS As Workbook, wshS As Worksheet, s As Long, m As Long, data_lastrow2 As Long
    sFolder = Sheets("Dashboard").Range("C18").Value
    OptimizedMode True
    Set wshT = ThisWorkbook.Sheets("Data")
    Set wshO = ThisWorkbook.Sheets("Dashboard")
    wshT.AutoFilterMode = False
    usedrow = wshT.Range("F" & wshT.Rows.Count).End(xlUp).Row + 2
    wshT.Range("A1:BC" & usedrow).Clear
    t = 2
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
        With FolderPicker
            .Title = "Please Choose One"
            .InitialFileName = sFolder
            .AllowMultiSelect = False
            .ButtonName = "Confirm"
                If .Show = -1 Then
                    mypath = .SelectedItems(1)
                    Else
                        End
                End If
        End With
    Sheets("Dashboard").Range("C18").Value = mypath
    sFolder = Sheets("Dashboard").Range("C18").Value & "\"
    sFile = Dir(sFolder & "*.xlsx*")
    ' Loop through the files
    Do While sFile <> ""
        On Error Resume Next
        Set wbkS = Workbooks.Open(sFolder & sFile)
        Set wshS = wbkS.Worksheets(1)
        ' Get the last used row
        m = wshS.Range("A" & wshS.Rows.Count).End(xlUp).Row
        If m < 2 Then
            wbkS.Close Savechanges:=False
            ' Get next filename
            sFile = Dir
        Else            ' Copy range
            wshS.Range("A" & wshT.[Top_Row_Number] + 1 & ":BB" & wshT.[Top_Row_Number] + 1).Copy Destination:=wshT.Range("B1")
            wshT.Range("A1").Value = "FileName"
            wshS.Range("A" & wshT.[Top_Row_Number] + 2 & ":BB" & m).Copy Destination:=wshT.Range("B" & t)
            wshT.Range("A" & t & ":A" & t + m - wshT.[Top_Row_Number] - 2).Value = sFile
            ' Increment target row
            t = t + m - wshT.[Top_Row_Number]
            Application.CutCopyMode = False
            wbkS.Close Savechanges:=False
            ' Get next filename
            sFile = Dir
        End If
    Loop
    wshT.Activate
    wshT.Range("A1:A99999").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    MsgBox "Data has been updated"
    OptimizedMode False
    wshO.Activate
    Set wshT = Nothing
    Set wshO = Nothing
    Set wbkS = Nothing
    Set wshS = Nothing
    Set FolderPicker = Nothing
    usedrow = Empty: m = Empty: t = Empty
endTime = Now(): timetaken = startTime - endTime
[Status].Value = "Success": [Start_Time].Value = startTime: [Time_Taken].Value = Format(timetaken, "HH:MM:SS"): [UserName].Value = Environ("UserName")
Call captureendtime
End Sub
Sub CopyData3()
Call capturetime
Call MyShape_Click
Dim startTime As Date, endTime As Date, timetaken As Date
startTime = Now()
    Dim sFolder As String, FolderPicker As FileDialog, usedrow As Long, mypath As String, sFile As String, wshT As Worksheet, wshO As Worksheet, t As Long, wbkS As Workbook, wshS As Worksheet, s As Long, m As Long, data_lastrow2 As Long
    sFolder = Sheets("Dashboard").Range("C18").Value
    OptimizedMode True
    Set wshT = ThisWorkbook.Sheets("Checking")
    Set wshO = ThisWorkbook.Sheets("Dashboard")
    wshT.AutoFilterMode = False
    usedrow = wshT.Range("A" & wshT.Rows.Count).End(xlUp).Row + 2
    wshT.Range("A1:BC" & usedrow).Clear
    ' First available target row
    t = 2
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
        With FolderPicker
            .Title = "Please Choose One"
            .InitialFileName = sFolder
            .AllowMultiSelect = False
            .ButtonName = "Confirm"
                If .Show = -1 Then
                    mypath = .SelectedItems(1)
                    Else
                        End
                End If
        End With
    Sheets("Dashboard").Range("C18").Value = mypath
    sFolder = Sheets("Dashboard").Range("C18").Value & "\"
    sFile = Dir(sFolder & "*.xlsx*")
    ' Loop through the files
    Do While sFile <> ""
        On Error Resume Next
        Set wbkS = Workbooks.Open(sFolder & sFile)
        Set wshS = wbkS.Worksheets(1)
        ' Get the last used row
        m = wshS.Range("A" & wshS.Rows.Count).End(xlUp).Row
        If m < 11 Then
            wbkS.Close Savechanges:=False
            ' Get next filename
            sFile = Dir
        Else        ' Copy range
            wshT.Range("A1").Value = "FileName"
            wshT.Range("B1").Value = "Data Count"
            wshT.Range("B" & t).Value = m - wshT.[Top_Row_Number] - 1
            wshT.Range("A" & t).Value = sFile
            ' Increment target row
            t = t + 1
            Application.CutCopyMode = False
            wbkS.Close Savechanges:=False
            ' Get next filename
            sFile = Dir
        End If
    Loop
    wshT.Activate
    wshT.Range("A1:A99999").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    data_lastrow2 = wshT.Range("A" & wshT.Rows.Count).End(xlUp).Row
    wshT.Range("D1").Value = "Countif - Data per file"
    wshT.Range("D2").Value = "=COUNTIF(Data!C[-3],Checking!RC[-3])"
    wshT.Range("E1").Value = "Sumif - Amount per file"
    wshT.Range("E2").Value = "=SUMIFS(Data!C[7],Data!C[-4],Checking!RC[-4])"
    wshT.Range("D2:E" & data_lastrow2).Select
    Selection.FillDown
    MsgBox "Data-Checking has been updated"
    OptimizedMode False
    wshO.Activate
    Set wshT = Nothing
    Set wshO = Nothing
    Set wbkS = Nothing
    Set wshS = Nothing
    Set FolderPicker = Nothing
endTime = Now(): timetaken = startTime - endTime
[Status].Value = "Success": [Start_Time].Value = startTime: [Time_Taken].Value = Format(timetaken, "HH:MM:SS"): [UserName].Value = Environ("UserName")
Call captureendtime
End Sub
