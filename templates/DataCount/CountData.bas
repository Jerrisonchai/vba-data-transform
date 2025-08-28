Attribute VB_Name = "CountData"
Option Explicit
Sub CountData()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wshO As Worksheet: Set wshO = wb.Sheets("Dashboard")
    Dim wshT As Worksheet: Set wshT = wb.Sheets("Checking")
    Dim wshD As Worksheet: Set wshD = wb.Sheets("Data")
    
    Dim sFolder As String: sFolder = wshO.Range("C20").Value & "\"
    Dim catCol As String: catCol = UCase(wshO.Range("C21").Value) ' Column to count categories from
    Dim countCol As String: countCol = UCase(wshO.Range("C22").Value) ' Column to count data from
    
    ' Get categories from Checking sheet (starting from C1)
    Dim categories() As Variant
    categories = GetCategories(wshT)
    
    ' Validate inputs
    If Dir(sFolder, vbDirectory) = "" Then
        MsgBox "Invalid folder path!", vbExclamation
        Exit Sub
    End If
    
    If Len(catCol) = 0 Or Len(countCol) = 0 Then
        MsgBox "Please specify category and count columns!", vbExclamation
        Exit Sub
    End If
    
    OptimizedMode True
    ResetResultsArea wshT ' Clear previous results
    
    Dim sFile As String: sFile = Dir(sFolder & "*.xlsx*")
    Dim t As Long: t = 2 ' Start row for results
    
    Do While sFile <> ""
        Dim wbkS As Workbook
        Set wbkS = Workbooks.Open(sFolder & sFile, ReadOnly:=True)
        
        Dim wshS As Worksheet: Set wshS = wbkS.Worksheets(1)
        Dim lastRow As Long: lastRow = GetLastRow(wshS, countCol)
        
        If lastRow > 1 Then ' Check if data exists
            ' Write filename and total count
            wshT.Range("A" & t).Value = sFile
            wshT.Range("B" & t).Value = lastRow - 1 ' Subtract header row
            
            ' Count categories
            CountCategories wshS, wshT, catCol, categories, t
            
            ' Copy headers to Data sheet (optional)
            CopyHeaders wshS, wshD
        End If
        
        wbkS.Close False
        t = t + 1
        sFile = Dir
    Loop
    
    OptimizedMode False
    wshO.Activate
    MsgBox "Process completed successfully!", vbInformation
End Sub

Private Function GetCategories(wshT As Worksheet) As Variant
    Dim lastCol As Long
    lastCol = wshT.Cells(1, Columns.Count).End(xlToLeft).Column
    If lastCol < 3 Then lastCol = 3 ' Minimum at column C
    GetCategories = wshT.Range("C1", wshT.Cells(1, lastCol)).Value
End Function

Private Sub CountCategories(wshS As Worksheet, wshT As Worksheet, catCol As String, categories() As Variant, t As Long)
    Dim i As Long
    For i = 1 To UBound(categories, 2)
        Dim catValue As String
        catValue = categories(1, i)
        
        If Len(catValue) > 0 Then
            Dim countResult As Long
            countResult = Application.CountIfs( _
                wshS.Columns(catCol), _
                catValue _
            )
            wshT.Cells(t, 2 + i).Value = countResult
        End If
    Next i
End Sub

Private Sub CopyHeaders(wshS As Worksheet, wshD As Worksheet)
    Static headersCopied As Boolean
    If Not headersCopied Then
        Dim lastRow As Long
        lastRow = wshD.Cells(wshD.Rows.Count, "A").End(xlUp).Row
        If lastRow = 1 Then
            wshS.Rows(1).Copy wshD.Range("A1")
        End If
        headersCopied = True
    End If
End Sub

Private Function GetLastRow(ws As Worksheet, col As String) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Private Sub ResetResultsArea(wshT As Worksheet)
    Dim lastRow As Long
    lastRow = wshT.Cells(wshT.Rows.Count, "A").End(xlUp).Row
    If lastRow > 1 Then
        wshT.Range("A2:Z" & lastRow).ClearContents
    End If
End Sub

