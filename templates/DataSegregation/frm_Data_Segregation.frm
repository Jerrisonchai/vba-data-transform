VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frm_Data_Segregation 
   Caption         =   "Data Segregation Tool"
   ClientHeight    =   8820.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11250
   OleObjectBlob   =   "frm_Data_Segregation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frm_Data_Segregation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CheckBox1_Change()

If Me.CheckBox1.Value = True Then
    Me.TextBox2.Enabled = True
    Me.CommandButton5.Enabled = True
Else
    Me.TextBox2.Enabled = False
    Me.CommandButton5.Enabled = False
End If

End Sub


Private Sub CommandButton1_Click()
Call Right_List_Box
Call Primay_Field_List
End Sub

Private Sub CommandButton2_Click()

Me.ListBox2.AddItem Me.ListBox1.Value
Me.ListBox1.RemoveItem Me.ListBox1.ListIndex
Call Primay_Field_List
End Sub

Private Sub CommandButton3_Click()

Me.ListBox1.AddItem Me.ListBox2.Value
Me.ListBox2.RemoveItem Me.ListBox2.ListIndex
Call Primay_Field_List
End Sub

Private Sub CommandButton4_Click()
Call Left_List_Box
Call Primay_Field_List
End Sub

Private Sub CommandButton5_Click()

Dim fldr As FileDialog
Set fldr = Application.FileDialog(msoFileDialogFolderPicker)

fldr.AllowMultiSelect = False
fldr.Title = "Please select a folder to save your files"

If fldr.Show = -1 Then
    Me.TextBox2.Value = fldr.SelectedItems(1)
End If

End Sub

Private Sub CommandButton6_Click()

'''validation

If Me.CheckBox1.Value = True Then
    If Me.TextBox2.Value = "" Or Dir(Me.TextBox2.Value, vbDirectory) = "" Then
    MsgBox "Invalid folder path"
    Exit Sub
    End If
End If

'''option button1
If Me.OptionButton1.Value = True Then
    Multiple_Sheet_in_one_file
ElseIf Me.OptionButton2.Value = True Then
    Call Multiple_file_with_one_sheet
ElseIf Me.OptionButton3.Value = True Then
    Call Multiple_file_with_multiple_sheet
End If



End Sub

Private Sub CommandButton7_Click()

Me.OptionButton1.Value = True
Me.CheckBox1.Value = False
Me.TextBox2.Value = ""

Call Left_List_Box
Call Primay_Field_List

End Sub

Private Sub OptionButton2_Change()
If Me.OptionButton2.Value = True Then
    Me.CheckBox2.Enabled = True
Else
    Me.CheckBox2.Enabled = False
    Me.CheckBox2.Value = False
End If
End Sub


Private Sub UserForm_Activate()
Call Left_List_Box
End Sub

Sub Left_List_Box()
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Data")

Dim i As Integer

Me.ListBox1.Clear
Me.ListBox2.Clear

For i = 1 To Application.WorksheetFunction.CountA(sh.Range("1:1"))
 Me.ListBox1.AddItem sh.Cells(1, i).Value
Next i

End Sub


Sub Right_List_Box()
Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Data")

Dim i As Integer

Me.ListBox1.Clear
Me.ListBox2.Clear

For i = 1 To Application.WorksheetFunction.CountA(sh.Range("1:1"))
 Me.ListBox2.AddItem sh.Cells(1, i).Value
Next i

End Sub


Sub Primay_Field_List()
Me.ComboBox1.Clear
Dim i As Integer
If Me.ListBox2.ListCount > 1 Then
    Me.ComboBox1.Enabled = True
    Me.OptionButton3.Enabled = True
    
    For i = 0 To Me.ListBox2.ListCount - 1
        Me.ComboBox1.AddItem Me.ListBox2.List(i)
    Next i
    
Else

    Me.ComboBox1.Enabled = False
    Me.OptionButton3.Enabled = False
    Me.OptionButton1.Value = True
End If

End Sub


Sub Multiple_Sheet_in_one_file()

Dim sh As Worksheet
Dim nwb As Workbook
Dim nsh As Worksheet

Dim i As Integer
Dim lr As Long

Dim support As Worksheet
Set support = ThisWorkbook.Sheets("Support")
support.Cells.Clear

Set sh = ThisWorkbook.Sheets("Data")
sh.AutoFilterMode = False

'''Copy unique list of selected field
lr = sh.Range("A" & Application.Rows.Count).End(xlUp).Row
Dim Column_Number As Integer

For i = 0 To Me.ListBox2.ListCount - 1
    Column_Number = Application.WorksheetFunction.Match(Me.ListBox2.List(i), sh.Range("1:1"), 0)
    sh.Range(sh.Cells(1, Column_Number), sh.Cells(lr, Column_Number)).AdvancedFilter xlFilterInPlace, , , True
    sh.Range(sh.Cells(1, Column_Number), sh.Cells(lr, Column_Number)).SpecialCells(xlCellTypeVisible).Copy
    support.Cells(1, i + 1).PasteSpecial xlPasteAll
    sh.ShowAllData
    
    
Next i

Dim irow, icol As Integer
Set nwb = Workbooks.Add

For icol = 1 To Application.WorksheetFunction.CountA(support.Range("1:1"))
    For irow = 2 To support.Cells(Application.Rows.Count, icol).End(xlUp).Row
        Column_Number = Application.WorksheetFunction.Match(support.Cells(1, icol), sh.Range("1:1"), 0)
        sh.UsedRange.AutoFilter Column_Number, support.Cells(irow, icol).Value
        sh.UsedRange.Copy
        Set nsh = nwb.Sheets.Add(after:=nwb.Sheets(nwb.Sheets.Count))
        sh.AutoFilterMode = False
        
        nsh.Range("A1").PasteSpecial xlPasteAll
        nsh.Name = support.Cells(irow, icol).Value
        
    Next irow
Next icol
    
If Me.CheckBox1.Value = True Then
    nwb.SaveAs Me.TextBox2.Value & Application.PathSeparator & "Segregated Data.xlsx"
    nwb.Worksheets(1).Delete
    nwb.Save
    nwb.Close
End If

End Sub


Sub Multiple_file_with_one_sheet()

Dim sh As Worksheet
Dim nwb As Workbook
Dim nsh As Worksheet

Dim i As Integer
Dim lr As Long

Dim support As Worksheet
Set support = ThisWorkbook.Sheets("Support")
support.Cells.Clear

Set sh = ThisWorkbook.Sheets("Data")
sh.AutoFilterMode = False

'''Copy unique list of selected field
lr = sh.Range("A" & Application.Rows.Count).End(xlUp).Row
Dim Column_Number As Integer

For i = 0 To Me.ListBox2.ListCount - 1
    Column_Number = Application.WorksheetFunction.Match(Me.ListBox2.List(i), sh.Range("1:1"), 0)
    sh.Range(sh.Cells(1, Column_Number), sh.Cells(lr, Column_Number)).AdvancedFilter xlFilterInPlace, , , True
    sh.Range(sh.Cells(1, Column_Number), sh.Cells(lr, Column_Number)).SpecialCells(xlCellTypeVisible).Copy
    support.Cells(1, i + 1).PasteSpecial xlPasteAll
    sh.ShowAllData
    
    
Next i

Dim irow, icol As Integer


For icol = 1 To Application.WorksheetFunction.CountA(support.Range("1:1"))
    For irow = 2 To support.Cells(Application.Rows.Count, icol).End(xlUp).Row
        Set nwb = Workbooks.Add
        Column_Number = Application.WorksheetFunction.Match(support.Cells(1, icol), sh.Range("1:1"), 0)
        sh.UsedRange.AutoFilter Column_Number, support.Cells(irow, icol).Value
        sh.UsedRange.Copy
        Set nsh = nwb.Sheets.Add(after:=nwb.Sheets(nwb.Sheets.Count))
        sh.AutoFilterMode = False
        
        nsh.Range("A1").PasteSpecial xlPasteAll
        nsh.Name = support.Cells(irow, icol).Value
        
        If Me.CheckBox1.Value = True Then
        
            If Me.CheckBox2.Value = True Then
                If Dir(Me.TextBox2.Value & Application.PathSeparator & support.Cells(1, icol), vbDirectory) = "" Then
                    MkDir (Me.TextBox2.Value & Application.PathSeparator & support.Cells(1, icol))
                    
                End If
                nwb.SaveAs Me.TextBox2.Value & Application.PathSeparator & support.Cells(1, icol) & Application.PathSeparator & support.Cells(irow, icol).Value
            Else
                nwb.SaveAs Me.TextBox2.Value & Application.PathSeparator & support.Cells(irow, icol).Value
            End If
            nwb.Worksheets(1).Delete
            nwb.Save
            nwb.Close
        End If

    Next irow
Next icol
    


End Sub


Sub Multiple_file_with_multiple_sheet()

Dim sh As Worksheet
Dim nwb As Workbook
Dim nsh As Worksheet

Dim i As Integer
Dim lr As Long

Dim support As Worksheet
Set support = ThisWorkbook.Sheets("Support")
support.Cells.Clear

Set sh = ThisWorkbook.Sheets("Data")
sh.AutoFilterMode = False

'''Copy unique list of selected field
lr = sh.Range("A" & Application.Rows.Count).End(xlUp).Row
Dim Column_Number As Integer

For i = 0 To Me.ListBox2.ListCount - 1
    Column_Number = Application.WorksheetFunction.Match(Me.ListBox2.List(i), sh.Range("1:1"), 0)
    sh.Range(sh.Cells(1, Column_Number), sh.Cells(lr, Column_Number)).AdvancedFilter xlFilterInPlace, , , True
    sh.Range(sh.Cells(1, Column_Number), sh.Cells(lr, Column_Number)).SpecialCells(xlCellTypeVisible).Copy
    support.Cells(1, i + 1).PasteSpecial xlPasteAll
    sh.ShowAllData
    
    
Next i

Dim irow, icol As Integer
Dim pri_row As Integer
Dim pri_col As Integer

pri_col = Application.WorksheetFunction.Match(Me.ComboBox1.Value, support.Range("1:1"), 0)

For pri_row = 2 To support.Cells(Application.Rows.Count, pri_col).End(xlUp).Row
    Set nwb = Workbooks.Add
    For icol = 1 To Application.WorksheetFunction.CountA(support.Range("1:1"))
        If pri_col <> icol Then
            For irow = 2 To support.Cells(Application.Rows.Count, icol).End(xlUp).Row
                sh.UsedRange.AutoFilter Application.WorksheetFunction.Match(Me.ComboBox1.Value, sh.Range("1:1"), 0), support.Cells(pri_row, pri_col).Value

                Column_Number = Application.WorksheetFunction.Match(support.Cells(1, icol), sh.Range("1:1"), 0)
                sh.UsedRange.AutoFilter Column_Number, support.Cells(irow, icol).Value
                
                
                If Application.WorksheetFunction.Subtotal(103, sh.Range("A:A")) > 1 Then
                    sh.UsedRange.Copy
                    Set nsh = nwb.Sheets.Add(after:=nwb.Sheets(nwb.Sheets.Count))
                    sh.AutoFilterMode = False
                    
                    nsh.Range("A1").PasteSpecial xlPasteAll
                    nsh.Name = support.Cells(irow, icol).Value
                End If

        
            Next irow
        End If
    Next icol
    If Me.CheckBox1.Value = True Then
        nwb.SaveAs Me.TextBox2.Value & Application.PathSeparator & support.Cells(pri_row, pri_col).Value
        nwb.Worksheets(1).Delete
        nwb.Save
        nwb.Close
    End If
Next pri_row

End Sub
