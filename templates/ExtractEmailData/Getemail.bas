Attribute VB_Name = "Getemail"
Option Explicit
Sub GetInboxItems()
    Dim ol As Outlook.Application, ns As Outlook.Namespace, fol As Outlook.Folder, i As Object, mi As Outlook.MailItem, n As Long
    Set ol = New Outlook.Application
    Set ns = ol.GetNamespace("MAPI")
    Set fol = ns.GetDefaultFolder(olFolderInbox)
    Set fol = fol.Folders(Sheets("Dashboard").Range("C16").Value)
    Sheets("Data").Range("A2", Sheets("Data").Range("A2").End(xlDown).End(xlToRight)).Clear
    n = 1
    For Each i In fol.Items
        If i.Class = olMail Then
            Set mi = i
            n = n + 1
            Sheets("Data").Cells(n, 1).Value = mi.SenderName
            Sheets("Data").Cells(n, 2).Value = mi.Subject
            Sheets("Data").Cells(n, 3).Value = mi.ReceivedTime
        End If
    Next i
End Sub
Sub GetEmailDetails()
    Dim ol As Outlook.Application, ns As Outlook.Namespace, fol As Outlook.Folder, i As Object, mi As Outlook.MailItem, FilterText As String, SubjectText As String, ReceivedTime As String, doc As Word.Document
    MsgBox "Please open Outlook apps first!"
    If Application.Intersect(ActiveCell, Range("A2", Range("A2").End(xlDown))) Is Nothing Then
        MsgBox "Please select a sender name first!", vbCritical
        Exit Sub
    End If
    Set ol = New Outlook.Application
    Set ns = ol.GetNamespace("MAPI")
    Set fol = ns.GetDefaultFolder(olFolderInbox)
    Set fol = fol.Folders(Sheets("Dashboard").Range("C16").Value)
    SubjectText = ActiveCell.Offset(0, 1).Value
    ReceivedTime = Format(ActiveCell.Offset(0, 2).Value, "d/m/yyyy")
    Select Case LCase(Left(SubjectText, 4))
        Case "re: ", "fw: "
            SubjectText = Mid(SubjectText, 5)
    End Select
    FilterText = "[SenderName] = '" & ActiveCell.Value & "'"
    FilterText = FilterText & " AND [Subject] = '" & SubjectText & "'"
    FilterText = FilterText & " AND [ReceivedTime] >= '#" & ReceivedTime & "#'"
    Set i = fol.Items.Find(FilterText)
    If i Is Nothing Then
        MsgBox "Nothing was found", vbExclamation
        Exit Sub
    End If
    If i.Class <> olMail Then
        MsgBox "Item is not an email", vbExclamation
        Exit Sub
    End If
    Set mi = i
    Set doc = mi.GetInspector.WordEditor
    doc.Range.Copy
    ActiveCell.Offset(0, 1).PasteSpecial
End Sub
