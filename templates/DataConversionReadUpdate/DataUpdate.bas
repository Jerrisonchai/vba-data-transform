Attribute VB_Name = "DataUpdate"
'Updating new data back into Raw Data based on validation
Sub UpdateFOC()
Call capturetime
Call MyShape_Click
Dim startTime As Date
Dim endTime As Date
Dim timetaken As Date

startTime = Now()
    Dim Slave As Worksheet
    Dim lrT As Long 'data last row
    Dim i As Long
    Set Slave = ThisWorkbook.Worksheets("Data")
    lrT = Slave.Cells(Rows.Count, 2).End(xlUp).Row
    With Slave
        For i = 2 To lrT
'        If Slave.Range("j" & i).Value = "FOC" Then
'            Else
            If Slave.Range("H" & i).Value Like "*FOC*" Or Slave.Range("H" & i).Value Like "*Reject*" Or Slave.Range("H" & i).Value Like "*Cancel*" Then
                Slave.Range("F" & i).Value = "FOC"
            End If
'        End If
        Next i
    End With
    Set Slave = Nothing
    lrT = 0
    i = 0
    
    
 MsgBox "Data FOC updated"
endTime = Now()
timetaken = startTime - endTime

[Status].Value = "Success"
[Start_Time].Value = startTime
[Time_Taken].Value = Format(timetaken, "HH:MM:SS")
[UserName].Value = Environ("Username")
 Call captureendtime
End Sub
