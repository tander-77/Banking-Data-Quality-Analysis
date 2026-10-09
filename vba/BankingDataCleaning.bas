Attribute VB_Name = "Module1"
Sub FIxMonthNames()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim corrections As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
    
    For i = 2 To lastRow
    
        Select Case LCase(Trim(ws.Cells(i, "K").Value))
        
            Case "january"
                ws.Cells(i, "K").Value = "jan"
                corrections = corrections + 1
            
            Case "june"
                ws.Cells(i, "K").Value = "jun"
                corrections = corrections + 1
                
            Case "july"
                ws.Cells(i, "K").Value = "jul"
                corrections = corrections + 1
            
            Case "august"
                ws.Cells(i, "K").Value = "aug"
                corrections = corrections + 1
                
            Case "november"
                ws.Cells(i, "K").Value = "nov"
                corrections = corrections + 1
                
        End Select
        
    Next i
        
    MsgBox corrections & " month values corrected."
        
End Sub

Sub FixDurationValues()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim corrections As Long
    Dim durationValue As Double
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    
    For i = 2 To lastRow
    
        If IsNumeric(ws.Cells(i, "L").Value) Then
    
        durationValue = ws.Cells(i, "L").Value
        
            If durationValue <> Int(durationValue) Then
        
            ws.Cells(i, "L").Value = Round(durationValue, 0)
            corrections = corrections + 1
            
            End If
        
        End If
    
    Next i
    
    MsgBox corrections & " duration values corrected."

End Sub

Sub FixMissingAges()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim corrections As Long
    Dim medianAge As Double
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    medianAge = Application.WorksheetFunction.Median(ws.Range("A2:A" & lastRow))
    
    For i = 2 To lastRow
    
        If Trim(ws.Cells(i, "A").Value) = "" Then
    
            ws.Cells(i, "A").Value = medianAge
            corrections = corrections + 1
        
        End If
    
    Next i
    
    MsgBox corrections & " missing age values corrected using a median age of " & medianAge & "."

End Sub

Sub FixMissingMaritalStatus()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim corrections As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    For i = 2 To lastRow
    
        If Trim(ws.Cells(i, "C").Value) = "" Then
        
            ws.Cells(i, "C").Value = "unknown"
            corrections = corrections + 1
            
        End If
        
    Next i
    
    MsgBox corrections & " missing marital status values corrected to unknown."
    
End Sub

Sub FixContactValues()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim corrections As Long
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    For i = 2 To lastRow
    
        If UCase(Trim(ws.Cells(i, "I").Value)) = "N/A" Then
        
            ws.Cells(i, "I").Value = "unknown"
            corrections = corrections + 1
        
        End If
        
    Next i
    
    MsgBox corrections & " contact N/A values corrected to unknown."
End Sub

Sub FixPreviousOutcome()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim corrections As Long
    Dim cellValue As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "P").End(xlUp).Row
    
    For i = 2 To lastRow
    
        cellValue = UCase(Trim(ws.Cells(i, "P").Value))
        
        If cellValue = "N/A" Or cellValue = "NULL" Or cellValue = "" Then
        
            ws.Cells(i, "P").Value = "unknown"
            corrections = corrections + 1
            
        End If
        
    Next i
    
    MsgBox corrections & " previous outcome values standardized to unknown."

End Sub

Sub ValidateCleanedData()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim issues As Long
    Dim cellValue As String
    
    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    For i = 2 To lastRow
    
        If Trim(ws.Cells(i, "A").Value) = "" Then
            issues = issues + 1
        End If
        
        If Trim(ws.Cells(i, "C").Value) = "" Then
            issues = issues + 1
        End If
        
        If UCase(Trim(ws.Cells(i, "I").Value)) = "N/A" Then
            issues = issues + 1
        End If
        
        cellValue = UCase(Trim(ws.Cells(i, "P").Value))
        
        If cellValue = "N/A" Or cellValue = "NULL" Or cellValue = "" Then
            issues = issues + 1
        End If
        
    Next i
    
    If issues = 0 Then
    
        MsgBox "Validation complete! No targeted data quality issues remain."
    
    Else
    
        MsgBox issues & " data quality issues still require review."
        
    End If
    
End Sub
