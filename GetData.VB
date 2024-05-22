Sub getData()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim partType As String
    Dim wsDict As Object
    Dim wsName As String
    
    Set wsSource = ThisWorkbook.Sheets("Sheet1")
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row 
    
    ' Create a dictionary to map part types to their respective worksheets
    Set wsDict = CreateObject("Scripting.Dictionary")
    
    ' Add part types and their corresponding worksheet names
    wsDict.Add "PRO", "PRO"
    wsDict.Add "CHK", "CHK"
    wsDict.Add "MBK", "MBK"
    wsDict.Add "MBS", "MBS"
    wsDict.Add "MEM", "MEM"
    wsDict.Add "MOD", "MOD"
    wsDict.Add "MMP", "MMP"
    
    ' Loop through each row in the source sheet
    For i = 2 To lastRow
        partType = Left(wsSource.Cells(i, 1).Value, 3) ' Get the first 3 characters
        
        ' Check if the part type is in the dictionary
        If wsDict.Exists(partType) Then
            wsName = wsDict(partType)
            
            ' Check if the worksheet already exists, if not, create it
            On Error Resume Next
            Set wsDest = ThisWorkbook.Sheets(wsName)
            If wsDest Is Nothing Then
                Set wsDest = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
                wsDest.Name = wsName
            End If
            On Error GoTo 0
            
            ' Copy the row to the destination worksheet
            wsSource.Rows(i).Copy
            wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Offset(1, 0).PasteSpecial xlPasteValues
        End If
    Next i
    
    Application.CutCopyMode = False
End Sub
