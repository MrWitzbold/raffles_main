Sub GenerateNumbers()
    Dim doc As Document
    Dim tbl As table
    Dim i As Long
    Dim j As Long
    Dim ji As Long
    
    Set doc = ActiveDocument

    ' Initialize counters
    j = 1001
    ji = 1010

    ' Iterate through all tables in the document
    For Each tbl In doc.Tables
    
        For i = 1 To tbl.Rows.Count
            tbl.Cell(i + 1, 1).Range.Text = CStr(j) & " - " & CStr(ji)
            j = j + 10
            ji = ji + 10
        Next i

    Next tbl
End Sub
