Private Sub ConvertPipeTablesToTable(ByVal doc As Document, ByVal rng As Range)
    Dim i As Long, j As Long
    Dim startIdx As Long, endIdx As Long
    Dim sepText As String, rowText As String
    Dim data() As String
    Dim rows As Long, cols As Long
    Dim guard As Long

    i = 1
    Do While i <= rng.Paragraphs.Count
        guard = guard + 1
        If guard > 20000 Then
            
            Exit Do
        End If

        rowText = CleanParaText(rng.Paragraphs(i).Range.Text)

        If IsPipeRow(rowText) And i + 1 <= rng.Paragraphs.Count Then
            sepText = CleanParaText(rng.Paragraphs(i + 1).Range.Text)
            If IsSeparatorLine(sepText) Then
                startIdx = i
                endIdx = i + 1
                j = i + 2
                Do While j <= rng.Paragraphs.Count
                    rowText = CleanParaText(rng.Paragraphs(j).Range.Text)
                    If IsPipeRow(rowText) Then
                        endIdx = j
                        j = j + 1
                    Else
                        Exit Do
                    End If
                Loop

                rows = endIdx - startIdx          ' header + data rows
                cols = PipeColumnCount(rng.Paragraphs(startIdx).Range.Text)

                If rows >= 1 And cols >= 2 Then
                    ' Snapshot every cell before touching the document, so
                    ' index shifts from the deletion cannot corrupt the data.
                    ReDim data(0 To rows - 1, 0 To cols - 1)
                    Dim cells As String()
                    For j = 0 To rows - 1
                        cells = SplitPipeRow(rng.Paragraphs(startIdx + j).Range.Text, cols)
                        Dim c As Long
                        For c = 0 To cols - 1
                            data(j, c) = cells(c)
                        Next c
                    Next j

                    ' Delete the block, then build the table from the snapshot.
                    Dim blockStart As Long
                    Dim blockRange As Range
                    blockStart = rng.Paragraphs(startIdx).Range.Start
                    Set blockRange = doc.Range( _
                        Start:=blockStart, _
                        End:=rng.Paragraphs(endIdx).Range.End)
                    blockRange.Delete

                    BuildTableFromData doc, blockStart, data, rows, cols

                    i = startIdx + 1
                Else
                    i = i + 1
                End If
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub