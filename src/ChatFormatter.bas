Attribute VB_Name = "ChatFormatter"
Option Explicit

'=======================================================================
'  Chat Formatter for Word
'  Converts text copied from AI chats (DeepSeek, ChatGPT, Gemini, ...)
'  into properly formatted Word content: headings, bold, italic,
'  inline code, code blocks, links, blockquotes, bullet/numbered
'  lists and real Word tables (Markdown pipe tables and TSV).
'
'  Entry points
'    FormatChatText            format the whole active document (Add-ins tab)
'    FormatSelection           format only the current selection
'    ShowSettings              show current configuration
'
'  Design notes
'    The document is read once into a snapshot of paragraph offsets and
'    text; every block is then rewritten from that snapshot, back to
'    front.  Nothing is indexed by live paragraph number inside a loop,
'    so the cost is proportional to document size rather than its square.
'
'    Fenced blocks are classified in that same single pass, so a
'    ```tsv fence becomes a real table and is never afterwards restyled
'    as a code block - which is what the previous pipeline did.
'
'    Settings live in a plain ini file under the user's home directory.
'    No registry, no WScript.Shell, no COM: the template runs the same
'    way on Windows and on macOS.
'=======================================================================

Public Const CF_VERSION As String = "2.0.0"
Public Const CF_TOOLBAR As String = "ChatFormatterToolbar"

Private Const FENCE As String = "```"

' Block kinds
Private Const BK_PLAIN As Long = 0      ' ordinary text, no tabs
Private Const BK_CODE As Long = 1       ' fenced block
Private Const BK_TSV As Long = 2        ' fenced block tagged tsv/tab/table
Private Const BK_TABRUN As Long = 3     ' unfenced run of tab separated rows
Private Const BK_PIPETABLE As Long = 4  ' markdown pipe table (| a | b |)

' Per-line roles produced by the plain-block pass
Private Const LK_TEXT As Long = 0
Private Const LK_HEADING As Long = 1
Private Const LK_BULLET As Long = 2
Private Const LK_NUMBERED As Long = 3
Private Const LK_QUOTE As Long = 4

' Inline style slots
Private Const IL_BOLD As Long = 1
Private Const IL_ITALIC As Long = 2
Private Const IL_STRIKE As Long = 3
Private Const IL_CODE As Long = 4
Private Const IL_LINK As Long = 5

'-----------------------------------------------------------------------
'  Module state
'-----------------------------------------------------------------------

Private mBusy As Boolean
Private mSilent As Boolean
Private mBullet As String
Private mTotalSteps As Long
Private mDoneSteps As Long
Private mBeatCount As Long

' One contiguous run of paragraphs.  First/Last are 1-based indexes into
' the snapshot; Lang is the fence language when the block came from one.
Private Type Blk
    Kind As Long
    First As Long
    Last As Long
    Lang As String
    Closed As Boolean
End Type


'=======================================================================
'  Public API
'=======================================================================

Public Sub FormatChatText()
    Dim doc As Document
    Dim resp As VbMsgBoxResult

    If Documents.Count = 0 Then
        Notify "No document is open."
        Exit Sub
    End If
    Set doc = ActiveDocument

    resp = MsgBox( _
        "Convert TSV and pipe tables to Word tables too?" & vbCrLf & vbCrLf & _
        "Yes = Markdown formatting + tables" & vbCrLf & _
        "No = Markdown formatting only", _
        vbYesNo + vbQuestion, "Chat Formatter")

    RunPipeline doc, doc.Content, (resp = vbYes), True
End Sub


Public Sub FormatSelection()
    Dim doc As Document
    Dim rng As Range

    If Documents.Count = 0 Then
        Notify "No document is open."
        Exit Sub
    End If
    Set doc = ActiveDocument
    Set rng = CurrentRange(doc)
    If rng Is Nothing Then
        Notify "No document available for formatting."
        Exit Sub
    End If

    RunPipeline doc, rng, (GetSetting("Tables", "1") = "1"), True
End Sub


Public Sub CF_SilentFormat()
    Dim doc As Document
    Dim rng As Range

    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument
    Set rng = CurrentRange(doc)
    If rng Is Nothing Then Exit Sub

    On Error GoTo Done
    mSilent = True
    RunPipeline doc, rng, (GetSetting("Tables", "1") = "1"), False
Done:
    mSilent = False
End Sub


Public Sub ShowSettings()
    Dim tables As String

    If GetSetting("Tables", "1") = "1" Then tables = "Enabled" Else tables = "Disabled"

    Notify "Chat Formatter " & CF_VERSION & vbCrLf & vbCrLf & _
           "Run via: Add-ins tab -> Format Chat Text (or Alt+F8)" & vbCrLf & _
           "Table conversion (FormatSelection): " & tables & vbCrLf & vbCrLf & _
           "Settings file: " & SettingsFile()
End Sub


Public Sub AutoExec()
    On Error Resume Next
    EnsureBullet
    CreateToolbarButton False
    On Error GoTo 0
End Sub


Public Sub AutoOpen()
    AutoExec
End Sub


Public Sub AutoSetup()
    On Error Resume Next
    EnsureBullet
    CreateToolbarButton True
    On Error GoTo 0
End Sub


Public Sub RemoveToolbar()
    Dim cb As CommandBar
    On Error Resume Next
    For Each cb In Application.CommandBars
        If cb.Name = CF_TOOLBAR Then cb.Delete
    Next cb
    On Error GoTo 0
End Sub


' Diagnostics: run the pipeline over the whole document with no dialogs.
Public Sub CF_SilentTest()
    If Documents.Count = 0 Then Exit Sub
    On Error GoTo Done
    mSilent = True
    RunPipeline ActiveDocument, ActiveDocument.Content, True, False
Done:
    mSilent = False
End Sub


Private Function CurrentRange(ByVal doc As Document) As Range
    Dim r As Range

    On Error Resume Next
    Set r = Selection.Range
    If Err.Number <> 0 Or r Is Nothing Then
        Err.Clear
        Set r = doc.Content
    ElseIf r.Start = r.End Then
        Set r = doc.Content
    End If
    Err.Clear
    On Error GoTo 0

    Set CurrentRange = r
End Function


'=======================================================================
'  Responsiveness helpers
'
'  The formatter runs on the UI thread, so Beat hands control back to the
'  host every few dozen paragraph visits and shows progress on the status
'  bar.  mBusy guards against re-entrant execution while Beat yields; it
'  must NOT gate Beat itself, or nothing would ever repaint.
'=======================================================================

Private Sub BeginRun(ByVal totalSteps As Long)
    mBusy = True
    mBeatCount = 0
    mTotalSteps = totalSteps
    mDoneSteps = 0
End Sub


Private Sub EndRun()
    mBusy = False
    mDoneSteps = 0
    mTotalSteps = 0
    On Error Resume Next
    Application.StatusBar = False
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub StepDone(ByVal label As String)
    mDoneSteps = mDoneSteps + 1
    Beat label
End Sub


Private Sub Beat(ByVal label As String)
    mBeatCount = mBeatCount + 1
    If mBeatCount < 40 Then Exit Sub
    mBeatCount = 0

    On Error Resume Next
    If mTotalSteps > 0 Then
        Application.StatusBar = "Chat Formatter: " & label & _
            " (" & mDoneSteps & " / " & mTotalSteps & ")"
    Else
        Application.StatusBar = "Chat Formatter: " & label & " ..."
    End If
    DoEvents
    Err.Clear
    On Error GoTo 0
End Sub


'=======================================================================
'  Pipeline
'=======================================================================

Private Sub RunPipeline(ByVal doc As Document, ByVal rng As Range, _
                        ByVal doTables As Boolean, ByVal notifyUser As Boolean)
    Dim pStart() As Long
    Dim pEnd() As Long
    Dim pText() As String
    Dim blks() As Blk
    Dim nBlocks As Long
    Dim nParas As Long
    Dim i As Long
    Dim lo As Long, hi As Long
    Dim undoStarted As Boolean
    Dim failed As Boolean
    Dim errNum As Long, errDesc As String
    Dim su As Boolean
    Dim pagin As Boolean, hasPagin As Boolean
    Dim spell As Boolean, hasSpell As Boolean
    Dim gram As Boolean, hasGram As Boolean

    If mBusy Then Exit Sub
    If rng Is Nothing Then Exit Sub

    EnsureBullet
    mBusy = True
    su = Application.ScreenUpdating
    On Error GoTo CleanUp

    ' One linear pass fills the snapshot and classifies the blocks.
    nBlocks = ScanBlocks(rng, pStart, pEnd, pText, blks)
    If nBlocks = 0 Then
        EndRun
        If notifyUser Then Notify "There is nothing here to format."
        Exit Sub
    End If
    nParas = UBound(pText)

    BeginRun nBlocks + 2

    On Error Resume Next
    pagin = Application.Options.Pagination
    hasPagin = (Err.Number = 0): Err.Clear
    spell = Application.Options.CheckSpellingAsYouType
    hasSpell = (Err.Number = 0): Err.Clear
    gram = Application.Options.CheckGrammarAsYouType
    hasGram = (Err.Number = 0): Err.Clear
    On Error GoTo CleanUp

    Application.ScreenUpdating = False

    ' Repagination and live proofing dominate the cost of restructuring a
    ' document, so they are suspended for the run and then restored to
    ' whatever the user actually had, rather than forced on.
    On Error Resume Next
    Application.Options.Pagination = False
    Application.Options.CheckSpellingAsYouType = False
    Application.Options.CheckGrammarAsYouType = False
    Err.Clear
    On Error GoTo CleanUp

    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "Chat Formatter"
    undoStarted = (Err.Number = 0)
    Err.Clear
    On Error GoTo CleanUp

    ' Back to front: replacing a block never moves a block still to visit.
    For i = nBlocks To 1 Step -1
        lo = pStart(blks(i).First)
        hi = ClampEnd(doc, pEnd(blks(i).Last))
        StepDone "formatting"
        If hi > lo Then
            Select Case blks(i).Kind
                Case BK_CODE
                    ApplyCodeBlock doc, lo, hi, blks(i), pText
                Case BK_TSV
                    ApplyFenceWithBody doc, lo, hi, blks(i), pText, doTables
                Case BK_TABRUN
                    ApplyTabRun doc, lo, hi, blks(i), pText, doTables
                Case BK_PIPETABLE
                    ApplyPipeTable doc, lo, hi, blks(i), pText, doTables
                Case Else
                    ApplyPlainBlock doc, lo, hi, blks(i), pText
            End Select
        End If
    Next i

CleanUp:
    failed = (Err.Number <> 0)
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If undoStarted Then Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = su
    If hasPagin Then Application.Options.Pagination = pagin
    If hasSpell Then Application.Options.CheckSpellingAsYouType = spell
    If hasGram Then Application.Options.CheckGrammarAsYouType = gram
    Application.StatusBar = False
    If Not failed Then doc.Repaginate
    Err.Clear
    On Error GoTo 0

    EndRun

    If failed Then
        Notify "Error during formatting:" & vbCrLf & errDesc & vbCrLf & _
               "Error code: " & errNum & vbCrLf & vbCrLf & _
               "Undo with Ctrl+Z (Windows) or Cmd+Z (Mac)."
    ElseIf notifyUser Then
        Notify "Formatting completed. Undo with Ctrl+Z (Windows) or Cmd+Z (Mac)."
    End If

End Sub


Private Function ClampEnd(ByVal doc As Document, ByVal v As Long) As Long
    Dim lastPos As Long
    lastPos = doc.Content.End - 1
    If v > lastPos Then v = lastPos
    If v < 0 Then v = 0
    ClampEnd = v
End Function


Private Function ContentEndBeforeCr(ByVal doc As Document, ByVal lo As Long, ByVal hi As Long) As Long
    Dim e As Long
    e = hi
    If e > lo Then
        If doc.Range(e - 1, e).Text = vbCr Then e = e - 1
    End If
    ContentEndBeforeCr = e
End Function


'=======================================================================
'  Snapshot and classification
'
'  Fills pStart(1..n), pEnd(1..n) with paragraph offsets and pText(1..n)
'  with raw paragraph text, then cuts the range into typed blocks.
'  Returns the block count.  This is the only pass that walks the live
'  document.
'=======================================================================

Private Function ScanBlocks(ByVal rng As Range, _
                            ByRef pStart() As Long, _
                            ByRef pEnd() As Long, _
                            ByRef pText() As String, _
                            ByRef blks() As Blk) As Long
    Dim p As Paragraph
    Dim n As Long, i As Long, j As Long
    Dim plainFrom As Long
    Dim s As String, lang As String
    Dim closeIdx As Long
    Dim kind As Long
    Dim fenceLang As String
    Dim count As Long

    If rng.Paragraphs.Count = 0 Then
        ScanBlocks = 0
        Exit Function
    End If

    ReDim pStart(1 To rng.Paragraphs.Count)
    ReDim pEnd(1 To rng.Paragraphs.Count)
    ReDim pText(1 To rng.Paragraphs.Count)
    ReDim blks(1 To 32)

    ' Text inside a table belongs to Word, not to us.  Leaving those
    ' paragraphs out of the snapshot is what stops a second run from
    ' dragging tables that are already finished back into the pipeline.
    n = 0
    For Each p In rng.Paragraphs
        If Not p.Range.Information(wdWithInTable) Then
            AppendParaSegments p.Range.Start, p.Range.End, p.Range.Text, _
                               pStart, pEnd, pText, n
            Beat "scanning"
        End If
    Next p
    If n < 1 Then
        ScanBlocks = 0
        Exit Function
    End If
    ReDim Preserve pStart(1 To n)
    ReDim Preserve pEnd(1 To n)
    ReDim Preserve pText(1 To n)

    count = 0
    plainFrom = 1
    i = 1
    Do While i <= n
        Beat "scanning"
        s = CleanParaText(pText(i))
        If IsFenceLine(s, lang) Then
            fenceLang = lang
            closeIdx = -1
            j = i + 1
            Do While j <= n
                If IsFenceLine(CleanParaText(pText(j)), lang) Then
                    closeIdx = j
                    Exit Do
                End If
                j = j + 1
            Loop

            If closeIdx = -1 Then
                count = FlushPlain(blks, count, plainFrom, i - 1, pStart, pEnd, pText)
                count = AddBlock(blks, count, BK_CODE, i, n, fenceLang, False)
                ScanBlocks = count
                Exit Function
            End If

            count = FlushPlain(blks, count, plainFrom, i - 1, pStart, pEnd, pText)

            kind = BK_CODE
            If IsTsvLang(fenceLang) Then
                If RangeHasTabs(pText, i + 1, closeIdx - 1) Then kind = BK_TSV
            End If
            count = AddBlock(blks, count, kind, i, closeIdx, fenceLang, True)

            i = closeIdx + 1
            plainFrom = i
        Else
            i = i + 1
        End If
    Loop

    count = FlushPlain(blks, count, plainFrom, n, pStart, pEnd, pText)
    ScanBlocks = count
End Function


' Splits Shift+Enter (Chr(11)) soft line breaks inside a paragraph into
' separate snapshot entries whenever raw range length matches text length.
Private Sub AppendParaSegments(ByVal rawStart As Long, ByVal rawEnd As Long, _
                               ByVal rawText As String, _
                               ByRef pStart() As Long, ByRef pEnd() As Long, _
                               ByRef pText() As String, ByRef n As Long)
    Dim segStart As Long
    Dim pos As Long
    Dim seg As String
    Dim segEnd As Long

    If InStr(rawText, Chr$(11)) > 0 And (rawEnd - rawStart) = Len(rawText) Then
        segStart = 1
        Do While segStart <= Len(rawText)
            pos = InStr(segStart, rawText, Chr$(11))
            If pos > 0 Then
                seg = Mid$(rawText, segStart, pos - segStart + 1)
                segEnd = rawStart + pos
                AppendSnapshotEntry rawStart + segStart - 1, segEnd, seg, pStart, pEnd, pText, n
                segStart = pos + 1
            Else
                seg = Mid$(rawText, segStart)
                AppendSnapshotEntry rawStart + segStart - 1, rawEnd, seg, pStart, pEnd, pText, n
                Exit Do
            End If
        Loop
    Else
        AppendSnapshotEntry rawStart, rawEnd, rawText, pStart, pEnd, pText, n
    End If
End Sub


Private Sub AppendSnapshotEntry(ByVal sPos As Long, ByVal ePos As Long, _
                                ByVal txt As String, _
                                ByRef pStart() As Long, ByRef pEnd() As Long, _
                                ByRef pText() As String, ByRef n As Long)
    n = n + 1
    If n > UBound(pStart) Then
        ReDim Preserve pStart(1 To n * 2)
        ReDim Preserve pEnd(1 To n * 2)
        ReDim Preserve pText(1 To n * 2)
    End If
    pStart(n) = sPos
    pEnd(n) = ePos
    pText(n) = txt
End Sub


' Emits the ordinary region between two fence blocks, split so that
' markdown pipe tables and unfenced runs of tab separated rows become
' their own blocks.  Keeping them separate means the text pass never has
' to reason about table insertion.
Private Function FlushPlain(ByRef blks() As Blk, ByVal count As Long, _
                            ByVal firstIdx As Long, ByVal lastIdx As Long, _
                            ByRef pStart() As Long, ByRef pEnd() As Long, _
                            ByRef pText() As String) As Long
    Dim i As Long
    Dim runFrom As Long
    Dim pipeEnd As Long
    Dim s As String
    Dim nextS As String
    Dim dummy As String
    Dim pairedSetext As Boolean

    i = firstIdx
    Do While i <= lastIdx
        s = CleanParaText(pText(i))
        If runFrom > 0 Then
            If InStr(s, vbTab) > 0 And Len(s) > 0 And Not IsListMarked(s) Then
                If pEnd(i - 1) <> pStart(i) Then
                    count = FlushTabRun(blks, count, runFrom, i - 1, pText)
                    runFrom = i
                End If
                i = i + 1
            Else
                count = FlushTabRun(blks, count, runFrom, i - 1, pText)
                runFrom = 0
            End If
        Else
            pipeEnd = FindPipeTableEnd(i, lastIdx, pStart, pEnd, pText)
            If pipeEnd > 0 Then
                count = AddBlock(blks, count, BK_PIPETABLE, i, pipeEnd, "", True)
                i = pipeEnd + 1
            ElseIf InStr(s, vbTab) > 0 And Len(s) > 0 And Not IsListMarked(s) Then
                runFrom = i
                i = i + 1
            ElseIf Len(s) = 0 Then
                If EndsWithSoftBreak(pText(i)) Then
                    count = AddBlock(blks, count, BK_PLAIN, i, i, "", True)
                End If
                i = i + 1
            ElseIf IsFenceLine(s, dummy) Then
                i = i + 1
            Else
                pairedSetext = False
                If i < lastIdx Then
                    If pEnd(i) = pStart(i + 1) Then
                        nextS = CleanParaText(pText(i + 1))
                        If IsSetextLine(nextS) And IsSetextSubject(s) Then
                            count = AddBlock(blks, count, BK_PLAIN, i, i + 1, "", True)
                            i = i + 2
                            pairedSetext = True
                        End If
                    End If
                End If
                If Not pairedSetext Then
                    count = AddBlock(blks, count, BK_PLAIN, i, i, "", True)
                    i = i + 1
                End If
            End If
        End If
    Loop
    If runFrom > 0 Then
        count = FlushTabRun(blks, count, runFrom, lastIdx, pText)
    End If
    FlushPlain = count
End Function


' An unfenced tab run needs at least 2 rows and 2 columns to become a
' table; a single stray tab line falls back to ordinary plain formatting so
' it never turns into a 1-row "columns only" grid.
Private Function FlushTabRun(ByRef blks() As Blk, ByVal count As Long, _
                             ByVal runFrom As Long, ByVal runTo As Long, _
                             ByRef pText() As String) As Long
    Dim k As Long
    Dim rows As Long, maxCols As Long, cols As Long
    Dim s As String

    For k = runFrom To runTo
        s = CleanParaText(pText(k))
        If Len(s) > 0 Then
            rows = rows + 1
            cols = CountChar(s, vbTab) + 1
            If cols > maxCols Then maxCols = cols
        End If
    Next k

    If rows >= 2 And maxCols >= 2 Then
        FlushTabRun = AddBlock(blks, count, BK_TABRUN, runFrom, runTo, "", True)
    Else
        For k = runFrom To runTo
            s = CleanParaText(pText(k))
            If Len(s) > 0 Or EndsWithSoftBreak(pText(k)) Then
                count = AddBlock(blks, count, BK_PLAIN, k, k, "", True)
            End If
        Next k
        FlushTabRun = count
    End If
End Function


Private Function FindPipeTableEnd(ByVal idx As Long, ByVal lastIdx As Long, _
                                  ByRef pStart() As Long, ByRef pEnd() As Long, _
                                  ByRef pText() As String) As Long
    Dim s0 As String, s1 As String, sj As String
    Dim j As Long, endIdx As Long
    Dim dummy As String

    s0 = CleanParaText(pText(idx))
    If Not IsPipeRow(s0) Then Exit Function
    If PipeColumnCount(s0) < 2 Then Exit Function

    ' Standard GFM pipe table: header row followed by | :--- | :--- |
    If idx < lastIdx Then
        If pEnd(idx) = pStart(idx + 1) Then
            s1 = CleanParaText(pText(idx + 1))
            If IsPipeSeparator(s1) Then
                endIdx = idx + 1
                j = idx + 2
                Do While j <= lastIdx
                    If pEnd(j - 1) <> pStart(j) Then Exit Do
                    sj = CleanParaText(pText(j))
                    If Len(sj) = 0 Or IsFenceLine(sj, dummy) Then Exit Do
                    If Not IsPipeRow(sj) Or IsPipeSeparator(sj) Then Exit Do
                    endIdx = j
                    j = j + 1
                Loop
                FindPipeTableEnd = endIdx
                Exit Function
            End If
        End If
    End If

    ' Bordered pipe rows (| a | b |) without a separator row: require >= 2 rows.
    If IsBorderedPipeRow(s0) And Not IsPipeSeparator(s0) Then
        endIdx = idx
        j = idx + 1
        Do While j <= lastIdx
            If pEnd(j - 1) <> pStart(j) Then Exit Do
            sj = CleanParaText(pText(j))
            If Len(sj) = 0 Or IsFenceLine(sj, dummy) Then Exit Do
            If IsBorderedPipeRow(sj) And PipeColumnCount(sj) >= 2 And Not IsPipeSeparator(sj) Then
                endIdx = j
                j = j + 1
            Else
                Exit Do
            End If
        Loop
        If endIdx > idx Then FindPipeTableEnd = endIdx
    End If
End Function


Private Function EndsWithSoftBreak(ByVal rawText As String) As Boolean
    If Len(rawText) = 0 Then Exit Function
    EndsWithSoftBreak = (Right$(rawText, 1) = Chr$(11))
End Function


Private Function IsSetextSubject(ByVal s As String) As Boolean
    Dim dummyNum As Long
    If Len(s) = 0 Then Exit Function
    If AtxLevel(s) > 0 Then Exit Function
    If IsHorizontalRule(s) Then Exit Function
    If Left$(s, 1) = ">" Then Exit Function
    If IsBulletStart(s) Then Exit Function
    If IsNumberedStart(s, dummyNum) Then Exit Function
    IsSetextSubject = True
End Function


Private Function AddBlock(ByRef blks() As Blk, ByVal count As Long, _
                          ByVal kind As Long, ByVal firstIdx As Long, _
                          ByVal lastIdx As Long, ByVal lang As String, _
                          ByVal closed As Boolean) As Long
    Dim n As Long
    n = count + 1
    If n > UBound(blks) Then ReDim Preserve blks(1 To UBound(blks) * 2)
    blks(n).Kind = kind
    blks(n).First = firstIdx
    blks(n).Last = lastIdx
    blks(n).Lang = lang
    blks(n).Closed = closed
    AddBlock = n
End Function


' ``` or ```language.  A longer run of backticks is treated as body text.
Private Function IsFenceLine(ByVal s As String, ByRef lang As String) As Boolean
    Dim rest As String
    If Left(s, Len(FENCE)) <> FENCE Then Exit Function
    rest = Mid(s, Len(FENCE) + 1)
    If InStr(rest, "`") > 0 Then Exit Function
    lang = LCase$(Trim$(rest))
    IsFenceLine = True
End Function


Private Function IsTsvLang(ByVal lang As String) As Boolean
    Select Case LCase$(lang)
        Case "tsv", "tab", "tabs", "table"
            IsTsvLang = True
    End Select
End Function


Private Function RangeHasTabs(ByRef pText() As String, _
                              ByVal firstIdx As Long, ByVal lastIdx As Long) As Boolean
    Dim i As Long
    For i = firstIdx To lastIdx
        If Len(Trim$(pText(i))) > 0 Then
            If InStr(pText(i), vbTab) > 0 Then
                RangeHasTabs = True
                Exit Function
            End If
        End If
    Next i
End Function


' A line this formatter produced on an earlier pass.  It must never be
' re-read as a table row, or every bullet would turn into a one-cell
' table on the second run.
Private Function IsListMarked(ByVal s As String) As Boolean
    If Len(s) < 2 Then Exit Function
    If InStr(s, vbTab) = 0 Then Exit Function
    Dim head As String
    head = Left(s, InStr(s, vbTab) - 1)
    head = Replace(head, " ", "")
    If Len(head) = 0 Then Exit Function
    If head = ChrW$(&H2022) Then IsListMarked = True: Exit Function
    If InStr("-*+", Left$(head, 1)) > 0 And Len(head) = 1 Then IsListMarked = True: Exit Function
    If Len(head) >= 2 And Right$(head, 1) = "." Then
        If IsAllDigits(Left$(head, Len(head) - 1)) Then IsListMarked = True
    End If
End Function


'=======================================================================
'  Block writers
'
'  Each writer owns exactly one block and receives absolute offsets.  The
'  shared idiom is: rewrite the block text once, then re-derive the range
'  as doc.Range(lo, lo + Len(newText)) and format inside it.  That keeps
'  the code free of index arithmetic that breaks when paragraphs vanish.
'=======================================================================

' Fenced block that is not tabular: drop the fence lines, style the body.
Private Sub ApplyCodeBlock(ByVal doc As Document, ByVal lo As Long, ByVal hi As Long, _
                           ByRef b As Blk, ByRef pText() As String)
    Dim nBody As Long, i As Long
    Dim joined As String
    Dim replaceEnd As Long
    Dim r As Range
    Dim endsWithSoftBreak As Boolean

    If hi > lo Then
        endsWithSoftBreak = (doc.Range(hi - 1, hi).Text = Chr$(11))
    End If

    If b.Closed Then
        nBody = b.Last - b.First - 1
    Else
        nBody = b.Last - b.First
    End If

    For i = 1 To nBody
        If i > 1 Then joined = joined & vbCr
        joined = joined & CleanParaText(pText(b.First + i))
    Next i

    If nBody < 1 Then
        doc.Range(lo, hi).Delete
        Exit Sub
    End If

    If endsWithSoftBreak Then
        doc.Range(lo, hi).Text = joined & vbCr
    Else
        replaceEnd = ContentEndBeforeCr(doc, lo, hi)
        doc.Range(lo, replaceEnd).Text = joined
    End If
    Set r = doc.Range(lo, lo + Len(joined))
    If r.End <= r.Start Then Exit Sub

    On Error Resume Next
    With r.Font
        .Name = "Consolas"
        .Size = 10
        .Color = RGB(45, 45, 45)
        .Bold = False
        .BoldBi = False
        .Italic = False
        .ItalicBi = False
    End With
    With r.ParagraphFormat
        .ReadingOrder = wdReadingOrderLtr
        .LeftIndent = 18
        .RightIndent = 18
        .SpaceBefore = 0
        .SpaceAfter = 6
    End With
    r.Shading.BackgroundPatternColor = RGB(242, 242, 242)
    Err.Clear
    On Error GoTo 0
End Sub


' Fenced block tagged tsv/tab/table.  The fence lines go, then the body is
' split into runs of tab separated rows (real tables) and prose lines.
' The runs are rewritten back to front so each one is still at its
' original offset when it is handled.
Private Sub ApplyFenceWithBody(ByVal doc As Document, ByVal lo As Long, ByVal hi As Long, _
                               ByRef b As Blk, ByRef pText() As String, _
                               ByVal doTables As Boolean)
    Dim nBody As Long, i As Long, a As Long
    Dim runStart() As Long
    Dim runEnd() As Long
    Dim runText() As String
    Dim rFrom() As Long
    Dim rTo() As Long
    Dim rTab() As Boolean
    Dim nRuns As Long
    Dim pos As Long
    Dim s As String
    Dim isTab As Boolean
    Dim canExtend As Boolean
    Dim shift As Long
    Dim cols As Long, rows As Long
    Dim data() As String
    Dim endPos As Long
    Dim closeSoftBreak As Boolean

    If b.Closed Then
        nBody = b.Last - b.First - 1
    Else
        nBody = b.Last - b.First
    End If
    If nBody < 1 Then
        doc.Range(lo, hi).Delete
        Exit Sub
    End If

    If hi > lo Then
        closeSoftBreak = (doc.Range(hi - 1, hi).Text = Chr$(11))
    End If

    ReDim runStart(1 To nBody)
    ReDim runEnd(1 To nBody)
    ReDim runText(1 To nBody)
    ReDim rFrom(1 To nBody)
    ReDim rTo(1 To nBody)
    ReDim rTab(1 To nBody)

    ' Paragraph.Range.Text already carries its own paragraph mark, so the
    ' first body line starts exactly one raw string length after lo.
    pos = lo + Len(pText(b.First))
    For i = 1 To nBody
        s = CleanParaText(pText(b.First + i))
        runText(i) = s
        runStart(i) = pos
        pos = pos + Len(pText(b.First + i))
        runEnd(i) = pos
    Next i
    If runEnd(nBody) > hi Then runEnd(nBody) = hi

    ' Group consecutive tab rows into table runs; everything else stays text.
    nRuns = 0
    For i = 1 To nBody
        isTab = doTables And Len(runText(i)) > 0 And InStr(runText(i), vbTab) > 0 _
                 And Not IsListMarked(runText(i))
        canExtend = False
        If nRuns > 0 Then canExtend = isTab And rTab(nRuns)
        If canExtend Then
            rTo(nRuns) = runEnd(i)
        Else
            nRuns = nRuns + 1
            rFrom(nRuns) = runStart(i)
            rTo(nRuns) = runEnd(i)
            rTab(nRuns) = isTab
        End If
    Next i

    ' Deleting the opening fence pulls the whole body left by its length,
    ' so the line and run offsets above have to follow before anything else is cut.
    shift = Len(pText(b.First))
    doc.Range(lo, runStart(1)).Delete
    If runEnd(nBody) < hi Then
        If closeSoftBreak Then
            doc.Range(runEnd(nBody) - shift, hi - shift).Text = vbCr
        Else
            doc.Range(runEnd(nBody) - shift, hi - shift).Delete
        End If
    End If
    For i = 1 To nBody
        runStart(i) = runStart(i) - shift
        runEnd(i) = runEnd(i) - shift
    Next i
    For a = 1 To nRuns
        rFrom(a) = rFrom(a) - shift
        rTo(a) = rTo(a) - shift
    Next a

    For a = nRuns To 1 Step -1
        If rTab(a) Then
            cols = 0
            rows = 0
            For i = 1 To nBody
                If runStart(i) >= rFrom(a) And runStart(i) < rTo(a) Then
                    If CountChar(runText(i), vbTab) + 1 > cols Then
                        cols = CountChar(runText(i), vbTab) + 1
                    End If
                    If Len(runText(i)) > 0 Then rows = rows + 1
                End If
            Next i
            If cols >= 2 And rows >= 1 Then
                ReDim data(0 To rows - 1, 0 To cols - 1)
                rows = 0
                For i = 1 To nBody
                    If runStart(i) >= rFrom(a) And runStart(i) < rTo(a) Then
                        If Len(runText(i)) > 0 Then
                            FillRow data, rows, runText(i), cols
                            rows = rows + 1
                        End If
                    End If
                Next i
                doc.Range(rFrom(a), rTo(a)).Delete
                BuildTableAt doc, rFrom(a), data, rows, cols
            End If
        Else
            If rTo(a) > rFrom(a) Then
                If doc.Range(rTo(a) - 1, rTo(a)).Text = Chr$(11) Then
                    doc.Range(rTo(a) - 1, rTo(a)).Text = vbCr
                End If
            End If
            endPos = ContentEndBeforeCr(doc, rFrom(a), rTo(a))
            If endPos > rFrom(a) Then
                On Error Resume Next
                CollapseDashesInRange doc.Range(rFrom(a), endPos)
                InlinePass doc.Range(rFrom(a), endPos)
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next a
End Sub


' Unfenced run of tab separated rows.
Private Sub ApplyTabRun(ByVal doc As Document, ByVal lo As Long, ByVal hi As Long, _
                        ByRef b As Blk, ByRef pText() As String, ByVal doTables As Boolean)
    Dim i As Long
    Dim cols As Long, rows As Long
    Dim data() As String
    Dim s As String

    If Not doTables Then Exit Sub

    For i = b.First To b.Last
        s = CleanParaText(pText(i))
        If Len(s) > 0 Then
            rows = rows + 1
            If CountChar(s, vbTab) + 1 > cols Then cols = CountChar(s, vbTab) + 1
        End If
    Next i
    If rows < 2 Or cols < 2 Then Exit Sub

    ReDim data(0 To rows - 1, 0 To cols - 1)
    rows = 0
    For i = b.First To b.Last
        s = CleanParaText(pText(i))
        If Len(s) > 0 Then
            FillRow data, rows, s, cols
            rows = rows + 1
        End If
    Next i

    doc.Range(lo, hi).Delete
    BuildTableAt doc, lo, data, rows, cols
End Sub


' Markdown pipe table (| a | b |).
Private Sub ApplyPipeTable(ByVal doc As Document, ByVal lo As Long, ByVal hi As Long, _
                           ByRef b As Blk, ByRef pText() As String, ByVal doTables As Boolean)
    Dim i As Long, c As Long
    Dim cols As Long, rows As Long, colCnt As Long
    Dim data() As String
    Dim cells() As String
    Dim s As String

    If Not doTables Then Exit Sub

    For i = b.First To b.Last
        s = CleanParaText(pText(i))
        If Len(s) > 0 And Not IsPipeSeparator(s) Then
            rows = rows + 1
            colCnt = PipeColumnCount(s)
            If colCnt > cols Then cols = colCnt
        End If
    Next i
    If rows < 1 Or cols < 2 Then Exit Sub

    ReDim data(0 To rows - 1, 0 To cols - 1)
    rows = 0
    For i = b.First To b.Last
        s = CleanParaText(pText(i))
        If Len(s) > 0 And Not IsPipeSeparator(s) Then
            cells = SplitPipeRow(s, cols)
            For c = 0 To cols - 1
                data(rows, c) = NormalizeDashes(cells(c))
            Next c
            rows = rows + 1
        End If
    Next i

    doc.Range(lo, hi).Delete
    BuildTableAt doc, lo, data, rows, cols
End Sub


Private Sub FillRow(ByRef data() As String, ByVal rowIdx As Long, _
                   ByVal s As String, ByVal cols As Long)
    Dim parts() As String
    Dim k As Long, n As Long

    s = Replace(s, Chr$(11), "")
    parts = Split(s, vbTab)
    n = UBound(parts)
    For k = 0 To cols - 1
        If k <= n Then
            data(rowIdx, k) = NormalizeDashes(Trim$(parts(k)))
        Else
            data(rowIdx, k) = ""
        End If
    Next k
End Sub


' Ordinary markdown text.  Block level markers are resolved first, then the
' whole block is rewritten once and formatted inside the new range.
Private Sub ApplyPlainBlock(ByVal doc As Document, ByVal lo As Long, ByVal hi As Long, _
                            ByRef b As Blk, ByRef pText() As String)
    Dim n As Long, i As Long, k As Long, m As Long
    Dim outText() As String
    Dim outKind() As Long
    Dim outArg() As Long
    Dim splitLines() As String
    Dim nSplit As Long
    Dim raw As String, s As String
    Dim joined As String
    Dim replaceEnd As Long
    Dim r As Range
    Dim p As Paragraph
    Dim idx As Long
    Dim lvl As Long
    Dim num As Long
    Dim numPrefix As String, numBody As String
    Dim isRtl As Boolean
    Dim endsWithSoftBreak As Boolean
    Dim origClean As String
    Dim hrDropped As Boolean

    If hi > lo Then
        endsWithSoftBreak = (doc.Range(hi - 1, hi).Text = Chr$(11))
    End If

    n = b.Last - b.First + 1
    ReDim outText(1 To n)
    ReDim outKind(1 To n)
    ReDim outArg(1 To n)

    m = 0
    For i = 1 To n
        raw = pText(b.First + i - 1)
        s = CleanParaText(raw)

        If Len(s) = 0 Then
            If Not endsWithSoftBreak Then
                AppendPlainOut outText, outKind, outArg, m, "", LK_TEXT, 0
            End If
        Else
            lvl = AtxLevel(s)
            If lvl > 0 Then
                AppendPlainOut outText, outKind, outArg, m, _
                    NormalizeDashes(Trim$(Mid$(s, lvl + 2))), LK_HEADING, lvl
            ElseIf IsSetextContinuation(s, m, outKind, outText) Then
                ' The line above this one becomes the heading.
                outKind(m) = LK_HEADING
                If Left$(s, 1) = "=" Then
                    outArg(m) = 1
                Else
                    outArg(m) = 2
                End If
            ElseIf IsHorizontalRule(s) Then
                hrDropped = True
            ElseIf Left$(s, 1) = ">" Then
                AppendPlainOut outText, outKind, outArg, m, _
                    NormalizeDashes(TrimQuoteMarker(s)), LK_QUOTE, 0
            ElseIf IsBulletStart(s) Then
                AppendPlainOut outText, outKind, outArg, m, _
                    mBullet & vbTab & NormalizeDashes(Mid$(s, 3)), LK_BULLET, 0
            Else
                nSplit = SplitSequentialNumbered(s, splitLines)
                If nSplit > 1 Then
                    For k = 1 To nSplit
                        If ParseNumberedPrefix(splitLines(k), numPrefix, num, numBody) Then
                            AppendPlainOut outText, outKind, outArg, m, _
                                numPrefix & "." & vbTab & NormalizeDashes(numBody), LK_NUMBERED, 0
                        Else
                            AppendPlainOut outText, outKind, outArg, m, _
                                NormalizeDashes(splitLines(k)), LK_TEXT, 0
                        End If
                    Next k
                ElseIf ParseNumberedPrefix(s, numPrefix, num, numBody) Then
                    AppendPlainOut outText, outKind, outArg, m, _
                        numPrefix & "." & vbTab & NormalizeDashes(numBody), LK_NUMBERED, 0
                Else
                    AppendPlainOut outText, outKind, outArg, m, _
                        NormalizeDashes(s), LK_TEXT, 0
                End If
            End If
        End If
    Next i

    If m = 0 Then
        If hrDropped Then
            doc.Range(lo, hi).Delete
        ElseIf endsWithSoftBreak Then
            doc.Range(lo, hi).Text = vbCr
        Else
            doc.Range(lo, hi).Delete
        End If
        Exit Sub
    End If

    If m = 1 And Not endsWithSoftBreak Then
        If outKind(1) = LK_TEXT And Not HasInlineMarkers(outText(1)) Then
            origClean = CleanParaText(pText(b.First))
            If outText(1) = origClean Then Exit Sub
            ' Only dash normalization changed this plain paragraph; update in
            ' place so existing heading/emoji formatting stays intact.
            replaceEnd = ContentEndBeforeCr(doc, lo, hi)
            If replaceEnd > lo Then
                Set r = doc.Range(lo, replaceEnd)
                CollapseDashesInRange r
                If NormalizeDashes(CleanParaText(r.Text)) <> CleanParaText(r.Text) Then
                    r.Text = outText(1)
                End If
            End If
            Exit Sub
        End If
    End If

    For i = 1 To m
        If i > 1 Then joined = joined & vbCr
        joined = joined & outText(i)
    Next i

    If endsWithSoftBreak Then
        doc.Range(lo, hi).Text = joined & vbCr
    Else
        replaceEnd = ContentEndBeforeCr(doc, lo, hi)
        doc.Range(lo, replaceEnd).Text = joined
    End If
    Set r = doc.Range(lo, lo + Len(joined))
    If r.End <= r.Start Then Exit Sub

    ' Paragraph level formatting first: it must not be wiped by the inline
    ' pass that follows.
    idx = 0
    For Each p In r.Paragraphs
        idx = idx + 1
        If idx > m Then Exit For
        On Error Resume Next
        Select Case outKind(idx)
            Case LK_HEADING
                ApplyHeading p, outArg(idx)
            Case LK_BULLET, LK_NUMBERED
                isRtl = IsRtlText(outText(idx))
                If isRtl Then
                    p.Range.ParagraphFormat.ReadingOrder = wdReadingOrderRtl
                    p.Range.ParagraphFormat.RightIndent = 24
                    p.Range.ParagraphFormat.LeftIndent = 0
                Else
                    p.Range.ParagraphFormat.ReadingOrder = wdReadingOrderLtr
                    p.Range.ParagraphFormat.LeftIndent = 24
                    p.Range.ParagraphFormat.RightIndent = 0
                End If
                p.Range.ParagraphFormat.FirstLineIndent = -18
            Case LK_QUOTE
                p.Range.Font.Italic = True
                p.Range.Font.ItalicBi = True
                p.Range.Font.Color = RGB(85, 85, 85)
                If IsRtlText(outText(idx)) Then
                    p.Range.ParagraphFormat.ReadingOrder = wdReadingOrderRtl
                End If
                p.Range.ParagraphFormat.LeftIndent = 24
                p.Range.ParagraphFormat.RightIndent = 24
                p.Range.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                p.Range.Borders(wdBorderRight).LineWidth = wdLineWidth150pt
                p.Range.Borders(wdBorderRight).Color = RGB(190, 190, 190)
        End Select
        Err.Clear
        On Error GoTo 0
    Next p

    On Error Resume Next
    InlinePass r
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub AppendPlainOut(ByRef outText() As String, ByRef outKind() As Long, _
                           ByRef outArg() As Long, ByRef m As Long, _
                           ByVal txt As String, ByVal kind As Long, ByVal arg As Long)
    m = m + 1
    If m > UBound(outText) Then
        ReDim Preserve outText(1 To m * 2)
        ReDim Preserve outKind(1 To m * 2)
        ReDim Preserve outArg(1 To m * 2)
    End If
    outText(m) = txt
    outKind(m) = kind
    outArg(m) = arg
End Sub


' VBA evaluates all operands of `And` (no short-circuiting), so check m >= 1
' before indexing outKind(m) / outText(m).
Private Function IsSetextContinuation(ByVal s As String, ByVal m As Long, _
                                      ByRef outKind() As Long, _
                                      ByRef outText() As String) As Boolean
    If m < 1 Then Exit Function
    If Not IsSetextLine(s) Then Exit Function
    If outKind(m) <> LK_TEXT Then Exit Function
    If Len(outText(m)) = 0 Then Exit Function
    If IsBulletOutput(outText(m)) Then Exit Function
    IsSetextContinuation = True
End Function


Private Sub ApplyHeading(ByVal p As Paragraph, ByVal level As Long)
    Dim styleId As WdBuiltinStyle
    Select Case level
        Case 1: styleId = wdStyleHeading1
        Case 2: styleId = wdStyleHeading2
        Case 3: styleId = wdStyleHeading3
        Case 4: styleId = wdStyleHeading4
        Case Else: styleId = wdStyleHeading5
    End Select
    p.Range.Style = p.Range.Document.Styles(styleId)
    If IsRtlText(p.Range.Text) Then
        p.Range.ParagraphFormat.ReadingOrder = wdReadingOrderRtl
    End If
    p.Range.ParagraphFormat.LeftIndent = 0
    p.Range.ParagraphFormat.RightIndent = 0
    p.Range.ParagraphFormat.FirstLineIndent = 0
End Sub


'-----------------------------------------------------------------------
'  Line classification helpers
'-----------------------------------------------------------------------

' 0 when the line is not an ATX heading, else 1..5.
Private Function AtxLevel(ByVal s As String) As Long
    Dim hashes As Long
    Dim ch As String
    Dim i As Long
    If Len(s) < 3 Then Exit Function
    If Left$(s, 1) <> "#" Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "#" Then
            hashes = hashes + 1
        ElseIf ch = " " Then
            If i = hashes + 1 And hashes >= 1 And hashes <= 5 Then AtxLevel = hashes
            Exit Function
        Else
            Exit Function
        End If
    Next i
End Function


Private Function IsSetextLine(ByVal s As String) As Boolean
    If Len(s) < 2 Then Exit Function
    If IsAllOf(s, "=") Then IsSetextLine = True: Exit Function
    If IsAllOf(s, "-") Then IsSetextLine = True
End Function


Private Function IsHorizontalRule(ByVal s As String) As Boolean
    Dim core As String
    If Len(s) < 3 Then Exit Function
    core = Replace(s, "-", "")
    core = Replace(core, "*", "")
    core = Replace(core, "_", "")
    core = Replace(core, "=", "")
    core = Replace(core, "~", "")
    core = Replace(core, "`", "")
    IsHorizontalRule = (Len(Trim$(core)) = 0)
End Function


Private Function IsAllOf(ByVal s As String, ByVal ch As String) As Boolean
    Dim i As Long
    If Len(s) < 2 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) <> ch Then Exit Function
    Next i
    IsAllOf = True
End Function


Private Function IsBulletStart(ByVal s As String) As Boolean
    If Len(s) < 3 Then Exit Function
    If Mid$(s, 2, 1) <> " " Then Exit Function
    If InStr("-*+", Left$(s, 1)) > 0 Then IsBulletStart = True
End Function


' Numeric value 0..9 for ASCII, Persian (06F0..06F9), and Arabic-Indic
' (0660..0669) digits, or -1 otherwise.
Private Function DigitValue(ByVal ch As String) As Long
    Dim code As Long
    If Len(ch) <> 1 Then
        DigitValue = -1
        Exit Function
    End If
    code = AscW(ch)
    If code < 0 Then code = code + 65536
    If code >= 48 And code <= 57 Then
        DigitValue = code - 48
    ElseIf code >= &H6F0 And code <= &H6F9 Then
        DigitValue = code - &H6F0
    ElseIf code >= &H660 And code <= &H669 Then
        DigitValue = code - &H660
    Else
        DigitValue = -1
    End If
End Function


Private Function ParseDigits(ByVal s As String) As Long
    Dim i As Long, d As Long, acc As Long
    For i = 1 To Len(s)
        d = DigitValue(Mid$(s, i, 1))
        If d < 0 Then
            ParseDigits = -1
            Exit Function
        End If
        acc = acc * 10 + d
    Next i
    ParseDigits = acc
End Function


Private Function IsSpaceChar(ByVal ch As String) As Boolean
    Dim code As Long
    If Len(ch) <> 1 Then Exit Function
    If ch = " " Or ch = vbTab Then
        IsSpaceChar = True
        Exit Function
    End If
    code = AscW(ch)
    If code < 0 Then code = code + 65536
    If code = &HA0 Or code = &H200C Or code = &H200E Or code = &H200F Then
        ' Note: ZWNJ (200C) stays within words; NBSP (A0) and directional marks act as separators.
        If code <> &H200C Then IsSpaceChar = True
    End If
End Function


' Accepts "12. text" (including Persian/Arabic-Indic numerals).
Private Function ParseNumberedPrefix(ByVal s As String, ByRef prefix As String, _
                                     ByRef num As Long, ByRef body As String) As Boolean
    Dim pos As Long
    Dim sep As String
    pos = InStr(s, ".")
    If pos < 2 Or pos > 5 Or pos >= Len(s) Then Exit Function
    sep = Mid$(s, pos + 1, 1)
    If Not IsSpaceChar(sep) Then Exit Function
    prefix = Left$(s, pos - 1)
    If Not IsAllDigits(prefix) Then Exit Function
    num = ParseDigits(prefix)
    If num < 0 Then Exit Function
    body = Mid$(s, pos + 2)
    ParseNumberedPrefix = True
End Function


Private Function IsNumberedStart(ByVal s As String, ByRef num As Long) As Boolean
    Dim prefix As String, body As String
    IsNumberedStart = ParseNumberedPrefix(s, prefix, num, body)
End Function


Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If DigitValue(Mid$(s, i, 1)) < 0 Then Exit Function
    Next i
    IsAllDigits = True
End Function


Private Function IsSentenceBoundaryChar(ByVal ch As String) As Boolean
    Dim code As Long
    If Len(ch) <> 1 Then Exit Function
    If InStr(".?!:;)]" & Chr$(34) & "'", ch) > 0 Then
        IsSentenceBoundaryChar = True
        Exit Function
    End If
    code = AscW(ch)
    If code < 0 Then code = code + 65536
    ' Arabic question mark (061F), semicolon (061B), comma (060C), guillemet (00BB), dashes
    If code = &H61F Or code = &H61B Or code = &H60C Or code = &HBB Or _
       code = &H2014 Or code = &H2013 Then
        IsSentenceBoundaryChar = True
    End If
End Function


Private Function HasSequentialNumbered(ByVal s As String) As Boolean
    Dim parts() As String
    HasSequentialNumbered = (SplitSequentialNumbered(s, parts) > 1)
End Function


' Splits inline sequential numbered lists such as "۱. اول. ۲. دوم. ۳. سوم."
' (or "مقدمه: ۱. اول. ۲. دوم.") into vertical items while avoiding prose
' numbers like "1. Read chapter 2. Today".
Private Function SplitSequentialNumbered(ByVal s As String, ByRef outLines() As String) As Long
    Dim nLen As Long, i As Long, j As Long
    Dim closeTick As Long, closeParen As Long
    Dim dLen As Long, numVal As Long
    Dim candPos() As Long, candNum() As Long
    Dim candPunct() As Boolean, candColon() As Boolean
    Dim nCand As Long
    Dim prevCh As String, prevNonWs As String
    Dim isPunct As Boolean, isColon As Boolean
    Dim chainIdx() As Long
    Dim chainLen As Long, startIdx As Long, wantNum As Long
    Dim allPunct As Boolean
    Dim outCount As Long, k As Long, segFrom As Long, segTo As Long
    Dim piece As String

    nLen = Len(s)
    If nLen < 6 Or InStr(s, ".") = 0 Then Exit Function

    ReDim candPos(1 To 16)
    ReDim candNum(1 To 16)
    ReDim candPunct(1 To 16)
    ReDim candColon(1 To 16)
    nCand = 0

    i = 1
    Do While i <= nLen
        If Mid$(s, i, 1) = "`" Then
            closeTick = InStr(i + 1, s, "`")
            If closeTick > 0 Then
                i = closeTick + 1
            Else
                i = i + 1
            End If
        ElseIf Mid$(s, i, 2) = "](" Then
            closeParen = InStr(i + 2, s, ")")
            If closeParen > 0 Then
                i = closeParen + 1
            Else
                i = i + 1
            End If
        Else
            prevCh = ""
            If i > 1 Then prevCh = Mid$(s, i - 1, 1)
            If i = 1 Or IsSpaceChar(prevCh) Then
                dLen = 0
                Do While i + dLen <= nLen And dLen < 4
                    If DigitValue(Mid$(s, i + dLen, 1)) >= 0 Then
                        dLen = dLen + 1
                    Else
                        Exit Do
                    End If
                Loop
                If dLen >= 1 And dLen <= 3 And i + dLen + 1 <= nLen Then
                    If Mid$(s, i + dLen, 1) = "." And IsSpaceChar(Mid$(s, i + dLen + 1, 1)) Then
                        If Len(Trim$(Mid$(s, i + dLen + 2))) > 0 Then
                            numVal = ParseDigits(Mid$(s, i, dLen))
                            isPunct = False
                            isColon = False
                            If i > 1 Then
                                j = i - 1
                                Do While j >= 1
                                    prevNonWs = Mid$(s, j, 1)
                                    If Not IsSpaceChar(prevNonWs) Then Exit Do
                                    j = j - 1
                                Loop
                                If j >= 1 Then
                                    isColon = (prevNonWs = ":")
                                    isPunct = IsSentenceBoundaryChar(prevNonWs)
                                End If
                            End If
                            nCand = nCand + 1
                            If nCand > UBound(candPos) Then
                                ReDim Preserve candPos(1 To nCand * 2)
                                ReDim Preserve candNum(1 To nCand * 2)
                                ReDim Preserve candPunct(1 To nCand * 2)
                                ReDim Preserve candColon(1 To nCand * 2)
                            End If
                            candPos(nCand) = i
                            candNum(nCand) = numVal
                            candPunct(nCand) = isPunct
                            candColon(nCand) = isColon
                            i = i + dLen + 2
                        Else
                            i = i + 1
                        End If
                    Else
                        i = i + 1
                    End If
                Else
                    i = i + 1
                End If
            Else
                i = i + 1
            End If
        End If
    Loop

    If nCand < 2 Then Exit Function

    startIdx = 0
    If candPos(1) = 1 And candNum(1) >= 1 And candNum(1) <= 50 Then
        startIdx = 1
    Else
        For i = 1 To nCand
            If candNum(i) = 1 And candColon(i) Then
                startIdx = i
                Exit For
            End If
        Next i
    End If
    If startIdx = 0 Then Exit Function

    ReDim chainIdx(1 To nCand)
    chainLen = 1
    chainIdx(1) = startIdx
    wantNum = candNum(startIdx) + 1
    allPunct = True

    For i = startIdx + 1 To nCand
        If candNum(i) = wantNum Then
            chainLen = chainLen + 1
            chainIdx(chainLen) = i
            If Not candPunct(i) Then allPunct = False
            wantNum = wantNum + 1
        End If
    Next i

    If chainLen < 2 Then Exit Function
    If chainLen = 2 Then
        If candNum(chainIdx(1)) <> 1 Or Not allPunct Then Exit Function
    End If

    ReDim outLines(1 To chainLen + 1)
    outCount = 0
    If candPos(chainIdx(1)) > 1 Then
        piece = Trim$(Left$(s, candPos(chainIdx(1)) - 1))
        If Len(piece) > 0 Then
            outCount = outCount + 1
            outLines(outCount) = piece
        End If
    End If
    For k = 1 To chainLen
        segFrom = candPos(chainIdx(k))
        If k < chainLen Then
            segTo = candPos(chainIdx(k + 1)) - 1
        Else
            segTo = nLen
        End If
        piece = Trim$(Mid$(s, segFrom, segTo - segFrom + 1))
        If Len(piece) > 0 Then
            outCount = outCount + 1
            outLines(outCount) = piece
        End If
    Next k

    If outCount > 1 Then SplitSequentialNumbered = outCount
End Function


'-----------------------------------------------------------------------
'  Dash normalization
'-----------------------------------------------------------------------

Private Function IsUnicodeDashChar(ByVal ch As String) As Boolean
    Dim code As Long
    If Len(ch) <> 1 Then Exit Function
    code = AscW(ch)
    If code < 0 Then code = code + 65536
    If code = &H2012 Or code = &H2013 Or code = &H2014 Or _
       code = &H2015 Or code = &H2500 Then
        IsUnicodeDashChar = True
    End If
End Function


' Normalizes Unicode dashes ("—", "–", "――", "──") and repeated prose
' hyphens ("——", " -- ", "word--word") outside inline code spans and link
' URLs to a single standard ASCII hyphen ("-").
Private Function NormalizeDashes(ByVal s As String) As String
    Dim nLen As Long, i As Long, j As Long
    Dim closeTick As Long, closeParen As Long
    Dim ch As String, prevCh As String, nextCh As String
    Dim hasUnicode As Boolean
    Dim outBuf As String

    If Len(s) = 0 Then
        NormalizeDashes = s
        Exit Function
    End If

    nLen = Len(s)
    i = 1
    Do While i <= nLen
        ch = Mid$(s, i, 1)
        If ch = "`" Then
            closeTick = InStr(i + 1, s, "`")
            If closeTick > 0 Then
                outBuf = outBuf & Mid$(s, i, closeTick - i + 1)
                i = closeTick + 1
            Else
                outBuf = outBuf & ch
                i = i + 1
            End If
        ElseIf Mid$(s, i, 2) = "](" Then
            closeParen = InStr(i + 2, s, ")")
            If closeParen > 0 Then
                outBuf = outBuf & Mid$(s, i, closeParen - i + 1)
                i = closeParen + 1
            Else
                outBuf = outBuf & ch
                i = i + 1
            End If
        ElseIf IsUnicodeDashChar(ch) Or ch = "-" Then
            hasUnicode = IsUnicodeDashChar(ch)
            j = i + 1
            Do While j <= nLen
                nextCh = Mid$(s, j, 1)
                If IsUnicodeDashChar(nextCh) Then
                    hasUnicode = True
                    j = j + 1
                ElseIf nextCh = "-" Then
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop
            If hasUnicode Then
                outBuf = outBuf & "-"
            ElseIf j - i >= 2 Then
                prevCh = ""
                If i > 1 Then prevCh = Mid$(s, i - 1, 1)
                nextCh = ""
                If j <= nLen Then nextCh = Mid$(s, j, 1)
                If prevCh = " " And nextCh = " " Then
                    outBuf = outBuf & "-"
                ElseIf Len(prevCh) > 0 And Len(nextCh) > 0 And _
                       prevCh <> " " And prevCh <> vbTab And _
                       nextCh <> " " And nextCh <> vbTab Then
                    outBuf = outBuf & "-"
                Else
                    outBuf = outBuf & Mid$(s, i, j - i)
                End If
            Else
                outBuf = outBuf & ch
            End If
            i = j
        Else
            outBuf = outBuf & ch
            i = i + 1
        End If
    Loop
    NormalizeDashes = outBuf
End Function


Private Sub CollapseDashesInRange(ByVal r As Range)
    Dim em As String
    em = ChrW$(&H2014)

    ReplaceAllInRange r, ChrW$(&H2012), em
    ReplaceAllInRange r, ChrW$(&H2013), em
    ReplaceAllInRange r, ChrW$(&H2015), em
    ReplaceAllInRange r, ChrW$(&H2500), em

    ReplaceAllInRange r, em & em, em
    ReplaceAllInRange r, em & "-", em
    ReplaceAllInRange r, "-" & em, em
    ReplaceAllInRange r, em, "-"

    ReplaceAllInRange r, " --- ", " - "
    ReplaceAllInRange r, " -- ", " - "
End Sub


Private Sub ReplaceAllInRange(ByVal r As Range, ByVal findStr As String, ByVal replStr As String)
    Dim guard As Long
    On Error Resume Next
    Do While InStr(r.Text, findStr) > 0 And guard < 20
        guard = guard + 1
        With r.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = findStr
            .Replacement.Text = replStr
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
    Loop
    Err.Clear
    On Error GoTo 0
End Sub


'-----------------------------------------------------------------------
'  Markdown pipe table helpers
'-----------------------------------------------------------------------

Private Function IsPipeRow(ByVal s As String) As Boolean
    If InStr(s, "|") = 0 Then Exit Function
    IsPipeRow = (PipeColumnCount(s) >= 2)
End Function


Private Function IsBorderedPipeRow(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    If Len(t) < 3 Then Exit Function
    If Left$(t, 1) = "|" And Right$(t, 1) = "|" And CountChar(t, "|") >= 2 Then
        IsBorderedPipeRow = True
    End If
End Function


Private Function IsPipeSeparator(ByVal s As String) As Boolean
    Dim t As String
    Dim parts() As String
    Dim i As Long, j As Long
    Dim cell As String, ch As String
    Dim hasDash As Boolean

    t = Trim$(s)
    If Len(t) < 3 Or InStr(t, "|") = 0 Or InStr(t, "-") = 0 Then Exit Function
    If Left$(t, 1) = "|" Then t = Mid$(t, 2)
    If Right$(t, 1) = "|" Then t = Left$(t, Len(t) - 1)
    parts = Split(t, "|")
    If UBound(parts) < 1 Then Exit Function

    For i = LBound(parts) To UBound(parts)
        cell = Trim$(parts(i))
        If Len(cell) = 0 Then Exit Function
        hasDash = False
        For j = 1 To Len(cell)
            ch = Mid$(cell, j, 1)
            If ch = "-" Then
                hasDash = True
            ElseIf ch <> ":" Then
                Exit Function
            End If
        Next j
        If Not hasDash Then Exit Function
    Next i
    IsPipeSeparator = True
End Function


Private Function PipeColumnCount(ByVal s As String) As Long
    Dim t As String
    t = Trim$(s)
    If Len(t) < 2 Or InStr(t, "|") = 0 Then Exit Function
    If Left$(t, 1) = "|" Then t = Mid$(t, 2)
    If Right$(t, 1) = "|" And Len(t) > 0 Then t = Left$(t, Len(t) - 1)
    If Len(t) = 0 Then Exit Function
    PipeColumnCount = CountChar(t, "|") + 1
End Function


Private Function SplitPipeRow(ByVal s As String, ByVal cols As Long) As String()
    Dim t As String
    Dim parts() As String
    Dim res() As String
    Dim k As Long, n As Long

    t = Trim$(s)
    If Len(t) >= 2 Then
        If Left$(t, 1) = "|" Then t = Mid$(t, 2)
        If Right$(t, 1) = "|" And Len(t) > 0 Then t = Left$(t, Len(t) - 1)
    End If
    parts = Split(t, "|")
    ReDim res(0 To cols - 1)
    n = UBound(parts)
    For k = 0 To cols - 1
        If k <= n Then
            res(k) = Trim$(parts(k))
        Else
            res(k) = ""
        End If
    Next k
    SplitPipeRow = res
End Function


Private Function HasInlineMarkers(ByVal s As String) As Boolean
    If InStr(s, "*") > 0 Then HasInlineMarkers = True: Exit Function
    If InStr(s, "~") > 0 Then HasInlineMarkers = True: Exit Function
    If InStr(s, "`") > 0 Then HasInlineMarkers = True: Exit Function
    If InStr(s, "[") > 0 And InStr(s, "](") > 0 Then HasInlineMarkers = True
End Function


Private Function TrimQuoteMarker(ByVal s As String) As String
    Dim t As String
    t = Mid$(s, 2)
    If Left$(t, 1) = " " Then t = Mid$(t, 2)
    TrimQuoteMarker = t
End Function


Private Function IsBulletOutput(ByVal s As String) As Boolean
    If Len(s) <= Len(mBullet) Then Exit Function
    If Left$(s, Len(mBullet)) <> mBullet Then Exit Function
    IsBulletOutput = (Mid$(s, Len(mBullet) + 1, 1) = vbTab)
End Function


'-----------------------------------------------------------------------
'  Inline stage
'
'  Applied with Word's own wildcard Find so the emphasis survives as real
'  character formatting instead of being flattened to plain text.  Order
'  matters: links and inline code first, bold before italic.
'-----------------------------------------------------------------------

Private Sub InlinePass(ByVal r As Range)
    InlineStyle r, "\[([!\[\]]@)\]\([!\(\)]@\)", "\1", IL_LINK
    InlineStyle r, "`([!\`]@)`", "\1", IL_CODE
    InlineStyle r, "\*\*([!\*]@)\*\*", "\1", IL_BOLD
    InlineStyle r, "\~\~([!\~]@)\~\~", "\1", IL_STRIKE
    InlineStyle r, "\*([!\*]@)\*", "\1", IL_ITALIC
End Sub

Private Sub InlineStyle(ByVal r As Range, ByVal pattern As String, _
                        ByVal repl As String, ByVal styleId As Long)
    On Error Resume Next
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = pattern
        .Replacement.Text = repl
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
        Select Case styleId
            Case IL_BOLD
                .Replacement.Font.Bold = True
                .Replacement.Font.BoldBi = True
            Case IL_ITALIC
                .Replacement.Font.Italic = True
                .Replacement.Font.ItalicBi = True
            Case IL_STRIKE
                .Replacement.Font.StrikeThrough = True
            Case IL_CODE
                .Replacement.Font.Name = "Consolas"
                .Replacement.Font.Size = 10
                .Replacement.Font.Color = RGB(150, 40, 40)
            Case IL_LINK
                .Replacement.Font.Color = RGB(0, 0, 200)
                .Replacement.Font.Underline = wdUnderlineSingle
        End Select
        .Execute Replace:=wdReplaceAll
    End With
    Err.Clear
    On Error GoTo 0
End Sub


'-----------------------------------------------------------------------
'  Tables
'-----------------------------------------------------------------------

' Replaces the caller-deleted range with a real Word table.  If the grid
' cannot be created the original text is written back, so a failure never
' destroys content.
Private Sub BuildTableAt(ByVal doc As Document, ByVal pos As Long, _
                         ByRef data() As String, ByVal rows As Long, ByVal cols As Long)
    Dim r As Range
    Dim cellRng As Range
    Dim t As Table
    Dim a As Long, b As Long
    Dim restored As String

    If rows < 1 Or cols < 1 Then Exit Sub

    For a = 0 To rows - 1
        For b = 0 To cols - 1
            restored = restored & data(a, b)
            If b < cols - 1 Then restored = restored & vbTab
        Next b
        If a < rows - 1 Then restored = restored & vbCr
    Next a

    On Error GoTo Fallback

    Set r = doc.Range(pos, pos)
    r.InsertParagraphAfter
    Set r = doc.Range(pos, pos)

    Set t = doc.Tables.Add( _
        Range:=r, _
        NumRows:=rows, _
        NumColumns:=cols, _
        DefaultTableBehavior:=wdWord9TableBehavior, _
        AutoFitBehavior:=wdAutoFitFixed)

    For a = 0 To rows - 1
        For b = 0 To cols - 1
            t.Cell(a + 1, b + 1).Range.Text = data(a, b)
        Next b
    Next a

    FormatTable t

    For a = 0 To rows - 1
        For b = 0 To cols - 1
            If HasInlineMarkers(data(a, b)) Then
                On Error Resume Next
                Set cellRng = t.Cell(a + 1, b + 1).Range
                cellRng.MoveEnd wdCharacter, -1
                If Err.Number = 0 Then
                    If cellRng.End > cellRng.Start Then InlinePass cellRng
                End If
                Err.Clear
                On Error GoTo 0
            End If
        Next b
    Next a
    Exit Sub

Fallback:
    On Error Resume Next
    doc.Range(pos, pos).InsertAfter restored
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub FormatTable(ByVal tbl As Table)
    Dim styled As Boolean
    Dim c As Cell

    On Error Resume Next

    If IsRtlText(tbl.Range.Text) Then
        tbl.TableDirection = wdTableDirectionRtl
        tbl.Range.ParagraphFormat.ReadingOrder = wdReadingOrderRtl
    End If

    ' "Table Grid" is English-only in most builds; fall back to direct
    ' borders when the localized style name does not resolve.
    tbl.Style = "Table Grid"
    If Err.Number = 0 Then styled = True
    Err.Clear

    If Not styled Then
        tbl.Borders.Enable = True
        tbl.Borders.OutsideLineStyle = wdLineStyleSingle
        tbl.Borders.InsideLineStyle = wdLineStyleSingle
        tbl.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        tbl.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        tbl.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        tbl.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        If tbl.Rows.Count > 1 Then tbl.Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
        If tbl.Columns.Count > 1 Then tbl.Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
    End If

    tbl.Range.Font.Bold = False
    tbl.Range.Font.BoldBi = False
    tbl.Range.Font.Size = 11
    tbl.Range.Font.SizeBi = 11
    tbl.Range.ParagraphFormat.SpaceBefore = 2
    tbl.Range.ParagraphFormat.SpaceAfter = 2
    tbl.Range.ParagraphFormat.LeftIndent = 0
    tbl.Range.ParagraphFormat.RightIndent = 0
    tbl.Range.ParagraphFormat.FirstLineIndent = 0

    If tbl.Rows.Count >= 1 Then
        tbl.Rows(1).Range.Font.Bold = True
        tbl.Rows(1).Range.Font.BoldBi = True
        tbl.Rows(1).Shading.Texture = wdTextureNone
        tbl.Rows(1).Shading.BackgroundPatternColor = RGB(219, 229, 241)
        For Each c In tbl.Rows(1).Cells
            c.Shading.Texture = wdTextureNone
            c.Shading.BackgroundPatternColor = RGB(219, 229, 241)
        Next c
        tbl.Rows(1).HeadingFormat = True
    End If

    tbl.PreferredWidthType = wdPreferredWidthPercent
    tbl.PreferredWidth = 100

    Err.Clear
    On Error GoTo 0
End Sub


'-----------------------------------------------------------------------
'  Small helpers
'-----------------------------------------------------------------------

Private Sub EnsureBullet()
    If Len(mBullet) = 0 Then mBullet = ChrW$(&H2022)
End Sub


Private Function CleanParaText(ByVal s As String) As String
    Do While Len(s) > 0
        If Right$(s, 1) = Chr$(13) Or Right$(s, 1) = Chr$(7) Or _
           Right$(s, 1) = Chr$(11) Or Right$(s, 1) = Chr$(10) Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    CleanParaText = Trim$(s)
End Function


Private Function CountChar(ByVal s As String, ByVal ch As String) As Long
    Dim k As Long, cnt As Long
    If Len(ch) = 0 Then Exit Function
    For k = 1 To Len(s)
        If Mid$(s, k, Len(ch)) = ch Then cnt = cnt + 1
    Next k
    CountChar = cnt
End Function


Private Function IsRtlText(ByVal s As String) As Boolean
    Dim k As Long, code As Long, rtlCount As Long, ltrCount As Long
    For k = 1 To Len(s)
        code = AscW(Mid$(s, k, 1))
        If code < 0 Then code = code + 65536
        If (code >= &H590 And code <= &H5FF) Or _
           (code >= &H600 And code <= &H6FF) Or _
           (code >= &H750 And code <= &H77F) Or _
           (code >= &HFB50 And code <= &HFDFF) Or _
           (code >= &HFE70 And code <= &HFEFF) Then
            rtlCount = rtlCount + 1
        ElseIf (code >= &H41 And code <= &H5A) Or (code >= &H61 And code <= &H7A) Then
            ltrCount = ltrCount + 1
        End If
    Next k
    IsRtlText = (rtlCount > ltrCount)
End Function


Private Sub Notify(ByVal msg As String)
    If mSilent Then Exit Sub
    MsgBox msg, vbInformation, "Chat Formatter"
End Sub


'-----------------------------------------------------------------------
'  Toolbar button
'-----------------------------------------------------------------------

Private Sub CreateToolbarButton(ByVal notifyUser As Boolean)
    Dim cb As CommandBar
    Dim btn As CommandBarButton
    Dim found As Boolean

    EnsureBullet

    On Error Resume Next
    For Each cb In Application.CommandBars
        If cb.Name = CF_TOOLBAR Then found = True
    Next cb
    If Not found Then
        Set cb = Application.CommandBars.Add(Name:=CF_TOOLBAR, Position:=msoBarTop, Temporary:=True)
    Else
        Set cb = Application.CommandBars(CF_TOOLBAR)
    End If
    If cb Is Nothing Then
        Err.Clear
        On Error GoTo 0
        Exit Sub                      ' command bars unavailable (Mac)
    End If

    Do While cb.Controls.Count > 0
        cb.Controls(1).Delete
    Loop

    Set btn = cb.Controls.Add(Type:=msoControlButton)
    With btn
        .Caption = "Format Chat Text"
        .OnAction = "FormatChatText"
        .Style = msoButtonCaption
        .FaceId = 293
        .TooltipText = "Convert pasted AI-chat markdown into Word formatting"
    End With
    cb.Visible = True
    Err.Clear
    On Error GoTo 0

    If notifyUser Then
        Notify "The ""Format Chat Text"" button was added to the toolbar." & vbCrLf & _
               "You can also run FormatChatText from the macro list."
    End If
End Sub


'-----------------------------------------------------------------------
'  Settings
'
'  A plain ini file under the user's home directory.  The previous
'  implementation used WScript.Shell and the Windows registry, which cannot
'  work on macOS at all.
'-----------------------------------------------------------------------

Private Function SettingsFile() As String
    Dim base As String
    base = Environ$("HOME")
    If Len(base) = 0 Then base = Environ$("USERPROFILE")
    If Len(base) = 0 Then base = Environ$("TMPDIR")
    If Len(base) = 0 Then base = CurDir$
    Do While Len(base) > 1 And Right$(base, 1) = Application.PathSeparator
        base = Left$(base, Len(base) - 1)
    Loop
    SettingsFile = base & Application.PathSeparator & ".chatformatter" & _
                   Application.PathSeparator & "settings.ini"
End Function


Private Function GetSetting(ByVal key As String, ByVal def As String) As String
    Dim f As Integer
    Dim isOpen As Boolean
    Dim line As String
    Dim path As String

    On Error GoTo UseDefault
    path = SettingsFile()
    f = FreeFile
    Open path For Input As #f
    isOpen = True
    Do While Not EOF(f)
        Line Input #f, line
        If InStr(line, "=") > 0 Then
            If Trim$(Left$(line, InStr(line, "=") - 1)) = key Then
                GetSetting = Trim$(Mid$(line, InStr(line, "=") + 1))
                Close #f
                Exit Function
            End If
        End If
    Loop
    Close #f
    isOpen = False
UseDefault:
    On Error Resume Next
    If isOpen Then Close #f
    Err.Clear
    On Error GoTo 0
    GetSetting = def
End Function


Private Sub SetSetting(ByVal key As String, ByVal val As String)
    Dim f As Integer
    Dim isOpen As Boolean
    Dim path As String, dirPath As String
    Dim line As String
    Dim keep() As String
    Dim nKeep As Long
    Dim i As Long
    Dim updated As Boolean

    On Error GoTo GiveUp

    path = SettingsFile()
    dirPath = Left$(path, Len(path) - Len("settings.ini") - 1)
    If Len(Dir$(dirPath, vbDirectory)) = 0 Then MkDir dirPath

    nKeep = 0
    ReDim keep(1 To 64)

    If Len(Dir$(path)) > 0 Then
        f = FreeFile
        Open path For Input As #f
        isOpen = True
        Do While Not EOF(f)
            Line Input #f, line
            If InStr(line, "=") > 0 Then
                If Trim$(Left$(line, InStr(line, "=") - 1)) = key Then
                    line = key & "=" & val
                    updated = True
                End If
                nKeep = nKeep + 1
                If nKeep > UBound(keep) Then ReDim Preserve keep(1 To nKeep * 2)
                keep(nKeep) = line
            End If
        Loop
        Close #f
        isOpen = False
    End If

    If Not updated Then
        nKeep = nKeep + 1
        If nKeep > UBound(keep) Then ReDim Preserve keep(1 To nKeep * 2)
        keep(nKeep) = key & "=" & val
    End If

    f = FreeFile
    Open path For Output As #f
    isOpen = True
    For i = 1 To nKeep
        Print #f, keep(i)
    Next i
    Close #f
    isOpen = False

GiveUp:
    On Error Resume Next
    If isOpen Then Close #f
    Err.Clear
    On Error GoTo 0
End Sub
