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
'    FormatChatText            format the whole active document
'    FormatSelection           format only the current selection
'    CF_SilentFormat           format the selection without any dialog
'    ToggleAutoFormat          enable/disable right-click auto format
'    ShowSettings              show current configuration
'
'  Automatic mode: when enabled, right-clicking a selection that came
'  from an AI chat formats it immediately.  Disable with ToggleAutoFormat.
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


' Used by the right-click hook: never interrupts with a dialog.
Public Sub CF_SilentFormat()
    Dim doc As Document
    Dim rng As Range

    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument
    Set rng = CurrentRange(doc)
    If rng Is Nothing Then Exit Sub

    mSilent = True
    RunPipeline doc, rng, (GetSetting("Tables", "1") = "1"), False
    mSilent = False
End Sub


Public Sub ToggleAutoFormat()
    If GetSetting("AutoFormat", "1") = "1" Then
        SetSetting "AutoFormat", "0"
        Notify "Auto-format mode disabled." & vbCrLf & _
               "Re-enable it from the macro list (Alt+F8, or Tools > Macro on Mac)."
    Else
        SetSetting "AutoFormat", "1"
        Notify "Auto-format mode enabled." & vbCrLf & _
               "Select AI text in Word and right-click to format it automatically."
    End If
End Sub


Public Sub ShowSettings()
    Dim autoFmt As String, tables As String

    If GetSetting("AutoFormat", "1") = "1" Then autoFmt = "Enabled" Else autoFmt = "Disabled"
    If GetSetting("Tables", "1") = "1" Then tables = "Enabled" Else tables = "Disabled"

    Notify "Chat Formatter " & CF_VERSION & vbCrLf & vbCrLf & _
           "Auto-format (right-click): " & autoFmt & vbCrLf & _
           "Table conversion: " & tables & vbCrLf & vbCrLf & _
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
    mSilent = True
    RunPipeline ActiveDocument, ActiveDocument.Content, True, False
    mSilent = False
End Sub


'=======================================================================
'  Event glue - called from ThisDocument
'=======================================================================

Public Sub CF_RightClickHandler(ByVal Target As Range)
    If mBusy Then Exit Sub
    If GetSetting("AutoFormat", "1") <> "1" Then Exit Sub
    If Documents.Count = 0 Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Target.Start = Target.End Then Exit Sub

    ' Cheap heuristic so ordinary prose is left alone.
    If Not LooksLikeAIChat(Target.Text) Then Exit Sub

    On Error Resume Next
    Application.ScreenUpdating = False
    CF_SilentFormat
    Application.ScreenUpdating = True
    Err.Clear
    On Error GoTo 0
End Sub


Private Function LooksLikeAIChat(ByVal s As String) As Boolean
    If Len(s) < 8 Then Exit Function
    If InStr(s, "**") > 0 Then LooksLikeAIChat = True: Exit Function
    If InStr(s, "##") > 0 Then LooksLikeAIChat = True: Exit Function
    If InStr(s, FENCE) > 0 Then LooksLikeAIChat = True: Exit Function
    If CountChar(s, "|") >= 2 And InStr(s, vbCr) > 0 Then LooksLikeAIChat = True: Exit Function
    If CountChar(s, vbTab) >= 1 And InStr(s, vbCr) > 0 Then LooksLikeAIChat = True: Exit Function
    If InStr(s, "- ") > 0 Or InStr(s, "> ") > 0 Then LooksLikeAIChat = True
End Function


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
'  bar.  mBusy is the re-entrancy guard for the right-click hook; it must
'  NOT gate Beat itself, or nothing would ever repaint.
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
    Dim pagin As Boolean
    Dim spell As Boolean
    Dim gram As Boolean

    If mBusy Then Exit Sub
    If rng Is Nothing Then Exit Sub

    EnsureBullet

    ' One linear pass fills the snapshot and classifies the blocks.
    nBlocks = ScanBlocks(rng, pStart, pText, blks)
    If nBlocks = 0 Then
        If notifyUser Then Notify "There is nothing here to format."
        Exit Sub
    End If
    nParas = UBound(pStart) - 1

    BeginRun nBlocks + 2

    su = Application.ScreenUpdating
    On Error Resume Next
    pagin = Application.Options.Pagination
    spell = Application.Options.CheckSpellingAsYouType
    gram = Application.Options.CheckGrammarAsYouType
    Err.Clear
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
        hi = pStart(blks(i).Last + 1)
        If blks(i).Last >= nParas Then hi = ClampEnd(doc, pStart(nParas + 1))
        Beat "formatting"
        If hi > lo Then
            Select Case blks(i).Kind
                Case BK_CODE
                    ApplyCodeBlock doc, lo, hi, blks(i), pText
                Case BK_TSV
                    ApplyFenceWithBody doc, lo, hi, blks(i), pText, doTables
                Case BK_TABRUN
                    ApplyTabRun doc, lo, hi, blks(i), pText, doTables
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
    Application.Options.Pagination = pagin
    Application.Options.CheckSpellingAsYouType = spell
    Application.Options.CheckGrammarAsYouType = gram
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


'=======================================================================
'  Snapshot and classification
'
'  Fills pStart(1..n+1) with paragraph offsets and pText(1..n) with raw
'  paragraph text, then cuts the range into typed blocks.  Returns the
'  block count.  This is the only pass that walks the live document.
'=======================================================================

Private Function ScanBlocks(ByVal rng As Range, _
                            ByRef pStart() As Long, _
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

    ReDim pStart(1 To rng.Paragraphs.Count + 1)
    ReDim pText(1 To rng.Paragraphs.Count)
    ReDim blks(1 To 32)

    ' Text inside a table belongs to Word, not to us.  Leaving those
    ' paragraphs out of the snapshot is what stops a second run from
    ' dragging tables that are already finished back into the pipeline.
    n = 0
    For Each p In rng.Paragraphs
        If Not p.Range.Information(wdWithInTable) Then
            n = n + 1
            pStart(n) = p.Range.Start
            pText(n) = p.Range.Text
            Beat "scanning"
        End If
    Next p
    If n < 1 Then
        ScanBlocks = 0
        Exit Function
    End If
    ReDim Preserve pStart(1 To n + 1)
    ReDim Preserve pText(1 To n)
    pStart(n + 1) = rng.End

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
                count = FlushPlain(blks, count, plainFrom, i - 1, pText)
                count = AddBlock(blks, count, BK_CODE, i, n, lang, False)
                ScanBlocks = count
                Exit Function
            End If

            count = FlushPlain(blks, count, plainFrom, i - 1, pText)

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

    count = FlushPlain(blks, count, plainFrom, n, pText)
    ScanBlocks = count
End Function


' Emits the ordinary region between two fence blocks, split so that runs
' of tab separated rows become their own blocks.  Keeping them separate
' means the text pass never has to reason about table insertion.
Private Function FlushPlain(ByRef blks() As Blk, ByVal count As Long, _
                            ByVal firstIdx As Long, ByVal lastIdx As Long, _
                            ByRef pText() As String) As Long
    Dim i As Long
    Dim runFrom As Long
    Dim s As String

    i = firstIdx
    Do While i <= lastIdx
        s = CleanParaText(pText(i))
        If InStr(s, vbTab) > 0 And Len(s) > 0 And Not IsListMarked(s) Then
            If runFrom = 0 Then runFrom = i
        Else
            If runFrom > 0 Then
                count = AddBlock(blks, count, BK_TABRUN, runFrom, i - 1, "", True)
                runFrom = 0
            End If
            If Len(s) = 0 Or IsFenceLine(s, s) Then
                ' blank lines travel with the text that follows them
            Else
                count = AddBlock(blks, count, BK_PLAIN, i, i, "", True)
            End If
        End If
        i = i + 1
    Loop
    If runFrom > 0 Then
        count = AddBlock(blks, count, BK_TABRUN, runFrom, lastIdx, "", True)
    End If
    FlushPlain = count
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
    If IsNumeric(head) Then IsListMarked = True
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
    Dim r As Range

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

    doc.Range(lo, hi).Text = joined
    Set r = doc.Range(lo, lo + Len(joined))
    If r.End <= r.Start Then Exit Sub

    On Error Resume Next
    With r.Font
        .Name = "Consolas"
        .Size = 10
        .Color = RGB(45, 45, 45)
        .Bold = False
        .Italic = False
    End With
    With r.ParagraphFormat
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
    Dim nBody As Long, i As Long, a As Long, c As Long
    Dim runStart() As Long
    Dim runEnd() As Long
    Dim runText() As String
    Dim rFrom() As Long
    Dim rTo() As Long
    Dim rTab() As Boolean
    Dim nRuns As Long
    Dim joined As String
    Dim pos As Long
    Dim s As String
    Dim isTab As Boolean
    Dim canExtend As Boolean
    Dim shift As Long
    Dim cols As Long, rows As Long
    Dim data() As String

    If b.Closed Then
        nBody = b.Last - b.First - 1
    Else
        nBody = b.Last - b.First
    End If
    If nBody < 1 Then
        doc.Range(lo, hi).Delete
        Exit Sub
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
        pos = pos + Len(s) + 1
        If i > 1 Then joined = joined & vbCr
        joined = joined & s
    Next i
    For i = 1 To nBody
        runEnd(i) = runStart(i) + Len(runText(i)) + 1
    Next i

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
    ' so the run offsets above have to follow before anything else is cut.
    shift = Len(pText(b.First))
    doc.Range(lo, runStart(1)).Delete
    If runEnd(nBody) < hi Then doc.Range(runEnd(nBody) - shift, hi - shift).Delete
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
        End If
    Next a

    ' Prose inside the fence still wants its inline formatting.
    On Error Resume Next
    InlinePass doc.Range(lo, lo + Len(joined))
    Err.Clear
    On Error GoTo 0
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
    If rows < 1 Or cols < 2 Then Exit Sub

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


Private Sub FillRow(ByRef data() As String, ByVal rowIdx As Long, _
                   ByVal s As String, ByVal cols As Long)
    Dim parts() As String
    Dim k As Long, n As Long

    s = Replace(s, Chr$(11), "")
    parts = Split(s, vbTab)
    n = UBound(parts)
    For k = 0 To cols - 1
        If k <= n Then
            data(rowIdx, k) = Trim$(parts(k))
        Else
            data(rowIdx, k) = ""
        End If
    Next k
End Sub


' Ordinary markdown text.  Block level markers are resolved first, then the
' whole block is rewritten once and formatted inside the new range.
Private Sub ApplyPlainBlock(ByVal doc As Document, ByVal lo As Long, ByVal hi As Long, _
                            ByRef b As Blk, ByRef pText() As String)
    Dim n As Long, i As Long, m As Long
    Dim outText() As String
    Dim outKind() As Long
    Dim outArg() As Long
    Dim raw As String, s As String
    Dim joined As String
    Dim r As Range
    Dim p As Paragraph
    Dim idx As Long
    Dim lvl As Long
    Dim num As Long
    Dim isRtl As Boolean

    n = b.Last - b.First + 1
    ReDim outText(1 To n)
    ReDim outKind(1 To n)
    ReDim outArg(1 To n)

    m = 0
    For i = 1 To n
        raw = pText(b.First + i - 1)
        s = CleanParaText(raw)

        If Len(s) = 0 Then
            m = m + 1
            outText(m) = ""
            outKind(m) = LK_TEXT
            outArg(m) = 0
        Else
            lvl = AtxLevel(s)
            If lvl > 0 Then
                m = m + 1
                outText(m) = Mid$(s, lvl + 2)
                outKind(m) = LK_HEADING
                outArg(m) = lvl
            ElseIf m > 0 And IsSetextLine(s) Then
                ' The line above this one becomes the heading.
                outKind(m) = LK_HEADING
                If Left$(s, 1) = "=" Then
                    outArg(m) = 1
                ElseIf Not IsBulletOutput(outText(m)) Then
                    outArg(m) = 2
                Else
                    outKind(m) = LK_TEXT
                    outArg(m) = 0
                End If
            ElseIf IsHorizontalRule(s) Then
                ' dropped entirely
            ElseIf Left$(s, 1) = ">" Then
                m = m + 1
                outText(m) = TrimQuoteMarker(s)
                outKind(m) = LK_QUOTE
                outArg(m) = 0
            ElseIf IsBulletStart(s) Then
                m = m + 1
                outText(m) = mBullet & vbTab & Mid$(s, 3)
                outKind(m) = LK_BULLET
                outArg(m) = 0
            ElseIf IsNumberedStart(s, num) Then
                m = m + 1
                outText(m) = CStr(num) & "." & vbTab & NumberedBody(s, num)
                outKind(m) = LK_NUMBERED
                outArg(m) = 0
            Else
                m = m + 1
                outText(m) = s
                outKind(m) = LK_TEXT
                outArg(m) = 0
            End If
        End If
    Next i

    If m = 0 Then
        doc.Range(lo, hi).Delete
        Exit Sub
    End If

    For i = 1 To m
        If i > 1 Then joined = joined & vbCr
        joined = joined & outText(i)
    Next i

    doc.Range(lo, hi).Text = joined
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
                    p.Range.ParagraphFormat.RightIndent = 24
                    p.Range.ParagraphFormat.LeftIndent = 0
                Else
                    p.Range.ParagraphFormat.LeftIndent = 24
                    p.Range.ParagraphFormat.RightIndent = 0
                End If
                p.Range.ParagraphFormat.FirstLineIndent = -18
            Case LK_QUOTE
                p.Range.Font.Italic = True
                p.Range.Font.Color = RGB(85, 85, 85)
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


' Accepts "12. text" and reports the number.
Private Function IsNumberedStart(ByVal s As String, ByRef num As Long) As Boolean
    Dim pos As Long
    pos = InStr(s, ". ")
    If pos < 2 Or pos > 5 Then Exit Function
    If Not IsNumeric(Left$(s, pos - 1)) Then Exit Function
    num = CLng(Left$(s, pos - 1))
    IsNumberedStart = True
End Function


Private Function NumberedBody(ByVal s As String, ByVal num As Long) As String
    NumberedBody = Mid$(s, Len(CStr(num)) + 3)
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
            Case IL_ITALIC
                .Replacement.Font.Italic = True
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
        AutoFitBehavior:=wdAutoFitContent)

    For a = 0 To rows - 1
        For b = 0 To cols - 1
            t.Cell(a + 1, b + 1).Range.Text = data(a, b)
        Next b
    Next a

    FormatTable t
    Exit Sub

Fallback:
    On Error Resume Next
    doc.Range(pos, pos).InsertAfter restored
    Err.Clear
    On Error GoTo 0
End Sub


Private Sub FormatTable(ByVal tbl As Table)
    Dim styled As Boolean

    On Error Resume Next

    ' "Table Grid" is English-only in most builds; fall back to direct
    ' borders when the localized style name does not resolve.
    tbl.Style = "Table Grid"
    If Err.Number = 0 Then styled = True
    Err.Clear

    If Not styled Then
        tbl.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        tbl.Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
        tbl.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        tbl.Borders(wdBorderRight).LineStyle = wdLineStyleSingle
        tbl.Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
        tbl.Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
    End If

    tbl.Range.Font.Size = 11
    tbl.Range.ParagraphFormat.SpaceBefore = 2
    tbl.Range.ParagraphFormat.SpaceAfter = 2
    tbl.Range.ParagraphFormat.LeftIndent = 0
    tbl.Range.ParagraphFormat.RightIndent = 0
    tbl.Range.ParagraphFormat.FirstLineIndent = 0

    If tbl.Rows.Count >= 1 Then
        tbl.Rows(1).Range.Font.Bold = True
        tbl.Rows(1).Shading.BackgroundPatternColor = RGB(219, 229, 241)
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
        If Right$(s, 1) = Chr$(13) Or Right$(s, 1) = Chr$(7) Then
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
    SettingsFile = base & Application.PathSeparator & ".chatformatter" & _
                   Application.PathSeparator & "settings.ini"
End Function


Private Function GetSetting(ByVal key As String, ByVal def As String) As String
    Dim f As Integer
    Dim line As String
    Dim path As String

    On Error GoTo UseDefault
    path = SettingsFile()
    f = FreeFile
    Open path For Input As #f
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
UseDefault:
    On Error Resume Next
    Err.Clear
    On Error GoTo 0
    GetSetting = def
End Function


Private Sub SetSetting(ByVal key As String, ByVal val As String)
    Dim f As Integer
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

    If Len(Dir$(path)) > 0 Then
        f = FreeFile
        Open path For Input As #f
        nKeep = 0
        ReDim keep(1 To 64)
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
    End If

    If Not updated Then
        nKeep = nKeep + 1
        If nKeep > UBound(keep) Then ReDim Preserve keep(1 To nKeep * 2)
        keep(nKeep) = key & "=" & val
    End If

    f = FreeFile
    Open path For Output As #f
    For i = 1 To nKeep
        Print #f, keep(i)
    Next i
    Close #f

GiveUp:
    On Error Resume Next
    Err.Clear
    On Error GoTo 0
End Sub


ute VB_Name = "ChatFormatter"
