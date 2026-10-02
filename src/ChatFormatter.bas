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
'    ToggleAutoFormat          enable/disable right-click auto format
'    ShowSettings              show current configuration
'
'  Automatic mode: when enabled, right-clicking a selection that
'  came from an AI chat formats it immediately. Disable with
'  ToggleAutoFormat.
'=======================================================================

Public Const CF_VERSION As String = "1.0.0"
Public Const CF_TOOLBAR As String = "ChatFormatterToolbar"

' Bullet glyph. Chr(149) is undefined in Windows-1256/Arabic code pages
' and renders as garbage; 0x2022 is the real Unicode bullet.
Private mBullet As String


'-----------------------------------------------------------------------
'  Public API
'-----------------------------------------------------------------------

' Converts text that looks like markdown/TSV pasted from an AI chat.
' Safe to call repeatedly: on already-formatted text every pass finds
' nothing to change and the document is left alone.
Public Sub FormatChatText()
    Dim doc As Document
    Dim rng As Range

    If Documents.Count = 0 Then
        MsgBox "No document is open.", vbExclamation, "Chat Formatter"
        Exit Sub
    End If
    Set doc = ActiveDocument

    Dim resp As VbMsgBoxResult
    resp = MsgBox( _
        "Convert TSV and Pipe tables to Word tables too?" & vbCrLf & vbCrLf & _
        "Yes = Markdown formatting + tables" & vbCrLf & _
        "No = Markdown formatting only", _
        vbYesNo + vbQuestion, "Chat Formatter")

    Set rng = doc.Content

    If Not RangeNeedsFormatting(rng) Then
        MsgBox "The selected text does not need formatting." & vbCrLf & _
               "The text is already clean.", vbInformation, "Chat Formatter"
        Exit Sub
    End If

    FormatRangeCore doc, rng, (resp = vbYes)
End Sub


Public Sub FormatSelection()
    Dim doc As Document
    Dim rng As Range
    If Documents.Count = 0 Then
        MsgBox "No document is open.", vbExclamation, "Chat Formatter"
        Exit Sub
    End If
    Set doc = ActiveDocument

    ' Use the current selection when Word is visible and a real selection
    ' exists. In headless/COM contexts Application.Selection can hang, so
    ' fall back to the whole document content.
    If Application.Visible Then
        On Error Resume Next
        Set rng = Selection.Range
        If Err.Number <> 0 Or rng Is Nothing Then
            Err.Clear
            Set rng = doc.Content
        ElseIf rng.Start = rng.End Then
            Set rng = doc.Content
        End If
        On Error GoTo 0
    Else
        Set rng = doc.Content
    End If

    If rng Is Nothing Then
        MsgBox "No document available for formatting.", vbExclamation, "Chat Formatter"
        Exit Sub
    End If

    If Not RangeNeedsFormatting(rng) Then
        MsgBox "The selected text does not need formatting." & vbCrLf & _
               "The text is already clean.", vbInformation, "Chat Formatter"
        Exit Sub
    End If

    FormatRangeCore doc, rng, (GetSetting("Tables", "1") = "1")
End Sub


' Cheap scan for any markdown construct we know how to fix. Used to make
' repeated runs idempotent instead of re-processing clean text.
Private Function RangeNeedsFormatting(ByVal rng As Range) As Boolean
    Dim s As String
    On Error GoTo NotPlain

    s = rng.Text
    If Len(s) = 0 Then Exit Function

    If InStr(s, "**") > 0 Then RangeNeedsFormatting = True: Exit Function
    If InStr(s, "~~") > 0 Then RangeNeedsFormatting = True: Exit Function
    If InStr(s, "```") > 0 Then RangeNeedsFormatting = True: Exit Function
    If InStr(s, "|") > 0 Then RangeNeedsFormatting = True: Exit Function
    If CountChar(s, vbTab) > 0 Then RangeNeedsFormatting = True: Exit Function
    If InStr(s, "[") > 0 And InStr(s, "](") > 0 Then RangeNeedsFormatting = True: Exit Function

    Dim k As Long
    Dim p As Paragraph
    Dim t As String
    For k = 1 To rng.Paragraphs.Count
        Set p = rng.Paragraphs(k)
        t = CleanParaText(p.Range.Text)
        If Left(t, 2) = "# " Or Left(t, 3) = "## " Or Left(t, 4) = "### " _
           Or Left(t, 5) = "#### " Or Left(t, 6) = "##### " Then
            RangeNeedsFormatting = True: Exit Function
        End If
        If Left(t, 2) = "> " Then RangeNeedsFormatting = True: Exit Function
        If Left(t, 2) = "- " Or Left(t, 2) = "* " Or Left(t, 2) = "+ " Then
            RangeNeedsFormatting = True: Exit Function
        End If
        If Left(t, 1) = "`" Then RangeNeedsFormatting = True: Exit Function
    Next k

    Exit Function
NotPlain:
    ' Unreadable range (mixed cell markers etc.) - assume it needs work.
    RangeNeedsFormatting = True
End Function


'-----------------------------------------------------------------------
'  Auto mode - right-click a pasted AI answer and it formats itself
'-----------------------------------------------------------------------

Public Sub ToggleAutoFormat()
    Dim v As Boolean
    v = GetSetting("AutoFormat", "1")
    If v Then
        SetSetting "AutoFormat", "0"
        MsgBox "Auto-format mode disabled." & vbCrLf & _
                     "To re-enable: Alt+F8 then ToggleAutoFormat", _
                  vbInformation, "Chat Formatter"
    Else
        SetSetting "AutoFormat", "1"
        MsgBox "Auto-format mode enabled." & vbCrLf & _
                     "Select AI text in Word and right-click to format it automatically.", _
                  vbInformation, "Chat Formatter"
    End If
End Sub


Public Sub ShowSettings()
    Dim autoFmt As String, tables As String
    If GetSetting("AutoFormat", "1") = "1" Then autoFmt = "Enabled"
    Else autoFmt = "Disabled"
    End If
    If GetSetting("Tables", "1") = "1" Then tables = "Enabled"
    Else tables = "Disabled"
    End If
    MsgBox "Chat Formatter " & CF_VERSION & vbCrLf & vbCrLf & _
           "Auto-format (right-click): " & autoFmt & vbCrLf & _
           "Table conversion: " & tables & vbCrLf & vbCrLf & _
           "Toolbar:", _
           vbInformation, "Chat Formatter Settings"
End Sub


' Called by ThisDocument when Word starts or the template loads.
Public Sub AutoExec()
    On Error Resume Next
    EnsureBullet
    CreateToolbarButton False
    On Error GoTo 0
End Sub


Public Sub AutoOpen()
    AutoExec
End Sub


' Rebuilds the toolbar button (used by install.bat and after repair).
Public Sub AutoSetup()
    On Error Resume Next
    EnsureBullet
    CreateToolbarButton True
    On Error GoTo 0
End Sub


Public Sub RemoveToolbar()
    On Error Resume Next
    Dim cb As CommandBar
    For Each cb In Application.CommandBars
        If cb.Name = CF_TOOLBAR Then cb.Delete
    Next cb
    On Error GoTo 0
End Sub


'-----------------------------------------------------------------------
'  Event glue - called from ThisDocument
'-----------------------------------------------------------------------

Public Sub CF_RightClickHandler(ByVal Target As Range)
    Static busy As Boolean
    Dim doc As Document
    Dim rng As Range

    If busy Then Exit Sub
    If GetSetting("AutoFormat", "1") <> "1" Then Exit Sub
    If Documents.Count = 0 Then Exit Sub
    If Target Is Nothing Then Exit Sub

    ' Only act on a non-empty selection.
    If Target.Start = Target.End Then Exit Sub
    If Target.Start < Target.Document.Content.Start Then Exit Sub

    Set doc = Target.Document
    Set rng = doc.Range(Start:=Target.Start, End:=Target.End)

    ' Bail out if the selection does not look like AI-chat markdown.
    If Not LooksLikeAIChat(rng.Text) Then Exit Sub

    busy = True
    On Error Resume Next
    Application.ScreenUpdating = False
    FormatRangeCore doc, rng, (GetSetting("Tables", "1") = "1")
    Application.ScreenUpdating = True
    Err.Clear
    On Error GoTo 0
    busy = False
End Sub


' Cheap heuristic to avoid re-formatting arbitrary prose on right-click.
Private Function LooksLikeAIChat(ByVal s As String) As Boolean
    Dim n As Long
    If Len(s) < 8 Then Exit Function
    If InStr(s, "**") > 0 Then LooksLikeAIChat = True: Exit Function
    If InStr(s, "##") > 0 Then LooksLikeAIChat = True: Exit Function
    If InStr(s, "`") > 0 Then LooksLikeAIChat = True: Exit Function
    If InStr(s, "```") > 0 Then LooksLikeAIChat = True: Exit Function
    n = CountChar(s, "|")
    If n >= 2 And CountChar(s, vbCr) >= 1 Then LooksLikeAIChat = True: Exit Function
    n = CountChar(s, vbTab)
    If n >= 1 And CountChar(s, vbCr) >= 1 Then LooksLikeAIChat = True: Exit Function
    If InStr(s, "- ") > 0 Or InStr(s, "> ") > 0 Then LooksLikeAIChat = True: Exit Function
End Function


Public Sub CF_SilentTest()
    Dim doc As Document
    Dim rng As Range
    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument
    Set rng = doc.Content
    silentMode = True
    FormatRangeCore doc, rng, True
    silentMode = False
End Sub


'-----------------------------------------------------------------------
'  Core formatter
'-----------------------------------------------------------------------

Private Sub FormatRangeCore(ByVal doc As Document, ByVal rng As Range, ByVal doTables As Boolean)
    Dim su As Boolean
    Dim undoStarted As Boolean
    Dim failed As Boolean
    Dim errNum As Long, errDesc As String

    RunStageCore doc, rng, doTables
End Sub




Private Sub RunStageCore(ByVal doc As Document, ByVal rng As Range, ByVal doTables As Boolean)
    Dim su As Boolean
    Dim undoStarted As Boolean
    Dim failed As Boolean
    Dim errNum As Long, errDesc As String


    EnsureBullet

    su = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error Resume Next
    undoStarted = False
    Err.Clear
    On Error GoTo CleanUp


    If doTables Then
        ConvertPipeTablesToTable doc, rng
        ConvertTSVToTable doc, rng
    End If


    RemoveHorizontalRules doc, rng
    ConvertCodeBlocks doc, rng
    ConvertHeadings doc, rng
    FormatBold doc, rng
    FormatItalic doc, rng
    FormatStrike doc, rng
    FormatInlineCode doc, rng
    FormatLinks doc, rng
    ConvertBlockquotes doc, rng
    ConvertBulletLists doc, rng
    ConvertNumberedLists doc, rng
    CleanStrayMarkers doc, rng

CleanUp:
    failed = (Err.Number <> 0)
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    Application.ScreenUpdating = su
    Err.Clear
    On Error GoTo 0

    If failed Then
        MsgBox "Error during formatting:" & vbCrLf & _
               errDesc & vbCrLf & "Error code: " & errNum & vbCrLf & vbCrLf & _
               "Undo: Ctrl+Z", vbExclamation, "Chat Formatter"
    Else
        MsgBox "Formatting completed successfully!" & vbCrLf & _
               "(Undo: Ctrl+Z)", vbInformation, "Chat Formatter"
    End If
End Sub


'-----------------------------------------------------------------------
'  1) Horizontal rules
'-----------------------------------------------------------------------

Private Sub RemoveHorizontalRules(ByVal doc As Document, ByVal rng As Range)
    Dim k As Long, p As Paragraph
    Dim t As String, body As String

    For k = doc.Paragraphs.Count To 1 Step -1
        Set p = doc.Paragraphs(k)
        t = CleanParaText(p.Range.Text)
        body = Replace(Replace(t, "-", ""), "*", "")
        body = Replace(body, "_", "")
        body = Replace(body, "=", "")
        body = Trim(body)
        If Len(t) >= 3 And Len(body) = 0 Then
            p.Range.Delete
        End If
    Next k
End Sub


'-----------------------------------------------------------------------
'  2) Fenced code blocks ``` ... ```
'-----------------------------------------------------------------------

Private Sub ConvertCodeBlocks(ByVal doc As Document, ByVal rng As Range)
    Dim i As Long, j As Long
    Dim inCode As Boolean
    Dim openIdx As Long
    Dim guard As Long

    inCode = False
    openIdx = 0

    For i = doc.Paragraphs.Count To 1 Step -1
        guard = guard + 1
        If guard > 20000 Then Exit For

        If InStr(doc.Paragraphs(i).Range.Text, "```") > 0 Then
            If Not inCode Then
                openIdx = i
                inCode = True
            Else
                ' opening fence = openIdx, closing fence = i
                For j = openIdx + 1 To i - 1
                    On Error Resume Next
                    With doc.Paragraphs(j).Range
                        .Font.Name = "Consolas"
                        .Font.Size = 10
                        .Font.Color = RGB(45, 45, 45)
                        .Font.Bold = False
                        .Font.Italic = False
                        .ParagraphFormat.LeftIndent = 18
                        .ParagraphFormat.RightIndent = 18
                        .ParagraphFormat.SpaceBefore = 0
                        .ParagraphFormat.SpaceAfter = 0
                        .Shading.BackgroundPatternColor = RGB(242, 242, 242)
                    End With
                    On Error GoTo 0
                Next j

                On Error Resume Next
                doc.Paragraphs(openIdx).Range.Shading.BackgroundPatternColor = wdColorAutomatic
                doc.Paragraphs(i).Range.Shading.BackgroundPatternColor = wdColorAutomatic
                doc.Paragraphs(openIdx).Range.Delete
                doc.Paragraphs(i - 1).Range.Delete
                Err.Clear
                On Error GoTo 0

                inCode = False
                openIdx = 0
            End If
        End If
    Next i

    ' Unterminated fence: strip the marker, leave the text as normal body.
    If inCode And openIdx > 0 Then
        On Error Resume Next
        doc.Paragraphs(openIdx).Range.Shading.BackgroundPatternColor = wdColorAutomatic
        doc.Paragraphs(openIdx).Range.Delete
        Err.Clear
        On Error GoTo 0
    End If
End Sub


'-----------------------------------------------------------------------
'  3) Headings: # .. ##### and setext (=== / ---)
'-----------------------------------------------------------------------

Private Sub ConvertHeadings(ByVal doc As Document, ByVal rng As Range)
    Dim k As Long
    Dim p As Paragraph
    Dim r As Range
    Dim txt As String
    Dim prevTxt As String

    For k = doc.Paragraphs.Count To 1 Step -1
        Set p = doc.Paragraphs(k)
        Set r = p.Range
        If r.End > r.Start Then
            If r.Characters.Last.Text = Chr(13) Then r.MoveEnd wdCharacter, -1
            txt = r.Text
        Else
            txt = ""
        End If

        If Len(txt) > 0 Then
            ' ATX headings, level 1..5
            If Left(txt, 6) = "##### " Then
                ApplyHeading doc, p, r, Mid(txt, 7), wdStyleHeading5
            ElseIf Left(txt, 5) = "#### " Then
                ApplyHeading doc, p, r, Mid(txt, 6), wdStyleHeading4
            ElseIf Left(txt, 4) = "### " Then
                ApplyHeading doc, p, r, Mid(txt, 5), wdStyleHeading3
            ElseIf Left(txt, 3) = "## " Then
                ApplyHeading doc, p, r, Mid(txt, 4), wdStyleHeading2
            ElseIf Left(txt, 2) = "# " Then
                ApplyHeading doc, p, r, Mid(txt, 3), wdStyleHeading1
            Else
                ' Setext: previous paragraph is a heading candidate and this
                ' line is all = or all -
                prevTxt = ""
                If k < doc.Paragraphs.Count Then prevTxt = CleanParaText(doc.Paragraphs(k + 1).Range.Text)
                If Len(prevTxt) > 0 And Len(txt) >= 2 Then
                    If IsAllOf(txt, "=") Then
                        ApplyHeading doc, p, r, prevTxt, wdStyleHeading1
                    ElseIf IsAllOf(txt, "-") And Not IsBulletLine(prevTxt) Then
                        ApplyHeading doc, p, r, prevTxt, wdStyleHeading2
                    End If
                End If
            End If
        End If
    Next k
End Sub


Private Sub ApplyHeading(ByVal doc As Document, ByVal p As Paragraph, ByVal r As Range, _
                         ByVal newText As String, ByVal styleId As WdBuiltinStyle)
    On Error Resume Next
    r.Text = newText
    p.Range.Style = doc.Styles(styleId)
    p.Range.ParagraphFormat.LeftIndent = 0
    p.Range.ParagraphFormat.FirstLineIndent = 0
    Err.Clear
    On Error GoTo 0
End Sub


Private Function IsAllOf(ByVal s As String, ByVal ch As String) As Boolean
    Dim k As Long
    If Len(Trim(s)) < 2 Then Exit Function
    For k = 1 To Len(s)
        If Mid(s, k, 1) <> ch Then Exit Function
    Next k
    IsAllOf = True
End Function


Private Function IsBulletLine(ByVal s As String) As Boolean
    If Left(s, 2) = "- " Or Left(s, 2) = "* " Or Left(s, 2) = "+ " Then IsBulletLine = True
End Function


'-----------------------------------------------------------------------
'  4-7) Inline markdown via wildcard Find
'-----------------------------------------------------------------------

Private Sub FormatBold(ByVal doc As Document, ByVal rng As Range)
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = "\*\*([!\*]@)\*\*"
        .Replacement.Text = "\1"
        .Replacement.Font.Bold = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub


Private Sub FormatItalic(ByVal doc As Document, ByVal rng As Range)
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = "\*([!\*]@)\*"
        .Replacement.Text = "\1"
        .Replacement.Font.Italic = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub


Private Sub FormatStrike(ByVal doc As Document, ByVal rng As Range)
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = "\~\~([!\~]@)\~\~"
        .Replacement.Text = "\1"
        .Replacement.Font.StrikeThrough = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub


Private Sub FormatInlineCode(ByVal doc As Document, ByVal rng As Range)
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = "`([!\`]@)`"
        .Replacement.Text = "\1"
        .Replacement.Font.Name = "Consolas"
        .Replacement.Font.Size = 10
        .Replacement.Font.Color = RGB(150, 40, 40)
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub


Private Sub FormatLinks(ByVal doc As Document, ByVal rng As Range)
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = "\[([!\[\]]@)\]\([!\(\)]@\)"
        .Replacement.Text = "\1"
        .Replacement.Font.Color = RGB(0, 0, 200)
        .Replacement.Font.Underline = wdUnderlineSingle
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub


'-----------------------------------------------------------------------
'  8) Blockquotes
'-----------------------------------------------------------------------

Private Sub ConvertBlockquotes(ByVal doc As Document, ByVal rng As Range)
    Dim k As Long
    Dim p As Paragraph
    Dim r As Range
    Dim txt As String

    For k = doc.Paragraphs.Count To 1 Step -1
        Set p = doc.Paragraphs(k)
        Set r = p.Range
        If r.End > r.Start Then
            If r.Characters.Last.Text = Chr(13) Then r.MoveEnd wdCharacter, -1
            txt = r.Text
            If Left(txt, 2) = "> " Then
                On Error Resume Next
                r.Text = Mid(txt, 3)
                With p.Range
                    .Font.Italic = True
                    .Font.Color = RGB(85, 85, 85)
                    .ParagraphFormat.LeftIndent = 24
                    .ParagraphFormat.RightIndent = 24
                    .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                    .Borders(wdBorderRight).LineWidth = wdLineWidth150pt
                    .Borders(wdBorderRight).Color = RGB(190, 190, 190)
                End With
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next k
End Sub


'-----------------------------------------------------------------------
'  9) Bullet lists - real Unicode bullet, RTL aware
'-----------------------------------------------------------------------

Private Sub ConvertBulletLists(ByVal doc As Document, ByVal rng As Range)
    Dim k As Long
    Dim p As Paragraph
    Dim r As Range
    Dim txt As String
    Dim body As String
    Dim isRtl As Boolean

    EnsureBullet

    For k = doc.Paragraphs.Count To 1 Step -1
        Set p = doc.Paragraphs(k)
        Set r = p.Range
        If r.End > r.Start Then
            If r.Characters.Last.Text = Chr(13) Then r.MoveEnd wdCharacter, -1
            txt = r.Text
            If Left(txt, 2) = "- " Or Left(txt, 2) = "* " Or Left(txt, 2) = "+ " Then
                body = Mid(txt, 3)
                isRtl = IsRtlText(body)
                On Error Resume Next
                r.Text = mBullet & vbTab & body
                If isRtl Then
                    p.Range.ParagraphFormat.RightIndent = 24
                    p.Range.ParagraphFormat.LeftIndent = 0
                    p.Range.ParagraphFormat.FirstLineIndent = -18
                Else
                    p.Range.ParagraphFormat.LeftIndent = 24
                    p.Range.ParagraphFormat.RightIndent = 0
                    p.Range.ParagraphFormat.FirstLineIndent = -18
                End If
                Err.Clear
                On Error GoTo 0
            End If
        End If
    Next k
End Sub


'-----------------------------------------------------------------------
'  10) Numbered lists - the original number is preserved as text
'      (auto-numbering restarts per block and breaks RTL layout)
'-----------------------------------------------------------------------

Private Sub ConvertNumberedLists(ByVal doc As Document, ByVal rng As Range)
    Dim k As Long
    Dim p As Paragraph
    Dim r As Range
    Dim txt As String
    Dim pos As Long
    Dim numPart As String
    Dim body As String
    Dim isRtl As Boolean

    For k = doc.Paragraphs.Count To 1 Step -1
        Set p = doc.Paragraphs(k)
        Set r = p.Range
        If r.End > r.Start Then
            If r.Characters.Last.Text = Chr(13) Then r.MoveEnd wdCharacter, -1
            txt = r.Text
            pos = InStr(txt, ". ")
            If pos > 1 And pos <= 4 Then
                numPart = Left(txt, pos - 1)
                If IsNumeric(numPart) Then
                    body = Mid(txt, pos + 2)
                    isRtl = IsRtlText(body)
                    On Error Resume Next
                    r.Text = numPart & "." & vbTab & body
                    If isRtl Then
                        p.Range.ParagraphFormat.RightIndent = 24
                        p.Range.ParagraphFormat.LeftIndent = 0
                        p.Range.ParagraphFormat.FirstLineIndent = -18
                    Else
                        p.Range.ParagraphFormat.LeftIndent = 24
                        p.Range.ParagraphFormat.RightIndent = 0
                        p.Range.ParagraphFormat.FirstLineIndent = -18
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
        End If
    Next k
End Sub


'-----------------------------------------------------------------------
'  11) Stray markers left over from partially formatted text
'-----------------------------------------------------------------------

' Row delimiters and cell markers left behind once the block becomes a
' table are stripped by deleting the block before insertion, so nothing
' extra is needed here beyond stray inline asterisks.
Private Sub CleanStrayMarkers(ByVal doc As Document, ByVal rng As Range)
    With rng.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = "^13\* "
        .Replacement.Text = "^13"
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub


'-----------------------------------------------------------------------
'  Tables - Markdown pipe tables
'-----------------------------------------------------------------------

Private Sub ConvertPipeTablesToTable(ByVal doc As Document, ByVal rng As Range)
    Dim i As Long, j As Long
    Dim startIdx As Long, endIdx As Long
    Dim sepText As String, rowText As String
    Dim data() As String
    Dim rows As Long, cols As Long
    Dim guard As Long

    i = 1
    Do While i <= doc.Paragraphs.Count
        guard = guard + 1
        If guard > 20000 Then Exit Do

        rowText = CleanParaText(doc.Paragraphs(i).Range.Text)

        If IsPipeRow(rowText) And i + 1 <= doc.Paragraphs.Count Then
            sepText = CleanParaText(doc.Paragraphs(i + 1).Range.Text)
            If IsSeparatorLine(sepText) Then
                startIdx = i
                endIdx = i + 1
                j = i + 2
                Do While j <= doc.Paragraphs.Count
                    rowText = CleanParaText(doc.Paragraphs(j).Range.Text)
                    If IsPipeRow(rowText) Then
                        endIdx = j
                        j = j + 1
                    Else
                        Exit Do
                    End If
                Loop

                rows = endIdx - startIdx
                cols = PipeColumnCount(doc.Paragraphs(startIdx).Range.Text)

                If rows >= 1 And cols >= 2 Then
                    ReDim data(0 To rows - 1, 0 To cols - 1)
                    For j = 0 To rows - 1
                        Dim cells() As String
                        cells = SplitPipeRow(doc.Paragraphs(startIdx + j).Range.Text, cols)
                        Dim c As Long
                        For c = 0 To cols - 1
                            data(j, c) = cells(c)
                        Next c
                    Next j

                    Dim blockStart As Long
                    Dim blockRange As Range
                    blockStart = doc.Paragraphs(startIdx).Range.Start
                    Set blockRange = doc.Range(Start:=blockStart, End:=doc.Paragraphs(endIdx).Range.End)
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


Private Function PipeColumnCount(ByVal s As String) As Long
    s = CleanParaText(s)
    If Len(s) < 2 Then Exit Function
    If Left(s, 1) = "|" Then s = Mid(s, 2)
    If Right(s, 1) = "|" Then s = Left(s, Len(s) - 1)
    PipeColumnCount = CountChar(s, "|") + 1
End Function


Private Function SplitPipeRow(ByVal s As String, ByVal cols As Long) As String()
    Dim parts() As String
    Dim res() As String
    Dim k As Long, n As Long

    s = CleanParaText(s)
    If Len(s) >= 2 Then
        If Left(s, 1) = "|" Then s = Mid(s, 2)
        If Right(s, 1) = "|" Then s = Left(s, Len(s) - 1)
    End If

    parts = Split(s, "|")

    ReDim res(0 To cols - 1)
    n = UBound(parts)
    For k = 0 To cols - 1
        If k <= n Then
            res(k) = Trim(parts(k))
        Else
            res(k) = ""
        End If
    Next k
    SplitPipeRow = res
End Function


' Builds a table with explicit row/column counts, then fills it. Passing
' both NumRows and NumColumns keeps Word from guessing wrong.
Private Sub BuildTableFromData(ByVal doc As Document, ByVal anchorStart As Long, _
                               ByRef data() As String, ByVal rows As Long, ByVal cols As Long)
    Dim r As Range
    Dim t As Table
    Dim a As Long, b As Long

    On Error GoTo Fallback

    ' Tables.Add needs a range covering at least one paragraph.
    Set r = doc.Range(Start:=anchorStart, End:=anchorStart)
    r.InsertParagraphAfter
    Set r = doc.Range(Start:=anchorStart, End:=anchorStart)

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

    Call FormatTable(t)
    Exit Sub

Fallback:
    ' The grid could not be created. Put the original text back so no
    ' content is silently destroyed.
    On Error Resume Next
    Dim restored As String
    For a = 0 To rows - 1
        For b = 0 To cols - 1
            restored = restored & data(a, b)
            If b < cols - 1 Then restored = restored & " "
        Next b
        restored = restored & vbCr
    Next a
    doc.Range(anchorStart, anchorStart).InsertAfter restored
    Err.Clear
    On Error GoTo 0
End Sub


'-----------------------------------------------------------------------
'  Tables - tab separated blocks (Excel, Sheets, ...)
'-----------------------------------------------------------------------

Private Sub ConvertTSVToTable(ByVal doc As Document, ByVal rng As Range)
    Dim i As Long, j As Long
    Dim startIdx As Long, endIdx As Long
    Dim paraText As String
    Dim maxTabs As Long, curTabs As Long
    Dim data() As String
    Dim rows As Long, cols As Long
    Dim r As Range
    Dim guard As Long

    i = 1
    Do While i <= doc.Paragraphs.Count
        guard = guard + 1
        If guard > 20000 Then Exit Do

        paraText = doc.Paragraphs(i).Range.Text
        curTabs = CountChar(paraText, vbTab)

        If curTabs >= 1 And Len(Trim(paraText)) > 1 Then
            startIdx = i
            maxTabs = curTabs
            endIdx = i
            j = i + 1
            Do While j <= doc.Paragraphs.Count
                paraText = doc.Paragraphs(j).Range.Text
                curTabs = CountChar(paraText, vbTab)
                If curTabs >= 1 And Len(Trim(paraText)) > 1 Then
                    If curTabs > maxTabs Then maxTabs = curTabs
                    endIdx = j
                    j = j + 1
                Else
                    Exit Do
                End If
            Loop

            rows = endIdx - startIdx + 1
            cols = maxTabs + 1

            If rows >= 1 And cols >= 2 Then
                ReDim data(0 To rows - 1, 0 To cols - 1)
                For j = 0 To rows - 1
                    Dim cells() As String
                    cells = SplitByTab(doc.Paragraphs(startIdx + j).Range.Text, cols)
                    Dim c As Long
                    For c = 0 To cols - 1
                        data(j, c) = cells(c)
                    Next c
                Next j

                Dim blockStart As Long
                blockStart = doc.Paragraphs(startIdx).Range.Start
                Set r = doc.Range(Start:=blockStart, End:=doc.Paragraphs(endIdx).Range.End)
                r.Delete

                BuildTableFromData doc, blockStart, data, rows, cols
                i = startIdx + 1
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
    Loop
End Sub


Private Function SplitByTab(ByVal s As String, ByVal cols As Long) As String()
    Dim parts() As String
    Dim res() As String
    Dim k As Long, n As Long

    s = CleanParaText(s)
    parts = Split(s, vbTab)

    ReDim res(0 To cols - 1)
    n = UBound(parts)
    For k = 0 To cols - 1
        If k <= n Then
            res(k) = Trim(Replace(parts(k), Chr(11), ""))
        Else
            res(k) = ""
        End If
    Next k
    SplitByTab = res
End Function


'-----------------------------------------------------------------------
'  Table formatting - locale safe
'-----------------------------------------------------------------------

Private Sub FormatTable(ByVal tbl As Table)
    Dim styled As Boolean
    styled = False

    ' "Table Grid" is English-only in most builds; fall back to direct
    ' borders when the localized name does not resolve.
    On Error Resume Next
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
    tbl.Range.ParagraphFormat.FirstLineIndent = 0

    If tbl.rows.Count >= 1 Then
        tbl.rows(1).Range.Font.Bold = True
        tbl.rows(1).Shading.BackgroundPatternColor = RGB(219, 229, 241)
        tbl.rows(1).HeadingFormat = True
    End If

    tbl.PreferredWidthType = wdPreferredWidthPercent
    tbl.PreferredWidth = 100

    Err.Clear
    On Error GoTo 0
End Sub


'-----------------------------------------------------------------------
'  Toolbar button
'-----------------------------------------------------------------------

Private Sub CreateToolbarButton(ByVal notify As Boolean)
    Dim cb As CommandBar
    Dim btn As CommandBarButton
    Dim found As Boolean

    EnsureBullet

    For Each cb In Application.CommandBars
        If cb.Name = CF_TOOLBAR Then
            Set cb = Application.CommandBars(CF_TOOLBAR)
            found = True
            Exit For
        End If
    Next cb

    If Not found Then
        Set cb = Application.CommandBars.Add(Name:=CF_TOOLBAR, Position:=msoBarTop, Temporary:=True)
    End If

    ' Remove stale buttons, then add exactly one.
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

    If notify Then
        MsgBox "The ""Format Chat Text"" button was added to the toolbar." & vbCrLf & _
               "You can use this button or the right-click menu to format text " & _
               "(or Alt+F8 then FormatChatText).", _
               vbInformation, "Chat Formatter"
    End If
End Sub


'-----------------------------------------------------------------------
'  Helpers
'-----------------------------------------------------------------------

Private Sub EnsureBullet()
    If mBullet = "" Then mBullet = ChrW(&H2022)
End Sub


Private Function CleanParaText(ByVal s As String) As String
    If Len(s) > 0 Then
        Do While Right(s, 1) = Chr(13) Or Right(s, 1) = Chr(7)
            s = Left(s, Len(s) - 1)
        Loop
    End If
    CleanParaText = Trim(s)
End Function


Private Function IsPipeRow(ByVal s As String) As Boolean
    s = CleanParaText(s)
    If Len(s) < 3 Then Exit Function
    If Left(s, 1) <> "|" Then Exit Function
    If Right(s, 1) <> "|" Then Exit Function
    IsPipeRow = (CountChar(s, "|") >= 2)
End Function


Private Function IsSeparatorLine(ByVal s As String) As Boolean
    Dim k As Long, ch As String
    Dim hasDash As Boolean

    s = CleanParaText(s)
    If Len(s) = 0 Then Exit Function
    If InStr(s, "|") = 0 Then Exit Function

    For k = 1 To Len(s)
        ch = Mid(s, k, 1)
        If ch = "-" Then
            hasDash = True
        ElseIf ch <> "|" And ch <> ":" And ch <> " " And ch <> vbTab Then
            Exit Function
        End If
    Next k
    IsSeparatorLine = hasDash
End Function


Private Function CountChar(ByVal s As String, ByVal ch As String) As Long
    Dim k As Long, cnt As Long
    If Len(ch) = 0 Then Exit Function
    For k = 1 To Len(s)
        If Mid(s, k, Len(ch)) = ch Then cnt = cnt + 1
    Next k
    CountChar = cnt
End Function


Private Function IsRtlText(ByVal s As String) As Boolean
    Dim k As Long, code As Long, rtlCount As Long, ltrCount As Long
    For k = 1 To Len(s)
        code = AscW(Mid(s, k, 1))
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


'-----------------------------------------------------------------------
'  Settings (registry, per user)
'-----------------------------------------------------------------------

Private Function GetSetting(ByVal key As String, ByVal def As String) As String
    Dim v As String
    On Error Resume Next
    v = GetSettingCore(key)
    On Error GoTo 0
    If Len(v) = 0 Then v = def
    GetSetting = v
End Function


Private Sub SetSetting(ByVal key As String, ByVal val As String)
    On Error Resume Next
    SetSettingCore key, val
    On Error GoTo 0
End Sub


Private Function GetSettingCore(ByVal key As String) As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    GetSettingCore = wsh.RegRead( _
        "HKEY_CURRENT_USER\Software\ChatFormatter\" & key)
End Function


Private Sub SetSettingCore(ByVal key As String, ByVal val As String)
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.RegWrite _
        "HKEY_CURRENT_USER\Software\ChatFormatter\" & key, val, "REG_SZ"
End Sub

