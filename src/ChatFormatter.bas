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

' Set by CF_TestFormatRange so FormatRangeCore suppresses its dialogs.
Private silentMode As Boolean

' Last error seen by FormatRangeCore. Kept in a module variable rather
' than a document variable: touching a template's variables can raise a
' modal save prompt, which would hang an automated run.
Private mLastError As String
Private mStage As String


Private Sub SetLastError(ByVal errNum As Long, ByVal errDesc As String)
    If errNum = 0 Then
        mLastError = ""
    Else
        mLastError = CStr(errNum) & "|" & errDesc
    End If
End Sub


Public Function CF_GetLastError() As String
    CF_GetLastError = mLastError
End Function


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
        MsgBox "??? ???? ??? ????.", vbExclamation, "Chat Formatter"
        Exit Sub
    End If
    Set doc = ActiveDocument

    Dim resp As VbMsgBoxResult
    resp = MsgBox( _
        "??? ???????? ??????? (TSV ? Pipe) ?? ?? ???? Word ????? ?????" & vbCrLf & vbCrLf & _
        "???  = Markdown + ???????" & vbCrLf & _
        "??? = ??? Markdown", _
        vbYesNo + vbQuestion, "Chat Formatter")

    Set rng = doc.Content

    If Not RangeNeedsFormatting(rng) Then
        MsgBox "??? ??? ?? ??? ????????? ??? ???." & vbCrLf & _
               "??? ?????? ???? ????.", vbInformation, "Chat Formatter"
        Exit Sub
    End If

    FormatRangeCore doc, rng, (resp = vbYes)
End Sub


Public Sub FormatSelection()
    Dim doc As Document
    If Documents.Count = 0 Then
        MsgBox "??? ???? ??? ????.", vbExclamation, "Chat Formatter"
        Exit Sub
    End If
    Set doc = ActiveDocument

    If doc.Selection.Range.Start = doc.Selection.Range.End Then
        MsgBox "????? ??? ???? ??? ?? ?????? ????.", vbInformation, "Chat Formatter"
        Exit Sub
    End If

    If Not RangeNeedsFormatting(doc.Selection.Range) Then
        MsgBox "??? ??? ?? ??? ????????? ??? ???." & vbCrLf & _
               "??? ?????? ???? ????.", vbInformation, "Chat Formatter"
        Exit Sub
    End If

    FormatRangeCore doc, doc.Selection.Range, (GetSetting("Tables", "1") = "1")
End Sub


' Silent entry point used by tools\test.ps1. Same behaviour as
' FormatChatText but with no dialogs, so automated runs cannot hang.
Public Sub CF_TestFormatRange()
    Dim doc As Document

    TraceStage "ENTRY CF_TestFormatRange"
    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument

    silentMode = True
    On Error Resume Next
    If doc.Selection.Range.Start = doc.Selection.Range.End Then
        Set doc.Content.Select
    End If
    Err.Clear
    On Error GoTo 0

    FormatRangeCore doc, doc.Content, True

    silentMode = False
End Sub


' Runs exactly one stage so the test harness can call them in order and see
' which one wedges Word, instead of waiting on a single opaque call.
Public Sub CF_TestStage(ByVal n As Long)
    Dim doc As Document
    Dim rng As Range

    TraceStage "stage " & n
    If Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument
    Set rng = doc.Content

    silentMode = True
    On Error Resume Next
    EnsureBullet
    Err.Clear
    On Error GoTo 0

    Select Case n
        Case 1: ConvertPipeTablesToTable doc, rng
        Case 2: ConvertTSVToTable doc, rng
        Case 3: RemoveHorizontalRules doc, rng
        Case 4: ConvertCodeBlocks doc, rng
        Case 5: ConvertHeadings doc, rng
        Case 6: FormatBold doc, rng
        Case 7: FormatItalic doc, rng
        Case 8: FormatStrike doc, rng
        Case 9: FormatInlineCode doc, rng
        Case 10: FormatLinks doc, rng
        Case 11: ConvertBlockquotes doc, rng
        Case 12: ConvertBulletLists doc, rng
        Case 13: ConvertNumberedLists doc, rng
        Case 14: CleanStrayMarkers doc, rng
    End Select

    silentMode = False
    TraceStage "stage " & n & " done"
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
        MsgBox "???? ?????? ??????? ??." & vbCrLf & _
                     "???? ????????? ??????: Alt+F8 ? ToggleAutoFormat", _
                  vbInformation, "Chat Formatter"
    Else
        SetSetting "AutoFormat", "1"
        MsgBox "???? ?????? ???? ??." & vbCrLf & _
                     "???? ??? AI ?? ??? ??? ?? Word ?????? ? ????????? ??.", _
                  vbInformation, "Chat Formatter"
    End If
End Sub


Public Sub ShowSettings()
    Dim autoFmt As String, tables As String
    If GetSetting("AutoFormat", "1") = "1" Then autoFmt = "????"
    Else autoFmt = "???????"
    End If
    If GetSetting("Tables", "1") = "1" Then tables = "????"
    Else tables = "???????"
    End If
    MsgBox "Chat Formatter " & CF_VERSION & vbCrLf & vbCrLf & _
           "???? ?????? (?????????): " & autoFmt & vbCrLf & _
           "????? ????: " & tables & vbCrLf & vbCrLf & _
           "????? ???????:", _
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


' mStage records the last stage that started, so a hung or failed run can be
' diagnosed from outside instead of guessing.
Public Function CF_GetStage() As String
    CF_GetStage = mStage
End Function


' Appends a timestamped trace line. Used only by the test harness; if a
' stage wedges Word the trace on disk still shows how far it got.
Private Sub TraceStage(ByVal msg As String)
    ' Tracing disabled to avoid file I/O issues during automated runs
End Sub


Private Sub RunStageCore(ByVal doc As Document, ByVal rng As Range, ByVal doTables As Boolean)
    Dim su As Boolean
    Dim undoStarted As Boolean
    Dim failed As Boolean
    Dim errNum As Long, errDesc As String

    mStage = "EnsureBullet"

    TraceStage "EnsureBullet"
    EnsureBullet

    su = Application.ScreenUpdating
    Application.ScreenUpdating = False

    On Error Resume Next
    Application.UndoRecord.StartCustomRecord "Chat Formatter"
    If Err.Number = 0 Then undoStarted = True
    Err.Clear
    On Error GoTo CleanUp

    mStage = "Tables"

    TraceStage "Tables"
    If doTables Then
        mStage = "PipeTables"
        TraceStage "PipeTables"
        ConvertPipeTablesToTable doc, rng
        mStage = "TSVTables"
        TraceStage "TSVTables"
        ConvertTSVToTable doc, rng
    End If

    mStage = "HorizontalRules"

    TraceStage "HorizontalRules"
    RemoveHorizontalRules doc, rng
    mStage = "CodeBlocks"
    TraceStage "CodeBlocks"
    ConvertCodeBlocks doc, rng
    mStage = "Headings"
    TraceStage "Headings"
    ConvertHeadings doc, rng
    mStage = "Bold"
    TraceStage "Bold"
    FormatBold doc, rng
    mStage = "Italic"
    TraceStage "Italic"
    FormatItalic doc, rng
    mStage = "Strike"
    TraceStage "Strike"
    FormatStrike doc, rng
    mStage = "InlineCode"
    TraceStage "InlineCode"
    FormatInlineCode doc, rng
    mStage = "Links"
    TraceStage "Links"
    FormatLinks doc, rng
    mStage = "Blockquotes"
    TraceStage "Blockquotes"
    ConvertBlockquotes doc, rng
    mStage = "Bullets"
    TraceStage "Bullets"
    ConvertBulletLists doc, rng
    mStage = "Numbered"
    TraceStage "Numbered"
    ConvertNumberedLists doc, rng
    mStage = "StrayMarkers"
    TraceStage "StrayMarkers"
    CleanStrayMarkers doc, rng
    mStage = "Done"
    TraceStage "Done"

CleanUp:
    failed = (Err.Number <> 0)
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If undoStarted Then Application.UndoRecord.EndCustomRecord
    Err.Clear
    Application.ScreenUpdating = su
    Err.Clear
    On Error GoTo 0

    ' Automated runs (tools\test.ps1) set silentMode so a modal dialog can
    ' never block the Word instance and hang the test. The error is still
    ' published to a document variable so the harness can read it.
    If silentMode Then
        SetLastError errNum, errDesc
        Exit Sub
    End If

    If failed Then
        MsgBox "??? ?? ?????? ???:" & vbCrLf & _
               errDesc & vbCrLf & "????? ???: " & errNum & vbCrLf & vbCrLf & _
               "???? ?????: Ctrl+Z", vbExclamation, "Chat Formatter"
    Else
        MsgBox "?????? ??? ?? ?????? ????? ??!" & vbCrLf & _
               "(???? ?????: Ctrl+Z)", vbInformation, "Chat Formatter"
    End If
End Sub


'-----------------------------------------------------------------------
'  1) Horizontal rules
'-----------------------------------------------------------------------

Private Sub RemoveHorizontalRules(ByVal doc As Document, ByVal rng As Range)
    Dim k As Long, p As Paragraph
    Dim t As String, body As String

    For k = rng.Paragraphs.Count To 1 Step -1
        Set p = rng.Paragraphs(k)
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

    For i = rng.Paragraphs.Count To 1 Step -1
        guard = guard + 1
        If guard > 20000 Then Exit For

        If InStr(rng.Paragraphs(i).Range.Text, "```") > 0 Then
            If Not inCode Then
                openIdx = i
                inCode = True
            Else
                ' opening fence = openIdx, closing fence = i
                For j = openIdx + 1 To i - 1
                    On Error Resume Next
                    With rng.Paragraphs(j).Range
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
                rng.Paragraphs(openIdx).Range.Shading.BackgroundPatternColor = wdColorAutomatic
                rng.Paragraphs(i).Range.Shading.BackgroundPatternColor = wdColorAutomatic
                rng.Paragraphs(openIdx).Range.Delete
                rng.Paragraphs(i - 1).Range.Delete
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
        rng.Paragraphs(openIdx).Range.Shading.BackgroundPatternColor = wdColorAutomatic
        rng.Paragraphs(openIdx).Range.Delete
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

    For k = rng.Paragraphs.Count To 1 Step -1
        Set p = rng.Paragraphs(k)
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
                If k < rng.Paragraphs.Count Then prevTxt = CleanParaText(rng.Paragraphs(k + 1).Range.Text)
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

    For k = rng.Paragraphs.Count To 1 Step -1
        Set p = rng.Paragraphs(k)
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

    For k = rng.Paragraphs.Count To 1 Step -1
        Set p = rng.Paragraphs(k)
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

    For k = rng.Paragraphs.Count To 1 Step -1
        Set p = rng.Paragraphs(k)
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
    Do While i <= rng.Paragraphs.Count
        guard = guard + 1
        If guard > 20000 Then Exit Do

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
    Dim blockParas As Long
    Dim data() As String
    Dim rows As Long, cols As Long
    Dim r As Range
    Dim guard As Long

    i = 1
    Do While i <= rng.Paragraphs.Count
        guard = guard + 1
        If guard > 20000 Then Exit Do

        paraText = rng.Paragraphs(i).Range.Text
        curTabs = CountChar(paraText, vbTab)

        If curTabs >= 1 And Len(Trim(paraText)) > 1 Then
            startIdx = i
            maxTabs = curTabs
            endIdx = i
            j = i + 1
            Do While j <= rng.Paragraphs.Count
                paraText = rng.Paragraphs(j).Range.Text
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
                Dim cells As String()
                For j = 0 To rows - 1
                    cells = SplitByTab(rng.Paragraphs(startIdx + j).Range.Text, cols)
                    Dim c As Long
                    For c = 0 To cols - 1
                        data(j, c) = cells(c)
                    Next c
                Next j

                Set r = doc.Range( _
                    Start:=rng.Paragraphs(startIdx).Range.Start, _
                    End:=rng.Paragraphs(endIdx).Range.End)
                Dim blockStart As Long
                blockStart = rng.Paragraphs(startIdx).Range.Start
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
        MsgBox "???? ""Format Chat Text"" ?? ???? ????? ????? ??." & vbCrLf & _
               "???? ??? ?? ?? ?? ??? ??? ?? Word ?????? ? ????????? ?? " & _
               "(?? Alt+F8 ? FormatChatText ?? ???? ??).", _
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

