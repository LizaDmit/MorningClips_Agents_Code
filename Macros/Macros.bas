Attribute VB_Name = "Formatting"
Option Explicit

' ================= FormatMorningClips (v11 - Summary + FULL ARTICLES) =================
' Input: raw Reporter+FullArticles .docx. Output: gold visual format.
' One-shot on a FRESH file; not idempotent.
' SINGLE MODULE - all procedures below. Delete any other module holding an older copy.
'
' v11: Full Articles headlines are NUMBERED (1. 2. 3., per-section, resetting after
' FULL ARTICLES), not dash-bulleted. Any leading "- " the prompt emits is stripped
' before numbering. Byline gets a gap ABOVE it; date sits tight below the byline,
' with the normal body gap below the date. Roster pruning and long-paragraph
' repair remain Summary-only; Full Articles stays fully verbatim. Divider is
' inserted before FULL ARTICLES, not at document end, when that section exists.

' ---- Indent constants (CENTIMETERS) ----
Private Const IND_TRAD    As Single = 0.04
Private Const IND_CAT     As Single = 0.83
Private Const IND_ITEM    As Single = 1.31
Private Const IND_NONEWS  As Single = 1.3
' ---- Margins (CENTIMETERS) ----
Private Const MARGIN_TOP    As Single = 2.49
Private Const MARGIN_BOTTOM As Single = 2.54
Private Const MARGIN_LEFT   As Single = 3.17
Private Const MARGIN_RIGHT  As Single = 3.17
' ---- Spacing (points) ----
Private Const SP_TITLE    As Single = 12
Private Const SP_DATE     As Single = 18
Private Const SP_TRAD     As Single = 6
Private Const SP_SECTION  As Single = 6
Private Const SP_HEAD     As Single = 10
Private Const SP_BODY     As Single = 10
Private Const SP_META     As Single = 10   ' gap above the byline line
Private Const DIV_LEN     As Integer = 40


Sub FormatMorningClips()

    Dim doc As Document
    Set doc = ActiveDocument
    Application.ScreenUpdating = False

    On Error Resume Next
    ActiveWindow.View.ShowFieldCodes = False
    On Error GoTo 0

    With doc.PageSetup
        .TopMargin = CentimetersToPoints(MARGIN_TOP)
        .BottomMargin = CentimetersToPoints(MARGIN_BOTTOM)
        .LeftMargin = CentimetersToPoints(MARGIN_LEFT)
        .RightMargin = CentimetersToPoints(MARGIN_RIGHT)
    End With

    RemoveContentControls
    UnlinkHyperlinksToUrlText
    RemoveHorizontalRules

    Dim para As Paragraph
    For Each para In doc.Paragraphs
        On Error Resume Next
        para.Range.ListFormat.RemoveNumbers
        On Error GoTo 0
    Next para

    ReplaceAll ChrW(8203), ""
    ReplaceAll ChrW(65279), ""
    ReplaceAll ChrW(160), " "
    ReplaceAll "^l", "^p"

    TrimParagraphSpaces

    Dim pass As Long, lenBefore As Long
    For pass = 1 To 30
        lenBefore = Len(doc.Content.Text)
        ReplaceAll "^p^p", "^p"
        If Len(doc.Content.Text) >= lenBefore Then Exit For
    Next pass

    DeleteParasEqualTo "Social channels:"
    DeleteParasEqualTo "Top of Form"
    DeleteParasEqualTo "Bottom of Form"

    SplitTitleDate
    SplitBylineDate
    SplitUrlOffHeadlines
    SplitSocials
    PrependSectionNumerals
    NumberHeadlines
    DenumberLeakedRosters
    PruneAttributionRoster
    EnsureDivider

    Dim i As Long, p As Paragraph, s As String
    Dim inFullArticles As Boolean
    inFullArticles = False

    For i = 1 To doc.Paragraphs.Count
        Set p = doc.Paragraphs(i)
        s = ParaText(p)

        If s = "FULL ARTICLES" Then inFullArticles = True

        With p.Range.Font
            .Name = "Calibri": .Size = 11: .Bold = False: .Underline = wdUnderlineNone
        End With
        p.Alignment = wdAlignParagraphLeft
        p.LeftIndent = 0: p.FirstLineIndent = 0
        p.SpaceBefore = 0: p.SpaceAfter = 0
        p.TabStops.ClearAll

        p.Borders(wdBorderTop).LineStyle = wdLineStyleNone
        p.Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        p.Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        p.Borders(wdBorderRight).LineStyle = wdLineStyleNone

        If s = "DAILY MEDIA MONITORING REPORT" Then
            p.Range.Font.Bold = True
            p.Range.Font.Underline = wdUnderlineSingle
            p.Alignment = wdAlignParagraphCenter
            p.SpaceAfter = SP_TITLE
        ElseIf IsDateLine(s) Then
            p.Range.Font.Bold = True
            p.Alignment = wdAlignParagraphCenter
            p.SpaceAfter = SP_DATE
        ElseIf s = "Traditional Media" Then
            p.Range.Font.Bold = True
            p.LeftIndent = CentimetersToPoints(IND_TRAD)
            p.SpaceAfter = SP_TRAD
        ElseIf s = "FULL ARTICLES" Then
            p.Range.Font.Bold = True
            p.Alignment = wdAlignParagraphCenter
            p.LeftIndent = 0
            p.SpaceBefore = SP_SECTION
            p.SpaceAfter = SP_TRAD
        ElseIf s = "-ENDS-" Then
            p.Alignment = wdAlignParagraphCenter
            p.LeftIndent = 0
            p.SpaceBefore = SP_SECTION
        ElseIf IsSection(s) Then
            p.Range.Font.Bold = True
            p.LeftIndent = CentimetersToPoints(IND_NONEWS)
            p.FirstLineIndent = CentimetersToPoints(IND_CAT - IND_NONEWS)
            p.SpaceBefore = SP_SECTION
            p.TabStops.ClearAll
            p.TabStops.Add Position:=CentimetersToPoints(IND_NONEWS), Alignment:=wdAlignTabLeft
        ElseIf IsSocial(s) Then
            p.Range.Font.Bold = True
            p.LeftIndent = CentimetersToPoints(IND_NONEWS)
            p.FirstLineIndent = CentimetersToPoints(IND_CAT - IND_NONEWS)
            p.SpaceBefore = SP_SECTION
            p.TabStops.ClearAll
            p.TabStops.Add Position:=CentimetersToPoints(IND_NONEWS), Alignment:=wdAlignTabLeft
        ElseIf IsRosterLine(s) Then
            p.LeftIndent = CentimetersToPoints(IND_ITEM)
            p.SpaceAfter = SP_BODY
        ElseIf IsHeadline(s) Then
            p.Range.Font.Bold = True
            p.LeftIndent = CentimetersToPoints(IND_ITEM)
            p.FirstLineIndent = CentimetersToPoints(IND_CAT - IND_ITEM)
            p.SpaceBefore = SP_HEAD
            p.TabStops.ClearAll
            p.TabStops.Add Position:=CentimetersToPoints(IND_ITEM), Alignment:=wdAlignTabLeft
        ElseIf inFullArticles And IsBylineLine(s) Then
            p.LeftIndent = CentimetersToPoints(IND_ITEM)
            p.SpaceBefore = SP_META
            p.SpaceAfter = 0
        ElseIf inFullArticles And IsLongFormDate(s) Then
            p.LeftIndent = CentimetersToPoints(IND_ITEM)
            p.SpaceBefore = 0
            p.SpaceAfter = SP_BODY
        ElseIf IsURL(s) Then
            p.LeftIndent = CentimetersToPoints(IND_ITEM)
            MakeHyperlink p, s
        ElseIf IsDivider(s) Then
            p.LeftIndent = 0
            p.SpaceBefore = SP_SECTION
        ElseIf s = "No relevant news" Then
            SetCleanText p, "No relevant news"
            p.LeftIndent = CentimetersToPoints(IND_NONEWS)
            p.FirstLineIndent = 0
            p.SpaceAfter = SP_BODY
        Else
            p.LeftIndent = CentimetersToPoints(IND_ITEM)
            p.SpaceAfter = SP_BODY
        End If
    Next i

    StripColorTags
    HighlightAll

    SplitRosterLines
    NumberRosterEntries
    SplitEntryHeaders
    RepairLongParagraphs

    Application.ScreenUpdating = True
    MsgBox "FormatMorningClips done.", vbInformation

End Sub


' ---------- FULL ARTICLES boundary ----------

Private Function FullArticlesStart() As Long
    ' Paragraph index of the "FULL ARTICLES" heading; 0 if absent (summary-only doc).
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long
    For i = 1 To doc.Paragraphs.Count
        If ParaText(doc.Paragraphs(i)) = "FULL ARTICLES" Then
            FullArticlesStart = i
            Exit Function
        End If
    Next i
    FullArticlesStart = 0
End Function


' ---------- structural transforms ----------

Private Sub RemoveContentControls()
    Dim doc As Document: Set doc = ActiveDocument
    On Error Resume Next
    Do While doc.ContentControls.Count > 0
        doc.ContentControls(1).Delete False
    Loop
    Do While doc.FormFields.Count > 0
        doc.FormFields(1).Delete
    Loop
    On Error GoTo 0
End Sub


Private Sub UnlinkHyperlinksToUrlText()
    Dim doc As Document: Set doc = ActiveDocument
    Dim h As Hyperlink, addr As String, r As Range, guard As Long
    guard = 0
    Do While doc.Hyperlinks.Count > 0
        Set h = doc.Hyperlinks(1)
        addr = h.Address
        Set r = h.Range
        h.Delete
        If Len(addr) > 0 Then
            r.Text = addr
        End If
        guard = guard + 1
        If guard > 500 Then Exit Do
    Loop
End Sub


Private Sub RemoveHorizontalRules()
    ' Delete Word AutoFormat horizontal-rule shapes (o:hr rects, the grey line).
    ' These are picture shapes in runs, not paragraph borders, so the border
    ' clear loop cannot remove them. The text divider is re-inserted
    ' deterministically by EnsureDivider.
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long

    For i = doc.InlineShapes.Count To 1 Step -1
        On Error Resume Next
        doc.InlineShapes(i).Delete
        On Error GoTo 0
    Next i

    For i = doc.Shapes.Count To 1 Step -1
        On Error Resume Next
        doc.Shapes(i).Delete
        On Error GoTo 0
    Next i
End Sub

Private Function StartsWithNumber(ByVal s As String) As Boolean
    ' True if the line already begins "12. " or "12: " - don't re-number it.
    Dim j As Long
    j = 1
    Do While j <= Len(s)
        If Mid$(s, j, 1) >= "0" And Mid$(s, j, 1) <= "9" Then j = j + 1 Else Exit Do
    Loop
    If j = 1 Or j > Len(s) Then Exit Function
    If Mid$(s, j, 1) = "." Or Mid$(s, j, 1) = ":" Then StartsWithNumber = True
End Function

Private Sub NumberRosterEntries()
    ' Stat-block entry names ("Doug Leone") lose their source rank because Word
    ' autoformats the pasted "1. Doug Leone" into a list and RemoveNumbers strips
    ' it. Ranks are always sequential from 1 within an article, so regenerate them.
    ' A name line is one whose NEXT paragraph starts with "Firm:".
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long, s As String, nxt As String, n As Long
    n = 0
    For i = 1 To doc.Paragraphs.Count - 1
        s = ParaText(doc.Paragraphs(i))
        nxt = ParaText(doc.Paragraphs(i + 1))

        ' reset the counter at each article headline or section header
        If IsHeadline(s) Or IsSection(s) Or s = "FULL ARTICLES" Then n = 0

        If Left(nxt, 5) = "Firm:" And Len(s) > 0 And Len(s) < 60 Then
            If InStr(s, ":") = 0 Then
                If StartsWithNumber(s) Then
                    n = LeadingNumber(s)          ' adopt the source rank
                Else
                    n = n + 1
                    doc.Paragraphs(i).Range.InsertBefore CStr(n) & ". "
                End If
            End If
        End If
    Next i
End Sub

Private Function LeadingNumber(ByVal s As String) As Long
    Dim j As Long
    j = 1
    Do While j <= Len(s)
        If Mid$(s, j, 1) >= "0" And Mid$(s, j, 1) <= "9" Then j = j + 1 Else Exit Do
    Loop
    If j > 1 Then LeadingNumber = CLng(Left$(s, j - 1))
End Function

Private Sub SplitEntryHeaders()
    ' Roundup entry headers sit on their own source line but get joined to the
    ' body, which repeats the name: "ByteDance ByteDance became one of..."
    ' (or "3. ByteDance ByteDance became..."). Split after the header name.
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long, raw As String, w() As String
    Dim cutAfter As Long, st As Long

    For i = doc.Paragraphs.Count To 1 Step -1
        raw = doc.Paragraphs(i).Range.Text
        If Len(raw) > 0 Then
            If Right$(raw, 1) = vbCr Then raw = Left$(raw, Len(raw) - 1)
        End If
        If Len(raw) < 6 Then GoTo NextP

        w = Split(raw, " ")
        cutAfter = 0

        If UBound(w) >= 2 Then
            If w(0) = w(1) And Len(w(0)) >= 3 Then
                cutAfter = Len(w(0))
            ElseIf UBound(w) >= 3 Then
                If w(1) = w(2) And Len(w(1)) >= 3 And IsNumToken(w(0)) Then
                    cutAfter = Len(w(0)) + 1 + Len(w(1))
                End If
            End If
        End If

        If cutAfter > 0 Then
            st = doc.Paragraphs(i).Range.Start
            doc.Range(st + cutAfter, st + cutAfter + 1).Text = vbCr
        End If
NextP:
    Next i
End Sub


Private Function IsNumToken(ByVal t As String) As Boolean
    ' "3." or "3:" style entry number
    If Len(t) < 2 Then Exit Function
    If Right$(t, 1) <> "." And Right$(t, 1) <> ":" Then Exit Function
    Dim j As Long
    For j = 1 To Len(t) - 1
        If Mid$(t, j, 1) < "0" Or Mid$(t, j, 1) > "9" Then Exit Function
    Next j
    IsNumToken = True
End Function

Private Sub TrimParagraphSpaces()
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long, r As Range, guard As Long
    For i = 1 To doc.Paragraphs.Count
        guard = 0
        Do
            Set r = doc.Paragraphs(i).Range
            r.MoveEnd wdCharacter, -1
            If r.End <= r.Start Then Exit Do
            If Right(r.Text, 1) <> " " Then Exit Do
            doc.Range(r.End - 1, r.End).Delete
            guard = guard + 1
        Loop While guard <= 50
    Next i
End Sub


Private Sub DeleteParasEqualTo(target As String)
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long
    For i = doc.Paragraphs.Count To 1 Step -1
        If ParaText(doc.Paragraphs(i)) = target Then
            doc.Paragraphs(i).Range.Delete
        End If
    Next i
End Sub


Private Sub SplitTitleDate()
    ' Title and date sometimes arrive glued in one paragraph:
    ' "DAILY MEDIA MONITORING REPORT May.28.2026". Split into two paragraphs.
    Dim doc As Document: Set doc = ActiveDocument
    Dim p As Paragraph, s As String, title As String, cutPos As Long
    title = "DAILY MEDIA MONITORING REPORT"
    Set p = doc.Paragraphs(1)
    s = ParaText(p)
    If Left(s, Len(title)) = title And Len(s) > Len(title) Then
        cutPos = p.Range.Start + Len(title)
        Do While doc.Range(cutPos, cutPos + 1).Text = " "
            doc.Range(cutPos, cutPos + 1).Delete
        Loop
        doc.Range(cutPos, cutPos).InsertBefore vbCr
    End If
End Sub


Private Sub SplitUrlOffHeadlines()
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long, p As Paragraph, s As String, pos As Long
    Dim startPos As Long, cutPos As Long, guard As Long
    For i = doc.Paragraphs.Count To 1 Step -1
        Set p = doc.Paragraphs(i)
        s = p.Range.Text
        If InStr(s, " | ") > 0 And InStr(s, "http") > 1 Then
            pos = InStr(s, "http")
            startPos = p.Range.Start
            cutPos = startPos + pos - 1
            guard = 0
            Do While cutPos > startPos And doc.Range(cutPos - 1, cutPos).Text = " "
                doc.Range(cutPos - 1, cutPos).Delete
                cutPos = cutPos - 1
                guard = guard + 1
                If guard > 50 Then Exit Do
            Loop
            doc.Range(cutPos, cutPos).InsertBefore vbCr
        End If
    Next i
End Sub

Private Sub SplitSocials()
    ' Summary only. Full Articles bodies can legitimately begin with
    ' "Facebook ..." (e.g. the Meta entry), which would otherwise be
    ' mis-split into a fake social-channel bullet.
    Dim doc As Document: Set doc = ActiveDocument
    Dim labels As Variant
    labels = Array("Facebook", "X (Twitter)", "YouTube")
    Dim i As Long, j As Long, p As Paragraph, s As String
    Dim lbl As String, cutPos As Long, guard As Long, faStart As Long

    faStart = FullArticlesStart()
    If faStart = 0 Then faStart = doc.Paragraphs.Count + 1

    For i = faStart - 1 To 1 Step -1
        Set p = doc.Paragraphs(i)
        s = ParaText(p)
        For j = LBound(labels) To UBound(labels)
            lbl = labels(j)
            If Len(s) > Len(lbl) Then
                If Left(s, Len(lbl)) = lbl And Mid(s, Len(lbl) + 1, 1) = " " Then
                    p.Range.InsertBefore ChrW(8226) & vbTab
                    cutPos = p.Range.Start + 2 + Len(lbl)
                    guard = 0
                    Do While doc.Range(cutPos, cutPos + 1).Text = " "
                        doc.Range(cutPos, cutPos + 1).Delete
                        guard = guard + 1
                        If guard > 50 Then Exit Do
                    Loop
                    doc.Range(cutPos, cutPos).InsertBefore vbCr
                    Exit For
                End If
            End If
        Next j
    Next i
End Sub


Private Sub PrependSectionNumerals()
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long, p As Paragraph, s As String
    For i = 1 To doc.Paragraphs.Count
        Set p = doc.Paragraphs(i)
        s = ParaText(p)
        If s = "Neil Shen" Then
            p.Range.InsertBefore "i." & vbTab
        ElseIf s = "Company News" Then
            p.Range.InsertBefore "ii." & vbTab
        ElseIf s = "VC Industry News" Then
            p.Range.InsertBefore "iii." & vbTab
        End If
    Next i
End Sub


Private Sub NumberHeadlines()
    ' Numbers headlines in BOTH the Summary and Full Articles, per-section,
    ' resetting the counter at "FULL ARTICLES" and at every section header.
    ' A Full Articles headline may arrive dash-prefixed from the prompt
    ' ("- Headline | Outlet") - strip that dash before numbering, so the final
    ' text reads "1.<tab>Headline | Outlet" exactly like the Summary.
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long, p As Paragraph, s As String, n As Long
    n = 0
    For i = 1 To doc.Paragraphs.Count
        Set p = doc.Paragraphs(i)
        s = ParaText(p)
        If s = "FULL ARTICLES" Then
            n = 0
        ElseIf IsSection(s) Then
            n = 1
        ElseIf IsHeadline(s) And Not IsRosterLine(s) Then
            If Left(s, 2) = "- " Then
                doc.Range(p.Range.Start, p.Range.Start + 2).Delete
            End If
            If n = 0 Then n = 1
            p.Range.InsertBefore CStr(n) & "." & vbTab
            n = n + 1
        End If
    Next i
End Sub


Sub DenumberLeakedRosters()
    ' Strip the "N." + tab/space that numbering puts on a roster line
    ' mis-classified as a headline. Pure backstop - only strips a leading
    ' numeric prefix + delimiter on a line matching roster markers, so it is
    ' safe to run over the whole document (Summary or Full Articles).
    Dim doc As Document, para As Paragraph, rng As Range
    Dim t As String, p As Long, j As Long
    Set doc = ActiveDocument

    For p = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(p)
        Set rng = para.Range
        t = rng.Text
        If Len(t) > 0 Then If Right$(t, 1) = vbCr Then t = Left$(t, Len(t) - 1)

        j = 1
        Do While j <= Len(t)
            If Mid$(t, j, 1) >= "0" And Mid$(t, j, 1) <= "9" Then j = j + 1 Else Exit Do
        Loop
        If j > 1 And j <= Len(t) - 1 Then
            If Mid$(t, j, 1) = "." And (Mid$(t, j + 1, 1) = " " Or Mid$(t, j + 1, 1) = vbTab) Then
                If InStr(t, "| #") > 0 Or InStr(t, "Where it counted most") > 0 Then
                    doc.Range(rng.Start, rng.Start + j + 1).Delete
                End If
            End If
        End If
    Next p
End Sub


Private Sub PruneAttributionRoster()
    ' Deterministic roster cleanup - SUMMARY ONLY. On a "Where it counted most... | #"
    ' line, keep the header + only the "Name | Firm | #rank" units naming a tracked
    ' keyword; drop the rest. Full Articles bodies keep rosters verbatim (gold).
    ' Runs BEFORE HighlightAll so the kept line still gets highlighted.
    Dim doc As Document: Set doc = ActiveDocument
    Dim para As Paragraph
    Dim t As String, p As Long, lastP As Long
    Dim tracked As Variant, k As Long
    Dim header As String, body As String, cutAt As Long
    Dim i As Long, ch As String, d As String
    Dim curUnit As String, kept As String

    tracked = Array("Neil Shen", "HongShan", "Hongshan", "HSG", "IDG", "ZhenFund", "Hillhouse", "Granite Asia", "Matrix Partners China", "Qiming Venture Partners", "KKR", "EQT", "TPG", "Walden International", "Carlyle")

    lastP = FullArticlesStart()
    If lastP = 0 Then lastP = doc.Paragraphs.Count + 1

    For p = lastP - 1 To 1 Step -1
        Set para = doc.Paragraphs(p)
        t = para.Range.Text
        If Len(t) > 0 Then If Right$(t, 1) = vbCr Then t = Left$(t, Len(t) - 1)

        If InStr(t, "| #") > 0 And InStr(t, "Where it counted most") > 0 Then
            cutAt = InStr(t, ":*")
            If cutAt > 0 Then
                header = Left$(t, cutAt + 1)
                body = Mid$(t, cutAt + 2)
            Else
                cutAt = InStr(t, ":")
                If cutAt > 0 Then
                    header = Left$(t, cutAt)
                    body = Mid$(t, cutAt + 1)
                Else
                    header = ""
                    body = t
                End If
            End If

            kept = ""
            curUnit = ""
            i = 1
            Do While i <= Len(body)
                ch = Mid$(body, i, 1)
                curUnit = curUnit & ch
                If ch = "#" Then
                    Do While i + 1 <= Len(body)
                        d = Mid$(body, i + 1, 1)
                        If d >= "0" And d <= "9" Then
                            curUnit = curUnit & d
                            i = i + 1
                        Else
                            Exit Do
                        End If
                    Loop
                    For k = LBound(tracked) To UBound(tracked)
                        If InStr(1, curUnit, tracked(k), vbTextCompare) > 0 Then
                            If Len(kept) > 0 Then kept = kept & " "
                            kept = kept & Trim(curUnit)
                            Exit For
                        End If
                    Next k
                    curUnit = ""
                End If
                i = i + 1
            Loop

            para.Range.Text = Trim(header & " " & kept) & vbCr
        End If
    Next p
End Sub


Private Sub EnsureDivider()
    ' The Reporter's own em-dash divider gets converted to an HR shape by Word
    ' AutoFormat and deleted by RemoveHorizontalRules. Re-insert it as plain text.
    ' Combined doc: divider goes immediately BEFORE the FULL ARTICLES heading
    ' (closing the Summary); the document ends with -ENDS-, no trailing divider.
    ' Summary-only doc: divider at document end, as before.
    Dim doc As Document: Set doc = ActiveDocument
    Dim faStart As Long
    Dim lastP As Paragraph

    faStart = FullArticlesStart()

    If faStart > 0 Then
        If faStart > 1 Then
            If IsDivider(ParaText(doc.Paragraphs(faStart - 1))) Then Exit Sub
        End If
        doc.Paragraphs(faStart).Range.InsertBefore String(DIV_LEN, ChrW(8212)) & vbCr
    Else
        Set lastP = doc.Paragraphs(doc.Paragraphs.Count)
        If IsDivider(ParaText(lastP)) Then Exit Sub
        If ParaText(lastP) = "" Then
            lastP.Range.InsertBefore String(DIV_LEN, ChrW(8212))
        Else
            doc.Content.InsertAfter vbCr & String(DIV_LEN, ChrW(8212))
        End If
    End If
End Sub


Sub SplitRosterLines()
    ' Deterministic roster re-splitting for flattened Midas-style stat-blocks.
    ' Wrap MUST be wdFindStop: the replacement text contains the search text,
    ' so wdFindContinue would wrap and re-match forever (infinite loop).
    ' MatchWildcards is set explicitly in every block - Word's Find object
    ' keeps that flag between calls and a leaked True breaks literal finds.
    Dim doc As Document
    Dim labels As Variant
    Dim i As Long
    Dim rng As Range

    Set doc = ActiveDocument

    ' 1) Break before a numbered roster entry header ("12. Neil Shen") glued to
    '    the tail of the previous line.
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = " ([0-9]{1,2}). ([A-Z])"
        .Replacement.Text = "^p\1. \2"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With

    ' 2) Break before each stat-block label. Bare label catches both the spaced
    '    form and the glued "NameFirm:" form.
    labels = Array("Firm:", "Total appearances:", "Average rank:", "Best rank:", "Net Worth:", "First year on the list:", "Last year on the list:", "Biggest deals:")

    For i = LBound(labels) To UBound(labels)
        Set rng = doc.Content
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = CStr(labels(i))
            .Replacement.Text = "^p" & CStr(labels(i))
            .Forward = True
            .Wrap = wdFindStop
            .MatchCase = True
            .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
    Next i

    ' 3) collapse any empty paragraphs the splits created
    ReplaceAll "^p^p", "^p"
End Sub


Sub RepairLongParagraphs()
    ' Two-threshold split. Summary: 700 (tuned on summary paragraphs).
    ' Full Articles: 1400 - high enough to leave genuine long verbatim
    ' paragraphs whole (a 700-900 char quote stays one paragraph) while still
    ' breaking up 2000+ char blobs where the Extractor lost every break.
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    Dim s As String
    Dim i As Long, p As Long, faIndex As Long
    Dim splitPos() As Long
    Dim nSplits As Long
    Dim paraStart As Long
    Dim maxLen As Long
    Const MAXLEN_SUMMARY As Long = 700
    Const MAXLEN_FULL    As Long = 1400

    Set doc = ActiveDocument

    faIndex = FullArticlesStart()
    If faIndex = 0 Then faIndex = doc.Paragraphs.Count + 1

    For p = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(p)
        Set rng = para.Range
        s = rng.Text
        If Len(s) > 0 Then
            If Right$(s, 1) = vbCr Then s = Left$(s, Len(s) - 1)
        End If

        If p >= faIndex Then maxLen = MAXLEN_FULL Else maxLen = MAXLEN_SUMMARY

        If Len(s) <= maxLen Then GoTo NextPara
        If InStr(s, vbTab) > 0 Then GoTo NextPara

        paraStart = rng.Start

        nSplits = 0
        ReDim splitPos(1 To Len(s))
        For i = 1 To Len(s) - 2
            If (Mid$(s, i, 1) = "." Or Mid$(s, i, 1) = "!" Or Mid$(s, i, 1) = "?") _
               And Mid$(s, i + 1, 1) = " " _
               And IsUpperLetter(Mid$(s, i + 2, 1)) Then
                If Not EndsWithAbbrev(Left$(s, i)) Then
                    nSplits = nSplits + 1
                    splitPos(nSplits) = i + 1
                End If
            End If
        Next i

        For i = nSplits To 1 Step -1
            Dim sp As Range
            Set sp = doc.Range(paraStart + splitPos(i) - 1, paraStart + splitPos(i))
            sp.Text = vbCr
        Next i
NextPara:
    Next p
End Sub


Private Sub StripColorTags()
    ' Backstop: delete literal color-tag markup if the Reporter/FullArticles
    ' prompt leaks it (e.g. "[TURQUOISE]Hillhouse[/TURQUOISE]").
    Dim tags As Variant, i As Long
    tags = Array("[GREEN]", "[/GREEN]", "[YELLOW]", "[/YELLOW]", "[TURQUOISE]", "[/TURQUOISE]")
    For i = LBound(tags) To UBound(tags)
        ReplaceAll CStr(tags(i)), ""
    Next i
End Sub


' ---------- helpers ----------

Private Sub ReplaceAll(findText As String, replText As String)
    With ActiveDocument.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replText
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub


Private Function ParaText(p As Paragraph) As String
    Dim t As String
    t = p.Range.Text
    Do While Len(t) > 0 And (Right(t, 1) = vbCr Or Right(t, 1) = Chr(7) Or Right(t, 1) = vbLf Or Right(t, 1) = Chr(11))
        t = Left(t, Len(t) - 1)
    Loop
    ParaText = Trim(t)
End Function


Private Function IsDateLine(s As String) As Boolean
    Dim a() As String
    a = Split(s, ".")
    If UBound(a) = 2 Then
        If IsNumeric(a(1)) And IsNumeric(a(2)) Then
            If Len(Trim(a(2))) = 4 And InStr(s, " ") = 0 Then IsDateLine = True
        End If
    End If
End Function


Private Function IsSection(s As String) As Boolean
    IsSection = (s Like "i." & vbTab & "*") Or (s Like "ii." & vbTab & "*") Or (s Like "iii." & vbTab & "*") Or (s Like "i. *") Or (s Like "ii. *") Or (s Like "iii. *")
End Function


Private Function IsSocial(s As String) As Boolean
    Dim bt As String, bs As String
    bt = ChrW(8226) & vbTab
    bs = ChrW(8226) & " "
    IsSocial = (s = bt & "Facebook") Or (s = bt & "X (Twitter)") Or (s = bt & "YouTube") Or (s = bs & "Facebook") Or (s = bs & "X (Twitter)") Or (s = bs & "YouTube")
End Function


Private Function IsRosterLine(ByVal s As String) As Boolean
    ' Attribution roster line: contains " | #" or the roster header text.
    ' Checked BEFORE IsHeadline so rosters get body styling, not bold headline styling.
    IsRosterLine = (InStr(s, "| #") > 0) Or (InStr(s, "Where it counted most") > 0)
End Function


Private Function IsHeadline(s As String) As Boolean
    ' Tolerates a leading "- " (Full Articles dash form) so both NumberHeadlines
    ' and the styling loop correctly recognise the line either before or after
    ' the dash has been stripped and replaced with a number.
    Dim t As String
    t = s
    If Left(t, 2) = "- " Then t = Mid(t, 3)
    IsHeadline = (InStr(t, " | ") > 0) And Not IsSection(t) And Not IsSocial(t)
End Function


Private Function IsBylineLine(ByVal s As String) As Boolean
    ' A Full Articles byline line: starts with "By " and is short, not a headline.
    IsBylineLine = (Left(s, 3) = "By ") And (Len(s) < 120) And (InStr(s, " | ") = 0)
End Function


Private Function IsLongFormDate(ByVal s As String) As Boolean
    ' A Full Articles date line: "May 20, 2026" style. Short, month-first, year-last.
    Dim months As Variant, i As Long
    If Len(s) > 20 Or InStr(s, ",") = 0 Then Exit Function
    months = Array("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    For i = LBound(months) To UBound(months)
        If Left(s, Len(months(i))) = CStr(months(i)) Then
            If IsNumeric(Right(s, 4)) Then IsLongFormDate = True
            Exit Function
        End If
    Next i
End Function


Private Function IsURL(s As String) As Boolean
    IsURL = (LCase(Left(s, 4)) = "http")
End Function


Private Function IsDivider(s As String) As Boolean
    If Len(s) >= 3 Then
        If Left(s, 3) = "---" Then IsDivider = True
        If Left(s, 1) = ChrW(8212) Then IsDivider = True
    End If
End Function


Private Function IsUpperLetter(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsUpperLetter = (ch >= "A" And ch <= "Z")
End Function


Private Function EndsWithAbbrev(ByVal leftPart As String) As Boolean
    Dim k As Long, ch As String, tok As String
    k = Len(leftPart) - 1
    tok = ""
    Do While k >= 1
        ch = Mid$(leftPart, k, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or ch = "." Then
            tok = ch & tok
            k = k - 1
        Else
            Exit Do
        End If
    Loop

    Dim abbr As Variant
    abbr = Array("Ltd", "Inc", "Co", "Corp", "U.S", "U.K", "A.N.C", "Mr", "Mrs", "Ms", "Dr", "St", "vs", "etc", "No", "Jr", "Sr", "Ph.D")
    Dim a As Long
    For a = LBound(abbr) To UBound(abbr)
        If StrComp(tok, CStr(abbr(a)), vbTextCompare) = 0 Then
            EndsWithAbbrev = True
            Exit Function
        End If
    Next a

    If Len(tok) = 1 And tok >= "A" And tok <= "Z" Then EndsWithAbbrev = True
End Function


Private Sub SetCleanText(p As Paragraph, newText As String)
    Dim r As Range
    Set r = p.Range
    r.MoveEnd wdCharacter, -1
    r.Text = newText
End Sub


Private Sub MakeHyperlink(p As Paragraph, url As String)
    Dim r As Range, h As Hyperlink
    Set r = p.Range
    r.MoveEnd wdCharacter, -1
    If r.Hyperlinks.Count > 0 Then Exit Sub
    On Error Resume Next
    Set h = ActiveDocument.Hyperlinks.Add(Anchor:=r, Address:=url, TextToDisplay:=url)
    If Not h Is Nothing Then
        h.Range.Font.Color = RGB(5, 99, 193)
        h.Range.Font.Underline = wdUnderlineSingle
    End If
    On Error GoTo 0
End Sub


' ---------- highlighting ----------

Private Sub HighlightAll()
    ' Runs over the WHOLE document - gold highlights keywords in Full Articles too.
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long, p As Paragraph, s As String
    Dim green() As String, yellow() As String, turq() As String
    green = Split("Neil Shen", "|")
    yellow = Split("HongShan|Hongshan|HSG", "|")
    turq = Split("IDG Capital|IDG|ZhenFund|Hillhouse|Granite Asia|Matrix Partners China|Qiming Venture Partners|KKR|EQT|TPG|Walden International|Carlyle", "|")

    For i = 1 To doc.Paragraphs.Count
        Set p = doc.Paragraphs(i)
        s = ParaText(p)
        If Not IsURL(s) And Not IsDivider(s) Then
            HiliteList p.Range, green, wdBrightGreen
            HiliteList p.Range, yellow, wdYellow
            HiliteList p.Range, turq, wdTurquoise
        End If
    Next i
End Sub


Private Sub HiliteList(rng As Range, terms() As String, clr As Long)
    Dim k As Long
    For k = LBound(terms) To UBound(terms)
        HiliteTerm rng, terms(k), clr
    Next k
End Sub


Private Sub HiliteTerm(rng As Range, term As String, clr As Long)
    Dim f As Range, lastPos As Long
    Set f = rng.Duplicate
    lastPos = -1
    With f.Find
        .ClearFormatting
        .Text = term
        .MatchCase = True
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindStop
    End With
    Do While f.Find.Execute
        If f.End > rng.End Then Exit Do
        If f.Font.Bold = False Then f.HighlightColorIndex = clr
        If f.End <= lastPos Then Exit Do
        lastPos = f.End
        f.Start = f.End
        f.End = rng.End
    Loop
End Sub
Private Sub SplitBylineDate()
    ' Byline and date arrive merged ("By Mark Sparrow May 19, 2026").
    ' Split before the trailing "Month D, YYYY" so each gets its own line.
    Dim doc As Document: Set doc = ActiveDocument
    Dim rng As Range
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "(By [!^13]@) ([JFMASOND][a-z]{2,8} [0-9]{1,2}, [0-9]{4})"
        .Replacement.Text = "\1^p\2"
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub
