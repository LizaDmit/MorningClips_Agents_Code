Attribute VB_Name = "Module4"
Option Explicit

' ================= FormatMorningClips (v9 - Windows, CM indents + margins) =================
' Input: raw Reporter .docx. Output: Report_20 visual format.
' One-shot on a FRESH file; not idempotent.
' SINGLE MODULE - all procedures below. Delete any other module holding an older copy.

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

    SplitUrlOffHeadlines
    SplitTitleDate
    SplitSocials
    PrependSectionNumerals
    NumberHeadlines
    DenumberLeakedRosters
    PruneAttributionRoster
    EnsureDivider

    Dim i As Long, p As Paragraph, s As String
    For i = 1 To doc.Paragraphs.Count
        Set p = doc.Paragraphs(i)
        s = ParaText(p)

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

    HighlightAll

    SplitRosterLines
    RepairLongParagraphs

    Application.ScreenUpdating = True
    MsgBox "FormatMorningClips done.", vbInformation

End Sub
Private Sub PruneAttributionRoster()
    ' Deterministic roster cleanup. On a "Where it counted most... | #" line,
    ' keep the header + only the "Name | Firm | #rank" units naming a tracked
    ' keyword; drop the rest. Preserves each kept unit verbatim, rank included.
    ' Runs BEFORE HighlightAll so the kept line still gets highlighted.
    Dim doc As Document: Set doc = ActiveDocument
    Dim para As Paragraph
    Dim t As String, p As Long
    Dim tracked As Variant, k As Long
    Dim header As String, body As String, cutAt As Long
    Dim i As Long, ch As String, d As String
    Dim curUnit As String, kept As String

    tracked = Array("Neil Shen", "HongShan", "HSG", "Hongshan", "IDG", "ZhenFund", "Hillhouse", "Granite Asia", "Matrix Partners China", "Qiming Venture Partners", "KKR", "EQT", "TPG", "Walden International", "Carlyle")

    For p = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(p)
        t = para.Range.Text
        If Len(t) > 0 Then If Right$(t, 1) = vbCr Then t = Left$(t, Len(t) - 1)

        If InStr(t, "| #") > 0 And InStr(t, "Where it counted most") > 0 Then
            ' split header (up to ":*" or ":") from roster body
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

            ' walk body; each unit ends at "#<digits>"; keep units with a tracked name
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

Sub SplitRosterLines()
    ' Deterministic roster re-splitting for flattened Midas-style stat-blocks.
    ' Inserts a paragraph mark before each roster label when it follows a space.
    Dim doc As Document
    Dim labels As Variant
    Dim i As Long
    Dim rng As Range

    Set doc = ActiveDocument
    labels = Array("Firm:", "Total appearances:", "Average rank:", "Best rank:", "Net Worth:", "First year on the list:", "Last year on the list:", "Biggest deals:")

    For i = LBound(labels) To UBound(labels)
        Set rng = doc.Content
        With rng.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = " " & CStr(labels(i))
            .Replacement.Text = "^p" & CStr(labels(i))
            .Forward = True
            .Wrap = wdFindContinue
            .MatchCase = True
            .Execute Replace:=wdReplaceAll
        End With
    Next i
End Sub


Sub DenumberLeakedRosters()
    ' Strip the "N." + tab/space that numbering puts on a roster line
    ' mis-classified as a headline. Runs after NumberHeadlines. Pure backstop.
    Dim doc As Document, para As Paragraph, rng As Range
    Dim t As String, p As Long, j As Long
    Set doc = ActiveDocument

    For p = 1 To doc.Paragraphs.Count
        Set para = doc.Paragraphs(p)
        Set rng = para.Range
        t = rng.Text
        If Len(t) > 0 Then If Right$(t, 1) = vbCr Then t = Left$(t, Len(t) - 1)

        ' leading digits then "." then space OR tab
        j = 1
        Do While j <= Len(t)
            If Mid$(t, j, 1) >= "0" And Mid$(t, j, 1) <= "9" Then j = j + 1 Else Exit Do
        Loop
        If j > 1 And j <= Len(t) - 1 Then
            If Mid$(t, j, 1) = "." And (Mid$(t, j + 1, 1) = " " Or Mid$(t, j + 1, 1) = vbTab) Then
                ' roster markers a real headline never carries
                If InStr(t, "| #") > 0 Or InStr(t, "Where it counted most") > 0 Then
                    doc.Range(rng.Start, rng.Start + j + 1).Delete   ' digits + "." + one delimiter
                End If
            End If
        End If
    Next p
End Sub


Private Function IsRosterLine(ByVal s As String) As Boolean
    ' Attribution roster line: contains " | #" or the roster header text.
    ' Checked BEFORE IsHeadline so rosters get body styling, not bold headline styling.
    IsRosterLine = (InStr(s, "| #") > 0) Or (InStr(s, "Where it counted most") > 0)
End Function


Sub RepairLongParagraphs()
    ' Deterministic body-paragraph re-splitting.
    ' Run LAST, after highlighting. Inserts paragraph marks only; preserves char formatting.
    Dim doc As Document
    Dim para As Paragraph
    Dim rng As Range
    Dim s As String
    Dim i As Long, p As Long
    Dim splitPos() As Long
    Dim nSplits As Long
    Dim paraStart As Long
    Const MAXLEN As Long = 700

    Set doc = ActiveDocument

    For p = doc.Paragraphs.Count To 1 Step -1
        Set para = doc.Paragraphs(p)
        Set rng = para.Range
        s = rng.Text
        If Len(s) > 0 Then
            If Right$(s, 1) = vbCr Then s = Left$(s, Len(s) - 1)
        End If

        ' guards
        If Len(s) <= MAXLEN Then GoTo NextPara
        If InStr(s, vbTab) > 0 Then GoTo NextPara   ' likely a pasted table row - leave alone

        paraStart = rng.Start

        ' collect inter-sentence space positions (1-based into s)
        nSplits = 0
        ReDim splitPos(1 To Len(s))
        For i = 1 To Len(s) - 2
            If (Mid$(s, i, 1) = "." Or Mid$(s, i, 1) = "!" Or Mid$(s, i, 1) = "?") _
               And Mid$(s, i + 1, 1) = " " _
               And IsUpperLetter(Mid$(s, i + 2, 1)) Then
                If Not EndsWithAbbrev(Left$(s, i)) Then
                    nSplits = nSplits + 1
                    splitPos(nSplits) = i + 1   ' the space char
                End If
            End If
        Next i

        ' apply back-to-front (length-neutral, but safe)
        For i = nSplits To 1 Step -1
            Dim sp As Range
            Set sp = doc.Range(paraStart + splitPos(i) - 1, paraStart + splitPos(i))
            sp.Text = vbCr
        Next i
NextPara:
    Next p
End Sub


Private Function IsUpperLetter(ByVal ch As String) As Boolean
    If Len(ch) = 0 Then Exit Function
    IsUpperLetter = (ch >= "A" And ch <= "Z")
End Function

Private Sub SplitTitleDate()
    ' Title and date sometimes arrive glued in one paragraph:
    ' "DAILY MEDIA MONITORING REPORT May.28.2026". Split into two paragraphs.
    Dim doc As Document: Set doc = ActiveDocument
    Dim p As Paragraph, s As String, title As String, cutPos As Long
    title = "DAILY MEDIA MONITORING REPORT"
    Set p = doc.Paragraphs(1)
    s = ParaText(p)
    If Left(s, Len(title)) = title And Len(s) > Len(title) Then
        ' cut right after the title, drop the single separating space
        cutPos = p.Range.Start + Len(title)
        Do While doc.Range(cutPos, cutPos + 1).Text = " "
            doc.Range(cutPos, cutPos + 1).Delete
        Loop
        doc.Range(cutPos, cutPos).InsertBefore vbCr
    End If
End Sub


Private Function EndsWithAbbrev(ByVal leftPart As String) As Boolean
    ' leftPart ends at the boundary period. Grab the trailing letters/dots token before it.
    Dim k As Long, ch As String, tok As String
    k = Len(leftPart) - 1        ' char just before the final period
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

    ' single initial like "J." -> skip
    If Len(tok) = 1 And tok >= "A" And tok <= "Z" Then EndsWithAbbrev = True
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
    Dim doc As Document: Set doc = ActiveDocument
    Dim labels As Variant
    labels = Array("Facebook", "X (Twitter)", "YouTube")
    Dim i As Long, j As Long, p As Paragraph, s As String
    Dim lbl As String, cutPos As Long, guard As Long
    For i = doc.Paragraphs.Count To 1 Step -1
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
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long, p As Paragraph, s As String, n As Long
    n = 0
    For i = 1 To doc.Paragraphs.Count
        Set p = doc.Paragraphs(i)
        s = ParaText(p)
        If IsSection(s) Then
            n = 1
        ElseIf IsHeadline(s) And Not IsRosterLine(s) Then
            If n = 0 Then n = 1
            p.Range.InsertBefore CStr(n) & "." & vbTab
            n = n + 1
        End If
    Next i
End Sub


Private Sub EnsureDivider()
    Dim doc As Document: Set doc = ActiveDocument
    Dim lastP As Paragraph
    Set lastP = doc.Paragraphs(doc.Paragraphs.Count)
    If IsDivider(ParaText(lastP)) Then Exit Sub
    If ParaText(lastP) = "" Then
        lastP.Range.InsertBefore String(DIV_LEN, ChrW(8212))
    Else
        doc.Content.InsertAfter vbCr & String(DIV_LEN, ChrW(8212))
    End If
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


Private Function IsHeadline(s As String) As Boolean
    IsHeadline = (InStr(s, " | ") > 0) And Not IsSection(s) And Not IsSocial(s)
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

Private Sub RemoveHorizontalRules()
    ' Delete Word AutoFormat horizontal-rule shapes (o:hr rects, grey line).
    ' These are inline shapes in runs, NOT paragraph borders, so the border
    ' clear loop can't remove them.
    Dim doc As Document: Set doc = ActiveDocument
    Dim sh As InlineShape, i As Long
    For i = doc.InlineShapes.Count To 1 Step -1
        Set sh = doc.InlineShapes(i)
        On Error Resume Next
        sh.Delete
        On Error GoTo 0
    Next i
End Sub
