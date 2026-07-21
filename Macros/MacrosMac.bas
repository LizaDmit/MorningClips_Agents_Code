Attribute VB_Name = "Module2"
Option Explicit

' ================= FormatMorningClips (v2 - raw Reporter .docx input) =================
' Input contract (AAAA.docx shape):
'   - Title / date / "Traditional Media" as plain paragraphs
'   - Section labels auto-numbered, text only: Neil Shen / Company News / VC Industry News
'   - Headlines auto-numbered, "Headline | Outlet https://..." with URL glued on as hyperlink
'   - Body paragraphs pre-merged (left verbatim, not split)
'   - "Social channels:" header line (deleted)
'   - Socials merged: "Facebook No relevant news" etc (split into bullet + line)
'   - No divider in input (added)
' Output: Report_20 visual format. One-shot on a fresh file; not idempotent.

' ---- Tunable constants ----
Private Const IND_SECTION As Single = 0.25   ' inches: i./ii./iii. + social bullets
Private Const IND_ITEM    As Single = 0.5    ' inches: headlines, url, body, "No relevant news"
Private Const SP_TITLE    As Single = 12     ' pt after title
Private Const SP_DATE     As Single = 18     ' pt after date
Private Const SP_TRAD     As Single = 6      ' pt after "Traditional Media"
Private Const SP_SECTION  As Single = 6      ' pt before a section / social heading
Private Const SP_HEAD     As Single = 10     ' pt before an article headline
Private Const SP_BODY     As Single = 10     ' pt after body / "No relevant news"
Private Const DIV_LEN     As Integer = 45    ' em-dashes in closing divider

Sub FormatMorningClips()
    Dim doc As Document
    Set doc = ActiveDocument
    Application.ScreenUpdating = False

' 1. Strip ALL Word auto-numbering FIRST (prevents loop stalls downstream)
    doc.Content.ListFormat.RemoveNumbers

    ' 2. Soft line breaks -> paragraph marks
    ReplaceAll "^l", "^p"

    ' 3. Strip empty paragraphs (bounded)
    Dim pass As Long, lenBefore As Long
    For pass = 1 To 30
        lenBefore = Len(doc.Content.Text)
        ReplaceAll "^p^p", "^p"
        If Len(doc.Content.Text) >= lenBefore Then Exit For
    Next pass

    ' 4. Delete "Social channels:" header line(s)
    DeleteParasEqualTo "Social channels:"

    ' 5. Split glued URL off each headline
    SplitUrlOffHeadlines

    ' 6. Split merged social lines
    SplitSocials

    ' 7. Literal section numerals
    PrependSectionNumerals

    ' 7b. Number headlines within each section (restarts at 1)
    NumberHeadlines

    ' 8. Ensure closing divider
    EnsureDivider

    ' 9. Per-paragraph formatting
    Dim i As Long, p As Paragraph, s As String
    For i = 1 To doc.Paragraphs.Count
        Set p = doc.Paragraphs(i)
        s = ParaText(p)

        ' reset
        With p.Range.Font
            .Name = "Calibri": .Size = 11: .Bold = False: .Underline = wdUnderlineNone
        End With
        p.Alignment = wdAlignParagraphLeft
        p.LeftIndent = 0: p.FirstLineIndent = 0
        p.SpaceBefore = 0: p.SpaceAfter = 0

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
            p.SpaceAfter = SP_TRAD
        ElseIf IsSection(s) Then
            p.Range.Font.Bold = True
            p.LeftIndent = InchesToPoints(IND_ITEM)
            p.FirstLineIndent = InchesToPoints(IND_SECTION - IND_ITEM)
            p.SpaceBefore = SP_SECTION
            p.TabStops.ClearAll
            p.TabStops.Add Position:=InchesToPoints(IND_ITEM), Alignment:=wdAlignTabLeft
        ElseIf IsSocial(s) Then
            p.Range.Font.Bold = True
            p.LeftIndent = InchesToPoints(IND_SECTION)
            p.SpaceBefore = SP_SECTION
        ElseIf IsHeadline(s) Then
            p.Range.Font.Bold = True
            p.LeftIndent = InchesToPoints(IND_ITEM)
            p.SpaceBefore = SP_HEAD
        ElseIf IsURL(s) Then
            p.LeftIndent = InchesToPoints(IND_ITEM)
            If p.Range.Hyperlinks.Count = 0 Then
                MakeHyperlink p, s
            Else
                p.Range.Font.Color = RGB(5, 99, 193)
                p.Range.Font.Underline = wdUnderlineSingle
            End If
        ElseIf IsDivider(s) Then
            p.SpaceBefore = SP_SECTION
        ElseIf s = "No relevant news" Then
            p.LeftIndent = InchesToPoints(IND_SECTION)
            p.SpaceAfter = SP_BODY
        Else   ' body
            p.LeftIndent = InchesToPoints(IND_ITEM)
            p.SpaceAfter = SP_BODY
        End If
    Next i

    ' 10. Highlight keywords by category (skips bold text + URL/divider paras)
    HighlightAll

    Application.ScreenUpdating = True
    MsgBox "FormatMorningClips done.", vbInformation
End Sub

' ---------- structural transforms ----------

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
                    p.Range.InsertBefore ChrW(8226) & " "
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
    Do While Len(t) > 0 And (Right(t, 1) = vbCr Or Right(t, 1) = Chr(7) _
             Or Right(t, 1) = vbLf Or Right(t, 1) = Chr(11))
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
    IsSection = (s Like "i. *") Or (s Like "ii. *") Or (s Like "iii. *") _
             Or (s Like "i." & vbTab & "*") Or (s Like "ii." & vbTab & "*") _
             Or (s Like "iii." & vbTab & "*")
End Function

Private Function IsSocial(s As String) As Boolean
    Dim b As String
    b = ChrW(8226) & " "
    IsSocial = (s = b & "Facebook") Or (s = b & "X (Twitter)") Or (s = b & "YouTube")
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

Private Sub MakeHyperlink(p As Paragraph, url As String)
    Dim r As Range, h As Hyperlink
    Set r = p.Range
    r.MoveEnd wdCharacter, -1
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
    yellow = Split("HongShan|HSG", "|")
    turq = Split("IDG Capital|IDG|ZhenFund|Hillhouse|Granite Asia|" & _
                 "Matrix Partners China|Qiming Venture Partners|KKR|EQT|" & _
                 "TPG|Walden International|Carlyle", "|")

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
Private Sub NumberHeadlines()
    ' Prepends "1. ", "2. "... to headlines; counter restarts at each section
    Dim doc As Document: Set doc = ActiveDocument
    Dim i As Long, p As Paragraph, s As String, n As Long
    n = 0
    For i = 1 To doc.Paragraphs.Count
        Set p = doc.Paragraphs(i)
        s = ParaText(p)
        If IsSection(s) Then
            n = 1
        ElseIf IsHeadline(s) Then
            If n = 0 Then n = 1
            p.Range.InsertBefore CStr(n) & ". "
            n = n + 1
        End If
    Next i
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
        If f.End <= lastPos Then Exit Do   ' progress guard - cannot hang
        lastPos = f.End
        f.Start = f.End
        f.End = rng.End
    Loop
End Sub

