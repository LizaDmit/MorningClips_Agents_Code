Attribute VB_Name = "Module1"
Option Explicit

 

' ================= FormatMorningClips (v8 - Windows, CM indents + margins) =================

' Input: raw Reporter .docx. Output: Report_20 visual format.

' One-shot on a FRESH file; not idempotent.

 

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

    SplitSocials

    PrependSectionNumerals

    NumberHeadlines

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

 

    Application.ScreenUpdating = True

    MsgBox "FormatMorningClips done.", vbInformation

End Sub

 

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

        ElseIf IsHeadline(s) Then

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

    yellow = Split("HongShan|HSG", "|")

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
