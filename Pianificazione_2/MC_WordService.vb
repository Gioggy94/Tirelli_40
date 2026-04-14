Imports System.IO
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Public Class MC_WordService

    Public Sub GeneraManuale(percorsoOutput As String,
                             macchina As MC_Macchina,
                             fotocellule As List(Of MC_Fotocellula),
                             errori As List(Of MC_CodiceErrore),
                             testoFotocellule As String,
                             testoErrori As String,
                             revisione As String,
                             lingua As String)

        Using doc = WordprocessingDocument.Create(percorsoOutput, WordprocessingDocumentType.Document)
            Dim mainPart = doc.AddMainDocumentPart()
            mainPart.Document = New Document()
            Dim body As New Body()
            AddStyles(mainPart)

            ' FRONTESPIZIO
            body.Append(Para(""))
            body.Append(Para(""))
            body.Append(Para("MANUALE DI USO E MANUTENZIONE", bold:=True, center:=True, size:=36))
            body.Append(Para(""))
            body.Append(Para(macchina.NomeMacchina, bold:=True, center:=True, size:=28))
            body.Append(Para(""))
            body.Append(InfoRow("Matricola:", macchina.Matricola))
            body.Append(InfoRow("Modello:", macchina.Modello))
            body.Append(InfoRow("Cliente:", macchina.ClienteFinale))
            body.Append(InfoRow("Anno:", If(macchina.AnnoCostruzione.HasValue, macchina.AnnoCostruzione.Value.ToString(), "")))
            body.Append(InfoRow("Revisione:", revisione))
            body.Append(InfoRow("Lingua:", lingua))
            body.Append(InfoRow("Data:", DateTime.Now.ToString("dd/MM/yyyy")))
            body.Append(PageBreakPara())

            ' CAP. 5.2 COMANDI E FOTOCELLULE
            body.Append(Para("5.2 Comandi e fotocellule", bold:=True, size:=24))
            body.Append(Para(""))
            If Not String.IsNullOrWhiteSpace(testoFotocellule) Then
                For Each linea In testoFotocellule.Split(New Char() {vbLf(0)})
                    body.Append(Para(linea.Trim()))
                Next
            ElseIf fotocellule.Count > 0 Then
                body.Append(TabellaFotocellule(fotocellule))
                body.Append(Para(""))
                For i As Integer = 0 To fotocellule.Count - 1
                    Dim f = fotocellule(i)
                    body.Append(Para($"5.2.{i + 1}  {f.Codice}", bold:=True, size:=20))
                    body.Append(Para($"Tipo: {f.TipoNome}"))
                    body.Append(Para(""))
                Next
            End If
            body.Append(PageBreakPara())

            ' CODICI ERRORE
            body.Append(Para("Messaggi di allarme e codici errore", bold:=True, size:=24))
            body.Append(Para(""))
            If Not String.IsNullOrWhiteSpace(testoErrori) Then
                For Each linea In testoErrori.Split(New Char() {vbLf(0)})
                    body.Append(Para(linea.Trim()))
                Next
            ElseIf errori.Count > 0 Then
                body.Append(TabellaErrori(errori))
            End If

            mainPart.Document.Body = body
            mainPart.Document.Save()
        End Using
    End Sub

    ' ──────────────────────────────────────────────
    ' PARAGRAPHS
    ' ──────────────────────────────────────────────

    Private Function Para(testo As String,
                          Optional bold As Boolean = False,
                          Optional center As Boolean = False,
                          Optional size As Integer = 0) As Paragraph
        Dim p As New Paragraph()
        If center Then
            Dim pPr As New ParagraphProperties()
            pPr.Append(New Justification() With {.Val = JustificationValues.Center})
            p.Append(pPr)
        End If
        If Not String.IsNullOrEmpty(testo) Then
            Dim run As New Run()
            If bold OrElse size > 0 Then
                Dim rPr As New RunProperties()
                If bold Then rPr.Append(New Bold())
                If size > 0 Then rPr.Append(New FontSize() With {.Val = CStr(size * 2)})
                run.Append(rPr)
            End If
            run.Append(New Text(testo) With {.Space = SpaceProcessingModeValues.Preserve})
            p.Append(run)
        End If
        Return p
    End Function

    Private Function InfoRow(etichetta As String, valore As String) As Paragraph
        Dim p As New Paragraph()
        Dim run1 As New Run()
        Dim rPr1 As New RunProperties()
        rPr1.Append(New Bold())
        run1.Append(rPr1)
        run1.Append(New Text(etichetta.PadRight(25)) With {.Space = SpaceProcessingModeValues.Preserve})
        Dim run2 As New Run()
        run2.Append(New Text(If(valore, "")) With {.Space = SpaceProcessingModeValues.Preserve})
        p.Append(run1)
        p.Append(run2)
        Return p
    End Function

    Private Function PageBreakPara() As Paragraph
        Dim p As New Paragraph()
        Dim run As New Run()
        run.Append(New Break() With {.Type = BreakValues.Page})
        p.Append(run)
        Return p
    End Function

    ' ──────────────────────────────────────────────
    ' TABELLE
    ' ──────────────────────────────────────────────

    Private Function TabellaFotocellule(fotocellule As List(Of MC_Fotocellula)) As Table
        Dim t As New Table()
        t.Append(TableProps())
        t.Append(BuildRow(True, "Codice", "Tipo"))
        For Each f In fotocellule
            t.Append(BuildRow(False, f.Codice, f.TipoNome))
        Next
        Return t
    End Function

    Private Function TabellaErrori(errori As List(Of MC_CodiceErrore)) As Table
        Dim t As New Table()
        t.Append(TableProps())
        t.Append(BuildRow(True, "Codice", "Gravita", "Titolo", "Causa", "Rimedio"))
        For Each e In errori
            t.Append(BuildRow(False, e.Codice, e.Gravita, e.Titolo, e.Causa, e.Rimedio))
        Next
        Return t
    End Function

    Private Function BuildRow(isHeader As Boolean, ParamArray celle() As String) As TableRow
        Dim row As New TableRow()
        For Each testo In celle
            Dim cell As New TableCell()
            Dim para As New Paragraph()
            Dim run As New Run()
            If isHeader Then
                Dim rPr As New RunProperties()
                rPr.Append(New Bold())
                run.Append(rPr)
            End If
            run.Append(New Text(If(testo, "")) With {.Space = SpaceProcessingModeValues.Preserve})
            para.Append(run)
            cell.Append(para)
            row.Append(cell)
        Next
        Return row
    End Function

    Private Function TableProps() As TableProperties
        Dim tp As New TableProperties()
        Dim borders As New TableBorders()
        borders.Append(New TopBorder() With {.Val = BorderValues.Single, .Size = 4UI})
        borders.Append(New BottomBorder() With {.Val = BorderValues.Single, .Size = 4UI})
        borders.Append(New LeftBorder() With {.Val = BorderValues.Single, .Size = 4UI})
        borders.Append(New RightBorder() With {.Val = BorderValues.Single, .Size = 4UI})
        borders.Append(New InsideHorizontalBorder() With {.Val = BorderValues.Single, .Size = 4UI})
        borders.Append(New InsideVerticalBorder() With {.Val = BorderValues.Single, .Size = 4UI})
        tp.Append(borders)
        tp.Append(New TableWidth() With {.Type = TableWidthUnitValues.Pct, .Width = "5000"})
        Return tp
    End Function

    ' ──────────────────────────────────────────────
    ' STILI
    ' ──────────────────────────────────────────────

    Private Sub AddStyles(mainPart As MainDocumentPart)
        Dim stylesPart = mainPart.AddNewPart(Of StyleDefinitionsPart)()
        Dim styles As New Styles()
        styles.Append(BuildStyleNormal())
        styles.Append(BuildStyleHeading("Heading1", "heading 1", "36", True))
        styles.Append(BuildStyleHeading("Heading2", "heading 2", "28", True))
        stylesPart.Styles = styles
    End Sub

    Private Function BuildStyleNormal() As Style
        Dim s As New Style()
        s.Type    = StyleValues.Paragraph
        s.StyleId = "Normal"
        s.Default = True
        Dim sn As New StyleName()
        sn.Val = "Normal"
        s.Append(sn)
        Dim srp As New StyleRunProperties()
        Dim rf As New RunFonts()
        rf.Ascii = "Arial"
        rf.HighAnsi = "Arial"
        srp.Append(rf)
        srp.Append(New FontSize() With {.Val = "22"})
        s.Append(srp)
        Return s
    End Function

    Private Function BuildStyleHeading(styleId As String, styleName As String,
                                       fontSize As String, isBold As Boolean) As Style
        Dim s As New Style()
        s.Type = StyleValues.Paragraph
        s.StyleId = styleId
        Dim sn As New StyleName()
        sn.Val = styleName
        s.Append(sn)
        Dim srp As New StyleRunProperties()
        Dim rf As New RunFonts()
        rf.Ascii = "Arial"
        rf.HighAnsi = "Arial"
        srp.Append(rf)
        If isBold Then srp.Append(New Bold())
        srp.Append(New FontSize() With {.Val = fontSize})
        s.Append(srp)
        Return s
    End Function

End Class
