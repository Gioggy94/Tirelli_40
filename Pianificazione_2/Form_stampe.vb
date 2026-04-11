Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Printing
Imports System.Runtime.InteropServices
Imports FOXITREADERLib
Imports AxFOXITREADERLib

Public Class Form_stampe


	Public magazzino_destinazione As String
	Public Preview As New PrintPreviewDialog
	Public Sel_Stampante As New PrintDialog
	Public Stampante_Selezionata As Boolean


	Public Codice_articolo As String
	Public id_da_stampare As Integer
	Public ODP_da_stampare As String
	Public stampa_tipo As String
	Public Stampa_ODP As String

	Public Stampa_Matricola As String
	Public Stampa_Ubicazione_Macchina As String
	Public Stampa_progressivo_commessa As String
	Public Stampa_Descrizione As String
	Public stampa_ubicazione_articolo As String
	Public stampa_fornitore As String
	Public stampa_numero_em As String


	Public stampa_tipo_OC As String
	Public stampa_numero_OC As String
	Public lavorazione As String


	Public Stampa_Descrizione_Articolo As String
	Public Stampa_Qta As String
	Public stampa_numero_scontrino As String
	Public N_magazzino_ferretto As String
	Public N_cassetto_ferretto As String
	Public stampa_cliente As String
	Public Stato_odp As String
	Public codice_gruppo_art_padre As String
	Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
		Me.Close()
	End Sub

	Sub start_form_stampe(par_datagridview As DataGridView, par_solo_aperti As Boolean, par_id_scontrino As String, par_utente_galileo As String, par_codice_sap As String)
		Dim filtro_aperti As String
		If par_solo_aperti = True Then
			filtro_aperti = " and YTPELE =1 "
		Else
			filtro_aperti = ""
		End If

		Dim filtro_scontrino As String
		If par_id_scontrino = "" Then
			filtro_scontrino = ""
		Else
			filtro_scontrino = " and Yidrce =''" & par_id_scontrino & "'' "
		End If

		par_datagridview.Rows.Clear()
		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()


		Dim CMD_SAP As New SqlCommand
		Dim cmd_SAP_reader As SqlDataReader

		CMD_SAP.Connection = Cnn

		If Homepage.ERP_provenienza = "SAP" Then


			CMD_SAP.CommandText = " "




		Else

			CMD_SAP.CommandText = "SELECT  
    t10.YIDRCE,
    t10.YTPELE,
    t10.YDTINS,
    t10.YDTEXS,
    t10.YPROFP,
    t10.YDSPRO,
    t10.YORPRP,
    trim(t10.YCDARP) as YCDARP,
    t10.YMGDES,
    t10.YQTAPR,
    t10.YDSARP,
    t10.YCDMAP,
    t10.YDSMAP,
    t10.YCLIEP,
    t10.YLOTTP,
    t10.YBAIAP
,coalesce(t10.yassep,'') as 'yassep'

FROM OPENQUERY(AS400, '
    SELECT  *
    FROM S786FAD1.TIR90VIS.YETMAT0F
where 0=0 " & filtro_scontrino & "
and upper(YDSPRO)  LIKE ''%" & par_utente_galileo & "%''
and upper(YCDARP)  LIKE ''%" & par_codice_sap & "%''

" & filtro_aperti & "
order by YIDRCE desc
limit 500
') AS t10 "

		End If
		cmd_SAP_reader = CMD_SAP.ExecuteReader

		Do While cmd_SAP_reader.Read()

			par_datagridview.Rows.Add(
		cmd_SAP_reader("YIDRCE"),
		cmd_SAP_reader("YTPELE"),
		cmd_SAP_reader("YDTINS"),
		cmd_SAP_reader("YDTEXS"),
		cmd_SAP_reader("YPROFP"),
		cmd_SAP_reader("YDSPRO"),
		cmd_SAP_reader("YORPRP"),
		cmd_SAP_reader("YCDARP"),
		cmd_SAP_reader("YMGDES"),
		cmd_SAP_reader("YQTAPR"),
		cmd_SAP_reader("YDSARP"),
		cmd_SAP_reader("YCDMAP"),
		cmd_SAP_reader("YDSMAP"),
		cmd_SAP_reader("YCLIEP"),
		cmd_SAP_reader("YLOTTP"),
		cmd_SAP_reader("YBAIAP"),
		cmd_SAP_reader("yassep")
	)

		Loop

		cmd_SAP_reader.Close()
		Cnn.Close()

	End Sub

	Private Sub Form_stampe_Load(sender As Object, e As EventArgs) Handles MyBase.Load
		start_form_stampe(DataGridView1, CheckBox1.Checked, TextBox4.Text, TextBox1.Text.ToUpper, TextBox2.Text.ToUpper)
		Timer1.Interval = Homepage.tempo_stampe_scontrini
		TextBox3.Text = Homepage.tempo_stampe_scontrini
	End Sub


	Sub Stampa_stampa(par_id_stampa As Integer)

		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()


		Dim CMD_SAP As New SqlCommand
		Dim cmd_SAP_reader As SqlDataReader

		CMD_SAP.Connection = Cnn

		If Homepage.ERP_provenienza = "SAP" Then

			CMD_SAP.CommandText = " "
		Else

			CMD_SAP.CommandText = "SELECT  
     t10.YIDRCE,
    t10.YTPELE,
    t10.YDTINS,
    t10.YDTEXS,
    t10.YPROFP,
    t10.YDSPRO,
    t10.YORPRP,
    t10.YCDARP,
    t10.YMGDES,
    t10.YQTAPR,
    t10.YDSARP,
    trim(t10.YCDMAP) as 'YCDMAP',
    t10.YDSMAP,
    t10.YCLIEP,
    t10.YLOTTP,
    t10.YBAIAP,
t10.YUBICP,
 t10.YFERRP,
 t10.YCASSP
,T10.YSTATP
,t10.YCLA2P
,t10.YFLG3P as 'impegno_cliente_principale'
,substring(t10.YFLG3P,1,1) as 'Tipo_OC'
,t10.YFLG4P as 'Lavorazione'
,t10.yassep


FROM OPENQUERY(AS400, '
    SELECT *
    FROM S786FAD1.TIR90VIS.YETMAT0F
where YIDRCE =''" & par_id_stampa & "''

') AS t10 "

		End If
		cmd_SAP_reader = CMD_SAP.ExecuteReader

		If cmd_SAP_reader.Read() Then

			'		par_datagridview.Rows.Add(
			'	cmd_SAP_reader("YIDRCE"),
			'	cmd_SAP_reader("YTPELE"),
			'	cmd_SAP_reader("YDTINS"),
			'	cmd_SAP_reader("YDTEXS"),
			'	cmd_SAP_reader("YPROFP"),
			'	cmd_SAP_reader("YDSPRO"),
			'	cmd_SAP_reader("YORPRP"),
			'	cmd_SAP_reader("YCDARP"),
			'	cmd_SAP_reader("YMGDES"),
			'	cmd_SAP_reader("YQTAPR"),
			'	cmd_SAP_reader("YDSARP"),
			'	cmd_SAP_reader("YCDMAP"),
			'	cmd_SAP_reader("YDSMAP"),
			'	cmd_SAP_reader("YCLIEP"),
			'	cmd_SAP_reader("YLOTTP"),
			'	cmd_SAP_reader("YBAIAP")
			')

			stampa_numero_scontrino = cmd_SAP_reader("YIDRCE")
			Stampa_Descrizione_Articolo = cmd_SAP_reader("YDSARP")
			stampa_numero_em = "Manca_num_EM"
			stampa_fornitore = "MANCA_fornitore"
			Stampa_Descrizione = "U_produzione"
			Stampa_progressivo_commessa = cmd_SAP_reader("YLOTTP")
			Stampa_Ubicazione_Macchina = cmd_SAP_reader("YBAIAP")
			Stampa_Matricola = cmd_SAP_reader("YCDMAP")
			Stampa_ODP = cmd_SAP_reader("YORPRP")
			stampa_ubicazione_articolo = cmd_SAP_reader("YUBICP")
			Stampa_Qta = cmd_SAP_reader("YQTAPR")
			N_magazzino_ferretto = cmd_SAP_reader("YFERRP")
			N_cassetto_ferretto = cmd_SAP_reader("YCASSP")
			stampa_cliente = cmd_SAP_reader("YCLIEP")
			Stato_odp = cmd_SAP_reader("YSTATP")
			codice_gruppo_art_padre = cmd_SAP_reader("YCLA2P")
			stampa_tipo = "Trasf " & cmd_SAP_reader("YMGDES")
			stampa_tipo_OC = cmd_SAP_reader("Tipo_oc")
			stampa_numero_OC = cmd_SAP_reader("impegno_cliente_principale")
			lavorazione = cmd_SAP_reader("lavorazione")

			If Stampa_ODP = "P" Then
				MsgBox("L'Ordine " & Stampa_ODP & " era pianificato, rilasciare ODP e poi stampare lo scontrino")
			End If

			Fun_Stampa(stampa_tipo, False, Stampante_Selezionata, Scontrino, cmd_SAP_reader("YMGDES"), "UBIC", cmd_SAP_reader("YCDARP"))
			'	Fun_Stampa(Stampa_Tipo, par_preview_scontrino, Stampante_Selezionata, Scontrino, magazzino_destinazione, "", par_codice_sap)

		End If

		cmd_SAP_reader.Close()
		Cnn.Close()
		SEGNA_CHE_è_STAMPATO(par_id_stampa)
	End Sub

	Sub Fun_Stampa(par_stampa_tipo As String, par_preview_scontrino As Boolean, par_stampante_selezionata As Boolean, par_scontrino As PrintDocument, par_magazzino_destinazione As String, par_ubicazione As String, par_codice_articolo As String)

		Codice_articolo = par_codice_articolo
		magazzino_destinazione = par_magazzino_destinazione
		stampa_tipo = par_stampa_tipo
		Dim altezza_scontrino As Integer
		If par_stampa_tipo <> "Trasferimento interno" Then
			altezza_scontrino = 215
		Else
			altezza_scontrino = 155
		End If

		'preview_scontrino = True

		If par_preview_scontrino = True Then
			If par_stampante_selezionata = False Then
				Sel_Stampante.AllowSomePages = False
				Sel_Stampante.ShowHelp = False
				Sel_Stampante.Document = Scontrino

				' Utilizza un PrintPreviewDialog per mostrare l'anteprima di stampa
				Dim previewDialog As New PrintPreviewDialog()
				previewDialog.Document = Scontrino

				Dim result As DialogResult = previewDialog.ShowDialog()

				If (result = DialogResult.OK) Then
					par_stampante_selezionata = True

					' Ora la stampante è selezionata, puoi chiamare Scontrino.Print()
					par_scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", 185, altezza_scontrino)
					par_scontrino.Print()
				End If
			Else
				' Se la stampante è già stata selezionata in precedenza, stampa direttamente
				par_scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", 185, altezza_scontrino)
				par_scontrino.Print()
			End If

		Else
			If par_stampante_selezionata = False Then
				Sel_Stampante.AllowSomePages = False
				Sel_Stampante.ShowHelp = False
				Sel_Stampante.Document = Scontrino
				Dim result As DialogResult = Sel_Stampante.ShowDialog()
				If (result = DialogResult.OK) Then
					Stampante_Selezionata = True
					' Utilizza un PrintPreviewDialog per mostrare l'anteprima di stampa
					par_scontrino.DefaultPageSettings.PaperSize = New System.Drawing.Printing.PaperSize("Paper Size Name", 185, altezza_scontrino)
					Dim previewDialog As New PrintPreviewDialog()
					previewDialog.Document = Scontrino
					par_scontrino.Print()
				End If
			Else
				par_scontrino.Print()
			End If
		End If


	End Sub


	Private Sub Scontrino_PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles Scontrino.PrintPage


		' === FONT ===
		Dim Penna As New Pen(Color.Black)
		Dim fDiciture As New Font("Calibri", 6, FontStyle.Italic)
		Dim fSmall As New Font("Calibri", 7, FontStyle.Regular)
		Dim fDesc As New Font("Calibri", 8, FontStyle.Italic)
		Dim fMatricola As New Font("Calibri", 12, FontStyle.Bold)
		Dim fQta As New Font("Calibri", 12, FontStyle.Bold)
		Dim fUbicazione As New Font("Calibri", 12, FontStyle.Italic)
		Dim fPosizione As New Font("Calibri", 12, FontStyle.Bold)
		Dim fODP As New Font("Calibri", 16, FontStyle.Bold)
		Dim fCodice As New Font("Calibri", 22, FontStyle.Italic)
		Dim fCodiceMini As New Font("Calibri", 18, FontStyle.Italic)

		Dim g As System.Drawing.Graphics = e.Graphics
		g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias

		' === BORDO ESTERNO ===
		g.DrawRectangle(Penna, 1, 1, 183, 189)

		g.DrawString("N° " & stampa_numero_scontrino, fDiciture, Brushes.Gray, 3, 0)
		' -------------------------------------------------------
		' RIGA 1: [MAG Destinazione (0-88)] [ODP / Ubicazione / CQ (90-182)]
		' -------------------------------------------------------

		' --- Magazzino Destinazione ---
		g.DrawRectangle(Penna, 3, 3, 85, 35)
		g.DrawString("MAG Destinazione", fDiciture, Brushes.Black, 5, 5)
		Dim magTroncato As String = magazzino_destinazione.Substring(0, Math.Min(magazzino_destinazione.Length, 4))
		g.DrawString(magTroncato, fMatricola, Brushes.Black,
				 Valore_centro(3, 85, magTroncato, fMatricola, e, "X"),
				 Valore_centro(3, 35, magTroncato, fMatricola, e, "Y") + 5)

		' --- Riquadro destra: ODP oppure Ubicazione oppure CQ ---
		g.DrawRectangle(Penna, 90, 3, 92, 35)

		Select Case True

			Case magazzino_destinazione = "TSC" OrElse magazzino_destinazione = "SCA" OrElse
			 magazzino_destinazione = "BSCA" OrElse magazzino_destinazione = "T03" OrElse
			 magazzino_destinazione = "03" OrElse magazzino_destinazione = "BSC" OrElse magazzino_destinazione = "T01"

				g.DrawString("Ubicazione", fDiciture, Brushes.Black, 92, 5)
				g.DrawString(stampa_ubicazione_articolo, fMatricola, Brushes.Black,
						 Valore_centro(90, 92, stampa_ubicazione_articolo, fMatricola, e, "X"),
						 Valore_centro(3, 35, stampa_ubicazione_articolo, fMatricola, e, "Y") + 5)

			Case magazzino_destinazione = "CQ"

				Dim docScontrino As String = "TR"
				g.DrawString(docScontrino & " " & stampa_numero_em, fDiciture, Brushes.Black, 92, 5)
				g.DrawString(Microsoft.VisualBasic.Left(stampa_fornitore, 28), fDiciture, Brushes.Black, 92, 15)

			Case Else ' TWP, WIP e tutti gli altri → mostra ODP

				g.DrawString(stampa_tipo, fDiciture, Brushes.Black, 92, 5)
				g.DrawString(Stampa_ODP, fMatricola, Brushes.Black,
						 Valore_centro(90, 92, Stampa_ODP, fMatricola, e, "X"),
						 Valore_centro(3, 35, Stampa_ODP, fMatricola, e, "Y") + 5)

				' -------------------------------------------------------
				' RIGA 4: Stato ODP (riquadro pieno larghezza)
				' -------------------------------------------------------
				g.DrawRectangle(Penna, 3, 156, 179, 22)
				g.DrawString("Stato ODP", fDiciture, Brushes.Black, 5, 158)
				g.DrawString(Stato_odp, fQta, Brushes.Black,
						 Valore_centro(3, 179, Stato_odp, fQta, e, "X"),
						 Valore_centro(156, 22, Stato_odp, fQta, e, "Y") + 2)

		End Select

		' -------------------------------------------------------
		' RIGA 2: Articolo (larghezza piena) + Qta sovrapposta
		' -------------------------------------------------------
		g.DrawRectangle(Penna, 3, 40, 179, 71)
		g.DrawString("Articolo", fDiciture, Brushes.Black, 4, 42)

		Dim fontCodice As Font = If(Len(Codice_articolo) >= 8, fCodice, fCodiceMini)
		g.DrawString(Codice_articolo, fontCodice, Brushes.Black, 4, 44)
		g.DrawString(Stampa_Descrizione_Articolo, fDesc, Brushes.Black, 4, 74)

		' Quantità in alto a destra nel riquadro articolo
		Dim numero_Qta As Decimal
		If Decimal.TryParse(Stampa_Qta.Replace(".", ","), numero_Qta) Then
			Stampa_Qta = Math.Round(numero_Qta).ToString("0")
		End If
		g.DrawString(Stampa_Qta & " PZ", fQta, Brushes.Black, 115, 42)

		' -------------------------------------------------------
		' RIGA 3: [Matricola (3-87)] [Posizione (90-182)]
		' -------------------------------------------------------

		' --- Matricola ---
		g.DrawRectangle(Penna, 3, 114, 84, 40)
		g.DrawString("Matricola", fDiciture, Brushes.Black, 6, 116)
		g.DrawString(Stampa_Matricola, fMatricola, Brushes.Black, 6, 121)

		Dim clienteTroncato As String = If(stampa_cliente.Length > 20,
									   stampa_cliente.Substring(0, 20),
									   stampa_cliente)
		g.DrawString(clienteTroncato, fDiciture, Brushes.Black, 6, 141)

		' Testo aggiuntivo nella cella Matricola (a destra) — tipo macchina / CDS / QE
		If codice_gruppo_art_padre = "63" Then
			g.DrawString("Q.E.", fMatricola, Brushes.Black,
					 Valore_centro(3, 84, "Q.E.", fMatricola, e, "X") + 110, 130)
		ElseIf stampa_tipo_OC = "" OrElse stampa_tipo_OC = " " OrElse stampa_tipo_OC = "B" Then
			g.DrawString(Stampa_Ubicazione_Macchina, fMatricola, Brushes.Black,
					 Valore_centro(3, 84, Stampa_Ubicazione_Macchina, fMatricola, e, "X") + 110, 130)
		Else
			g.DrawString("CDS " & stampa_tipo_OC, fMatricola, Brushes.Black,
					 Valore_centro(3, 84, "CDS " & stampa_tipo_OC, fMatricola, e, "X") + 110, 130)
			g.DrawString("OC " & stampa_numero_OC, fDesc, Brushes.Black,
					 Valore_centro(3, 84, "OC " & stampa_numero_OC, fDesc, e, "X") + 90, 155)
		End If

		' --- Posizione (solo per TWP / WIP) ---
		If magazzino_destinazione = "TWP" OrElse magazzino_destinazione = "WIP" Then
			g.DrawRectangle(Penna, 90, 114, 92, 40)
			g.DrawString("Posizione", fDiciture, Brushes.Black, 93, 116)
			g.DrawString(Stampa_progressivo_commessa, fPosizione, Brushes.Black,
					 Valore_centro(90, 92, Stampa_progressivo_commessa, fPosizione, e, "X") + 20,
					 Valore_centro(114, 40, Stampa_progressivo_commessa, fPosizione, e, "Y") - 10)
		End If

		' --- Ferretto (solo per T15 / FER / 15) ---
		If magazzino_destinazione = "T15" OrElse magazzino_destinazione = "FER" OrElse magazzino_destinazione = "15" Then
			g.DrawRectangle(Penna, 3, 114, 84, 40)
			g.DrawString("N° Fer", fDiciture, Brushes.Black, 6, 116)
			g.DrawString(N_magazzino_ferretto, fMatricola, Brushes.Black, 6, 121)

			g.DrawRectangle(Penna, 90, 114, 92, 40)
			g.DrawString("N° CAS", fDiciture, Brushes.Black, 93, 116)
			g.DrawString(N_cassetto_ferretto, fPosizione, Brushes.Black,
					 Valore_centro(90, 92, N_cassetto_ferretto, fPosizione, e, "X"),
					 Valore_centro(114, 40, N_cassetto_ferretto, fPosizione, e, "Y") + 5)
		End If



		' -------------------------------------------------------
		' FOOTER: dipendente + data/ora
		' -------------------------------------------------------
		Dim DataOggi As String = DateTime.Now.ToString("dd/MM/yy HH:mm")
		g.DrawString(DataOggi, fSmall, Brushes.Black, 120, 179)

		Dim NomeDipendente As String = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).COGNOME &
								   " " & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).NOME
		Dim NomeTroncato As String = NomeDipendente.Substring(0, Math.Min(12, NomeDipendente.Length))
		g.DrawString(NomeTroncato, fSmall, Brushes.Black, 5, 179)

	End Sub

	Function Valore_centro(par_x_iniziale As Integer, par_lunghezza_x As Integer, par_contenuto As String, par_font_size As Font, e As System.Drawing.Printing.PrintPageEventArgs, par_asse As String)


		Dim stringSize As SizeF = e.Graphics.MeasureString(par_contenuto, par_font_size)
		Dim centro As Integer
		If par_asse = "X" Then
			centro = (par_x_iniziale + par_x_iniziale + par_lunghezza_x) / 2 - stringSize.Width / 2
		Else
			centro = (par_x_iniziale + par_x_iniziale + par_lunghezza_x) / 2 - stringSize.Height / 2
			centro = (par_x_iniziale + par_x_iniziale + par_lunghezza_x) / 2 - stringSize.Height / 2
		End If

		Return centro
	End Function



	Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick


		If e.RowIndex >= 0 Then


			If e.ColumnIndex = DataGridView1.Columns.IndexOf(ODP) Then



				ODP_Form.docnum_odp = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ODP").Value
				ODP_Form.Show()
				ODP_Form.inizializza_form(DataGridView1.Rows(e.RowIndex).Cells(columnName:="ODP").Value)



			End If



			id_da_stampare = DataGridView1.Rows(e.RowIndex).Cells(columnName:="id_stampa").Value
			RichTextBox1.Text = id_da_stampare
			ODP_da_stampare = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ODP").Value
			Stampa_ODP = DataGridView1.Rows(e.RowIndex).Cells(columnName:="ODP").Value
		End If

	End Sub

	Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
		Stampa_stampa(id_da_stampare)
	End Sub

	Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
		stampe_automatiche(Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_galileo)
		Timer1.Start()
		Button3.Visible = True
		Panel1.BackColor = Color.Lime
		Button2.Visible = False
	End Sub

	Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
		Timer1.Stop()
		Button2.Visible = True
		Button3.Visible = False
		Panel1.BackColor = Color.IndianRed
	End Sub

	Sub stampe_automatiche(par_utente As String)

		Dim Cnn As New SqlConnection
		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()


		Dim CMD_SAP As New SqlCommand
		Dim cmd_SAP_reader As SqlDataReader

		CMD_SAP.Connection = Cnn

		If Homepage.ERP_provenienza = "SAP" Then

			CMD_SAP.CommandText = " "

		Else

			CMD_SAP.CommandText = "SELECT TOP 1 
    t10.YIDRCE,
    t10.YTPELE,
    t10.YDTINS,
    t10.YDTEXS,
    t10.YPROFP,
    t10.YDSPRO,
    t10.YORPRP,
    t10.YCDARP,
    t10.YMGDES,
    t10.YQTAPR,
    t10.YDSARP,
    t10.YCDMAP,
    t10.YDSMAP,
    t10.YCLIEP,
    t10.YLOTTP,
    t10.YBAIAP
FROM OPENQUERY(AS400, '
    SELECT *
    FROM S786FAD1.TIR90VIS.YETMAT0F
where YTPELE=1 and YPROFP=''" & par_utente & "''
order by YIDRCE
') AS t10 "

		End If
		cmd_SAP_reader = CMD_SAP.ExecuteReader

		If cmd_SAP_reader.Read() Then

			Stampa_stampa(cmd_SAP_reader("YIDRCE"))
			'Await Task.Delay(3000)
		End If

		cmd_SAP_reader.Close()
		Cnn.Close()

	End Sub

	Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
		stampe_automatiche(Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_galileo)
	End Sub

	Sub SEGNA_CHE_è_STAMPATO(PAR_ID_STAMPA As String)




		Dim Cnn As New SqlConnection

		Cnn.ConnectionString = Homepage.sap_tirelli
		Cnn.Open()
		Dim CMD_SAP As New SqlCommand
		CMD_SAP.Connection = Cnn

		CMD_SAP.CommandText = "update
[AS400].[S786FAD1].[TIR90VIS].[YETMAT0F] 
set ytpele=2
from
[AS400].[S786FAD1].[TIR90VIS].[YETMAT0F] 
where yidrce=" & PAR_ID_STAMPA & "

"
		CMD_SAP.ExecuteNonQuery()
		Cnn.Close()





	End Sub

	Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
		start_form_stampe(DataGridView1, CheckBox1.Checked, TextBox4.Text, TextBox1.Text.ToUpper, TextBox2.Text.ToUpper)
	End Sub

	Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

	End Sub

	Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
		Dim par_datagridview As DataGridView
		par_datagridview = DataGridView1

		Dim row As DataGridViewRow = par_datagridview.Rows(e.RowIndex)

		' Condizione: colonna "Stampa" è nulla E colonna "Mag" = "01"
		Dim stampaVuota As Boolean = row.Cells(columnName:="Stampa").Value = "" OrElse
								  row.Cells(columnName:="Stampa").Value = "" OrElse
								  String.IsNullOrWhiteSpace(row.Cells(columnName:="Stampa").Value.ToString())

		Dim magE01 As Boolean = row.Cells(columnName:="Mag").Value IsNot Nothing AndAlso
							 row.Cells(columnName:="Mag").Value.ToString().Trim() = "T01"

		If stampaVuota AndAlso magE01 Then
			row.DefaultCellStyle.ForeColor = Color.Red
		ElseIf row.Cells(columnName:="Stato").Value IsNot Nothing AndAlso
		   CInt(row.Cells(columnName:="Stato").Value) = 2 Then
			row.Cells(columnName:="Stato").Style.BackColor = Color.Lime
			row.DefaultCellStyle.ForeColor = Color.Black ' reset colore testo
		Else
			row.Cells(columnName:="Stato").Style.ForeColor = Color.IndianRed
		End If
	End Sub

	Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
		Me.WindowState = FormWindowState.Minimized
	End Sub

	Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
		If TextBox3.Text = "" Then
			MsgBox("definire un tempo")
		Else
			Homepage.tempo_stampe_scontrini = TextBox3.Text

			Homepage.Aggiorna_INI_COMPUTER()
			Homepage.Enabled = True
			MsgBox("tempo salvato con successo")
		End If
	End Sub

	Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
		Dim testo As String = TextBox3.Text

		If testo = "" Then Exit Sub

		If Not IsNumeric(testo) Then
			TextBox3.Text = testo.Substring(0, testo.Length - 1)
			TextBox3.SelectionStart = TextBox3.Text.Length
		End If
	End Sub

	Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

	End Sub

	Private Sub TextBox4_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
		' Permette solo numeri e backspace
		If Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> ControlChars.Back Then
			e.Handled = True
		End If
	End Sub

	Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

	End Sub

	Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
		ODP_Form.testata_odp(Stampa_ODP)  ' già ottimizzata
		ODP_Form.Fun_Stampa()              ' verifica anche qui
	End Sub
End Class