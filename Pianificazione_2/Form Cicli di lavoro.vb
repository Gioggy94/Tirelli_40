Imports System.Data.SqlClient
Imports System.IO
Imports System.Reflection.Emit






Public Class Form_Cicli_di_lavoro

    Public ultimo_lancio As Integer
    Public ultimo_lanciatore As String

    Public id_ultimo_lancio_mrp_TIR As Integer
    Public id_ultimo_lanciatore_mrp_TIR As String
    Public id_ultimo_lancio_mrp_BRB As Integer
    Public id_ultimo_lanciatore_mrp_brb As String

    Public codice_sap As String
    Public codice_bp As String
    Public riga As Integer
    Private filtro_gruppo As String
    Private id_selezionato As Integer = 0
    Private magazzino_dest As String

    Sub ultimo_lancio_MRP(par_brand As String)


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "SELECT TOP (1) [ID_LANCIO]
      ,[Data_INIZIO]
      ,[Data_FINE]
      ,[UTENTE]
      ,[BRAND]
      ,[STATO]
  FROM [Tirelli_40].[dbo].[MRP_LANCIO]
  where stato='OK' and brand='" & par_brand & "'
  order by id_lancio desc
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() Then
            ultimo_lancio = cmd_SAP_reader("ID_LANCIO")
            ultimo_lanciatore = cmd_SAP_reader("utente")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub

    Public Async Sub ControllaEVisualizzaPDF(par_checkbox As CheckBox, par_disegno As String, par_web_browser As WebBrowser)
        If par_checkbox.Checked Then
            Await VisualizzaPDFInBackground(Homepage.percorso_disegni_generico, par_disegno, par_web_browser)
        End If
    End Sub

    Public Async Function VisualizzaPDFInBackground(par_percorso_base As String, par_disegno As String, par_web_browser As WebBrowser) As Task
        ' Componi il percorso del file PDF
        Dim pdfPath As String = par_percorso_base & "PDF\" & par_disegno & ".PDF"

        ' Verifica l'esistenza del file in modo asincrono
        Dim fileExists As Boolean = Await Task.Run(Function() File.Exists(pdfPath))

        If fileExists Then
            ' Aggiungi i parametri per nascondere barra degli strumenti e pannelli laterali
            pdfPath &= "#toolbar=0&zoom=50&navpanes=0"

            ' Esegui l'operazione di navigazione sul thread dell'UI
            Await Task.Run(Sub()
                               par_web_browser.Invoke(Sub()
                                                          par_web_browser.Show()
                                                          par_web_browser.Navigate(pdfPath)
                                                      End Sub)
                           End Sub)
        Else
            ' Nascondi il WebBrowser sul thread dell'UI
            par_web_browser.Invoke(Sub() par_web_browser.Hide())
        End If
    End Function

    Private Sub Form_Cicli_di_lavoro_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Acquisti.Inserimento_fasi(ComboBox4)
        inizializza_form()
    End Sub

    Sub inizializza_form()
        UT.inserimento_gruppi(ComboBox1)
        MRP.riempi_datagridview_log(DataGridView4)
        ultimo_lancio_MRP("TIR01")
        id_ultimo_lancio_mrp_TIR = ultimo_lancio
        id_ultimo_lanciatore_mrp_TIR = ultimo_lanciatore
        ultimo_lancio_MRP("BRB01")
        id_ultimo_lancio_mrp_BRB = ultimo_lancio
        id_ultimo_lanciatore_mrp_brb = ultimo_lanciatore
        TROVA_CODICi_da_ciclare(DataGridView2, 183)
        UT.inserimento_gruppi(ComboBox1)
        Magazzino.Inserimento_produttore(ComboBox3)
    End Sub

    Sub TROVA_CODICi_da_ciclare(par_datagridview As DataGridView, par_gruppo_articoli As Integer)


        par_datagridview.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "SELECT 
    t10.codice, 
    t11.itemname, 
    t11.frgnname, 
    t11.createdate
FROM
(
    SELECT 
        t1.[itemcode] AS 'Codice'
    FROM oitm t1 inner join oitw t2 on t2.itemcode=t1.itemcode and t2.IsCommited>0
    WHERE 
        SUBSTRING(t1.itemcode, 1, 1) = 'D' 
        AND (t1.[ItmsGrpCod] = " & par_gruppo_articoli & " OR t1.[ItmsGrpCod] = 100)
    GROUP BY t1.[itemcode]
) AS t10
INNER JOIN oitm t11 ON t10.codice = t11.itemcode
WHERE 
    CONVERT(DATETIME, t11.createdate, 120) >= CONVERT(DATETIME, '2025-01-07', 120)
GROUP BY 
    t10.codice, 
    t11.itemname, 
    t11.frgnname, 
    t11.createdate
ORDER BY t11.createdate;

"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader("codice"), cmd_SAP_reader("itemname"), cmd_SAP_reader("frgnname"), cmd_SAP_reader("createdate"))

        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub

    Sub TROVA_CODICi_MRP(par_datagridview As DataGridView, numero_MRP As Integer, par_filtro_gruppo As String, par_codice As String, par_descrizione As String, par_materiale As String, par_da_ord As Integer, par_Disp As Integer, par_data As Date, par_rof As Boolean)
        Dim filtro_codice As String
        If par_codice = "" Then
            filtro_codice = ""
        Else
            filtro_codice = " and t10.cod LIKE '%" & par_codice & "%'"
        End If

        Dim filtro_desc As String
        If par_descrizione = "" Then
            filtro_desc = ""
        Else
            filtro_desc = " and t10.Descrizione LIKE '%" & par_descrizione & "%'"
        End If

        Dim filtro_materiale As String
        If par_materiale = "" Then
            filtro_materiale = ""
        Else
            filtro_materiale = " and t10.materiale LIKE '%" & par_materiale & "%'"
        End If

        Dim filtro_da_ord As String
        If par_da_ord = 0 Then
            filtro_da_ord = ""
        Else
            filtro_da_ord = " and t10.da_ord> " & par_da_ord & ""
        End If
        Dim filtro_rof As String

        If par_rof = True Then
            filtro_rof = " and t10.rof <>'Y'"
        Else
            filtro_rof = ""
        End If

        Dim filtro_disp As String = ""


        filtro_disp = " and t10.disp < " & par_Disp & ""



        par_datagridview.Rows.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "
             select  t10.[contatore],
    t10.[ID],
    t10.[TKT],
    t10.[M_B],
    t10.[X_],
    t10.[DB],
    t10.[COD],
    t10.[DESCRIZIONE],
    t10.[Desc_supp],
    t10.[GRUPPO_ART],
    t10.[DISEGNO],
    t10.[MATERIALE],
    t10.[TRATTAMENTO],
    t10.[MAG],
    t10.[CONF],
    t10.[ORD],
    t10.[DISP],
    t10.[MIN_MAG],
    t10.[MIN_ORD],
    t10.[QTY_BRO],
    t10.[DA_ORD],
    t10.[MU],
    t10.[ROF],
    t10.[CAUSALE_],
    t10.[Motivo_stock],
    t10.[COMM],
    t10.[CLIENTE],
    t10.[causale],
    t10.[TIPO MONT],
    t10.[FASE IMP],
    t10.[QTY_x_COMM],
    t10.[ORD_x_COMM],
    t10.[CONSEGNA],
    t10.[ULTIMO_FORNIT],
    t10.[FORNIT_PREFER],
    t10.[FORNIT_OA_APERTO],
    t10.[Importo]
from
(SELECT 
t0.[contatore],
      t0.[ID]
      ,t0.[N_Ticket] as 'TKT'
,T1.PRCRMNTMTD AS 'M_B'
	  ,substring(t0.[codice],1,1) as 'X_'
	  ,t1.TreeType as 'DB'

      ,t0.[codice] as 'COD'
	  ,t1.itemname as 'DESCRIZIONE'
,coalesce(t1.frgnname,'') as 'Desc_supp'
	  ,t2.ItmsGrpNam as 'GRUPPO_ART'
	  ,t1.u_disegno AS 'DISEGNO'
	  ,t1.U_PRG_TIR_Materiale AS 'MATERIALE'
	  ,t1.U_PRG_TIR_trattamento AS 'TRATTAMENTO'
	  ,t0.[MAG]
      ,t0.[CONF]
      ,t0.[ORD]
      ,t0.[DISP]
	  ,t0.minimo AS 'MIN_MAG'
	  ,coalesce(t1.MinOrdrQty,0) as 'MIN_ORD'
      ,t0.[inter] AS 'QTY_BRO'
      , case
	  when coalesce(t1.MinOrdrQty,0)>=
	  case when t0.[disp]-t0.minimo<0 then -t0.disp-t0.minimo else 0 end
	  +coalesce(a.q_agg,0) then coalesce(t1.MinOrdrQty,0)
	 --+coalesce(0,0) then coalesce(t1.MinOrdrQty,0)
	  else
	  case when t0.[disp]-t0.minimo<0 then -t0.disp+t0.minimo else 0 end+coalesce(a.q_agg,0) end as 'DA_ORD'
	 ,COALESCE(B.u_produzione,'') AS 'MU'
	 ,coalesce(c.rof,'') as 'ROF'
	 ,t0.[MOTIVO] AS 'CAUSALE_'
	 ,coalesce(t1.u_ubimag,'') as 'Motivo_stock'
      ,t0.[commessa] AS 'COMM'
      ,t0.[CLIENTE]
,t0.causale
  ,t0.tipo_montaggio AS 'TIPO MONT'
      ,T0.FASE AS 'FASE IMP'
      ,t0.[quantity] AS 'QTY_x_COMM'
      ,t0.[ord_per_commessa] AS 'ORD_x_COMM'
      ,t0.[CONSEGNA]
      ,t0.[ULTIMO_FORNITORE] AS 'ULTIMO_FORNIT'
	  ,t3.cardname as 'FORNIT_PREFER'
	  ,coalesce(D.CARDNAME,'') as 'FORNIT_OA_APERTO'
, t4.price*case
	  when coalesce(t1.MinOrdrQty,0)>=
	  case when t0.[disp]-t0.minimo<0 then -t0.disp-t0.minimo else 0 end
	  +coalesce(a.q_agg,0) then coalesce(t1.MinOrdrQty,0)
	 --+coalesce(0,0) then coalesce(t1.MinOrdrQty,0)
	  else
	  case when t0.[disp]-t0.minimo<0 then -t0.disp+t0.minimo else 0 end+coalesce(a.q_agg,0) end as 'Importo'
      
      
  FROM [tirelli_40].[dbo].[MRP" & numero_MRP & "] t0
  left join oitm t1 on t0.codice=t1.itemcode
  left join oitb t2 on t1.ItmsGrpcod=t2.ItmsGrpcod
  left join ocrd t3 on t3.cardcode=t1.cardcode
  left join itm1 t4 on t4.itemcode=t0.codice and t4.pricelist=2
  
  left join (select t0.codice, sum(t0.[quantity]) as 'Q_agg'
  from [tirelli_40].[dbo].MRP" & numero_MRP & " t0 
  where t0.id   Like '%%.%%' 
  group by t0.codice ) A on a.codice=t0.codice

   left join (select T0.ITEMCODE,t0.u_produzione
  from owor t0 
  where (t0.status='P' or t0.status='R') and t0.u_produzione   Like '%%INT%%' 
  group by T0.ITEMCODE,t0.u_produzione ) B on B.ITEMCODE=t0.codice

  left join (SELECT  T1.[ItemCode] as 'Codice',  t1.u_prg_azs_commessa as 'ROF'

FROM OPQT T0  INNER JOIN PQT1 T1 ON T0.[DocEntry] = T1.[DocEntry] 


WHERE t0.docstatus='O' AND T1.[LineStatus]='O' AND T0.[DocType]='I' and (substring(t1.itemcode,1,1)='C' or substring(t1.itemcode,1,1)='D' or substring(t1.itemcode,1,1)='0') 

group by T1.[ItemCode] ,  t1.u_prg_azs_commessa ) C on c.codice=t0.codice

 left join (SELECT  T1.[ItemCode] as 'Codice',  T0.CARDNAME

FROM OPOR T0  INNER JOIN POR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 


WHERE t0.docstatus='O' AND T1.[LineStatus]='O' AND T0.[DocType]='I' and (substring(t1.itemcode,1,1)='C' or substring(t1.itemcode,1,1)='D' or substring(t1.itemcode,1,1)='0') 

group by T1.[ItemCode] ,  T0.CARDNAME ) D on D.codice=t0.codice
where  substring(t0.codice,1,1)='D' and t0.disp- case
	  when coalesce(t1.MinOrdrQty,0)>=
	  case when t0.[disp]-t0.minimo<0 then -t0.disp-t0.minimo else 0 end
	  +coalesce(a.q_agg,0) then coalesce(t1.MinOrdrQty,0)
	
	  else
	  case when t0.[disp]-t0.minimo<0 then -t0.disp+t0.minimo else 0 end+coalesce(a.q_agg,0) end<0
)
as t10
  LEFT JOIN  [Tirelli_40].dbo.Valutazioni_mrp_mu T11 ON T11.CODICE=T10.COD AND T11.COMMESSA = T10.COMM AND  CAST(T11.DA_ORD AS VARCHAR)=CAST(T10.DA_ORD AS VARCHAR)
where T11.CODICE IS NULL " & par_filtro_gruppo & filtro_codice & filtro_desc & filtro_materiale & filtro_da_ord & filtro_disp & filtro_rof & "
 order by [cod]
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader("cod"), cmd_SAP_reader("Descrizione"), cmd_SAP_reader("Desc_supp"), cmd_SAP_reader("GRUPPO_ART"), cmd_SAP_reader("Materiale"), cmd_SAP_reader("Disp"), cmd_SAP_reader("Da_ord"), cmd_SAP_reader("Comm"), cmd_SAP_reader("Fase IMP"), cmd_SAP_reader("Importo"), cmd_SAP_reader("Consegna"))

        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub




    Sub info_anagrafiche_distinta(par_codice_padre As String)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = Cnn1

        CMD_SAP.CommandText = "SELECT T0.[qauntity]
FROM oitt T0 left join oitm t1 on t0.code=t1.itemcode

WHERE T0.[code] ='" & par_codice_padre & "' "

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() Then

            TextBox1.Text = cmd_SAP_reader("qauntity")

        Else

            TextBox1.Text = 1

        End If


        cmd_SAP_reader.Close()
        Cnn1.Close()


    End Sub

    Sub compila_anagrafica(par_codice_sap As String)
        Button8.Text = par_codice_sap
        Label1.Text = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione
        Label2.Text = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Descrizione_SUP
        Button12.Text = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Disegno
        TextBox2.Text = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).unita_misura
        ComboBox1.Text = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Gruppo
        ComboBox2.Text = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Approvvigionamento
        Label3.Text = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).nome_fornitore
        codice_bp = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).codice_fornitore


    End Sub

    Sub anteprima_disegno()

        Dim percorso_disegni As String = Homepage.percorso_disegni_generico
        Dim pdfFile As String = percorso_disegni & "PDF\" & Button12.Text & ".PDF"

        ControllaEVisualizzaPDF(CheckBox1, Button12.Text, WebBrowser1)

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Magazzino.Codice_SAP = Button8.Text

        ' Ripristina la finestra se è minimizzata
        If Magazzino.WindowState = FormWindowState.Minimized Then
            Magazzino.WindowState = FormWindowState.Normal
        End If

        ' Porta la finestra in primo piano
        Magazzino.BringToFront()
        Magazzino.Activate()
        Magazzino.Show()

        Magazzino.TextBox2.Text = Button8.Text
        Magazzino.OttieniDettagliAnagrafica(Button8.Text)
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click


        Magazzino.visualizza_disegno(Button12.Text)
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        If ComboBox3.SelectedIndex < 0 Then
            MsgBox("Selezionare un produttore")
            Return
        End If
        aggiorna_anagrafica()
        TROVA_CODICi_da_ciclare(DataGridView2, 183)
    End Sub

    Sub aggiorna_anagrafica()
        Dim answer As Integer
        answer = MsgBox("Confermare di aggiornare le informazioni anagrafiche ?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")


        If answer = vbYes Then

            If ComboBox3.SelectedIndex = -1 Then
                ComboBox3.SelectedIndex = 0

            End If

            Magazzino.aggiornare_descrizione_desc_sup_osservazioni(codice_sap, Label1.Text, Label2.Text, Magazzino.OttieniDettagliAnagrafica(codice_sap).Osservazioni, Button12.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, Magazzino.Elenco_produttori(ComboBox3.SelectedIndex), Magazzino.OttieniDettagliAnagrafica(codice_sap).Catalogo, TextBox2.Text, codice_bp, UT.Elenco_gruppi(ComboBox1.SelectedIndex), Magazzino.OttieniDettagliAnagrafica(codice_sap).gestione_magazzino, Magazzino.OttieniDettagliAnagrafica(codice_sap).motivazione_stock, "")

            Magazzino.update_AITM(codice_sap, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)

            MsgBox("Anagrafica aggiornata con successo")
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Business_partner.Provenienza = "Ciclo_di_lavoro"
        Business_partner.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click


        If riga = 0 Then
            Button7.Visible = False
        End If
        Distinta_base_form.SpostaRigaSu(DataGridView1, riga)
        riga -= 1
    End Sub

    Private Sub DataGridView1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellValueChanged
        If e.RowIndex >= 0 Then
            Dim contatore As Integer

            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Codice) Then


                'Try
                Distinta_base_form.itemcode_riga = UCase(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice").Value)
                Distinta_base_form.informazioni_articolo(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Codice").Value, DataGridView1, e.RowIndex)



            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(Quantità) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Attrezzaggio) Then





                contatore = e.RowIndex

                If InStr(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value, ",") > 1 Then
                    Distinta_base_form.quantità_itt1 = LSet(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value, InStr(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value), InStr(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value), ",") - 1))
                Else
                    Distinta_base_form.quantità_itt1 = Replace(DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value, ",", ".")
                End If

                If InStr(DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value, ",") > 1 Then


                    Distinta_base_form.prezzo_unitario_itt1 = LSet(DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value, InStr(DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value), InStr(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value), ",") - 1))

                Else
                    Distinta_base_form.prezzo_unitario_itt1 = Replace(DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value, ",", ".")
                End If

                If InStr(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, ",") > 1 Then
                    Distinta_base_form.attrezzaggio_itt1 = LSet(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, InStr(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, ",") - 1) & "." & StrReverse(LSet(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value), InStr(StrReverse(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value), ",") - 1))

                Else
                    Distinta_base_form.attrezzaggio_itt1 = Replace(DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, ",", ".")

                End If


                DataGridView1.Rows(riga).Cells(columnName:="Totale").Value = (DataGridView1.Rows(riga).Cells(columnName:="Quantità").Value + DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value) * DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value

                Dim costoTotale As Decimal = 0

                ' Scorre tutte le righe del DataGridView e somma i valori della colonna "Totale"
                For Each row As DataGridViewRow In DataGridView1.Rows
                    If Not IsDBNull(row.Cells("Totale").Value) Then
                        costoTotale += Convert.ToDecimal(row.Cells("Totale").Value)
                    End If
                Next
                TextBox3.Text = costoTotale.ToString("C2") / TextBox1.Text ' Format currency
            End If
        End If

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        riga = e.RowIndex
        If riga = 0 Then
            Button7.Visible = False
        End If
        If riga = DataGridView1.RowCount - 1 Then
            Button6.Visible = False
        End If
        If riga > 0 Then
            Button7.Visible = True
        End If
        If riga < DataGridView1.RowCount - 1 Then
            Button6.Visible = True
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Distinta_base_form.SpostaRigaGiù(DataGridView1, riga)
        riga += 1
        If riga = DataGridView1.RowCount - 2 Then
            Button6.Visible = False
            Button7.Visible = True
        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        aggiorna_distinta_base()


    End Sub

    Sub aggiorna_distinta_base()
        If TextBox1.Text >= 1 Then
            Distinta_base_form.delete_oitt(codice_sap)
            Distinta_base_form.delete_itt1(codice_sap)
            Dim loginstanc As Integer = Distinta_base_form.Trova_ultimo_loginstanc_distinta(codice_sap)
            Distinta_base_form.INSERT_INTO_OITT(codice_sap, TextBox1.Text, Label1.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, loginstanc)
            Distinta_base_form.INSERT_INTO_AITT(codice_sap, TextBox1.Text, Label1.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato, loginstanc)
            Dim contatore As Integer = 0
            Do While contatore <= DataGridView1.Rows.Count - 2


                Distinta_base_form.INSERT_INTO_ITT1(codice_sap, DataGridView1.Rows(contatore).Cells(columnName:="Codice").Value, DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value, DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value, DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, contatore, DataGridView1.Rows(contatore).Cells(columnName:="Magazzino_").Value, DataGridView1.Rows(contatore).Cells(columnName:="Descrizione").Value, loginstanc, DataGridView1.Rows(contatore).Cells(columnName:="Importazione").Value)
                Distinta_base_form.INSERT_INTO_ATT1(codice_sap, DataGridView1.Rows(contatore).Cells(columnName:="Codice").Value, DataGridView1.Rows(contatore).Cells(columnName:="Quantità").Value, DataGridView1.Rows(contatore).Cells(columnName:="prezzo_unitario").Value, DataGridView1.Rows(contatore).Cells(columnName:="Attrezzaggio").Value, contatore, DataGridView1.Rows(contatore).Cells(columnName:="Magazzino_").Value, DataGridView1.Rows(contatore).Cells(columnName:="Descrizione").Value, loginstanc, DataGridView1.Rows(contatore).Cells(columnName:="Importazione").Value)
                contatore = contatore + 1

            Loop
            Distinta_base_form.MEttere_db_produzione_oitm(codice_sap)
            MsgBox("Distinta base creata con successo")




        Else
            MsgBox("Selezionare una quantità > 0")

        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If Not IsNumeric(TextBox1.Text) Then
            MsgBox("Selezionare un valore numerico valido")
            TextBox1.Text = 1
        End If

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        aggiorna_anagrafica()
        aggiorna_distinta_base()
        TROVA_CODICi_da_ciclare(DataGridView2, 183)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        ' Inserisci "R00500" nella prima colonna (indice 0) della nuova riga
        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00500"
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00554"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00550"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00563"
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00562"
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00561"
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00613"
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00564"
    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00505"
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00503"
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00504"
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00526"
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00502"
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00540"
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00572"
    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00587"
    End Sub

    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00600"
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00598"
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00599"
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        ' Aggiungi una nuova riga
        Dim rowIndex As Integer = DataGridView1.Rows.Add()

        DataGridView1.Rows(rowIndex).Cells(3).Value = "R00610"
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click
        UT.Show()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged

    End Sub

    Private Sub CancellareRigaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CancellareRigaToolStripMenuItem.Click
        DataGridView1.Rows.RemoveAt(riga)
    End Sub

    Private Sub DatiAnagraficiArticoloToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DatiAnagraficiArticoloToolStripMenuItem.Click

        Magazzino.Codice_SAP = DataGridView1.Rows(riga).Cells(columnName:="Codice").Value
        ' Ripristina la finestra se è minimizzata
        If Magazzino.WindowState = FormWindowState.Minimized Then
            Magazzino.WindowState = FormWindowState.Normal
        End If

        ' Porta la finestra in primo piano
        Magazzino.BringToFront()
        Magazzino.Activate()
        Magazzino.Show()

        Magazzino.TextBox2.Text = Magazzino.Codice_SAP
        Magazzino.OttieniDettagliAnagrafica(Magazzino.Codice_SAP)
    End Sub

    Private Sub DistintaBaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DistintaBaseToolStripMenuItem.Click

        Dim new_form_distinta_form As New Distinta_base_form

        ' Imposta il valore della TextBox
        new_form_distinta_form.TextBox1.Text = DataGridView1.Rows(riga).Cells("Codice").Value.ToString()

        ' Calcola 1 cm in pixel (conversione da cm a pixel, 96 dpi)
        Dim oneCmInPixels As Integer = CInt(1 / 2.54 * 96) ' 1 cm ≈ 37,8 pixel a 96 dpi

        ' Imposta la posizione della nuova form rispetto alla posizione dell'attuale form
        new_form_distinta_form.StartPosition = FormStartPosition.Manual
        new_form_distinta_form.Location = New Point(Me.Location.X + oneCmInPixels, Me.Location.Y + oneCmInPixels)

        ' Mostra la form
        new_form_distinta_form.Show()
    End Sub

    Private Sub DataGridView1_MouseDown(sender As Object, e As MouseEventArgs) Handles DataGridView1.MouseDown
        ' Controlla se il tasto destro del mouse è stato premuto
        If e.Button = MouseButtons.Right Then
            ' Ottiene l'indice della riga in base alla posizione del mouse
            Dim hit As DataGridView.HitTestInfo = DataGridView1.HitTest(e.X, e.Y)

            ' Verifica se una cella valida è stata cliccata
            If hit.RowIndex >= 0 Then
                ' Seleziona la riga corrispondente
                DataGridView1.ClearSelection()
                DataGridView1.Rows(hit.RowIndex).Selected = True
            End If
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        ' Verifica se c'è una riga selezionata
        If DataGridView1.SelectedRows.Count > 0 Then
            ' Ottieni l'indice della prima riga selezionata
            Dim rowIndex As Integer = DataGridView1.SelectedRows(0).Index

            ' Fai qualcosa con rowIndex, ad esempio visualizzalo o utilizzalo per ulteriori operazioni
            riga = rowIndex.ToString()
        End If
    End Sub







    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim par_datagridview As DataGridView = DataGridView2
        codice_sap = par_datagridview.Rows(e.RowIndex).Cells(columnName:="Codice_lista").Value
        compila_anagrafica(codice_sap)
        anteprima_disegno()
        info_anagrafiche_distinta(codice_sap)
        Distinta_base_form.Riempi_distinta_base(DataGridView1, codice_sap, TextBox3, TextBox1)
    End Sub



    Private Sub DataGridView2_CellContentClick_1(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub

    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick


        Dim par_datagridview As DataGridView = DataGridView3
        If e.RowIndex >= 0 Then
            codice_sap = par_datagridview.Rows(e.RowIndex).Cells(columnName:="Cod_MRP").Value
            compila_anagrafica(codice_sap)
            anteprima_disegno()
            info_anagrafiche_distinta(codice_sap)
            Distinta_base_form.Riempi_distinta_base(DataGridView1, codice_sap, TextBox3, TextBox1)
            compila_odp(par_datagridview.Rows(e.RowIndex).Cells(columnName:="Cod_MRP").Value, Magazzino.OttieniDettagliAnagrafica(par_datagridview.Rows(e.RowIndex).Cells(columnName:="Cod_MRP").Value).Descrizione, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Da_ord").Value, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Commessa").Value, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Fase").Value, par_datagridview.Rows(e.RowIndex).Cells(columnName:="Consegna").Value)
            anagrafiche_min_disp(par_datagridview.Rows(e.RowIndex).Cells(columnName:="Cod_MRP").Value)
        End If
    End Sub

    Sub anagrafiche_min_disp(par_codice_sap As String)
        Label6.Text = Acquisti.disponibilità(par_codice_sap)
        Label5.Text = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Minimo
        Label4.Text = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).minordrqty
        Label7.Text = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Distinta_base

        If Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Distinta_base = "Y" Then
            Button5.BackColor = Color.Lime
            Button7.Hide()
        Else
            Button5.BackColor = Color.Red
            Button7.Show()
        End If
    End Sub

    Sub compila_odp(par_codice_articolo As String, par_descrizione As String, par_quantità As Integer, par_commessa As String, par_fase As String, par_consegna As String)
        Button35.Text = par_codice_articolo
        Label8.Text = par_descrizione
        TextBox10.Text = par_quantità
        TextBox9.Text = par_commessa
        DateTimePicker2.Value = par_consegna
        ComboBox4.Text = par_fase
    End Sub



    Sub carica_checkedlistbox_status(par_checkedlistbox As CheckedListBox, par_utente_mrp As Integer)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select t10.ItmsGrpCod, t10.GRUPPO_ART
from
(
SELECT 
t2.ItmsGrpCod
	  ,t2.ItmsGrpNam as 'GRUPPO_ART'
	 
  FROM [tirelli_40].[dbo].[MRP" & par_utente_mrp & "] t0
  left join oitm t1 on t0.codice=t1.itemcode
  left join oitb t2 on t1.ItmsGrpcod=t2.ItmsGrpcod
  left join ocrd t3 on t3.cardcode=t1.cardcode
  
  left join (select t0.codice, sum(t0.[quantity]) as 'Q_agg'
  from [tirelli_40].[dbo].MRP" & par_utente_mrp & " t0 
  where t0.id   Like '%%.%%' 
  group by t0.codice ) A on a.codice=t0.codice

   left join (select T0.ITEMCODE,t0.u_produzione
  from owor t0 
  where (t0.status='P' or t0.status='R') and t0.u_produzione   Like '%%INT%%' 
  group by T0.ITEMCODE,t0.u_produzione ) B on B.ITEMCODE=t0.codice

  left join (SELECT  T1.[ItemCode] as 'Codice',  t1.u_prg_azs_commessa as 'ROF'

FROM OPQT T0  INNER JOIN PQT1 T1 ON T0.[DocEntry] = T1.[DocEntry] 


WHERE t0.docstatus='O' AND T1.[LineStatus]='O' AND T0.[DocType]='I' and (substring(t1.itemcode,1,1)='C' or substring(t1.itemcode,1,1)='D' or substring(t1.itemcode,1,1)='0') 

group by T1.[ItemCode] ,  t1.u_prg_azs_commessa ) C on c.codice=t0.codice

 left join (SELECT  T1.[ItemCode] as 'Codice',  T0.CARDNAME

FROM OPOR T0  INNER JOIN POR1 T1 ON T0.[DocEntry] = T1.[DocEntry] 


WHERE t0.docstatus='O' AND T1.[LineStatus]='O' AND T0.[DocType]='I' and (substring(t1.itemcode,1,1)='C' or substring(t1.itemcode,1,1)='D' or substring(t1.itemcode,1,1)='0') 

group by T1.[ItemCode] ,  T0.CARDNAME ) D on D.codice=t0.codice
where  substring(t0.codice,1,1)='D' and t0.disp- case
	  when coalesce(t1.MinOrdrQty,0)>=
	  case when t0.[disp]-t0.minimo<0 then -t0.disp-t0.minimo else 0 end
	  +coalesce(a.q_agg,0) then coalesce(t1.MinOrdrQty,0)
	
	  else
	  case when t0.[disp]-t0.minimo<0 then -t0.disp+t0.minimo else 0 end+coalesce(a.q_agg,0) end<0
	  )
	  as t10

	  group by t10.ItmsGrpCod,t10.GRUPPO_ART"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()

            par_checkedlistbox.Items.Add(cmd_SAP_reader_2("Gruppo_ART"), True)



        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()
        'componi_filtro_status()

    End Sub

    Sub componi_filtro_gruppo_articolo(par_checkedlistbox As CheckedListBox)

        filtro_gruppo = "and "
        For Each selectedItem As Object In par_checkedlistbox.CheckedItems
            Dim selectedText As String = selectedItem.ToString()
            Dim firstPart As String = selectedText

            ' Utilizza il valore di firstPart come desideri

            If filtro_gruppo = "and " Then
                filtro_gruppo = filtro_gruppo & "(t10.GRUPPO_ART= '" & firstPart & "'"
            Else
                filtro_gruppo = filtro_gruppo & " or t10.GRUPPO_ART= '" & firstPart & "'"
            End If
        Next
        filtro_gruppo = filtro_gruppo & ")"


    End Sub

    Private Sub CheckedListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        componi_filtro_gruppo_articolo(CheckedListBox1)
    End Sub

    Private Sub Button34_Click(sender As Object, e As EventArgs) Handles Button34.Click
        If id_selezionato = 0 Then
            MsgBox("Selezionare un MRP da caricare")
        Else
            TROVA_CODICi_MRP(DataGridView3, id_selezionato, filtro_gruppo, TextBox5.Text, TextBox6.Text, TextBox7.Text, TextBox8.Text, TextBox4.Text, DateTimePicker1.Value, CheckBox2.Checked)

        End If
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub Button32_Click_1(sender As Object, e As EventArgs) Handles Button32.Click
        If id_selezionato = 0 Then
            MsgBox("Selezionare un MRP da caricare")
        Else
            TabControl2.SelectedTab = TabPage3
            TROVA_CODICi_MRP(DataGridView3, id_selezionato, filtro_gruppo, TextBox5.Text, TextBox6.Text, TextBox7.Text, TextBox8.Text, TextBox4.Text, DateTimePicker1.Value, CheckBox2.Checked)
            carica_checkedlistbox_status(CheckedListBox1, id_selezionato)



        End If


    End Sub

    Private Sub Button33_Click(sender As Object, e As EventArgs) Handles Button33.Click
        MRP.Show()
    End Sub

    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        id_selezionato = DataGridView4.Rows(e.RowIndex).Cells(columnName:="utente_sap").Value
    End Sub

    Private Sub Button35_Click(sender As Object, e As EventArgs) Handles Button35.Click
        Magazzino.Codice_SAP = Button35.Text

        ' Ripristina la finestra se è minimizzata
        If Magazzino.WindowState = FormWindowState.Minimized Then
            Magazzino.WindowState = FormWindowState.Normal
        End If

        ' Porta la finestra in primo piano
        Magazzino.BringToFront()
        Magazzino.Activate()
        Magazzino.Show()

        Magazzino.TextBox2.Text = Button35.Text
        Magazzino.OttieniDettagliAnagrafica(Button35.Text)
    End Sub

    Private Sub DataGridView3_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView3.CellFormatting
        Dim par_datagridview As DataGridView = DataGridView3
        Try
            If par_datagridview.Rows(e.RowIndex).Cells(columnName:="Disp").Value < 0 Then

                par_datagridview.Rows(e.RowIndex).Cells(columnName:="Disp").Style.BackColor = Color.Red


            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button36_Click(sender As Object, e As EventArgs) Handles Button36.Click
        Dim answer As Integer
        If Label7.Text = "N" Then
            MsgBox("Prima creare una distinta base")

        Else
            If ComboBox4.SelectedIndex < 0 Then
                MsgBox("scegliere una fase")
            Else

                If Button35.Text = Nothing Or Button35.Text = "" Then
                    MsgBox("Non risulta alcun articolo")
                Else

                    If Acquisti.disponibilità(Button35.Text) >= 0 And TextBox8.Text <> "STOCK" And TextBox8.Text <> "SCORTA" Then
                        answer = MsgBox("Il codice risulta disponibile e non risulta essere ordinato a STOCK. Confermare l'ordine?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
                        If answer = vbYes Then


                            If ComboBox5.Text = "INT" Or ComboBox5.Text = "INT_SALD" Then
                                Magazzino_dest = "CAP2"
                            ElseIf ComboBox5.Text = "ASSEMBL" Then
                                Magazzino_dest = "02"
                            End If



                            Acquisti.procedura_lancio_odp(Button35.Text, ComboBox5.Text, Acquisti.codice_fase_inserimento, DateTimePicker2.Value, DateTimePicker2.Value, TextBox9.Text, Acquisti.Cliente_relativo_alla_commessa(TextBox9.Text).Cardname, TextBox10.Text, Acquisti.Cliente_relativo_alla_commessa(TextBox9.Text).DocNum, Acquisti.Cliente_relativo_alla_commessa(TextBox9.Text).Cardcode, magazzino_dest)


                            anagrafiche_min_disp(DataGridView2.Rows(riga).Cells(columnName:="Cod_MPR").Value)


                            MsgBox("Ordine lanciato con successo")

                        Else
                            MsgBox("Ordine annullato")

                        End If

                    ElseIf TextBox10.Text < -Acquisti.disponibilità(Button35.Text) Then
                        answer = MsgBox("Il codice risulta ordinato per una minore quantità del necessario. Confermare la quantità?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
                        If answer = vbYes Then

                            If Button35.Text = Nothing Or Button35.Text = "" Then
                                MsgBox("Risulta un errore nel codice da lanciare, non risulta selezionato nessun codice")
                            Else
                                If ComboBox5.Text = "INT" Or ComboBox5.Text = "INT_SALD" Then
                                    Magazzino_dest = "CAP2"
                                ElseIf ComboBox5.Text = "ASSEMBL" Then
                                    Magazzino_dest = "02"
                                End If


                                Acquisti.procedura_lancio_odp(Button35.Text, ComboBox5.Text, Acquisti.codice_fase_inserimento, DateTimePicker2.Value, DateTimePicker2.Value, TextBox9.Text, Acquisti.Cliente_relativo_alla_commessa(TextBox9.Text).Cardname, TextBox10.Text, Acquisti.Cliente_relativo_alla_commessa(TextBox9.Text).DocNum, Acquisti.Cliente_relativo_alla_commessa(TextBox9.Text).Cardcode, magazzino_dest)

                                MsgBox("Ordine lanciato con successo")
                            End If


                            anagrafiche_min_disp(DataGridView2.Rows(riga).Cells(columnName:="Cod_MPR").Value)


                        Else
                            MsgBox("Ordine annullato")

                        End If


                    Else


                        If ComboBox5.Text = "INT" Or ComboBox5.Text = "INT_SALD" Then
                            Magazzino_dest = "CAP2"
                        ElseIf ComboBox5.Text = "ASSEMBL" Then
                            Magazzino_dest = "02"
                        End If



                        Acquisti.procedura_lancio_odp(Button35.Text, ComboBox5.Text, Acquisti.codice_fase_inserimento, DateTimePicker2.Value, DateTimePicker2.Value, TextBox9.Text, Acquisti.Cliente_relativo_alla_commessa(TextBox9.Text).Cardname, TextBox10.Text, Acquisti.Cliente_relativo_alla_commessa(TextBox9.Text).DocNum, Acquisti.Cliente_relativo_alla_commessa(TextBox9.Text).Cardcode, magazzino_dest)

                        anagrafiche_min_disp(Button35.Text)

                        MsgBox("Ordine lanciato con successo")


                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Try
            Acquisti.codice_fase_inserimento = Acquisti.elenco_fasi(ComboBox4.SelectedIndex)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub Button37_Click(sender As Object, e As EventArgs) Handles Button37.Click
        Dim par_datagridview As DataGridView = DataGridView3
        For Each row As DataGridViewRow In par_datagridview.Rows
            insert_into_Valutazioni_mrp_mu(row.Cells("Cod_MRP").Value, row.Cells("commessa").Value, row.Cells("Da_ord").Value, id_selezionato)

        Next
        Beep()
        MsgBox("VALUTAZIONI AGGIORNATE CON SUCCESO")
    End Sub

    Sub insert_into_Valutazioni_mrp_mu(par_codice As String, par_commessa As String, par_da_ord As String, par_utente As Integer)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand

        CMD_SAP_1.Connection = Cnn1
        CMD_SAP_1.CommandText = "INSERT INTO [TIRELLI_40].DBO.Valutazioni_mrp_mu
(Utente,codice,commessa,da_ord,data)
values (" & par_utente & ", '" & par_codice & "' , '" & par_commessa & "', '" & par_da_ord & "', getdate())"

        CMD_SAP_1.ExecuteNonQuery()

        Cnn1.Close()
    End Sub

    Sub cancella_valutazioni_mrp_mu()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand

        CMD_SAP_1.Connection = Cnn1
        CMD_SAP_1.CommandText = "DELETE [TIRELLI_40].DBO.Valutazioni_mrp_mu
"

        CMD_SAP_1.ExecuteNonQuery()

        Cnn1.Close()
    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub Button38_Click(sender As Object, e As EventArgs) Handles Button38.Click
        cancella_valutazioni_mrp_mu()
    End Sub
End Class