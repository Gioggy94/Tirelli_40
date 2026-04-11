Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word
Imports System.Runtime.InteropServices.ComTypes



Public Class Form_scheda_intervento

    Public Elenco_dipendenti(1000) As String
    Public Elenco_owner(1000) As String

    Public codicedip As Integer
    Public esito_controllo As String
    Public inserimento As String

    Public Codice_BP As String
    Public Codice_BP_finale As String

    Public Codice_BP_jG As String
    Public Codice_BP_finale_jG As String



    Public id_intervento
    Public id_intervento_da_modificare
    Private nota_spese As Object

    Public giorni_venduti As Integer
    Public Property giorni_venduti_utilizzati As Integer



    Sub inizializza_form(par_id)
        informazioni_intervento(par_id)
        trova_giorni_venduti_progetto(par_id)
        trova_giorni_venduti_40_progetto(par_id)
        trova_giorni_venduti_utilizzati_progetto(par_id)
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Sub Inserimento_dipendenti()
        ComboBox3.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome'
        FROM [TIRELLI_40].[DBO].OHEM T0 
left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 
where t0.active='Y' order by T0.[lastName]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")


            ComboBox3.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Private Sub TextBox5KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox5.KeyPress
        ' Accetto solo l'immissione di numeri interi e decimali

        ' Recupero il codice ascii del tasto digitato
        ' il tasto digitato è memorizzato nella proprietà "KeyChar"
        ' dell'oggetto System.Windows.Forms.KeyPressEventArgs

        Dim KeyAscii As Short = Asc(e.KeyChar)

        ' In questo caso oltre a consentire numeri, tasto Canc
        ' e tasto BackSpace, devo consentire anche l'immissione
        ' del punto e della virgola
        If KeyAscii < 48 And KeyAscii <> 24 And KeyAscii <> 8 And e.KeyChar <> "." And e.KeyChar <> "," Then
            KeyAscii = 0
        ElseIf KeyAscii > 57 Then
            KeyAscii = 0
        End If

        ' Faccio in modo che se l'utente digita la virgola
        ' mi appaia il punto
        If e.KeyChar = "," Then
            KeyAscii = 46 ' 46 è il codice ascii del punto
        End If

        ' Il punto è si consentito
        ' ma non come primo carattere
        If TextBox1.TextLength = 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If

        ' ovviamente se c'è già un punto
        ' non è consentito digitarne altri
        If (KeyAscii = 46) And
            TextBox1.Text.IndexOf(".") > 0 Then
            KeyAscii = 0
        End If

        ' Reimposto il keychar
        e.KeyChar = Chr(KeyAscii)
    End Sub 'permetto solo numeri come input

    Private Sub TextBox3KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        ' Accetto solo l'immissione di numeri interi e decimali

        ' Recupero il codice ascii del tasto digitato
        ' il tasto digitato è memorizzato nella proprietà "KeyChar"
        ' dell'oggetto System.Windows.Forms.KeyPressEventArgs

        Dim KeyAscii As Short = Asc(e.KeyChar)

        ' In questo caso oltre a consentire numeri, tasto Canc
        ' e tasto BackSpace, devo consentire anche l'immissione
        ' del punto e della virgola
        If KeyAscii < 48 And KeyAscii <> 24 And KeyAscii <> 8 And e.KeyChar <> "." And e.KeyChar <> "," Then
            KeyAscii = 0
        ElseIf KeyAscii > 57 Then
            KeyAscii = 0
        End If

        ' Faccio in modo che se l'utente digita la virgola
        ' mi appaia il punto
        If e.KeyChar = "," Then
            KeyAscii = 46 ' 46 è il codice ascii del punto
        End If

        ' Il punto è si consentito
        ' ma non come primo carattere
        If TextBox1.TextLength = 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If

        ' ovviamente se c'è già un punto
        ' non è consentito digitarne altri
        If (KeyAscii = 46) And
            TextBox1.Text.IndexOf(".") > 0 Then
            KeyAscii = 0
        End If

        ' Reimposto il keychar
        e.KeyChar = Chr(KeyAscii)
    End Sub 'permetto solo numeri come input

    Private Sub TextBox4KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox4.KeyPress
        ' Accetto solo l'immissione di numeri interi e decimali

        ' Recupero il codice ascii del tasto digitato
        ' il tasto digitato è memorizzato nella proprietà "KeyChar"
        ' dell'oggetto System.Windows.Forms.KeyPressEventArgs

        Dim KeyAscii As Short = Asc(e.KeyChar)

        ' In questo caso oltre a consentire numeri, tasto Canc
        ' e tasto BackSpace, devo consentire anche l'immissione
        ' del punto e della virgola
        If KeyAscii < 48 And KeyAscii <> 24 And KeyAscii <> 8 And e.KeyChar <> "." And e.KeyChar <> "," Then
            KeyAscii = 0
        ElseIf KeyAscii > 57 Then
            KeyAscii = 0
        End If

        ' Faccio in modo che se l'utente digita la virgola
        ' mi appaia il punto
        If e.KeyChar = "," Then
            KeyAscii = 46 ' 46 è il codice ascii del punto
        End If

        ' Il punto è si consentito
        ' ma non come primo carattere
        If TextBox1.TextLength = 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If

        ' ovviamente se c'è già un punto
        ' non è consentito digitarne altri
        If (KeyAscii = 46) And
            TextBox1.Text.IndexOf(".") > 0 Then
            KeyAscii = 0
        End If

        ' Reimposto il keychar
        e.KeyChar = Chr(KeyAscii)
    End Sub 'permetto solo numeri come input

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Try
            codicedip = Elenco_dipendenti(ComboBox3.SelectedIndex)
        Catch ex As Exception

        End Try


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Business_partner.Show()


        Business_partner.Provenienza = "Help_desk_interventi_BP"
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Business_partner.Show()

        Business_partner.Provenienza = "Help_desk_interventi_BP_finale"
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        MonthCalendar1.SelectionStart = DateTimePicker1.Value
        Dim difference As TimeSpan = DateTimePicker2.Value - DateTimePicker1.Value
        Dim daysDifference As Integer = difference.Days

        TextBox10.Text = daysDifference.ToString() + 1


    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        MonthCalendar1.SelectionEnd = DateTimePicker2.Value
        Dim difference As TimeSpan = DateTimePicker2.Value - DateTimePicker1.Value
        Dim daysDifference As Integer = difference.Days

        TextBox10.Text = daysDifference.ToString() + 1
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ComboBox5.SelectedIndex < 0 Then
            MsgBox("Scegliere tipo intervento")
            Return
        End If

        If ComboBox6.SelectedIndex < 0 Then
            MsgBox("Scegliere divisione")
            Return
        End If



        If ComboBox2.SelectedIndex < 0 Then
            MsgBox("Selezionare un owner")
        ElseIf ComboBox3.SelectedIndex < 0 Then
            MsgBox("Selezionare un dipendente")
        ElseIf TextBox6.Text = "" Then
            MsgBox("Inserire una commessa")
        ElseIf Codice_BP = "" Then
            MsgBox("Selezionare un cliente")
        ElseIf DateTimePicker1.Value = DateTime.MinValue Then
            MsgBox("Errore nella data inizio")
        ElseIf DateTimePicker2.Value = DateTime.MinValue Then
            MsgBox("Errore nella data fine")
        ElseIf TextBox5.Text = "" Then
            MsgBox("Inserire il numero di ore")
        ElseIf ComboBox1.SelectedIndex < 0 Then
            MsgBox("Selezionare una causale")
        Else
            Trova_ID()
            inserisci_RECORD_intervento_effettuato(id_intervento, Codice_BP_jG, Codice_BP_finale_jG)
            MsgBox("Intervento registrato con successo")
            Help_Desk_Interventi.inizializza_form()
            Me.Close()
        End If
    End Sub

    Sub Trova_ID()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select max(id)+1 as 'ID' 
from [TIRELLI_40].[DBO].Help_desk_interventi_effettuati_jc"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                id_intervento = cmd_SAP_reader_2("ID")
            Else
                id_intervento = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub

    Sub inserisci_RECORD_intervento_effettuato(par_id_intervento As String, par_codice_cliente_jG As String, par_codice_cliente_finale_jG As String)

        Dim par_n_fattura As Integer
        If TextBox1.Text = "" Then
            par_n_fattura = 0
        Else
            par_n_fattura = TextBox1.Text
        End If

        Dim par_n_note_spese As Integer
        If TextBox2.Text = "" Then
            par_n_note_spese = 0
        Else
            par_n_note_spese = TextBox2.Text
        End If

        Dim par_costo_vitto_alloggio As String
        If TextBox4.Text = "" Then
            par_costo_vitto_alloggio = 0
        Else
            par_costo_vitto_alloggio = TextBox4.Text
        End If

        Dim par_costo_trasporto As String
        If TextBox3.Text = "" Then
            par_costo_trasporto = 0
        Else
            par_costo_trasporto = TextBox3.Text
        End If

        Dim commento As String
        commento = Replace(RichTextBox1.Text, "'", " ")
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn
        'CONVERT(DATETIME, '" & ComboBox7.Text & ComboBox6.Text & ComboBox5.Text & "', 112)
        CMD_SAP.CommandText = "insert into 
[TIRELLI_40].[DBO].help_desk_interventi_effettuati_jc (owner, id,dipendente,commessa,cardcode,cardcode_final, Data_inizio , Data_fine, ore, causale, comments,fattura,nota_spese,costo_vitto_alloggio,Costo_trasporto,stato, tipo,ocrcode,Data_ultima_modifica,[cardcode_JG]
      ,[cardcode_Final_JG]) 
values ('" & Elenco_owner(ComboBox2.SelectedIndex) & "','" & par_id_intervento & "','" & codicedip & "','" & TextBox6.Text & "','" & Codice_BP & "', '" & Codice_BP_finale & "',cast (CONVERT(DATETIME,'" & DateTimePicker1.Value & "', 105) as date),cast (CONVERT(DATETIME,'" & DateTimePicker2.Value & "', 105) as date),'" & TextBox5.Text & "','" & ComboBox1.Text & "','" & commento & "','" & par_n_fattura & "','" & par_n_note_spese & "','" & par_costo_vitto_alloggio & "','" & par_costo_trasporto & "','" & ComboBox4.Text & "','" & ComboBox5.Text & "','" & ComboBox6.Text & "',getdate(),'" & par_codice_cliente_jG & "','" & par_codice_cliente_finale_jG & "') "
        CMD_SAP.ExecuteNonQuery()

        Cnn.Close()
    End Sub

    Sub informazioni_intervento(par_id_intervento)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        If Homepage.ERP_provenienza = "SAP" Then
            CMD_SAP_2.CommandText = "select  t0.id,concat(t1.lastname,' ',t1.firstname) as 'owner', t0.dipendente as 'codicedip', t2.lastname+ ' '+ t2.firstname as 'dipendente' ,t0.commessa,case when t5.itemname is null then '' else t5.itemname end as 'itemname',t3.cardcode as 'codicebp',t3.cardname, case when t0.cardcode_final is null then '' else t0.cardcode_final end as 'cardcode_final',case when t4.cardname is null then '' else t4.cardname end as 'cardname_final',t0.data_inizio,t0.data_fine,t0.ore,t0.causale,t0.stato, case when t0.comments is null then '' else t0.comments end as 'comments',
YEAR(T0.DATA_INIZIO) AS 'YEAR_DATA_INIZIO', MONTH(DATA_INIZIO) AS 'MONTH_DATA_INIZIO', DAY(DATA_INIZIO) AS 'DAY_DATA_INIZIO' , YEAR(T0.DATA_FINE) AS 'YEAR_DATA_FINE', MONTH(DATA_FINE) AS 'MONTH_DATA_FINE', DAY(DATA_FINE) AS 'DAY_DATA_FINE'
, case when t5.u_progetto is null then '' else t5.u_progetto end as 'U_progetto'
, 
CASE WHEN t0.fattura IS NULL THEN 0 ELSE T0.FATTURA END AS 'FATTURA',
CASE WHEN t0.nota_spese IS NULL THEN 0 ELSE T0.NOTA_SPESE END AS 'NOTA_SPESE',
CASE WHEN t0.costo_vitto_alloggio IS NULL THEN 0 ELSE T0.COSTO_VITTO_ALLOGGIO END AS 'COSTO_VITTO_ALLOGGIO' ,
CASE WHEN t0.costo_trasporto IS NULL THEN 0 ELSE T0.COSTO_TRASPORTO END AS 'COSTO_TRASPORTO'
,coalesce(t0.ocrcode,'') as 'ocrcode'
,coalesce(t0.tipo,'') as 'tipo'

from [TIRELLI_40].[DBO].help_desk_interventi_effettuati t0 
left join [TIRELLI_40].[DBO].ohem t1 on t1.empid=t0.owner 

left join [TIRELLI_40].[DBO].ohem t2 on t2.empid=t0.dipendente
left join ocrd t3 on t0.cardcode=t3.cardcode
left join ocrd t4 on t0.cardcode_final=t4.cardcode
left join oitm t5 on t5.itemcode=t0.commessa

where t0.id ='" & par_id_intervento & "'
"
        Else
            CMD_SAP_2.CommandText = "SELECT  
    t0.id,
    CONCAT(t1.lastname, ' ', t1.firstname) AS owner,
    t0.dipendente AS codicedip,
    CONCAT(t2.lastname, ' ', t2.firstname) AS dipendente,
    t0.commessa,
    'MANCA' AS itemname,
    coalesce(t3.codesap,'') AS codicebp,
    coalesce(t3.ds_conto,'') AS 'Cardname',
    COALESCE(t0.cardcode_final, '') AS cardcode_final,
    COALESCE(t4.ds_conto, '') AS cardname_final,
    t0.data_inizio,
    t0.data_fine,
    t0.ore,
    t0.causale,
    t0.stato,
    COALESCE(t0.comments, '') AS comments,

    YEAR(t0.data_inizio)  AS YEAR_DATA_INIZIO,
    MONTH(t0.data_inizio) AS MONTH_DATA_INIZIO,
    DAY(t0.data_inizio)   AS DAY_DATA_INIZIO,

    YEAR(t0.data_fine)  AS YEAR_DATA_FINE,
    MONTH(t0.data_fine) AS MONTH_DATA_FINE,
    DAY(t0.data_fine)   AS DAY_DATA_FINE,

    COALESCE(t5.itemcode, '') AS U_progetto,

    COALESCE(t0.fattura, 0)                AS FATTURA,
    COALESCE(t0.nota_spese, 0)             AS NOTA_SPESE,
    COALESCE(t0.costo_vitto_alloggio, 0)   AS COSTO_VITTO_ALLOGGIO,
    COALESCE(t0.costo_trasporto, 0)        AS COSTO_TRASPORTO,

    COALESCE(t0.ocrcode, '') AS ocrcode,
    COALESCE(t0.tipo, '')    AS tipo
	,t0.cardcode_final
,t0.cardcode_jg
,t0.cardcode_Final_JG
	
FROM [TIRELLI_40].[dbo].help_desk_interventi_effettuati_jc t0
left JOIN [TIRELLI_40].[dbo].ohem t1 ON t1.empid = t0.owner
left JOIN [TIRELLI_40].[dbo].ohem t2 ON t2.empid = t0.dipendente

left JOIN [AS400].[S786FAD1].[TIR90VIS].[JGALACF] t3
    ON t0.[cardcode_JG] = t3.conto

LEFT JOIN [AS400].[S786FAD1].[TIR90VIS].[JGALACF] t4
    ON t0.[cardcode_Final_JG]  = t4.conto AND COALESCE(t0.cardcode_final_jg,'') <>''

LEFT JOIN [AS400].[S786FAD1].[TIR90VIS].[JGALCOM] t5
    ON CONCAT('M', t5.subitemcode) = t0.commessa

WHERE t0.id = '" & par_id_intervento & "' "
        End If


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then

            Label1.Text = par_id_intervento

            ComboBox3.Text = cmd_SAP_reader_2("dipendente")

            codicedip = cmd_SAP_reader_2("codicedip")

            Label3.Text = cmd_SAP_reader_2("cardname")
            Codice_BP = cmd_SAP_reader_2("codicebp")

            Label4.Text = cmd_SAP_reader_2("cardname_final")
            Codice_BP_finale = cmd_SAP_reader_2("cardcode_final")
            Codice_BP_jG = cmd_SAP_reader_2("cardcode_jg")
            Codice_BP_finale_jG = cmd_SAP_reader_2("cardcode_final_jg")
            TextBox5.Text = cmd_SAP_reader_2("ore")

            ComboBox1.Text = cmd_SAP_reader_2("causale")

            ComboBox4.Text = cmd_SAP_reader_2("STATO")


            TextBox6.Text = cmd_SAP_reader_2("commessa")
            TextBox7.Text = cmd_SAP_reader_2("u_progetto")
            RichTextBox1.Text = cmd_SAP_reader_2("comments")

            ComboBox2.Text = cmd_SAP_reader_2("owner")


            TextBox1.Text = cmd_SAP_reader_2("fattura")

            TextBox2.Text = cmd_SAP_reader_2("nota_spese")
            TextBox4.Text = cmd_SAP_reader_2("costo_vitto_alloggio")
            TextBox3.Text = cmd_SAP_reader_2("costo_trasporto")

            DateTimePicker1.Value = cmd_SAP_reader_2("data_inizio")
            DateTimePicker2.Value = cmd_SAP_reader_2("data_fine")
            ComboBox6.Text = cmd_SAP_reader_2("ocrcode")
            ComboBox5.Text = cmd_SAP_reader_2("tipo")

        End If

        Cnn1.Close()


    End Sub

    Sub trova_giorni_venduti_progetto(par_id_intervento)
        If Homepage.ERP_provenienza = "SAP" Then


            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_2 As New SqlCommand
            Dim cmd_SAP_reader_2 As SqlDataReader


            CMD_SAP_2.Connection = Cnn1
            CMD_SAP_2.CommandText = "select   coalesce(SUM(coalesce(T3.Quantity , 0 )),0)   as 'Venduti'
from [TIRELLI_40].[DBO].help_desk_interventi_effettuati t0 
inner join TIRELLISRLDB.DBO.oitm t1 on t0.commessa = t1.itemcode
inner join TIRELLISRLDB.DBO.rdr1 t3 on t3.U_PRG_AZS_Commessa=t1.itemcode and (t3.itemcode='L00540' or t3.itemcode='L00508')
inner join TIRELLISRLDB.DBO.ordr t4 on t4.docentry=t3.docentry and T4.CANCELED <>'Y'
where  t0.ID=" & par_id_intervento & "
"

            cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

            If cmd_SAP_reader_2.Read() Then

                giorni_venduti = cmd_SAP_reader_2("Venduti")
                TextBox8.Text = giorni_venduti


            End If

            Cnn1.Close()

        End If

    End Sub

    Sub trova_giorni_venduti_40_progetto(par_id_intervento)
        If Homepage.ERP_provenienza = "SAP" Then


            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_2 As New SqlCommand
            Dim cmd_SAP_reader_2 As SqlDataReader


            CMD_SAP_2.Connection = Cnn1
            CMD_SAP_2.CommandText = "select coalesce(SUM(coalesce(T3.giorni , 0 )),0) as 'Venduti'
from [TIRELLI_40].[DBO].help_desk_interventi_effettuati t0 
inner join TIRELLISRLDB.DBO.oitm t1 on t0.commessa = t1.itemcode
inner join TIRELLISRLDB.DBO.oitm t2 on t2.u_progetto=t1.u_progetto
left join [TIRELLI_40].DBO.giorni_venduti_4_0 t3 on t3.MATRICOLA=t2.itemcode and (t3.itemcode='L00540' or t3.itemcode='L00508')
where t0.ID=" & par_id_intervento & "
"

            cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

            If cmd_SAP_reader_2.Read() Then

                giorni_venduti = cmd_SAP_reader_2("Venduti")
                TextBox11.Text = giorni_venduti


            End If

            Cnn1.Close()

        End If

    End Sub

    Sub trova_giorni_venduti_utilizzati_progetto(par_id_intervento)
        If Homepage.ERP_provenienza = "SAP" Then


            Dim Cnn1 As New SqlConnection
            Cnn1.ConnectionString = Homepage.sap_tirelli
            Cnn1.Open()

            Dim CMD_SAP_2 As New SqlCommand
            Dim cmd_SAP_reader_2 As SqlDataReader


            CMD_SAP_2.Connection = Cnn1
            CMD_SAP_2.CommandText = "select sum(case when datediff(DD,[Data_inizio],[Data_fine])+1 is null then 0 else datediff(DD,[Data_inizio],[Data_fine])+1 end) as 'Giorni'
	 
from
(
SELECT  
        
     
	  
	  t0.Commessa
      
  FROM [TIRELLI_40].[DBO].[Help_desk_interventi_effettuati] t0 
  left join TIRELLISRLDB.DBO.oitm t1 on t0.commessa=t1.itemcode

  WHERE 

  
  t0.id=" & par_id_intervento & "
)
as t10 left join [TIRELLI_40].[DBO].[Help_desk_interventi_effettuati] t11 on t11.Commessa=t10.commessa and t11.CAUSALE='VENDITA'
  
"

            cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

            If cmd_SAP_reader_2.Read() Then

                giorni_venduti_utilizzati = cmd_SAP_reader_2("Giorni")
                TextBox9.Text = giorni_venduti_utilizzati


            End If

            Cnn1.Close()

        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        ELIMINA_intervento(id_intervento_da_modificare)
        MsgBox("Intevento eliminato")
        Help_Desk_Interventi.inizializza_form()
        Me.Close()

    End Sub

    Sub ELIMINA_intervento(par_id_intervento)
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn

        CMD_SAP.CommandText = "DELETE [TIRELLI_40].[DBO].help_desk_interventi_effettuati_jc WHERE ID='" & id_intervento_da_modificare & "'"
        CMD_SAP.ExecuteNonQuery()

        cnn.Close()
    End Sub

    Sub modifica_dati_intervento(par_id_intervento)
        Dim owner As String
        If ComboBox2.SelectedIndex >= 0 Then
            owner = ", owner ='" & Elenco_owner(ComboBox2.SelectedIndex) & "'"
        Else
            owner = ""
        End If

        TextBox4.Text = Replace(TextBox4.Text, ",", ".")
        TextBox3.Text = Replace(TextBox3.Text, ",", ".")

        Dim comments As String
        comments = Replace(RichTextBox1.Text, "'", " ")
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn


        CMD_SAP.CommandText = "UPDATE [TIRELLI_40].[DBO].help_desk_interventi_effettuati_jc 
SET DIPENDENTE='" & codicedip & "'
,commessa='" & TextBox6.Text & "'
, Causale = '" & ComboBox1.Text & "'
, data_inizio= cast (CONVERT(DATETIME,'" & DateTimePicker1.Value & "', 105) as date)
, data_fine= cast (CONVERT(DATETIME,'" & DateTimePicker2.Value & "', 105) as date)
, ore ='" & TextBox5.Text & "',COMMENTS='" & comments & "'
, cardcode='" & Codice_BP & "'
, cardcode_final='" & Codice_BP_finale & "'
,[cardcode_JG]='" & Codice_BP_jG & "'
      ,[cardcode_Final_JG]='" & Codice_BP_finale_jG & "'
,fattura='" & TextBox1.Text & "'
,nota_spese='" & TextBox2.Text & "'
,costo_vitto_alloggio='" & TextBox4.Text & "'
,costo_trasporto='" & TextBox3.Text & "'
,stato='" & ComboBox4.Text & "'
, tipo ='" & ComboBox5.Text & "'
,ocrcode='" & ComboBox6.Text & "'
,Data_ultima_modifica =getdate()
 " & owner & "

        WHERE ID='" & par_id_intervento & "'"
        CMD_SAP.ExecuteNonQuery()

        Cnn.Close()
    End Sub

    Sub Inserimento_owner(par_combobox As ComboBox)
        par_combobox.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
FROM [TIRELLI_40].[DBO].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code 

where t0.active='Y' 
order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()

            Elenco_owner(Indice) = cmd_SAP_reader("Codice dipendenti")

            par_combobox.Items.Add(cmd_SAP_reader("Nome"))
            If cmd_SAP_reader("Codice dipendenti") = Homepage.ID_SALVATO Then
                par_combobox.SelectedIndex = Indice
            End If
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        If ComboBox5.SelectedIndex < 0 Then
            MsgBox("Scegliere tipo intervento")
            Return
        End If
        If ComboBox6.SelectedIndex < 0 Then
            MsgBox("Scegliere divisione")
            Return
        End If
        If TextBox6.Text = Nothing Then
            MsgBox("Inserire una commessa")
        Else
            If Codice_BP = Nothing Then
                MsgBox("Selezionare un cliente")
            Else
                If DateTimePicker1.Value = Nothing Then
                    MsgBox("Errore nella data inizio")
                Else
                    If DateTimePicker2.Value = Nothing Then
                        MsgBox("Errore nella data fine")
                    Else
                        If TextBox5.Text = Nothing Then
                            MsgBox("Inserire il numero di ore")
                        Else
                            If ComboBox1.SelectedIndex < 0 Then
                                MsgBox("Selezionare una causale")

                            Else

                                modifica_dati_intervento(Label1.Text)
                                MsgBox("Intervento aggiornato con successo")
                                Me.Close()
                                Help_Desk_Interventi.inizializza_form()


                            End If
                        End If
                    End If

                End If
            End If
        End If




    End Sub

    Private Sub Form_scheda_intervento_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Inserimento_owner(ComboBox2)
        Inserimento_dipendenti()
    End Sub

    Sub trova_nota_spese()
        If Homepage.ERP_provenienza = "SAP" Then


            Dim Cnn As New SqlConnection
            Cnn.ConnectionString = Homepage.sap_tirelli
            Cnn.Open()



            Dim CMD_SAP As New SqlCommand
            Dim cmd_SAP_reader As SqlDataReader

            CMD_SAP.Connection = Cnn
            CMD_SAP.CommandText = "select  case when max(t0.u_nrnotaspese)>=max(t1.[Nota_spese]) then max(t0.u_nrnotaspese) else max(t1.[Nota_spese]) end +1 as 'Max'

from
ordr t0, [TIRELLI_40].[DBO].[Help_desk_interventi_effettuati] t1
 "

            cmd_SAP_reader = CMD_SAP.ExecuteReader


            If cmd_SAP_reader.Read() Then
                nota_spese = cmd_SAP_reader("Max")

            End If
            cmd_SAP_reader.Close()
            Cnn.Close()
        End If
    End Sub 'Inserisco le risorse nella combo box

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        trova_nota_spese()
        TextBox2.Text = nota_spese
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged

    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Form_giorni_venduti_4_0.Show()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Pianificazione_offerte.Show()
        Pianificazione_offerte.tabella_intestazione = "OINV"
        Pianificazione_offerte.tabella_righe = "INV1"
        Pianificazione_offerte.inizializza_form()
        Pianificazione_offerte.TextBox_cliente.Text = Label3.Text
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        ComboBox6.Text = Homepage.Trova_regola_dist(TextBox6.Text)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged

    End Sub
End Class