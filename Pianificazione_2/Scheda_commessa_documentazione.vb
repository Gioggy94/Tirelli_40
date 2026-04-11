Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop


Public Class Scheda_commessa_documentazione
    Public commessa As String
    Public Elenco_reparti(1000) As String
    Public variabile_cambiata(1000) As String
    Public variabile As Integer = 0
    Public id_utente As Integer

    Public textbox0_modificato As Integer = 0
    Public textbox1_modificato As Integer = 0
    Public textbox2_modificato As Integer = 0
    Public textbox3_modificato As Integer = 0
    Public combobox4_modificato As Integer = 0
    Public textbox5_modificato As Integer = 0
    Public textbox6_modificato As Integer = 0
    Public textbox7_modificato As Integer = 0
    Public textbox8_modificato As Integer = 0
    Public textbox9_modificato As Integer = 0
    Public textbox10_modificato As Integer = 0
    Public textbox11_modificato As Integer = 0
    Public textbox12_modificato As Integer = 0
    Public combobox13_modificato As Integer = 0
    Public textbox14_modificato As Integer = 0
    Public Combobox15_modificato As Integer = 0
    Public textbox16_modificato As Integer = 0
    Public textbox17_modificato As Integer = 0
    Public textbox18_modificato As Integer = 0
    Public richtextbox19_modificato As Integer = 0
    Public textbox20_modificato As Integer = 0
    Public Combobox21_modificato As Integer = 0
    Public Combobox22_modificato As Integer = 0
    Public textbox23_modificato As Integer = 0
    Public textbox24_modificato As Integer = 0
    Public textbox25_modificato As Integer = 0
    Public textbox26_modificato As Integer = 0
    Public textbox27_modificato As Integer = 0
    Public textbox28_modificato As Integer = 0
    Public textbox29_modificato As Integer = 0
    Public textbox30_modificato As Integer = 0
    Public textbox31_modificato As Integer = 0
    Public textbox32_modificato As Integer = 0
    Public textbox33_modificato As Integer = 0
    Public textbox34_modificato As Integer = 0
    Public textbox35_modificato As Integer = 0
    Public textbox36_modificato As Integer = 0
    Public textbox37_modificato As Integer = 0
    Public textbox38_modificato As Integer = 0
    Public Combobox39_modificato As Integer = 0
    Public Combobox40_modificato As Integer = 0
    Public Combobox41_modificato As Integer = 0
    Public Combobox42_modificato As Integer = 0
    Public Combobox43_modificato As Integer = 0
    Public Richtextbox44_modificato As Integer = 0

    Public Richtextbox45_modificato As Integer = 0
    Public Richtextbox46_modificato As Integer = 0
    Public Richtextbox47_modificato As Integer = 0
    Public Richtextbox48_modificato As Integer = 0
    Public Richtextbox49_modificato As Integer = 0
    Public Richtextbox50_modificato As Integer = 0
    Public Richtextbox51_modificato As Integer = 0
    Public Richtextbox52_modificato As Integer = 0






    Public textbox0_old As String
    Public textbox1_old As String
    Public textbox2_old As String
    Public textbox3_old As String
    Public combobox4_old As String
    Public textbox5_old As String
    Public textbox6_old As String
    Public textbox7_old As String
    Public textbox8_old As String
    Public textbox9_old As String
    Public textbox10_old As String
    Public textbox11_old As String
    Public textbox12_old As String
    Public combobox13_old As String
    Public textbox14_old As String
    Public combobox15_old As String
    Public textbox16_old As String
    Public textbox17_old As String
    Public textbox18_old As String
    Public richtextbox19_old As String
    Public textbox20_old As String
    Public combobox21_old As String
    Public combobox22_old As String
    Public textbox23_old As String
    Public textbox24_old As String
    Public textbox25_old As String
    Public textbox26_old As String
    Public textbox27_old As String
    Public textbox28_old As String
    Public textbox29_old As String
    Public textbox30_old As String
    Public textbox31_old As String
    Public textbox32_old As String
    Public textbox33_old As String
    Public textbox34_old As String
    Public textbox35_old As String
    Public textbox36_old As String
    Public textbox37_old As String
    Public textbox38_old As String
    Public combobox39_old As String
    Public combobox40_old As String
    Public combobox41_old As String
    Public combobox42_old As String
    Public combobox43_old As String
    Public richtextbox44_old As String

    Public richtextbox45_old As String
    Public richtextbox46_old As String
    Public richtextbox47_old As String
    Public richtextbox48_old As String
    Public richtextbox49_old As String
    Public richtextbox50_old As String
    Public richtextbox51_old As String
    Public richtextbox52_old As String


    Public bp_code As String
    Public final_bp_code As String
    Public cartella_macchina As String
    Public nuovo_campo(1000) As String
    Private contatore As Integer = 1
    Public ID As Integer
    Private c As Control()
    Public carico_iniziale As Integer = 0
    Public Elenco_dipendenti(1000) As String
    Public numero_campi As Integer

    Public numero_campo_errore As Integer
    Public Descrizione_campo_errore As String

    Public riga_documenti As Integer

    Public campo As Integer
    Public contenuto As String
    Public contenuto_1 As String
    Public attendibilita As Integer
    Public campo_1 As Integer
    Public inizializzazione As Integer = 0
    Public campo_eliminazione As Integer
    Public codice_bp_campione As String


    Sub compila_anagrafica(par_commessa As String)

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "Select t10.itemcode,coalesce(t15.absentry,0) as 'N_progetto', coalesce(t15.name,'') as 'Nome_progetto', t13.itemname, case when t12.cardcode is null then '' else t12.cardcode end as 'Cardcode', case when t12.cardname is null then '' else t12.cardname end as 'Cardname', t12.docduedate,case when t12.u_destinazione is null then '' else t12.u_destinazione end as 'u_destinazione', case when t12.u_codicebp is null then '' else t12.u_codicebp end as 'codice_Cliente_finale', case when t14.cardname is null then '' else t14.cardname end  as 'Cliente_F'


from
(
SELECT t99.itemcode, max(t0.docentry) as 'Docentry'
from [TIRELLISRLDB].[DBO].oitm t99 
left join [TIRELLISRLDB].[DBO].rdr1 t0 on t99.itemcode=t0.itemcode
where substring(t99.itemcode,1,1)='M' and t99.itemcode ='" & par_commessa & "'
group by t99.itemcode
)
as t10 left join [TIRELLISRLDB].[DBO].rdr1 t11 on t11.itemcode=t10.itemcode and t11.docentry=t10.docentry
left join [TIRELLISRLDB].[DBO].ordr t12 on t12.docentry=t11.docentry
left join [TIRELLISRLDB].[DBO].oitm t13 on t13.itemcode=t10.itemcode
left join [TIRELLISRLDB].[DBO].ocrd t14 on t14.cardcode=t12.u_codicebp
left join [TIRELLISRLDB].[DBO].opmg t15 on t15.absentry=t13.u_progetto
order by t10.docentry DESC, t10.itemcode"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            Label1.Text = cmd_SAP_reader_2("itemcode")
            Label2.Text = cmd_SAP_reader_2("itemname")
            Label3.Text = cmd_SAP_reader_2("cardname")
            Label4.Text = cmd_SAP_reader_2("Cliente_F")
            Label16.Text = cmd_SAP_reader_2("u_destinazione")
            If cmd_SAP_reader_2("docduedate") IsNot DBNull.Value Then
                Dim data As DateTime = Convert.ToDateTime(cmd_SAP_reader_2("docduedate"))
                Label17.Text = data.ToString("dd/MM/yyyy") ' Converte la data nel formato "gg/mm/YYYY" e la assegna a Label17.Text
            Else
                ' Se il valore è DBNull, puoi assegnare un valore predefinito o vuoto a Label17.Text
                Label17.Text = "Valore non disponibile"
                ' oppure Label17.Text = String.Empty
            End If
            'Label17.Text = cmd_SAP_reader_2("docduedate")
            bp_code = cmd_SAP_reader_2("cardcode")
            final_bp_code = cmd_SAP_reader_2("codice_Cliente_finale")
            'cartella_macchina = cmd_SAP_reader_2("u_cartella_macchina")
            cartella_macchina = Scheda_tecnica.trova_percorso_documenti(Label1.Text, "COMMESSA", "")
            LinkLabel2.Text = cartella_macchina


            Button26.Text = cmd_SAP_reader_2("N_progetto")
            Label5.Text = cmd_SAP_reader_2("Nome_progetto")

        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        If cartella_macchina = "-" Then
            trova_cartella_macchina()
        End If

    End Sub


    Private Sub Scheda_commessa_documentazione_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        ComboBox_reparto.Text = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).nome_reparto
        carica_reparti()




    End Sub

    Sub carica_reparti()
        Dim indice = 0
        ComboBox_reparto.Items.Clear()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "select *
from [TIRELLI_40].[DBO].coll_reparti"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            ComboBox_reparto.Items.Add(cmd_SAP_reader_2("Descrizione"))
            Elenco_reparti(indice) = cmd_SAP_reader_2("Id_reparto")
            indice = indice + 1
        Loop
        cmd_SAP_reader_2.Close()
        cnn1.Close()
        ComboBox_reparto.Text = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).nome_reparto

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_reparto.SelectedIndexChanged

        'Homepage.codice_reparto = Elenco_reparti(ComboBox_reparto.SelectedIndex)
        'Homepage.nome_reparto = ComboBox_reparto.Text


        Inserimento_dipendenti()


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub



    Sub check_campo_passato()


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = cnn1
        CMD_SAP_2.CommandText = "Select t10.itemcode, t13.itemname, t12.cardname, case when t14.cardname is null then '' else t14.cardname end  as 'Cliente_F'
from
(
SELECT t0.itemcode, max(t0.docentry) as 'Docentry'
from rdr1 t0
where substring(t0.itemcode,1,1)='M' and t0.itemcode ='" & commessa & "'
group by t0.itemcode
)
as t10 inner join rdr1 t11 on t11.itemcode=t10.itemcode and t11.docentry=t10.docentry
inner join ordr t12 on t12.docentry=t11.docentry
inner join [TIRELLISRLDB].[DBO].oitm t13 on t13.itemcode=t10.itemcode
left join ocrd t14 on t14.cardcode=t12.u_codicebp order by t10.docentry DESC, t10.itemcode"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            Label1.Text = cmd_SAP_reader_2("itemcode")
            Label2.Text = cmd_SAP_reader_2("itemname")
            Label3.Text = cmd_SAP_reader_2("cardname")
            Label4.Text = cmd_SAP_reader_2("Cliente_F")
        End If
        cmd_SAP_reader_2.Close()
        cnn1.Close()
    End Sub

    Sub Aggiorna_record()
        Dim attendibilita As Integer
        Try

            c = Me.Controls.Find("combobox" & variabile_cambiata(contatore), True)
            attendibilita = c(0).Text

        Catch ex As Exception
            attendibilita = 0
        End Try

        Trova_ID()
        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "INSERT INTO [Tirelli_40].[dbo].[Scheda_commesse_Record] (id,Commessa,campo,reparto,dipendente,data,ora,contenuto,attendibilita) Values ('" & ID & "','" & Label1.Text & "', '" & variabile_cambiata(contatore) & "', '" & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto & "', '" & id_utente & "', getdate(),convert(varchar, getdate(), 108),'" & Replace(nuovo_campo(contatore), "'", " ") & "','" & attendibilita & "') "

        CMD_SAP_3.ExecuteNonQuery()

        cnn3.Close()


    End Sub

    Sub Trova_ID()
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "select case when max(id)+1 is null then 1 else max(id)+1 end as 'ID' 
from [Tirelli_40].[dbo].[Scheda_tecnica_record]"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() = True Then
            If Not cmd_SAP_reader_2("ID") Is System.DBNull.Value Then
                ID = cmd_SAP_reader_2("ID")
            Else
                ID = 1
            End If
            cmd_SAP_reader_2.Close()
        End If
        Cnn1.Close()
    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim salta As String = "NO"
        If id_utente = Nothing Then

        Else
            Dim i As Integer = 0

            Do While i <= numero_campi
                c = Me.Controls.Find("combobox" & i, True)

                Try

                    If c(0).Text = Nothing Then
                        numero_campo_errore = i
                        TROVA_CAMPO()
                        MsgBox("Per Il campo " & Descrizione_campo_errore & " Non è stato compilato il valore di ATTENDIBILITA' dell'informazione")
                        salta = "YES"

                    End If
                Catch ex As Exception

                End Try
                i = i + 1
            Loop

            If salta = "NO" Then

                Do While contatore <= variabile

                    Aggiorna_record()
                    Ultimo_aggiornamento()
                    contatore = contatore + 1
                Loop
                contatore = 1


                Scheda_commessa_Pianificazione.carica_commesse(Scheda_commessa_Pianificazione.DataGridView, Scheda_commessa_Pianificazione.TextBox1.Text, Scheda_commessa_Pianificazione.TextBox2.Text, Scheda_commessa_Pianificazione.filtro_cliente_f, Scheda_commessa_Pianificazione.filtro_n_progetto, Scheda_commessa_Pianificazione.filtro_nome_progetto_commessa, Scheda_commessa_Pianificazione.TextBox16.Text.ToUpper, Scheda_commessa_Pianificazione.filtro_desc_sup, "", "")
                MsgBox("Aggiornato con successo")
            End If
        End If


    End Sub

    Sub TROVA_CAMPO()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "select desc_campo 
from [Tirelli_40].[dbo].[Schede_Commesse_Campi]
where id_campo='" & numero_campo_errore & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            Descrizione_campo_errore = cmd_SAP_reader("desc_campo")
        End If
        cmd_SAP_reader.Close()
        cnn.Close()




    End Sub 'Inserisco le risorse nella combo box

    Sub Azzera_campi()
        Dim i As Integer = 0
        Do While i <= numero_campi
            Try
                c = Me.Controls.Find("TextBox_campi_" & i, True)
                c(0).Text = Nothing
            Catch ex As Exception

            End Try


            Try

                c = Me.Controls.Find("combobox" & i, True)

                c(0).Text = 4

            Catch ex As Exception

            End Try


            Try
                c = Me.Controls.Find("combobox_campi_" & i, True)
                c(0).Text = Nothing
            Catch ex As Exception

            End Try

            Try
                c = Me.Controls.Find("RichTextBox" & i, True)
                c(0).Text = Nothing
            Catch ex As Exception

            End Try

            i = i + 1
        Loop



    End Sub

    Sub Inserimento_dipendenti()
        ComboBox_utente.Items.Clear()
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM  [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code inner join [TIRELLI_40].[DBO].[coll_reparti] t2 on (t2.sap_id_reparto =t1.code or t2.sap_id_reparto_2 =t1.code)   where t0.active='Y' and t2.id_reparto='" & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto & "'  order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            ComboBox_utente.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()
        Try
            id_utente = Homepage.ID_SALVATO
            ' ComboBox_utente.Text = Homepage.UTENTE_NOME_SALVATO
        Catch ex As Exception

        End Try





    End Sub 'Inserisco le risorse nella combo box



    Private Sub ComboBox_utente_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox_utente.SelectedIndexChanged
        Homepage.ID_SALVATO = Elenco_dipendenti(ComboBox_utente.SelectedIndex)
        'Homepage.UTENTE_NOME_SALVATO = ComboBox_utente.Text
        id_utente = Elenco_dipendenti(ComboBox_utente.SelectedIndex)

        Homepage.Aggiorna_INI_COMPUTER()


    End Sub

    Sub Ultimo_aggiornamento()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT CONCAT(T12.LASTNAME,' ', T12.firstName) AS 'NOME', T11.DATA, T11.ORA
FROM
(
SELECT max(id) AS 'ID'
  FROM [Tirelli_40].[dbo].[Scheda_tecnica_record] where commessa='" & commessa & "'
  )
  AS T10 INNER JOIN [Tirelli_40].[dbo].[Scheda_tecnica_record] T11 ON T11.ID=T10.ID
  LEFT JOIN [TIRELLI_40].[dbo].OHEM T12 ON T12.EMPID=T11.Dipendente"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        If cmd_SAP_reader_2.Read() Then
            Label14.Text = cmd_SAP_reader_2("nome")
            Label15.Text = cmd_SAP_reader_2("Data") & " " & cmd_SAP_reader_2("ORA")
        Else
            Label14.Text = "-"
            Label15.Text = "-"
        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub




    Sub riempi_datagridview_combinazioni()
        DataGridView1.Rows.Clear()
        DataGridView1.Columns(columnName:="vel_richiesta").Visible = False
        DataGridView1.Columns(columnName:="immagine_1").Visible = False
        DataGridView1.Columns(columnName:="immagine_2").Visible = False
        DataGridView1.Columns(columnName:="immagine_3").Visible = False
        DataGridView1.Columns(columnName:="immagine_3").Visible = False
        DataGridView1.Columns(columnName:="immagine_4").Visible = False
        DataGridView1.Columns(columnName:="immagine_5").Visible = False
        DataGridView1.Columns(columnName:="immagine_6").Visible = False
        DataGridView1.Columns(columnName:="immagine_6").Visible = False
        DataGridView1.Columns(columnName:="immagine_7").Visible = False
        DataGridView1.Columns(columnName:="immagine_8").Visible = False
        DataGridView1.Columns(columnName:="immagine_9").Visible = False
        DataGridView1.Columns(columnName:="immagine_10").Visible = False

        DataGridView1.Columns(columnName:="nome_1").Visible = False
        DataGridView1.Columns(columnName:="nome_2").Visible = False
        DataGridView1.Columns(columnName:="nome_3").Visible = False
        DataGridView1.Columns(columnName:="nome_4").Visible = False
        DataGridView1.Columns(columnName:="nome_5").Visible = False
        DataGridView1.Columns(columnName:="nome_6").Visible = False
        DataGridView1.Columns(columnName:="nome_7").Visible = False
        DataGridView1.Columns(columnName:="nome_8").Visible = False
        DataGridView1.Columns(columnName:="nome_9").Visible = False
        DataGridView1.Columns(columnName:="nome_10").Visible = False

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.id_combinazione,t0.vel_richiesta, t0.campione_1, t11.INIZIALE_SIGLA + T1.NOME   as 'Nome_1',t1.immagine as 'Immagine_1', t0.campione_2, t12.INIZIALE_SIGLA + T2.NOME  as 'Nome_2', t2.immagine as 'Immagine_2', t0.campione_3,t13.INIZIALE_SIGLA + T3.NOME  as 'Nome_3',t3.immagine as 'immagine_3', t0.campione_4,t14.INIZIALE_SIGLA + T4.NOME  as 'Nome_4',t4.immagine as 'immagine_4', t0.campione_5,t15.INIZIALE_SIGLA + T5.NOME  as 'Nome_5',t5.immagine as 'immagine_5', t0.campione_6,t16.INIZIALE_SIGLA + T6.NOME  as 'Nome_6' ,t6.immagine as 'immagine_6', t0.campione_7, t17.INIZIALE_SIGLA + T7.NOME  as 'Nome_7',t7.immagine as 'immagine_7', t0.campione_8,t18.INIZIALE_SIGLA + T8.NOME  as 'Nome_8',t8.immagine as 'immagine_8', t0.campione_9,t19.INIZIALE_SIGLA + T9.NOME  as 'Nome_9',t9.immagine as 'immagine_9', t0.campione_10
,t20.INIZIALE_SIGLA + T10.NOME  as 'Nome_10',t10.immagine as 'immagine_10'

FROM [TIRELLI_40].[DBO].COLL_Combinazioni t0
left join [TIRELLI_40].[DBO].coll_campioni t1 on t0.campione_1=t1.ID_CAMPIONE
left join [TIRELLI_40].[DBO].coll_campioni t2 on t0.campione_2=t2.ID_CAMPIONE
left join [TIRELLI_40].[DBO].coll_campioni t3 on t0.campione_3=t3.ID_CAMPIONE
left join [TIRELLI_40].[DBO].coll_campioni t4 on t0.campione_4=t4.ID_CAMPIONE
left join [TIRELLI_40].[DBO].coll_campioni t5 on t0.campione_5=t5.ID_CAMPIONE
left join [TIRELLI_40].[DBO].coll_campioni t6 on t0.campione_6=t6.ID_CAMPIONE
left join [TIRELLI_40].[DBO].coll_campioni t7 on t0.campione_7=t7.ID_CAMPIONE
left join [TIRELLI_40].[DBO].coll_campioni t8 on t0.campione_8=t8.ID_CAMPIONE
left join [TIRELLI_40].[DBO].coll_campioni t9 on t0.campione_9=t9.ID_CAMPIONE
left join [TIRELLI_40].[DBO].coll_campioni t10 on t0.campione_10=t10.ID_CAMPIONE

left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t11 on t1.TIPO_campione= T11.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t12 on t2.TIPO_campione= T12.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t13 on t3.TIPO_campione= T13.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t14 on t4.TIPO_campione= T14.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t15 on t5.TIPO_campione= T15.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t16 on t6.TIPO_campione= T16.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t17 on t7.TIPO_campione= T17.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t18 on t8.TIPO_campione= T18.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t19 on t9.TIPO_campione= T19.ID_TIPO_CAMPIONE
left join [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t20 on t10.TIPO_campione= T20.ID_TIPO_CAMPIONE

where t0.commessa='" & Label1.Text & "'
order by

t11.INIZIALE_SIGLA ,  cast(substring(T1.NOME,1,99) as integer)"


        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Dim contatore As Integer = 0


        Do While cmd_SAP_reader_2.Read()
            DataGridView1.Columns(columnName:="vel_richiesta").Visible = True
            DataGridView1.Rows.Add(cmd_SAP_reader_2("id_combinazione"))

            If Not cmd_SAP_reader_2("campione_1") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="campione_1").Value = cmd_SAP_reader_2("campione_1")

            End If
            If Not cmd_SAP_reader_2("campione_2") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="campione_2").Value = cmd_SAP_reader_2("campione_2")

            End If
            If Not cmd_SAP_reader_2("campione_3") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="campione_3").Value = cmd_SAP_reader_2("campione_3")

            End If
            If Not cmd_SAP_reader_2("campione_4") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="campione_4").Value = cmd_SAP_reader_2("campione_4")

            End If
            If Not cmd_SAP_reader_2("campione_5") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="campione_5").Value = cmd_SAP_reader_2("campione_5")

            End If
            If Not cmd_SAP_reader_2("campione_6") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="campione_6").Value = cmd_SAP_reader_2("campione_6")

            End If
            If Not cmd_SAP_reader_2("campione_7") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="campione_7").Value = cmd_SAP_reader_2("campione_7")

            End If
            If Not cmd_SAP_reader_2("campione_8") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="campione_8").Value = cmd_SAP_reader_2("campione_8")

            End If
            If Not cmd_SAP_reader_2("campione_9") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="campione_9").Value = cmd_SAP_reader_2("campione_9")

            End If
            If Not cmd_SAP_reader_2("campione_10") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="campione_10").Value = cmd_SAP_reader_2("campione_10")

            End If

            If Not cmd_SAP_reader_2("Immagine_1") Is System.DBNull.Value Then
                Try
                    DataGridView1.Rows(contatore).Cells(columnName:="immagine_1").Value = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine_1"))
                Catch ex As Exception

                End Try

                DataGridView1.Columns(columnName:="immagine_1").Visible = True
            End If

            If Not cmd_SAP_reader_2("Immagine_2") Is System.DBNull.Value Then
                Try
                    DataGridView1.Rows(contatore).Cells(columnName:="immagine_2").Value = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine_2"))
                Catch ex As Exception

                End Try

                DataGridView1.Columns(columnName:="immagine_2").Visible = True
            End If
            If Not cmd_SAP_reader_2("Immagine_3") Is System.DBNull.Value Then
                Try

                    DataGridView1.Rows(contatore).Cells(columnName:="immagine_3").Value = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine_3"))
                    DataGridView1.Columns(columnName:="immagine_3").Visible = True
                Catch ex As Exception

                End Try
            End If
            If Not cmd_SAP_reader_2("Immagine_4") Is System.DBNull.Value Then
                Try
                    DataGridView1.Rows(contatore).Cells(columnName:="immagine_4").Value = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine_4"))
                    DataGridView1.Columns(columnName:="immagine_4").Visible = True
                Catch ex As Exception

                End Try

            End If
            If Not cmd_SAP_reader_2("Immagine_5") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="immagine_5").Value = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine_5"))
                DataGridView1.Columns(columnName:="immagine_5").Visible = True
            End If
            If Not cmd_SAP_reader_2("Immagine_6") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="immagine_6").Value = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine_6"))
                DataGridView1.Columns(columnName:="immagine_6").Visible = True
            End If
            If Not cmd_SAP_reader_2("Immagine_7") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="immagine_7").Value = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine_7"))
                DataGridView1.Columns(columnName:="immagine_7").Visible = True
            End If
            If Not cmd_SAP_reader_2("Immagine_8") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="immagine_8").Value = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine_8"))
                DataGridView1.Columns(columnName:="immagine_8").Visible = True
            End If
            If Not cmd_SAP_reader_2("Immagine_9") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="immagine_9").Value = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine_9"))
                DataGridView1.Columns(columnName:="immagine_9").Visible = True
            End If
            If Not cmd_SAP_reader_2("Immagine_10") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="immagine_10").Value = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine_10"))
                DataGridView1.Columns(columnName:="immagine_10").Visible = True
            End If


            If Not cmd_SAP_reader_2("nome_1") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="nome_1").Value = cmd_SAP_reader_2("nome_1")
                DataGridView1.Columns(columnName:="nome_1").Visible = True
            End If

            If Not cmd_SAP_reader_2("nome_2") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="nome_2").Value = cmd_SAP_reader_2("nome_2")
                DataGridView1.Columns(columnName:="nome_2").Visible = True
            End If
            If Not cmd_SAP_reader_2("nome_3") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="nome_3").Value = cmd_SAP_reader_2("nome_3")
                DataGridView1.Columns(columnName:="nome_3").Visible = True
            End If
            If Not cmd_SAP_reader_2("nome_4") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="nome_4").Value = cmd_SAP_reader_2("nome_4")
                DataGridView1.Columns(columnName:="nome_4").Visible = True
            End If
            If Not cmd_SAP_reader_2("nome_5") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="nome_5").Value = cmd_SAP_reader_2("nome_5")
                DataGridView1.Columns(columnName:="nome_5").Visible = True
            End If
            If Not cmd_SAP_reader_2("nome_6") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="nome_6").Value = cmd_SAP_reader_2("nome_6")
                DataGridView1.Columns(columnName:="nome_6").Visible = True
            End If
            If Not cmd_SAP_reader_2("nome_7") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="nome_7").Value = cmd_SAP_reader_2("nome_7")
                DataGridView1.Columns(columnName:="nome_7").Visible = True
            End If
            If Not cmd_SAP_reader_2("nome_8") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="nome_8").Value = cmd_SAP_reader_2("nome_8")
                DataGridView1.Columns(columnName:="nome_8").Visible = True
            End If
            If Not cmd_SAP_reader_2("nome_9") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="nome_9").Value = cmd_SAP_reader_2("nome_9")
                DataGridView1.Columns(columnName:="nome_9").Visible = True
            End If
            If Not cmd_SAP_reader_2("nome_10") Is System.DBNull.Value Then
                DataGridView1.Rows(contatore).Cells(columnName:="nome_10").Value = cmd_SAP_reader_2("nome_10")
                DataGridView1.Columns(columnName:="nome_10").Visible = True


            End If
            DataGridView1.Rows(contatore).Cells(columnName:="vel_richiesta").Value = cmd_SAP_reader_2("vel_richiesta")

            contatore = contatore + 1

        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        DataGridView1.ClearSelection()

    End Sub


    Sub riempi_datagridview_campioni()

        DataGridView3.Rows.Clear()
        DataGridView3.Columns(columnName:="Campione_").Visible = False
        DataGridView3.Columns(columnName:="Tipo_").Visible = False
        DataGridView3.Columns(columnName:="Nome_").Visible = False
        DataGridView3.Columns(columnName:="immagine_").Visible = False
        DataGridView3.Columns(columnName:="dato_6").Visible = False
        DataGridView3.Columns(columnName:="desc").Visible = False


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT t0.ID_CAMPIONE,  t1.INIZIALE_SIGLA + T0.NOME as 'Nome', case when (t0.immagine is null or t0.immagine ='') then 'N_A.JPG' else t0.immagine end as 'immagine', t1.descrizione as 'Tipo'
, t0.Dato_6, t0.descrizione
from [tirelli_40].[DBO].coll_campioni t0 left  join  [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t1 on t0.TIPO_campione= T1.ID_TIPO_CAMPIONE
where (t0.codice_bp=cast('" & codice_bp_campione & "' as integer) or t0.codice_bp=cast('" & bp_code & "' as integer) or  t0.codice_bp=cast('" & final_bp_code & "' as integer))

order by

t1.INIZIALE_SIGLA ,  cast(substring(T0.NOME,1,99) as integer)
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        ' Dim contatore As Integer = 0
        Dim i As Integer = 0
        Do While cmd_SAP_reader_2.Read()


            DataGridView3.Columns(columnName:="Tipo_").Visible = True
            DataGridView3.Columns(columnName:="Nome_").Visible = True
            DataGridView3.Columns(columnName:="immagine_").Visible = True
            DataGridView3.Columns(columnName:="dato_6").Visible = True
            DataGridView3.Columns(columnName:="desc").Visible = True


            Dim MyImage As Bitmap

            Try
                MyImage = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine"))
            Catch ex As Exception
                MyImage = Image.FromFile(Homepage.Percorso_immagini & "N_A.JPG")

            End Try
            Console.WriteLine(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine"))
            DataGridView3.Rows.Add(cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Nome"), cmd_SAP_reader_2("Tipo"), MyImage, cmd_SAP_reader_2("dato_6"), cmd_SAP_reader_2("descrizione")) 'Image.FromFile(cmd_SAP_reader_2("immagine"))

            i = i + 1
        Loop


        cmd_SAP_reader_2.Close()
        Cnn1.Close()


        DataGridView3.ClearSelection()

    End Sub




    Sub cerca_file()
        DataGridView2.Rows.Clear()
        Dim PERCORSO As String = Homepage.percorso_cartelle_macchine & LinkLabel2.Text
        Try
            For Each foundFile As String In My.Computer.FileSystem.GetFiles(PERCORSO)


                DataGridView2.Rows.Add(foundFile, "M", Mid(foundFile, Len(LinkLabel2.Text) + 2, 999999), Mid(foundFile, InStr(foundFile, "."), 999999))
            Next
        Catch ex As Exception

        End Try
    End Sub




    Sub filtra()
        Dim i = 0
        Dim parola0 As String
        Dim parola1 As String
        Dim parola2 As String


        Do While i < DataGridView2.RowCount

            Try
                parola0 = UCase(DataGridView2.Rows(i).Cells(columnName:="Tipo").Value)
                parola1 = UCase(DataGridView2.Rows(i).Cells(columnName:="Nome").Value)
                parola2 = UCase(DataGridView2.Rows(i).Cells(columnName:="File").Value)


                If parola0.Contains(UCase(ComboBox_filtro_tipo.Text)) Then
                    DataGridView2.Rows(i).Visible = True
                    If parola1.Contains(UCase(TextBox_filtro_nome.Text)) Then
                        DataGridView2.Rows(i).Visible = True


                        If parola2.Contains(UCase(TextBox_filtro_estensione.Text)) Then
                            DataGridView2.Rows(i).Visible = True

                        Else
                            DataGridView2.Rows(i).Visible = False

                        End If


                    Else
                        DataGridView2.Rows(i).Visible = False

                    End If

                Else
                    DataGridView2.Rows(i).Visible = False

                End If

            Catch ex As Exception


            End Try
            i = i + 1
        Loop
    End Sub










    Private Sub TextBox_campi_38_TextChanged(sender As Object, e As EventArgs) Handles TextBox_campi_38.TextChanged
        If carico_iniziale = 1 Then
            If textbox38_modificato = 0 Then
                textbox38_old = TextBox_campi_38.Text
                variabile = variabile + 1
            End If


            nuovo_campo(variabile) = TextBox_campi_38.Text
            variabile_cambiata(variabile) = 38


            textbox38_modificato = 1
        End If
    End Sub



    Private Sub RichTextBox44_TextChanged(sender As Object, e As EventArgs) Handles RichTextBox44.TextChanged
        If carico_iniziale = 1 Then
            If Richtextbox44_modificato = 0 Then
                richtextbox44_old = RichTextBox44.Text
                variabile = variabile + 1
            End If


            nuovo_campo(variabile) = RichTextBox44.Text
            variabile_cambiata(variabile) = 44


            Richtextbox44_modificato = 1

        End If
    End Sub



    Private Sub Cmd_Inserisci_Click(sender As Object, e As EventArgs) Handles Cmd_Inserisci.Click

        Form_nuovo_campione.Show()
        Form_nuovo_campione.inizializza_form()

        If final_bp_code = "" Then
            'Form_Inserisci_Campioni.Codice_BP = bp_code
            Form_nuovo_campione.Codice_BP_selezionato = bp_code
        Else
            Form_nuovo_campione.Codice_BP_selezionato = final_bp_code

        End If


        'Form_Inserisci_Campioni.Cerca_BP_Codice()
        'Form_Inserisci_Campioni.Show()
        'Form_Inserisci_Campioni.inizializzazione_form()

    End Sub

    Private Sub Cmd_Inserimento_Combinazioni_Click(sender As Object, e As EventArgs) Handles Cmd_Inserimento_Combinazioni.Click
        If final_bp_code = "" Then
            Form_Nuova_combinazione.codice_bp = bp_code
        Else
            Form_Nuova_combinazione.codice_bp = final_bp_code
        End If
        Form_Nuova_combinazione.Label4.Text = Form_Nuova_combinazione.TROVA_MAX_COMBINAZIONE(Label1.Text)
        Form_Nuova_combinazione.codice_commessa = Label1.Text
        Form_Nuova_combinazione.Show()


    End Sub





    Private Sub LinkLabel2_LinkClicked_1(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        Try
            Process.Start(Homepage.percorso_cartelle_macchine & LinkLabel2.Text)
        Catch ex As Exception
            MsgBox("Il percorso non esiste")
        End Try

    End Sub





    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        riga_documenti = e.RowIndex

    End Sub

    Private Sub DataGridView2_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellDoubleClick
        If riga_documenti >= 0 Then
            Process.Start(DataGridView2.Rows(riga_documenti).Cells(columnName:="Link").Value)
        End If
    End Sub

    Private Sub ComboBox_filtro_tipo_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox_filtro_tipo.SelectedIndexChanged
        filtra()
    End Sub

    Private Sub TextBox_filtro_nome_TextChanged(sender As Object, e As EventArgs) Handles TextBox_filtro_nome.TextChanged
        filtra()
    End Sub

    Private Sub TextBox_filtro_estensione_TextChanged(sender As Object, e As EventArgs) Handles TextBox_filtro_estensione.TextChanged
        filtra()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            LinkLabel2.Text = Replace(FolderBrowserDialog1.SelectedPath, Homepage.percorso_progetti, "")
            Scheda_tecnica.Aggiorna_percorso_macchina(LinkLabel2.Text, Label1.Text, "COMMESSA")
            cerca_file()
        End If
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub






    Sub Inserisci_record()
        If inizializzazione = 1 Then
            If ComboBox_reparto.SelectedIndex < 0 Then


            ElseIf ComboBox_utente.SelectedIndex < 0 Then

            Else
                Trova_ID()
                Dim Cnn As New SqlConnection
                Cnn.ConnectionString = Homepage.sap_tirelli


                Cnn.Open()

                Dim Cmd_SAP As New SqlCommand

                Cmd_SAP.Connection = Cnn
                Cmd_SAP.CommandText = "INSERT INTO [Tirelli_40].[dbo].[Scheda_tecnica_record] (ID,COMMESSA,CAMPO,REPARTO,DIPENDENTE, DATA,ORA,CONTENUTO,CONTENUTO_1,ATTENDIBILITA,campo_1) VALUES(" & ID & ",'" & Label1.Text & "', '" & campo & "', '" & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto & "', '" & id_utente & "', getdate(),convert(varchar, getdate(), 108),'" & Replace(contenuto, "'", " ") & "','" & Replace(contenuto_1, "'", " ") & "','" & Attendibilità_info_popup.valore_attendibilità & "','" & campo_1 & "')"
                Cmd_SAP.ExecuteNonQuery()

                Cnn.Close()
            End If
            Ultimo_aggiornamento()
        End If


    End Sub


    Sub elimina_record()
        If inizializzazione = 1 Then
            If ComboBox_reparto.SelectedIndex < 0 Then


            ElseIf ComboBox_utente.SelectedIndex < 0 Then

            Else
                Dim Cnn As New SqlConnection
                Cnn.ConnectionString = Homepage.sap_tirelli


                Cnn.Open()

                Dim Cmd_SAP As New SqlCommand

                Cmd_SAP.Connection = Cnn
                Cmd_SAP.CommandText = "delete [Tirelli_40].[dbo].[Scheda_tecnica_record] where commessa='" & Label1.Text & "' and campo='" & campo_eliminazione & "'"
                Cmd_SAP.ExecuteNonQuery()

                Cnn.Close()
            End If
        End If
    End Sub



    Private Sub ComboBox1_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        campo = 1
        campo_1 = 0
        contenuto = ComboBox1.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Sub COMPILA_RECORD_INIZIALI()

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "Select t11.campo, t11.contenuto, t11.contenuto_1, t11.campo_1
from
(
SELECT max(t0.id) as 'ID', t0.campo,t0.campo_1, t0.commessa
FROM [Tirelli_40].[dbo].[Scheda_tecnica_record] t0
where t0.commessa='" & Label1.Text & "'
group by t0.campo, t0.commessa, t0.campo_1
)
as t10 inner join [Tirelli_40].[dbo].[Scheda_tecnica_record] t11 on t10.id=t11.id"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader

        Do While cmd_SAP_reader_2.Read()
            If cmd_SAP_reader_2("campo") = 0 Then

                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox13.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then
                    ComboBox11.Text = cmd_SAP_reader_2("contenuto_1")
                End If


            ElseIf cmd_SAP_reader_2("campo") = 1 Then
                ComboBox1.Text = cmd_SAP_reader_2("contenuto")
            ElseIf cmd_SAP_reader_2("campo") = 2 Then
                ComboBox2.Text = cmd_SAP_reader_2("contenuto")
            ElseIf cmd_SAP_reader_2("campo") = 3 Then

                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox3.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then
                    TextBox1.Text = cmd_SAP_reader_2("contenuto_1")

                End If

            ElseIf cmd_SAP_reader_2("campo") = 4 Then
                ComboBox4.Text = cmd_SAP_reader_2("contenuto")
                If cmd_SAP_reader_2("campo_1") = 1 Then
                    TextBox13.Text = cmd_SAP_reader_2("contenuto_1")
                End If
            ElseIf cmd_SAP_reader_2("campo") = 5 Then

                If cmd_SAP_reader_2("campo_1") = 1 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox24.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 2 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox25.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 3 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox26.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 4 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox27.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 5 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox28.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 6 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox29.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 7 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox30.Checked = True
                    End If

                End If





            ElseIf cmd_SAP_reader_2("campo") = 6 Then
                ComboBox6.Text = cmd_SAP_reader_2("contenuto")
            ElseIf cmd_SAP_reader_2("campo") = 7 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox7.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then
                    RichTextBox1.Text = cmd_SAP_reader_2("contenuto_1")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 8 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox8.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then
                    RichTextBox2.Text = cmd_SAP_reader_2("contenuto_1")
                End If
            ElseIf cmd_SAP_reader_2("campo") = 9 Then
                TextBox2.Text = cmd_SAP_reader_2("contenuto")
            ElseIf cmd_SAP_reader_2("campo") = 10 Then
                ComboBox9.Text = cmd_SAP_reader_2("contenuto")
            ElseIf cmd_SAP_reader_2("campo") = 11 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox10.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then
                    TextBox3.Text = cmd_SAP_reader_2("contenuto_1")
                End If
            ElseIf cmd_SAP_reader_2("campo") = 12 Then


            ElseIf cmd_SAP_reader_2("campo") = 13 Then
                ComboBox12.Text = cmd_SAP_reader_2("contenuto")

                'ElseIf cmd_SAP_reader_2("campo") = 14 Then

            ElseIf cmd_SAP_reader_2("campo") = 15 Then
                ComboBox12.Text = cmd_SAP_reader_2("contenuto")

            ElseIf cmd_SAP_reader_2("campo") = 16 Then

                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox14.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then
                    TextBox4.Text = cmd_SAP_reader_2("contenuto_1")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 17 Then

                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox15.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then

                End If


            ElseIf cmd_SAP_reader_2("campo") = 18 Then

                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox16.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then
                    TextBox5.Text = cmd_SAP_reader_2("contenuto_1")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 19 Then

                If cmd_SAP_reader_2("campo_1") = 1 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox43.Checked = True
                    End If
                End If
                If cmd_SAP_reader_2("campo_1") = 2 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox44.Checked = True
                    End If

                End If

                If cmd_SAP_reader_2("campo_1") = 3 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox45.Checked = True
                    End If

                End If
                If cmd_SAP_reader_2("campo_1") = 4 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox46.Checked = True
                    End If

                End If

                If cmd_SAP_reader_2("campo_1") = 5 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox47.Checked = True
                    End If

                End If

                If cmd_SAP_reader_2("campo_1") = 6 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox48.Checked = True
                    End If

                End If

                If cmd_SAP_reader_2("campo_1") = 7 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox49.Checked = True
                    End If

                End If

                If cmd_SAP_reader_2("campo_1") = 8 Then

                    TextBox6.Text = cmd_SAP_reader_2("contenuto_1")


                End If



            ElseIf cmd_SAP_reader_2("campo") = 20 Then

                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox22.Text = cmd_SAP_reader_2("contenuto")

                End If

            ElseIf cmd_SAP_reader_2("campo") = 53 Then


                TextBox12.Text = cmd_SAP_reader_2("contenuto")



            ElseIf cmd_SAP_reader_2("campo") = 1000 Then

                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox18.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then
                    ComboBox20.Text = cmd_SAP_reader_2("contenuto_1")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 1001 Then

                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox21.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then

                End If

            ElseIf cmd_SAP_reader_2("campo") = 1002 Then
                If cmd_SAP_reader_2("campo_1") = 1 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox5.Checked = True
                    End If
                End If
                If cmd_SAP_reader_2("campo_1") = 2 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox6.Checked = True
                    End If

                End If

                If cmd_SAP_reader_2("campo_1") = 3 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox7.Checked = True
                    End If

                End If

                If cmd_SAP_reader_2("campo_1") = 4 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox8.Checked = True
                    End If

                End If
            ElseIf cmd_SAP_reader_2("campo") = 2000 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox23.Text = cmd_SAP_reader_2("contenuto")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 2001 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox24.Text = cmd_SAP_reader_2("contenuto")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 2002 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox25.Text = cmd_SAP_reader_2("contenuto")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 2003 Then
                If cmd_SAP_reader_2("campo_1") = 1 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox1.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 2 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox2.Checked = True
                    End If
                End If

            ElseIf cmd_SAP_reader_2("campo") = 2004 Then
                If cmd_SAP_reader_2("campo_1") = 1 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox3.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 2 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox4.Checked = True
                    End If
                End If

            ElseIf cmd_SAP_reader_2("campo") = 2005 Then
                If cmd_SAP_reader_2("campo_1") = 1 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox9.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 2 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox10.Checked = True
                    End If
                ElseIf cmd_SAP_reader_2("campo_1") = 3 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox11.Checked = True
                    End If
                End If


            ElseIf cmd_SAP_reader_2("campo") = 2006 Then
                If cmd_SAP_reader_2("campo_1") = 1 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox12.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 2 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox13.Checked = True
                    End If

                End If

            ElseIf cmd_SAP_reader_2("campo") = 2007 Then
                If cmd_SAP_reader_2("campo_1") = 1 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox14.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 2 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox15.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 3 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox16.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 4 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox17.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 5 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox18.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 6 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox19.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 7 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox20.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 8 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox21.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 9 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox22.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 10 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox23.Checked = True
                    End If

                End If

            ElseIf cmd_SAP_reader_2("campo") = 3000 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox26.Text = cmd_SAP_reader_2("contenuto")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 3001 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox27.Text = cmd_SAP_reader_2("contenuto")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 3002 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox28.Text = cmd_SAP_reader_2("contenuto")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 3003 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox29.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then
                    TextBox7.Text = cmd_SAP_reader_2("contenuto_1")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 3004 Then

                If cmd_SAP_reader_2("campo_1") = 1 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox39.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 2 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox40.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 3 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox38.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 4 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox37.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 5 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox36.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 6 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox35.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 7 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox34.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 8 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox33.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 9 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox32.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 10 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox31.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 11 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox41.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 12 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox42.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 13 Then

                    RichTextBox4.Text = cmd_SAP_reader_2("contenuto_1")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 4000 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox30.Text = cmd_SAP_reader_2("contenuto")

                End If

            ElseIf cmd_SAP_reader_2("campo") = 4001 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox31.Text = cmd_SAP_reader_2("contenuto")

                End If

            ElseIf cmd_SAP_reader_2("campo") = 4002 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox32.Text = cmd_SAP_reader_2("contenuto")

                End If

            ElseIf cmd_SAP_reader_2("campo") = 4003 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    TextBox9.Text = cmd_SAP_reader_2("contenuto")

                End If

            ElseIf cmd_SAP_reader_2("campo") = 4004 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox37.Text = cmd_SAP_reader_2("contenuto")

                End If

            ElseIf cmd_SAP_reader_2("campo") = 5000 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    RichTextBox44.Text = cmd_SAP_reader_2("contenuto")

                End If

            ElseIf cmd_SAP_reader_2("campo") = 6000 Then

                If cmd_SAP_reader_2("campo_1") = 1 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox52.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 2 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox53.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 3 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox54.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 4 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox55.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 5 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox56.Checked = True
                    End If
                End If

            ElseIf cmd_SAP_reader_2("campo") = 6001 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox5.Text = cmd_SAP_reader_2("contenuto")

                End If

            ElseIf cmd_SAP_reader_2("campo") = 6002 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox19.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then
                    TextBox10.Text = cmd_SAP_reader_2("contenuto_1")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 6003 Then
                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox33.Text = cmd_SAP_reader_2("contenuto")
                ElseIf cmd_SAP_reader_2("campo_1") = 1 Then
                    TextBox11.Text = cmd_SAP_reader_2("contenuto_1")
                End If

            ElseIf cmd_SAP_reader_2("campo") = 6004 Then

                TextBox8.Text = cmd_SAP_reader_2("contenuto")



            ElseIf cmd_SAP_reader_2("campo") = 7000 Then

                ComboBox34.Text = cmd_SAP_reader_2("contenuto")

            ElseIf cmd_SAP_reader_2("campo") = 7001 Then

                ComboBox35.Text = cmd_SAP_reader_2("contenuto")

            ElseIf cmd_SAP_reader_2("campo") = 8001 Then



                If cmd_SAP_reader_2("campo_1") = 1 Then

                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox50.Checked = True
                    End If

                ElseIf cmd_SAP_reader_2("campo_1") = 2 Then
                    ComboBox36.Text = cmd_SAP_reader_2("contenuto_1")
                ElseIf cmd_SAP_reader_2("campo_1") = 3 Then
                    TextBox14.Text = cmd_SAP_reader_2("contenuto_1")
                End If



            ElseIf cmd_SAP_reader_2("campo") = 9001 Then

                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox30.Text = cmd_SAP_reader_2("contenuto")

                End If

            ElseIf cmd_SAP_reader_2("campo") = 9002 Then

                If cmd_SAP_reader_2("campo_1") = 0 Then
                    ComboBox39.Text = cmd_SAP_reader_2("contenuto")

                End If

            ElseIf cmd_SAP_reader_2("campo") = 9003 Then

                If cmd_SAP_reader_2("campo_1") = 1 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox51.Checked = True
                    End If
                End If
                If cmd_SAP_reader_2("campo_1") = 2 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox57.Checked = True
                    End If
                End If
                If cmd_SAP_reader_2("campo_1") = 3 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox58.Checked = True
                    End If
                End If
                If cmd_SAP_reader_2("campo_1") = 4 Then
                    If cmd_SAP_reader_2("contenuto_1") = "Y" Then
                        CheckBox59.Checked = True
                    End If
                    If cmd_SAP_reader_2("campo_1") = 5 Then

                        RichTextBox3.Text = cmd_SAP_reader_2("contenuto_1")

                    End If
                End If



            End If
        Loop
        cmd_SAP_reader_2.Close()
        Cnn1.Close()

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        campo = 2
        campo_1 = 0
        contenuto = ComboBox2.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged


        campo = 3
        If ComboBox3.Text = "Altro" Then
            contenuto = ComboBox3.Text
            TextBox1.Enabled = True
            Button2.Enabled = True
        Else
            contenuto = ComboBox3.Text
            contenuto_1 = Nothing
            TextBox1.Text = contenuto_1
            TextBox1.Enabled = False
            Button2.Enabled = False
        End If
        campo_1 = 0
        Inserisci_record()

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

        campo = 4
        If ComboBox4.Text = "Altro" Then
            contenuto = ComboBox4.Text
            TextBox13.Enabled = True
            Button23.Enabled = True
        Else
            contenuto = ComboBox4.Text
            contenuto_1 = Nothing
            TextBox13.Text = contenuto_1
            TextBox13.Enabled = False
            Button23.Enabled = False
        End If
        campo_1 = 0
        Inserisci_record()


    End Sub



    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        campo = 6
        campo_1 = 0
        contenuto = ComboBox6.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        'If ComboBox_reparto.SelectedIndex > 0 And ComboBox_utente.SelectedIndex > 0 Then
        campo = 7
        If ComboBox7.Text = "Altro" Then
            contenuto = ComboBox7.Text
            RichTextBox1.Enabled = True
            Button7.Enabled = True

        Else
            contenuto = ComboBox7.Text
            contenuto_1 = Nothing
            RichTextBox1.Text = contenuto_1
            RichTextBox1.Enabled = False
            Button7.Enabled = False
        End If
        campo_1 = 0
        Inserisci_record()
        'End If
    End Sub


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        campo = 3
        contenuto = ComboBox3.Text
        contenuto_1 = TextBox1.Text
        campo_1 = 1
        Inserisci_record()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        campo = 7
        contenuto = ComboBox7.Text
        contenuto_1 = RichTextBox1.Text
        campo_1 = 1
        Inserisci_record()

    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        campo = 8
        If ComboBox8.Text = "Altro" Then
            contenuto = ComboBox8.Text
            RichTextBox2.Enabled = True
            Button9.Enabled = True

        Else
            contenuto = ComboBox8.Text
            contenuto_1 = Nothing
            RichTextBox2.Text = contenuto_1
            RichTextBox2.Enabled = False
            Button9.Enabled = False
        End If
        campo_1 = 0
        Inserisci_record()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        campo = 8
        contenuto = ComboBox8.Text
        contenuto_1 = RichTextBox2.Text
        campo_1 = 1
        Inserisci_record()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click

        campo = 9
        campo_1 = 0
        contenuto = TextBox2.Text
        contenuto_1 = Nothing




        Inserisci_record()
        ' MsgBox("Campo aggiornato con successo")
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        campo = 10
        contenuto = ComboBox9.Text
        contenuto_1 = Nothing
        campo_1 = 0
        Inserisci_record()

    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged


        campo = 11
        If ComboBox10.Text = "Altro" Then
            contenuto = ComboBox10.Text
            TextBox3.Enabled = True
            Button11.Enabled = True
        Else
            contenuto = ComboBox10.Text
            contenuto_1 = Nothing
            TextBox3.Text = contenuto_1
            TextBox3.Enabled = False
            Button11.Enabled = False
        End If
        campo_1 = 0
        Inserisci_record()


    End Sub


    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        campo = 11
        contenuto = ComboBox10.Text
        contenuto_1 = TextBox3.Text
        campo_1 = 1

        Inserisci_record()
    End Sub



    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs)
        campo = 13
        campo_1 = 0
        contenuto = ComboBox12.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub







    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedIndexChanged




        campo = 0
        campo_1 = 1
        contenuto = Nothing
        contenuto_1 = ComboBox11.Text
        Inserisci_record()
        GroupBox35.Visible = True

        GroupBox35.Text = "N° Rubinetti"
        ComboBox12.Items.Clear()
        ComboBox12.Items.Add("-")
        ComboBox12.Items.Add("2")
        ComboBox12.Items.Add("3")
        ComboBox12.Items.Add("4")
        ComboBox12.Items.Add("6")
        ComboBox12.Items.Add("8")
        ComboBox12.Items.Add("10")
        ComboBox12.Items.Add("12")
        ComboBox12.Items.Add("14")
        ComboBox12.Items.Add("16")
        ComboBox12.Items.Add("20")
        ComboBox12.Items.Add("30")
        ComboBox12.Items.Add("36")
        ComboBox12.Items.Add("40")

        GroupBox37.Visible = True
        GroupBox37.Text = "Tipologia"
        ComboBox14.Items.Clear()
        ComboBox14.Items.Add("-")
        ComboBox14.Items.Add("Pompe peristaltiche ")
        ComboBox14.Items.Add("Pistoni")
        ComboBox14.Items.Add("Flow meter massico")
        ComboBox14.Items.Add("Flow meter magnetico")
        ComboBox14.Items.Add("Vuoto")

        GroupBox38.Visible = True
        GroupBox38.Text = "Carrello di riempimento"
        ComboBox15.Items.Clear()
        ComboBox15.Items.Add("-")
        ComboBox15.Items.Add("Si")
        ComboBox15.Items.Add("No")

        GroupBox39.Visible = True
        GroupBox39.Text = "Carrello di lavaggio"
        ComboBox16.Items.Clear()
        ComboBox16.Items.Add("-")
        ComboBox16.Items.Add("Si")
        ComboBox16.Items.Add("No")
        ComboBox16.Items.Add("altro")

        GroupBox63.Visible = True




        GroupBox44.Visible = True
        GroupBox44.Text = "Movimentazione rubinetti"
        ComboBox22.Items.Clear()
        ComboBox22.Items.Add("-")
        ComboBox22.Items.Add("Inseguimento")
        ComboBox22.Items.Add("Fissa")


    End Sub

    Private Sub ComboBox13_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox13.SelectedIndexChanged
        campo = 0
        campo_1 = Nothing
        contenuto = ComboBox13.Text
        contenuto_1 = Nothing
        ComboBox11.Text = ""

        Inserisci_record()
        ComboBox11.Items.Clear()
        GroupBox34.Visible = True
        If ComboBox13.Text = "Accessorio" Then
            ComboBox11.Items.Add("Controllo peso")
            ComboBox11.Items.Add("Estrattore flaconi")
            ComboBox11.Items.Add("Nastri di raffreddamento")
            ComboBox11.Items.Add("Piatto di Alimentazione")
            ComboBox11.Items.Add("Piatto di Raccolta")
            ComboBox11.Items.Add("Sistema di alimentazione automatica con elevatore")
            ComboBox11.Items.Add("Telaio INKJET")
        ElseIf ComboBox13.Text = "Riempimento" Then
            ComboBox11.Items.Add("Riempitrice")
            ComboBox11.Items.Add("Dosatore da banco")
            ComboBox11.Items.Add("Dosatore a terra")
        ElseIf ComboBox13.Text = "Tappatura" Then
            ComboBox11.Items.Add("Tappatore rotativo (Ro)")
            ComboBox11.Items.Add("Tappatore lineare")
            ComboBox11.Items.Add("Tappatore a inseguimento")
        ElseIf ComboBox13.Text = "Etichettatura" Then
            ComboBox11.Items.Add("Etichettatrice Linare")
            ComboBox11.Items.Add("Etichettatrice Rotativa")
        ElseIf ComboBox13.Text = "Termosaldatura" Then
            ComboBox11.Items.Add("Termosaldatrice Lineare")
            ComboBox11.Items.Add("Termosaldatrice rotativa")

        ElseIf ComboBox13.Text = "Macchina combinata" Then
            ComboBox11.Items.Add("Monoblocco Sigma")
            ComboBox11.Items.Add("Monoblocco Maxima")
            ComboBox11.Items.Add("Oscar RO")
            ComboBox11.Items.Add("Miniblocco")

        End If
    End Sub
    Private Sub TABPAGE6_Click(sender As Object, e As EventArgs) Handles TabPage6.Enter

        Scheda_tecnica.trova_gruppi_etichettaggio(DataGridView4, Label1.Text)
    End Sub
    Private Sub ComboBox12_SelectedIndexChanged_2(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        campo = 15
        campo_1 = 0
        contenuto = ComboBox12.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub ComboBox14_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox14.SelectedIndexChanged
        campo = 16
        campo_1 = 0
        contenuto = ComboBox14.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub Button12_Click_1(sender As Object, e As EventArgs) Handles Button12.Click
        campo = 16
        campo_1 = 1
        contenuto = Nothing
        contenuto_1 = TextBox4.Text
        Inserisci_record()
        MsgBox("Campo aggiornato con successo")
    End Sub

    Private Sub ComboBox15_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles ComboBox15.SelectedIndexChanged
        campo = 17
        campo_1 = 0
        contenuto = ComboBox15.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub ComboBox16_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox16.SelectedIndexChanged
        campo = 18
        campo_1 = 0
        contenuto = ComboBox16.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        campo = 18
        campo_1 = 1
        contenuto = Nothing
        contenuto_1 = TextBox5.Text
        Inserisci_record()
        MsgBox("Campo aggiornato con successo")
    End Sub



    Private Sub ComboBox17_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox17.SelectedIndexChanged

    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        campo = 19
        campo_1 = 8
        contenuto = Nothing
        contenuto_1 = TextBox6.Text
        Inserisci_record()
    End Sub

    Private Sub ComboBox18_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox18.SelectedIndexChanged
        campo = 1000
        campo_1 = Nothing
        contenuto = ComboBox18.Text
        contenuto_1 = Nothing


        Inserisci_record()
        GroupBox42.Visible = True
        GroupBox43.Visible = True
        CheckBox5.Visible = True
        CheckBox6.Visible = True
        CheckBox7.Visible = True
        CheckBox8.Visible = True
        ComboBox20.Items.Clear()

        If ComboBox18.Text = "Accessorio" Then
            ComboBox20.Items.Add("Controllo peso")
            ComboBox20.Items.Add("Elevatore")
            ComboBox20.Items.Add("Estrattore flaconi")
            ComboBox20.Items.Add("Nastri di raffreddamento")
            ComboBox20.Items.Add("Piatto di Alimentazione")
            ComboBox20.Items.Add("Piatto di Raccolta")
            ComboBox20.Items.Add("Sistema di alimentazione automatica con elevatore")
            ComboBox20.Items.Add("Telaio INKJET")
        End If
    End Sub

    Private Sub ComboBox20_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox20.SelectedIndexChanged
        CheckBox5.Checked = False
        CheckBox6.Checked = False
        CheckBox7.Checked = False
        CheckBox8.Checked = False

        CheckBox5.Visible = False
        CheckBox6.Visible = False
        CheckBox7.Visible = False
        CheckBox8.Visible = False



        campo = 1000
        campo_1 = 1
        contenuto = Nothing
        contenuto_1 = ComboBox20.Text
        Inserisci_record()


        If ComboBox20.Text = "Piatto di Alimentazione" Then
            GroupBox43.Visible = True
            GroupBox43.Text = "Diametro piatto"
            ComboBox21.Items.Clear()
            ComboBox21.Items.Add("-")
            ComboBox21.Items.Add("Diam. 1000")
            ComboBox21.Items.Add("Diam. 1200")
            ComboBox21.Items.Add("Diam. 1500")
        ElseIf ComboBox20.Text = "Piatto di Raccolta" Then
            GroupBox43.Visible = True
            GroupBox43.Text = "Diametro piatto"
            ComboBox21.Items.Clear()
            ComboBox21.Items.Add("-")
            ComboBox21.Items.Add("Diam. 1000")
            ComboBox21.Items.Add("Diam. 1200")
            ComboBox21.Items.Add("Diam. 1500")
        ElseIf ComboBox20.Text = "Nastri di raffreddamento" Then
            GroupBox43.Visible = True
            GroupBox43.Text = "Tipo di nastri"
            ComboBox21.Items.Clear()
            ComboBox21.Items.Add("-")
            ComboBox21.Items.Add("Sistema di nastri")
            ComboBox21.Items.Add("Camera coibentata")
            ComboBox21.Items.Add("Raffreddamento forzato con Chiller ")
        ElseIf ComboBox20.Text = "Controllo peso" Then
            GroupBox43.Visible = False
            CheckBox5.Visible = True
            CheckBox5.Text = "Nastrini"
            CheckBox6.Visible = True
            CheckBox6.Text = "Scarto"


        ElseIf ComboBox20.Text = "Sistema di alimentazione automatica con elevatore" Then
            GroupBox43.Visible = False
            CheckBox5.Visible = True
            CheckBox5.Text = "Tazza vibrante"
            CheckBox6.Visible = True
            CheckBox6.Text = "Sorter"
            CheckBox7.Visible = True
            CheckBox7.Text = "Alimentatore Trigger"
            CheckBox8.Visible = True
            CheckBox8.Text = "Elevatore"

        End If
    End Sub

    Private Sub ComboBox21_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox21.SelectedIndexChanged
        campo = 1001
        campo_1 = 0
        contenuto = ComboBox21.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub


    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        campo = 1002
        campo_1 = 1
        contenuto = Nothing
        If CheckBox5.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub


    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged


        campo = 1002
        campo_1 = 2
        contenuto = Nothing
        If CheckBox6.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub


    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        campo = 1002
        campo_1 = 3
        contenuto = Nothing
        If CheckBox7.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        campo = 1002
        campo_1 = 4
        contenuto = Nothing
        If CheckBox8.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub ComboBox22_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox22.SelectedIndexChanged
        campo = 20
        campo_1 = 0
        contenuto = ComboBox22.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub ComboBox23_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox23.SelectedIndexChanged
        campo = 2000
        campo_1 = Nothing
        contenuto = ComboBox23.Text
        contenuto_1 = Nothing
        ComboBox24.Text = ""

        Inserisci_record()
        ComboBox24.Items.Clear()
        ComboBox25.Items.Clear()

        GroupBox48.Visible = False
        GroupBox49.Visible = False
        GroupBox50.Visible = False
        GroupBox51.Visible = False
        GroupBox52.Visible = False

        If ComboBox23.Text = "Tappatore lineare" Then
            If CheckBox1.Checked = True Then
                CheckBox1.Checked = False
            End If
            If CheckBox2.Checked = True Then
                CheckBox2.Checked = False
            End If
            If CheckBox3.Checked = True Then
                CheckBox3.Checked = False
            End If
            If CheckBox4.Checked = True Then
                CheckBox4.Checked = False
            End If

            If CheckBox9.Checked = True Then
                CheckBox9.Checked = False
            End If
            If CheckBox10.Checked = True Then
                CheckBox10.Checked = False
            End If
            If CheckBox11.Checked = True Then
                CheckBox11.Checked = False
            End If

            If CheckBox12.Checked = True Then
                CheckBox12.Checked = False
            End If

            If CheckBox13.Checked = True Then
                CheckBox13.Checked = False
            End If


            GroupBox46.Visible = True
            GroupBox46.Text = "Sottotipo di macchina"
            ComboBox24.Items.Add("-")
            ComboBox24.Items.Add("Mininebel")
            ComboBox24.Items.Add("Nebel")
            ComboBox24.Items.Add("RO 1 E")
            ComboBox24.Items.Add("RO 1 E S")
            ComboBox24.Items.Add("RO 1 E C")

            GroupBox47.Visible = True
            GroupBox47.Text = "Numero di teste"
            ComboBox25.Items.Add("-")
            ComboBox25.Items.Add("1")
            ComboBox25.Items.Add("2")
            ComboBox25.Items.Add("3")

            GroupBox52.Visible = True
            GroupBox52.Text = "Caricamento manuale"
            CheckBox14.Text = "Caricamento automatico"
            CheckBox15.Text = "Centratori"
            CheckBox16.Text = "Testa bordatrice"
            CheckBox17.Text = "Testa trigger"
            CheckBox18.Text = "Blocco ingresso"
            CheckBox19.Text = "Antirotazione"
            CheckBox20.Text = "Pressetta"
            CheckBox21.Text = "Sistema di termosaldatura"
            CheckBox22.Text = "Con Stella"
            CheckBox23.Text = "a Inseguimento"



        ElseIf ComboBox23.Text = "Tappatore rotativo (Ro)" Then

            If CheckBox1.Checked = True Then
                CheckBox1.Checked = False
            End If
            If CheckBox2.Checked = True Then
                CheckBox2.Checked = False
            End If
            If CheckBox3.Checked = True Then
                CheckBox3.Checked = False
            End If
            If CheckBox4.Checked = True Then
                CheckBox4.Checked = False
            End If


            If CheckBox14.Checked = True Then
                CheckBox14.Checked = False
            End If
            If CheckBox15.Checked = True Then
                CheckBox15.Checked = False
            End If
            If CheckBox16.Checked = True Then
                CheckBox16.Checked = False
            End If
            If CheckBox17.Checked = True Then
                CheckBox17.Checked = False
            End If

            If CheckBox18.Checked = True Then
                CheckBox18.Checked = False
            End If
            If CheckBox19.Checked = True Then
                CheckBox19.Checked = False
            End If
            If CheckBox20.Checked = True Then
                CheckBox20.Checked = False
            End If

            If CheckBox21.Checked = True Then
                CheckBox21.Checked = False
            End If

            If CheckBox22.Checked = True Then
                CheckBox22.Checked = False
            End If

            If CheckBox23.Checked = True Then
                CheckBox23.Checked = False
            End If


            GroupBox48.Visible = False
            GroupBox49.Visible = False

            GroupBox46.Visible = True
            GroupBox46.Text = "Tipo di azionamento"
            ComboBox24.Items.Add("-")
            ComboBox24.Items.Add("Meccanico")
            ComboBox24.Items.Add("Elettronico")
            ComboBox24.Items.Add("Camma virtuale")

            GroupBox47.Visible = True
            GroupBox47.Text = "Numero di teste"
            ComboBox25.Items.Add("-")
            ComboBox25.Items.Add("3")
            ComboBox25.Items.Add("4")
            ComboBox25.Items.Add("6")
            ComboBox25.Items.Add("8")
            ComboBox25.Items.Add("10")
            ComboBox25.Items.Add("12")
            ComboBox25.Items.Add("16")
            ComboBox25.Items.Add("20")

            GroupBox50.Visible = True
            GroupBox50.Text = "Oggetto trattato"
            CheckBox9.Text = "Tappo"
            CheckBox10.Text = "Sottotappo"
            CheckBox11.Text = "Pompetta"

            GroupBox51.Visible = True
            GroupBox51.Text = "Tipologia stiratura"
            CheckBox12.Text = "Stella di stiratura pompetta elettronica"
            CheckBox13.Text = "Stella di stiratura pompetta pneumatica"




        ElseIf ComboBox23.Text = "Tappatore a inseguimento" Then
            If CheckBox9.Checked = True Then
                CheckBox9.Checked = False
            End If
            If CheckBox10.Checked = True Then
                CheckBox10.Checked = False
            End If
            If CheckBox11.Checked = True Then
                CheckBox11.Checked = False
            End If

            If CheckBox12.Checked = True Then
                CheckBox12.Checked = False
            End If
            If CheckBox13.Checked = True Then
                CheckBox13.Checked = False
            End If

            If CheckBox14.Checked = True Then
                CheckBox14.Checked = False
            End If
            If CheckBox15.Checked = True Then
                CheckBox15.Checked = False
            End If
            If CheckBox16.Checked = True Then
                CheckBox16.Checked = False
            End If
            If CheckBox17.Checked = True Then
                CheckBox17.Checked = False
            End If

            If CheckBox18.Checked = True Then
                CheckBox18.Checked = False
            End If
            If CheckBox19.Checked = True Then
                CheckBox19.Checked = False
            End If
            If CheckBox20.Checked = True Then
                CheckBox20.Checked = False
            End If

            If CheckBox21.Checked = True Then
                CheckBox21.Checked = False
            End If

            If CheckBox22.Checked = True Then
                CheckBox22.Checked = False
            End If

            If CheckBox23.Checked = True Then
                CheckBox23.Checked = False
            End If


            GroupBox46.Visible = True
            GroupBox46.Text = "Tipologia di oscar"
            ComboBox24.Items.Add("-")
            ComboBox24.Items.Add("Oscar 13 azionamenti")
            ComboBox24.Items.Add("Oscar 19 azionamenti")

            GroupBox47.Visible = True
            GroupBox47.Text = "Numero di teste tappanti"
            ComboBox25.Items.Add("-")
            ComboBox25.Items.Add("1")
            ComboBox25.Items.Add("2")
            ComboBox25.Items.Add("3")

            GroupBox48.Visible = True
            GroupBox48.Text = "Oggetto trattato"
            CheckBox1.Visible = True
            CheckBox2.Visible = True

            CheckBox1.Text = "Trigger"
            CheckBox2.Text = "Pompetta"

            GroupBox49.Visible = True
            GroupBox49.Text = "Tappo di chiusura"
            CheckBox3.Visible = True
            CheckBox4.Visible = True

            CheckBox3.Text = "A vite"
            CheckBox4.Text = "A pressione"


        End If
    End Sub


    Private Sub ComboBox24_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox24.SelectedIndexChanged
        campo = 2001
        campo_1 = 0
        contenuto = ComboBox24.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub ComboBox25_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox25.SelectedIndexChanged
        campo = 2002
        campo_1 = 0
        contenuto = ComboBox25.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        campo = 2003
        campo_1 = 1
        contenuto = Nothing
        If CheckBox1.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        campo = 2003
        campo_1 = 2
        contenuto = Nothing
        If CheckBox2.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        campo = 2004
        campo_1 = 1
        contenuto = Nothing
        If CheckBox3.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        campo = 2004
        campo_1 = 2
        contenuto = Nothing
        If CheckBox4.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        campo = 2005
        campo_1 = 1
        contenuto = Nothing
        If CheckBox9.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        campo = 2005
        campo_1 = 2
        contenuto = Nothing
        If CheckBox10.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub



    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        campo = 2005
        campo_1 = 3
        contenuto = Nothing
        If CheckBox11.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        campo = 2006
        campo_1 = 1
        contenuto = Nothing
        If CheckBox12.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        campo = 2006
        campo_1 = 2
        contenuto = Nothing
        If CheckBox13.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        campo = 2007
        campo_1 = 1
        contenuto = Nothing
        If CheckBox14.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        campo = 2007
        campo_1 = 2
        contenuto = Nothing
        If CheckBox15.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        campo = 2007
        campo_1 = 3
        contenuto = Nothing
        If CheckBox16.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox17_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox17.CheckedChanged
        campo = 2007
        campo_1 = 4
        contenuto = Nothing
        If CheckBox17.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox18_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox18.CheckedChanged
        campo = 2007
        campo_1 = 5
        contenuto = Nothing
        If CheckBox18.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox19_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox19.CheckedChanged
        campo = 2007
        campo_1 = 6
        contenuto = Nothing
        If CheckBox19.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox20_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox20.CheckedChanged
        campo = 2007
        campo_1 = 7
        contenuto = Nothing
        If CheckBox20.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox21_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox21.CheckedChanged
        campo = 2007
        campo_1 = 8
        contenuto = Nothing
        If CheckBox21.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox22_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox22.CheckedChanged
        campo = 2007
        campo_1 = 9
        contenuto = Nothing
        If CheckBox22.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox23_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox23.CheckedChanged
        campo = 2007
        campo_1 = 10
        contenuto = Nothing
        If CheckBox23.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub ComboBox26_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox26.SelectedIndexChanged
        campo = 3000
        campo_1 = Nothing
        contenuto = ComboBox26.Text
        contenuto_1 = Nothing
        ComboBox27.Text = ""

        Inserisci_record()
        ComboBox27.Items.Clear()
        ComboBox29.Items.Clear()

        GroupBox54.Visible = True
        GroupBox56.Visible = True
        GroupBox62.Visible = True

        ComboBox29.Items.Add("Avery")
        ComboBox29.Items.Add("Altro")

        If ComboBox26.Text = "Etichettatrice Linare" Then

            ComboBox27.Items.Add("-")
            ComboBox27.Items.Add("Miniecho")
            ComboBox27.Items.Add("Delta")
            ComboBox27.Items.Add("Bravo")

        ElseIf ComboBox26.Text = "Etichettatrice Rotativa" Then
            ComboBox27.Items.Add("-")
            ComboBox27.Items.Add("Tango")

        End If



    End Sub

    Private Sub ComboBox27_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox27.SelectedIndexChanged

        campo = 3001
        campo_1 = Nothing
        contenuto = ComboBox27.Text
        contenuto_1 = Nothing
        ComboBox28.Text = ""

        Inserisci_record()
        ComboBox28.Items.Clear()


        GroupBox55.Visible = True


        If ComboBox27.Text = "Miniecho" Then


            ComboBox28.Items.Add("-")
            ComboBox28.Items.Add("1")

        ElseIf ComboBox27.Text = "Delta" Then
            ComboBox28.Items.Add("-")
            ComboBox28.Items.Add("1")
            ComboBox28.Items.Add("2")

        ElseIf ComboBox27.Text = "Bravo" Then
            ComboBox28.Items.Add("-")
            ComboBox28.Items.Add("1 sopra/sotto")

        ElseIf ComboBox27.Text = "Tango" Then
            ComboBox28.Items.Add("-")
            ComboBox28.Items.Add("1")
            ComboBox28.Items.Add("2")
            ComboBox28.Items.Add("3")
            ComboBox28.Items.Add("4")


        End If
    End Sub

    Private Sub ComboBox28_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox28.SelectedIndexChanged
        campo = 3002
        campo_1 = Nothing
        contenuto = ComboBox28.Text
        contenuto_1 = Nothing


        Inserisci_record()
    End Sub

    Private Sub ComboBox29_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox29.SelectedIndexChanged
        campo = 3003
        campo_1 = Nothing
        contenuto = ComboBox29.Text
        contenuto_1 = Nothing


        Inserisci_record()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        campo = 3003
        campo_1 = 1
        contenuto = Nothing
        contenuto_1 = TextBox7.Text
        Inserisci_record()
        MsgBox("Campo aggiornato con successo")
    End Sub

    Private Sub ComboBox30_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox30.SelectedIndexChanged
        campo = 4000
        campo_1 = Nothing
        contenuto = ComboBox30.Text
        contenuto_1 = Nothing

        Inserisci_record()


        GroupBox58.Visible = True
        GroupBox59.Visible = True
        GroupBox60.Visible = True

        ComboBox31.Items.Clear()
        ComboBox31.Items.Add("-")
        ComboBox31.Items.Add("1")
        ComboBox31.Items.Add("2")
        ComboBox31.Items.Add("3")
        ComboBox31.Items.Add("4")

        ComboBox32.Items.Clear()
        ComboBox32.Items.Add("-")
        ComboBox32.Items.Add("Bobina")
        ComboBox32.Items.Add("Opercoli")
        ComboBox32.Items.Add("Prefustellati")

    End Sub

    Private Sub ComboBox31_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox31.SelectedIndexChanged
        campo = 4001
        campo_1 = Nothing
        contenuto = ComboBox31.Text
        contenuto_1 = Nothing

        Inserisci_record()
    End Sub

    Private Sub ComboBox32_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox32.SelectedIndexChanged
        campo = 4002
        campo_1 = Nothing
        contenuto = ComboBox31.Text
        contenuto_1 = Nothing

        Inserisci_record()
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click
        campo = 4003
        campo_1 = Nothing
        contenuto = TextBox9.Text
        contenuto_1 = Nothing
        Inserisci_record()
        MsgBox("Campo aggiornato con successo")
    End Sub





    Private Sub CheckBox24_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox24.CheckedChanged

        campo = 5
        campo_1 = 1
        contenuto = Nothing
        If CheckBox24.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox25_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox25.CheckedChanged
        campo = 5
        campo_1 = 2
        contenuto = Nothing
        If CheckBox25.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox26_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox26.CheckedChanged
        campo = 5
        campo_1 = 3
        contenuto = Nothing
        If CheckBox26.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox27_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox27.CheckedChanged
        campo = 5
        campo_1 = 4
        contenuto = Nothing
        If CheckBox27.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox28_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox28.CheckedChanged
        campo = 5
        campo_1 = 5
        contenuto = Nothing
        If CheckBox28.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox29_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox29.CheckedChanged
        campo = 5
        campo_1 = 6
        contenuto = Nothing
        If CheckBox29.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox30_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox30.CheckedChanged
        campo = 5
        campo_1 = 7
        contenuto = Nothing
        If CheckBox30.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub



    Private Sub CheckBox39_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox39.CheckedChanged
        campo = 3004
        campo_1 = 1
        contenuto = Nothing
        If CheckBox39.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox40_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox40.CheckedChanged
        campo = 3004
        campo_1 = 2
        contenuto = Nothing
        If CheckBox40.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox38_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox38.CheckedChanged
        campo = 3004
        campo_1 = 3
        contenuto = Nothing
        If CheckBox38.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox37_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox37.CheckedChanged
        campo = 3004
        campo_1 = 4
        contenuto = Nothing
        If CheckBox37.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox36_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox36.CheckedChanged
        campo = 3004
        campo_1 = 5
        contenuto = Nothing
        If CheckBox36.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox35_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox35.CheckedChanged
        campo = 3004
        campo_1 = 6
        contenuto = Nothing
        If CheckBox35.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox34_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox34.CheckedChanged
        campo = 3004
        campo_1 = 7
        contenuto = Nothing
        If CheckBox34.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox33_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox33.CheckedChanged
        campo = 3004
        campo_1 = 8
        contenuto = Nothing
        If CheckBox33.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox32_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox32.CheckedChanged
        campo = 3004
        campo_1 = 9
        contenuto = Nothing
        If CheckBox32.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox31_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox31.CheckedChanged
        campo = 3004
        campo_1 = 10
        contenuto = Nothing
        If CheckBox31.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox41_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox41.CheckedChanged
        campo = 3004
        campo_1 = 11
        contenuto = Nothing
        If CheckBox41.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox42_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox42.CheckedChanged
        campo = 3004
        campo_1 = 12
        contenuto = Nothing
        If CheckBox42.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        campo = 3004
        campo_1 = 13
        contenuto = Nothing
        contenuto_1 = RichTextBox4.Text
        Inserisci_record()
        MsgBox("Campo aggiornato con successo")
    End Sub

    Private Sub CheckBox43_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox43.CheckedChanged
        campo = 19
        campo_1 = 1
        contenuto = Nothing
        If CheckBox43.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox44_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox44.CheckedChanged
        campo = 19
        campo_1 = 2
        contenuto = Nothing
        If CheckBox44.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox45_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox45.CheckedChanged
        campo = 19
        campo_1 = 3
        contenuto = Nothing
        If CheckBox45.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox46_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox46.CheckedChanged
        campo = 19
        campo_1 = 4
        contenuto = Nothing
        If CheckBox46.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox47_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox47.CheckedChanged
        campo = 19
        campo_1 = 5
        contenuto = Nothing
        If CheckBox47.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox48_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox48.CheckedChanged
        campo = 19
        campo_1 = 6
        contenuto = Nothing
        If CheckBox48.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox49_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox49.CheckedChanged
        campo = 19
        campo_1 = 7
        contenuto = Nothing
        If CheckBox49.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox52_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox52.CheckedChanged
        campo = 6000
        campo_1 = 1
        contenuto = Nothing
        If CheckBox52.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox53_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox53.CheckedChanged
        campo = 6000
        campo_1 = 2
        contenuto = Nothing
        If CheckBox53.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox54_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox54.CheckedChanged
        campo = 6000
        campo_1 = 3
        contenuto = Nothing
        If CheckBox54.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox55_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox55.CheckedChanged
        campo = 6000
        campo_1 = 4
        contenuto = Nothing
        If CheckBox55.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox56_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox56.CheckedChanged
        campo = 6000
        campo_1 = 5
        contenuto = Nothing
        If CheckBox56.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        campo = 6001
        campo_1 = 0
        contenuto = ComboBox5.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub ComboBox19_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox19.SelectedIndexChanged
        campo = 6002
        campo_1 = 0
        contenuto = ComboBox19.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub Button20_Click(sender As Object, e As EventArgs)
        campo = 6002
        contenuto = Nothing
        contenuto_1 = TextBox10.Text
        campo_1 = 1
        Inserisci_record()
    End Sub

    Private Sub ComboBox33_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox33.SelectedIndexChanged
        campo = 6003
        campo_1 = 0
        contenuto = ComboBox33.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs)
        campo = 6003
        contenuto = Nothing
        contenuto_1 = TextBox11.Text
        campo_1 = 1
        Inserisci_record()
    End Sub

    Private Sub ComboBox34_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox34.SelectedIndexChanged
        campo = 7000
        campo_1 = 0
        contenuto = ComboBox34.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub

    Private Sub ComboBox35_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox35.SelectedIndexChanged
        campo = 7001
        campo_1 = 0
        contenuto = ComboBox35.Text
        contenuto_1 = Nothing
        Inserisci_record()
    End Sub









    Private Sub Button17_Click_1(sender As Object, e As EventArgs) Handles Button17.Click
        campo = 5000
        campo_1 = Nothing
        contenuto = RichTextBox44.Text
        contenuto_1 = Nothing
        Inserisci_record()
        MsgBox("Campo aggiornato con successo")
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub Button18_Click(sender As Object, e As EventArgs)
        campo = 6004
        campo_1 = Nothing
        contenuto = TextBox8.Text
        contenuto_1 = Nothing
        Inserisci_record()
        MsgBox("Campo aggiornato con successo")
    End Sub

    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click

        campo = 53
        campo_1 = Nothing
        contenuto = TextBox12.Text
        contenuto_1 = Nothing

        Inserisci_record()
        MsgBox("Campo aggiornato con successo")
    End Sub

    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        campo = 4
        contenuto = ComboBox4.Text
        contenuto_1 = TextBox13.Text
        campo_1 = 1
        Inserisci_record()
    End Sub



    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click


        trova_cartella_macchina()


    End Sub

    Sub trova_cartella_macchina()
        Dim cartella_padre As String
        Dim cliente As String
        Dim cartella_esistente As String = ""
        If Label4.Text = "" Then
            cliente = Label3.Text
        Else
            cliente = Label4.Text
        End If

        Try


            Dim rootDirectory As String = Homepage.percorso_cartelle_macchine

            Dim directories As String() = System.IO.Directory.GetDirectories(rootDirectory, Strings.Left(commessa, 4) & "*")

            For Each directory As String In directories
                cartella_padre = directory
            Next
            Dim sottocartella As String
            sottocartella = Replace(cartella_padre, rootDirectory, "")

            rootDirectory = cartella_padre

            directories = System.IO.Directory.GetDirectories(rootDirectory, commessa & "*")

            For Each directory As String In directories
                cartella_esistente = directory
            Next


            If cartella_esistente = "" Then




                Directory.CreateDirectory(cartella_padre & "\" & commessa & " " & Label2.Text & " - " & cliente)
                LinkLabel2.Text = sottocartella & "\" & commessa & " " & Label2.Text & " - " & cliente

                Scheda_tecnica.Aggiorna_percorso_macchina(LinkLabel2.Text, Label1.Text, "COMMESSA")
                cerca_file()


            Else

                ' LinkLabel2.Text = cartella_esistente
                LinkLabel2.Text = Replace(cartella_esistente, Homepage.percorso_cartelle_macchine, "")
                Scheda_tecnica.Aggiorna_percorso_macchina(LinkLabel2.Text, Label1.Text, "COMMESSA")
                cerca_file()

            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub CheckBox50_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox50.CheckedChanged
        campo = 8001
        campo_1 = 1
        contenuto = Nothing
        If CheckBox50.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub ComboBox36_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox36.SelectedIndexChanged
        campo = 8001
        campo_1 = 2
        contenuto = Nothing
        contenuto_1 = ComboBox36.Text
        Inserisci_record()
    End Sub

    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        campo = 8001
        contenuto = Nothing
        contenuto_1 = TextBox14.Text
        campo_1 = 3
        Inserisci_record()
    End Sub

    Private Sub ComboBox37_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox37.SelectedIndexChanged
        campo = 4004
        campo_1 = Nothing
        contenuto = ComboBox37.Text
        contenuto_1 = Nothing

        Inserisci_record()
    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)

    End Sub

    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        If Button26.Text <> 0 Then




            Progetto.Show()
            Progetto.BringToFront()
            Progetto.absentry = Button26.Text
            Progetto.inizializza_progetto()

        Else
            MsgBox("Nessun progetto è assegnato a questa commessa")

        End If
    End Sub

    Private Sub ComboBox38_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox38.SelectedIndexChanged
        campo = 9001
        campo_1 = Nothing
        contenuto = ComboBox38.Text
        contenuto_1 = Nothing

        Inserisci_record()
    End Sub

    Private Sub ComboBox39_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox39.SelectedIndexChanged
        campo = 9002
        campo_1 = Nothing
        contenuto = ComboBox39.Text
        contenuto_1 = Nothing

        Inserisci_record()
    End Sub

    Private Sub CheckBox51_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox51.CheckedChanged
        campo = 9003
        campo_1 = 1
        contenuto = Nothing
        If CheckBox51.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox57_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox57.CheckedChanged
        campo = 9003
        campo_1 = 2
        contenuto = Nothing
        If CheckBox57.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox58_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox58.CheckedChanged
        campo = 9003
        campo_1 = 3
        contenuto = Nothing
        If CheckBox58.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub CheckBox59_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox59.CheckedChanged
        campo = 9003
        campo_1 = 4
        contenuto = Nothing
        If CheckBox59.Checked = True Then
            contenuto_1 = "Y"
        Else
            contenuto_1 = "N"
        End If

        Inserisci_record()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        campo = 9003
        campo_1 = 5
        contenuto = Nothing
        contenuto_1 = RichTextBox3.Text
        Inserisci_record()
    End Sub

    Private Sub GroupBox44_Enter(sender As Object, e As EventArgs) Handles GroupBox44.Enter

    End Sub

    Private Sub DataGridView1_CellContextMenuStripChanged(sender As Object, e As DataGridViewCellEventArgs)

    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Form_Gruppi_Etichettaggio.commessa = Label1.Text
        Form_Gruppi_Etichettaggio.stato_gruppo = "Nuovo"
        Form_Gruppi_Etichettaggio.inizializza_form()
        Form_Gruppi_Etichettaggio.Show()

    End Sub

    Private Sub DataGridView4_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellContentClick

    End Sub

    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        If e.ColumnIndex = DataGridView4.Columns.IndexOf(N_) Then
            Form_Gruppi_Etichettaggio.commessa = Label1.Text
            Form_Gruppi_Etichettaggio.id = DataGridView4.Rows(e.RowIndex).Cells(columnName:="ID_").Value
            Form_Gruppi_Etichettaggio.N = DataGridView4.Rows(e.RowIndex).Cells(columnName:="N_").Value
            Form_Gruppi_Etichettaggio.stato_gruppo = "Visualizza"
            Form_Gruppi_Etichettaggio.inizializza_form()
            Form_Gruppi_Etichettaggio.Show()
        End If
    End Sub



    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Dim par_datagrdiview As DataGridView = DataGridView3
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = par_datagrdiview.Columns.IndexOf(Immagine_) Then



                Form_campione_visualizza.id_campione = par_datagrdiview.Rows(e.RowIndex).Cells(columnName:="Campione_").Value
                Form_campione_visualizza.Show()
                Form_campione_visualizza.inizializza_form()

            End If


        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        If DataGridView1.Rows(e.RowIndex).Cells("Collaudo").Value = 1 Then
            DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
        End If
        If DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Value = "M" Then
            DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Style.BackColor = Color.Orange
        ElseIf DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Value = "CDS" Then
            DataGridView1.Rows(e.RowIndex).Cells("Tipo_combinazione").Style.BackColor = Color.YellowGreen
        End If
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then




            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_1) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_2) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_3) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_4) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_5) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_6) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_7) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_8) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_9) Or e.ColumnIndex = DataGridView1.Columns.IndexOf(Immagine_10) Then


                Form_campione_visualizza.id_campione = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex - 2).Value
                Form_campione_visualizza.Show()
                Form_campione_visualizza.inizializza_form()






            ElseIf e.ColumnIndex = DataGridView1.Columns.IndexOf(Numero) Then


                If final_bp_code = "" Then
                    Form_Nuova_combinazione.codice_bp = bp_code
                Else
                    Form_Nuova_combinazione.codice_bp = final_bp_code
                End If

                Form_Nuova_combinazione.codice_commessa = Label1.Text
                Form_Nuova_combinazione.Show()
                Form_Nuova_combinazione.DataGridView2.Rows.Clear()
                Form_Nuova_combinazione.ID_combinazione_salvata = DataGridView1.Rows(e.RowIndex).Cells(columnName:="id_combinazione").Value
                Form_Nuova_combinazione.info_combinazioni(Form_Nuova_combinazione.DataGridView2, DataGridView1.Rows(e.RowIndex).Cells(columnName:="id_combinazione").Value)

            End If

        End If
    End Sub
End Class