Imports System.Data.SqlClient
Imports System.Reflection.Emit
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Word
Imports Tirelli.ODP_Form

Public Class Form_gestione_campioni

    Public tipo_campione As Integer = 100
    Private isShiftKeyDown As Boolean = False
    Private startIndex As Integer = -1
    Private Sub Form_gestione_campioni_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        inizializza_form()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Sub inizializza_form()
        riempi_datagridview_campioni_new(0, 30, DataGridView1, TextBox6.Text, TextBox17.Text, TextBox1.Text, TextBox3.Text)
        riempi_datagridview_richieste_campioni(TextBox9.Text, TextBox4.Text, DataGridView3, TextBox5.Text, TextBox7.Text, TextBox8.Text, TextBox2.Text)
    End Sub

    Sub filtra_datagridview_campioni()
        riempi_datagridview_campioni_new(0, TextBox16.Text, DataGridView1, TextBox6.Text, TextBox17.Text, TextBox1.Text, TextBox3.Text)
    End Sub

    Sub filtra_datagridview_richieste()
        riempi_datagridview_richieste_campioni(TextBox9.Text, TextBox4.Text, DataGridView3, TextBox5.Text, TextBox7.Text, TextBox8.Text, TextBox2.Text)
    End Sub



    Sub riempi_datagridview_campioni_new(par_tipo_campione As Integer, par_numero_risultati As Integer, par_datagridview As DataGridView, par_cliente As String, par_nome As String, par_tipo As String, par_id_campione As String)
        Dim par_filtro_Cliente As String = ""
        If par_cliente = "" Then

            par_filtro_Cliente = ""
        Else
            par_filtro_Cliente = " and coalesce( t12.cardname,'') Like '%%" & par_cliente & "%%'  "


        End If

        Dim par_filtro_nome As String = ""
        If par_nome = "" Then

            par_filtro_nome = ""

        Else
            par_filtro_nome = " and coalesce( T0.[Nome],'') Like '%%" & par_nome & "%%'  "


        End If

        Dim par_filtro_tipo As String = ""
        If par_tipo = "" Then

            par_filtro_tipo = ""

        Else
            par_filtro_tipo = " and t10.descrizione Like '%%" & par_tipo & "%%'  "


        End If

        Dim par_filtro_id_campione As String = ""
        If par_id_campione = "" Then

            par_filtro_id_campione = ""

        Else
            par_filtro_id_campione = " and T0.[Id_Campione] = '" & par_id_campione & "'  "


        End If





        par_datagridview.Rows.Clear()


        Dim Cnn6 As New SqlConnection
        Cnn6.ConnectionString = Homepage.sap_tirelli
        Cnn6.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn6
        CMD_SAP_2.CommandText = "SELECT TOP " & par_numero_risultati & "    T0.[Id_Campione],
CASE WHEN T0.[Codice_BP] IS NULL THEN '' ELSE T0.[Codice_BP] END AS 'Codice_BP',
  coalesce( T0.[Nome],'') AS 'Nome',
t13.iniziale_sigla,
    CASE WHEN T0.[Descrizione] IS NULL THEN '' ELSE T0.[Descrizione] END AS 'Note',
    CASE WHEN T0.[Codice_SAP] IS NULL THEN '' ELSE T0.[Codice_SAP] END AS 'Codice_SAP',
    CASE WHEN T0.[Tipo_Campione] IS NULL THEN '' ELSE T0.[Tipo_Campione] END AS 'Tipo_Campione',
	t10.descrizione as 'Tipo_nome',
	t0.ownerid,
	CONCAT(T17.LASTNAME,' ',T17.firstname) as 'Owner',
	t0.insertdate,
	t0.updatedate,
    CASE WHEN (t0.immagine is null or t0.immagine ='' ) then 'N_A.JPG' ELSE t0.immagine END AS 'immagine'


 ,CASE WHEN t1.[Altezza] IS NULL THEN 0 ELSE t1.[Altezza] END AS 'Altezza_t1',
    CASE WHEN t1.[Larghezza] IS NULL THEN 0 ELSE t1.[Larghezza] END AS 'Larghezza_t1',
    CASE WHEN t1.[Profondita] IS NULL THEN 0 ELSE t1.[Profondita] END AS 'Profondita_t1',
    CASE WHEN t1.[Diametro_Interno] IS NULL THEN 0 ELSE t1.[Diametro_Interno] END AS 'Diametro_Interno_t1',
  CASE WHEN t1.[Diametro_Esterno] IS NULL THEN 0 ELSE t1.[Diametro_Esterno] END AS 'Diametro_Esterno_t1',
CASE WHEN t1.[Volume] IS NULL THEN 0 ELSE t1.[Volume] END AS 'Volume_t1'


,CASE WHEN t1.[Spazio_Testa] IS NULL THEN 0 ELSE t1.[Spazio_Testa] END AS 'Spazio_Testa_t1',
CASE WHEN t1.[Materiale] IS NULL THEN '' ELSE t1.[Materiale] END AS 'Materiale_t1',
CASE WHEN t1.[Forma] IS NULL THEN '' ELSE t1.[Forma] END AS 'Forma_t1',
CASE WHEN t1.[Sezione] IS NULL THEN '' ELSE t1.[Sezione] END AS 'Sezione_t1'


,CASE WHEN t1.[Superficie] IS NULL THEN '' ELSE t1.[Superficie] END AS 'Superficie_t1',
CASE WHEN t1.[Produttore] IS NULL THEN '' ELSE t1.[Produttore] END AS 'Produttore_t1'

,CASE WHEN t1.[Codice_Produttore] IS NULL THEN '' ELSE t1.[Codice_Produttore] END AS 'Codice_Produttore_t1'

,CASE WHEN t1.[Collo_Centrato] IS NULL THEN '' ELSE t1.[Collo_Centrato]  END AS 'Collo_Centrato_t1'


,CASE WHEN t1.[Tipo_Tappo] IS NULL THEN '' ELSE t1.[Tipo_Tappo] END AS 'Tipo_Tappo_t1'


,CASE WHEN t1.[Filettatura] IS NULL THEN '' ELSE t1.[Filettatura] END AS 'Filettatura_t1',
CASE WHEN t1.[Diametro_Esterno_Fil] IS NULL THEN 0 ELSE t1.[Diametro_Esterno_Fil] END AS 'Diametro_Esterno_Fil_t1',
CASE WHEN t1.[Passo] IS NULL THEN 0 ELSE t1.[Passo] END AS 'Passo_t1',
CASE WHEN t1.[Num_Principi] IS NULL THEN 0 ELSE t1.[Num_Principi] END AS 'Num_Principi_t1'

,CASE WHEN t2.[Altezza] IS NULL THEN 0 ELSE t2.[Altezza] END AS 'Altezza_t2',
CASE WHEN t2.[Larghezza] IS NULL THEN 0 ELSE t2.[Larghezza] END AS 'Larghezza_t2',
CASE WHEN t2.[Profondità] IS NULL THEN 0 ELSE t2.[Profondità] END AS 'Profondità_t2',
CASE WHEN t2.[Diametro_Interno] IS NULL THEN 0 ELSE t2.[Diametro_Interno] END AS 'Diametro_Interno_t2',
CASE WHEN t2.[Fissaggio] IS NULL THEN '' ELSE t2.[Fissaggio] END AS 'Fissaggio_t2',
CASE WHEN t2.[Forma] IS NULL THEN '' ELSE t2.[Forma] END AS 'Forma_t2',
CASE WHEN t2.[Materiale] IS NULL THEN '' ELSE t2.[Materiale] END AS 'Materiale_t2',
CASE WHEN t2.[Superficie] IS NULL THEN '' ELSE t2.[Superficie] END AS 'Superficie_t2',
CASE WHEN t2.[Produttore] IS NULL THEN '' ELSE t2.[Produttore] END AS 'Produttore_t2',
CASE WHEN t2.[Codice_Produttore] IS NULL THEN '' ELSE t2.[Codice_Produttore] END AS 'Codice_Produttore_t2'

-- Campi t3
,CASE WHEN t3.[Altezza] IS NULL THEN 0 ELSE t3.[Altezza] END AS 'Altezza_t3',
CASE WHEN t3.[Larghezza] IS NULL THEN 0 ELSE t3.[Larghezza] END AS 'Larghezza_t3',
CASE WHEN t3.[Profondità] IS NULL THEN 0 ELSE t3.[Profondità] END AS 'Profondità_t3',
CASE WHEN t3.[Diametro_Interno] IS NULL THEN 0 ELSE t3.[Diametro_Interno] END AS 'Diametro_Interno_t3',
CASE WHEN t3.[Vite_Pressione] IS NULL THEN '' ELSE t3.[Vite_Pressione] END AS 'Vite_Pressione_t3',
CASE WHEN t3.[Forma] IS NULL THEN '' ELSE t3.[Forma] END AS 'Forma_t3',
CASE WHEN t3.[Materiale] IS NULL THEN '' ELSE t3.[Materiale] END AS 'Materiale_t3',


   CASE WHEN t4.[A] IS NULL THEN 0 ELSE t4.[A] END AS 'A_t4',
CASE WHEN t4.[B] IS NULL THEN 0 ELSE t4.[B] END AS 'B_t4',
CASE WHEN t4.[C] IS NULL THEN 0 ELSE t4.[C] END AS 'C_t4',
CASE WHEN t4.[D] IS NULL THEN 0 ELSE t4.[D] END AS 'D_t4',
CASE WHEN t4.[Quota_A] IS NULL THEN 0 ELSE t4.[Quota_A] END AS 'Quota_A_t4',
CASE WHEN t4.[Quota_B] IS NULL THEN 0 ELSE t4.[Quota_B] END AS 'Quota_B_t4',
CASE WHEN t4.[Quota_C] IS NULL THEN 0 ELSE t4.[Quota_C] END AS 'Quota_C_t4',
CASE WHEN t4.[Quota_D] IS NULL THEN 0 ELSE t4.[Quota_D] END AS 'Quota_D_t4',
CASE WHEN t4.[Quota_E] IS NULL THEN 0 ELSE t4.[Quota_E] END AS 'Quota_E_t4',
CASE WHEN t4.[Quota_F] IS NULL THEN 0 ELSE t4.[Quota_F] END AS 'Quota_F_t4',
CASE WHEN t4.[Quota_L] IS NULL THEN 0 ELSE t4.[Quota_L] END AS 'Quota_L_t4',
CASE WHEN t4.[SP] IS NULL THEN 0 ELSE t4.[SP] END AS 'SP_t4',
CASE WHEN t4.[Materiale] IS NULL THEN '' ELSE t4.[Materiale] END AS 'Materiale_t4',
CASE WHEN t4.[Tipologia] IS NULL THEN '' ELSE t4.[Tipologia] END AS 'Tipologia_t4',
CASE WHEN t4.[Superficie] IS NULL THEN '' ELSE t4.[Superficie] END AS 'Superficie_t4',
CASE WHEN t4.[Produttore] IS NULL THEN '' ELSE t4.[Produttore] END AS 'Produttore_t4',
CASE WHEN t4.[cod_produttore] IS NULL THEN '' ELSE t4.[cod_produttore] END AS 'Codice_Produttore_t4',
CASE WHEN t4.[Fissaggio] IS NULL THEN '' ELSE t4.[Fissaggio] END AS 'Fissaggio_t4',
CASE WHEN t4.[Ghiera] IS NULL THEN '' ELSE t4.[Ghiera] END AS 'Ghiera_t4',
CASE WHEN t4.[Copritappo] IS NULL THEN '' ELSE t4.[Copritappo] END AS 'Copritappo_t4'

-- Campi t5
,CASE WHEN t5.[Altezza] IS NULL THEN 0 ELSE t5.[Altezza] END AS 'Altezza_t5',
CASE WHEN t5.[Larghezza] IS NULL THEN 0 ELSE t5.[Larghezza] END AS 'Larghezza_t5',
CASE WHEN t5.[Trasparenza] IS NULL THEN '' ELSE t5.[Trasparenza] END AS 'Trasparenza_t5',
CASE WHEN t5.[Forma] IS NULL THEN '' ELSE t5.[Forma] END AS 'Forma_t5',
CASE WHEN t5.[Diametro_Esterno_Bobina] IS NULL THEN 0 ELSE t5.[Diametro_Esterno_Bobina] END AS 'Diametro_Esterno_Bobina_t5',
CASE WHEN t5.[Diametro_Interno_Bobina] IS NULL THEN 0 ELSE t5.[Diametro_Interno_Bobina] END AS 'Diametro_Interno_Bobina_t5',
CASE WHEN t5.[Avvolgimento_Bobina] IS NULL THEN '' ELSE t5.[Avvolgimento_Bobina] END AS 'Avvolgimento_Bobina_t5',
CASE WHEN t5.[Materiale] IS NULL THEN '' ELSE t5.[Materiale] END AS 'Materiale_t5',

CASE WHEN t6.[A] IS NULL THEN 0 ELSE t6.[A] END AS 'A_t6',
CASE WHEN t6.[B] IS NULL THEN 0 ELSE t6.[B] END AS 'B_t6',
CASE WHEN t6.[Quota_S] IS NULL THEN 0 ELSE t6.[Quota_S] END AS 'Quota_S_t6',
CASE WHEN t6.[Quota_H] IS NULL THEN 0 ELSE t6.[Quota_H] END AS 'Quota_H_t6',
CASE WHEN t6.[Quota_L] IS NULL THEN 0 ELSE t6.[Quota_L] END AS 'Quota_L_t6',
CASE WHEN t6.[Quota_W] IS NULL THEN 0 ELSE t6.[Quota_W] END AS 'Quota_W_t6',
CASE WHEN t6.[Quota_V] IS NULL THEN 0 ELSE t6.[Quota_V] END AS 'Quota_V_t6',
CASE WHEN t6.[Pressione/Vite] IS NULL THEN '' ELSE t6.[Pressione/Vite] END AS 'Pressione/Vite_t6',
CASE WHEN t6.[Produttore] IS NULL THEN '' ELSE t6.[Produttore] END AS 'Produttore_t6',
CASE WHEN t6.[Codice_produttore] IS NULL THEN '' ELSE t6.[Codice_produttore] END AS 'Codice_Produttore_t6',
CASE WHEN t6.[Materiale] IS NULL THEN '' ELSE t6.[Materiale] END AS 'Materiale_t6',
CASE WHEN t6.[SP] IS NULL THEN 0 ELSE t6.[SP] END AS 'SP_t6',
CASE WHEN t6.[T] IS NULL THEN 0 ELSE t6.[T] END AS 'T_t6',
CASE WHEN t6.[Fissaggio] IS NULL THEN '' ELSE t6.[Fissaggio] END AS 'Fissaggio_t6',
CASE WHEN t6.[Ghiera] IS NULL THEN '' ELSE t6.[Ghiera] END AS 'Ghiera_t6',
CASE WHEN t6.[Grileltto] IS NULL THEN '' ELSE t6.[Grileltto] END AS 'Grileltto_t6',
CASE WHEN t6.[Protezione] IS NULL THEN '' ELSE t6.[Protezione] END AS 'Protezione_t6',
CASE WHEN t6.[Note] IS NULL THEN '' ELSE t6.[Note] END AS 'Note_t6',
CASE WHEN t6.[Cannuccia] IS NULL THEN '' ELSE t6.[Cannuccia] END AS 'Cannuccia_t6',

-- Campi t7
CASE WHEN t7.[Densita] IS NULL THEN 0 ELSE t7.[Densita] END AS 'Densita_t7',
CASE WHEN t7.[Viscosita_Dinamica] IS NULL THEN 0 ELSE t7.[Viscosita_Dinamica] END AS 'Viscosita_Dinamica_t7',
CASE WHEN t7.[Conducibilita_Elettrica] IS NULL THEN 0 ELSE t7.[Conducibilita_Elettrica] END AS 'Conducibilita_Elettrica_t7',
CASE WHEN t7.[Categoria] IS NULL THEN '' ELSE t7.[Categoria] END AS 'Categoria_t7',
CASE WHEN t7.[Infiammabile] IS NULL THEN '' ELSE t7.[Infiammabile] END AS 'Infiammabile_t7',
CASE WHEN t7.[Nome_Commerciale] IS NULL THEN '' ELSE t7.[Nome_Commerciale] END AS 'Nome_Commerciale_t7',
CASE WHEN t7.[Viscosità_Cinematica] IS NULL THEN 0 ELSE t7.[Viscosità_Cinematica] END AS 'Viscosità_Cinematica_t7',
CASE WHEN t7.[Corrosivo] IS NULL THEN '' ELSE t7.[Corrosivo] END AS 'Corrosivo_t7',
CASE WHEN t7.[Nocivo/Tossico] IS NULL THEN '' ELSE t7.[Nocivo/Tossico] END AS 'Nocivo/Tossico_t7',
CASE WHEN t7.[Note] IS NULL THEN '' ELSE t7.[Note] END AS 'Note_t7',

-- Campi t8
CASE WHEN t8.[Larghezza] IS NULL THEN 0 ELSE t8.[Larghezza] END AS 'Larghezza_t8',
CASE WHEN t8.[Diametro_Fulcro] IS NULL THEN 0 ELSE t8.[Diametro_Fulcro] END AS 'Diametro_Fulcro_t8',
CASE WHEN t8.[Materiale] IS NULL THEN '' ELSE t8.[Materiale] END AS 'Materiale_t8',
CASE WHEN t8.[Temperatura_Saldatura] IS NULL THEN 0 ELSE t8.[Temperatura_Saldatura] END AS 'Temperatura_Saldatura_t8',
CASE WHEN t8.[Diametro_Esterno] IS NULL THEN 0 ELSE t8.[Diametro_Esterno] END AS 'Diametro_Esterno_t8',

 CASE WHEN t9.[Altezza] IS NULL THEN 0 ELSE t9.[Altezza] END AS 'Altezza_t9',
CASE WHEN t9.[Larghezza] IS NULL THEN 0 ELSE t9.[Larghezza] END AS 'Larghezza_t9',
CASE WHEN t9.[Profondità] IS NULL THEN 0 ELSE t9.[Profondità] END AS 'Profondità_t9',
CASE WHEN t9.[Diametro_Interno] IS NULL THEN 0 ELSE t9.[Diametro_Interno] END AS 'Diametro_Interno_t9',
CASE WHEN t9.[Fissaggio] IS NULL THEN '' ELSE t9.[Fissaggio] END AS 'Fissaggio_t9',
CASE WHEN t9.[Forma] IS NULL THEN '' ELSE t9.[Forma] END AS 'Forma_t9',
CASE WHEN t9.[Materiale] IS NULL THEN '' ELSE t9.[Materiale] END AS 'Materiale_t9',
CASE WHEN t9.[Superficie] IS NULL THEN '' ELSE t9.[Superficie] END AS 'Superficie_t9',
CASE WHEN t9.[Produttore] IS NULL THEN '' ELSE t9.[Produttore] END AS 'Produttore_t9',
CASE WHEN t9.[Codice_produttore] IS NULL THEN '' ELSE t9.[Codice_produttore] END AS 'Codice_Produttore_t9',

-- Campi t10
CASE WHEN t10.iniziale_sigla IS NULL THEN '' ELSE t10.iniziale_sigla END AS 'iniziale_sigla',

-- Campi t11
CASE WHEN t11.onhand IS NULL THEN 0 ELSE cast(t11.onhand as integer) END AS 'onhand',
CASE WHEN t11.u_ubicazione IS NULL THEN '' ELSE t11.u_ubicazione END AS 'u_ubicazione',

-- Campi t12
coalesce( t12.cardname,'') AS 'cardname',

t13.immagine_descrizione

,T14.CARDNAME
,case when t15.cardcode is null then '' else t15.cardcode end as 'codice_bp_principale'
,case when t15.cardname is null then '' else t15.cardname end as 'Cliente_principale'

FROM [TIRELLI_40].[DBO].[coll_campioni] AS T0
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_flaconi] t1 ON t0.id_campione = t1.codice_campione AND t0.tipo_campione = 100
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_tappi] t2 ON t0.id_campione = t2.codice_campione AND t0.tipo_campione = 101
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_sottotappi] t3 ON t0.id_campione = t3.codice_campione AND t0.tipo_campione = 102
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_pompette] t4 ON t0.id_campione = t4.codice_campione AND t0.tipo_campione = 103
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_etichette] t5 ON t0.id_campione = t5.codice_campione AND t0.tipo_campione = 104
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_trigger] t6 ON t0.id_campione = t6.codice_campione AND t0.tipo_campione = 105
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_prodotti] t7 ON t0.id_campione = t7.codice_campione AND t0.tipo_campione = 106
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_film] t8 ON t0.id_campione = t8.codice_campione AND t0.tipo_campione = 107
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_copritappi] t9 ON t0.id_campione = t9.codice_campione AND t0.tipo_campione = 108
LEFT JOIN [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t10 ON t10.Id_Tipo_Campione = t0.Tipo_Campione
LEFT JOIN [TIRELLISRLDB].[dbo].oitm t11 ON t11.itemcode = t0.codice_sap 
LEFT JOIN [TIRELLISRLDB].[dbo].ocrd t12 ON t12.cardcode = t0.codice_BP 
LEFT JOIN [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t13 ON t13.id_tipo_campione = t0.tipo_campione
LEFT JOIN [TIRELLISRLDB].[dbo].OCRD T14 ON CAST(T14.CARDCODE AS VARCHAR) = CAST(t0.codice_bp AS VARCHAR) 
LEFT JOIN [TIRELLISRLDB].[dbo].ocrd t15 ON CAST(t15.u_bp_riferimento AS VARCHAR) = CAST(t14.cardcode AS VARCHAR) 
LEFT JOIN [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t16 ON t0.TIPO_campione = T16.ID_TIPO_CAMPIONE
LEFT JOIN [TIRELLI_40].[dbo].OHEM t17 ON T17.EMPID = T0.ownerid


WHERE 0=0 " & par_filtro_Cliente & par_filtro_tipo & par_filtro_nome & par_filtro_id_campione & "

        order by T14.CARDNAME ,t10.INIZIALE_SIGLA ,  cast(T0.NOME as integer)"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(False, cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Tipo_nome"), cmd_SAP_reader_2("Iniziale_sigla") & cmd_SAP_reader_2("Nome"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("cliente_principale"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("codice_bp_principale"))

        Loop


        par_datagridview.ClearSelection()

        cmd_SAP_reader_2.Close()
        Cnn6.Close()
    End Sub

    Sub riempi_datagridview_richieste_campioni(par_owner As String, par_n_doc As String, par_datagridview As DataGridView, par_cliente As String, par_nome As String, par_tipo As String, par_id_campione As String)

        Dim par_filtro_owner As String = ""
        If par_owner = "" Then

            par_filtro_owner = ""
        Else
            par_filtro_owner = " and concat(t8.lastname,' ', t8.firstname) Like '%%" & par_owner & "%%'  "
        End If

        Dim par_filtro_id_Campione As String = ""
        If par_id_campione = "" Then

            par_filtro_id_Campione = ""
        Else
            par_filtro_id_Campione = " and t0.[id_campione] = '" & par_id_campione & "'  "
        End If

        Dim par_filtro_n_doc As String = ""
        If par_n_doc = "" Then

            par_filtro_n_doc = ""
        Else
            par_filtro_n_doc = " and [id_doc_richiesta] = '" & par_n_doc & "'  "
        End If

        Dim par_filtro_cliente As String = ""
        If par_cliente = "" Then

            par_filtro_cliente = ""
        Else
            par_filtro_cliente = " and t3.cardname Like '%%" & par_cliente & "%%'   "
        End If

        Dim filtro_nome_campione As String = ""
        If par_nome = "" Then

            filtro_nome_campione = ""
        Else
            filtro_nome_campione = " and  concat(t4.iniziale_sigla,coalesce( T1.[Nome],'')) Like '%%" & par_nome & "%%'   "
        End If

        Dim filtro_tipo_campione As String = ""
        If par_tipo = "" Then

            filtro_tipo_campione = ""
        Else
            filtro_tipo_campione = " and  t4.descrizione Like '%%" & par_tipo & "%%'   "
        End If






        par_datagridview.Rows.Clear()


        Dim Cnn6 As New SqlConnection
        Cnn6.ConnectionString = Homepage.sap_tirelli
        Cnn6.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn6
        CMD_SAP_2.CommandText = "select top 50 *

from
(
SELECT  
t0.[id_doc_richiesta]
,t0.[Id_richiesta]
      ,t0.[id_campione]
      ,t0.[Insertdate]
      ,t0.[Duedate]
      ,t0.[Owner] as 'Owner_ric'
, concat(t8.lastname,' ', t8.firstname) as 'Cognome_owner'
      ,t0.[Q_tot]
      ,t0.[Q_open]- SUM(COALESCE(T9.Q,0)) AS 'Q_OPEN'
      ,t0.[Status]
,coalesce(T1.[Codice_BP], '') AS 'Codice_BP',
t3.cardname,
  coalesce( T1.[Nome],'') AS 'Nome',
t4.iniziale_sigla,
  coalesce(T1.[Tipo_Campione] ,'') AS 'Tipo_Campione',
	t4.descrizione as 'Tipo_nome',
	t1.ownerid,
	CONCAT(T7.LASTNAME,' ',T7.firstname) as 'Owner',
	t1.insertdate as 'Data_codifica_campione',
	t1.updatedate,
    CASE WHEN COALESCE(t1.immagine,'')='' then 'N_A.JPG' ELSE t1.immagine END AS 'immagine',
t1.ubicazione
,coalesce(t8.email,'') as 'Email'

  FROM [TIRELLI_40].[DBO].[coll_campioni_richieste] t0
LEFT JOIN [TIRELLI_40].[DBO].coll_campioni AS T1 ON t0.id_campione = t1.id_campione
LEFT JOIN [TIRELLISRLDB].[dbo].ocrd t3 ON t3.cardcode = t1.codice_BP 
LEFT JOIN [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t4 ON t4.id_tipo_campione = t1.tipo_campione
LEFT JOIN [TIRELLISRLDB].[dbo].ocrd t5 ON CAST(t5.u_bp_riferimento AS VARCHAR) = CAST(t1.codice_bp AS VARCHAR) 
LEFT JOIN [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t6 ON t1.TIPO_campione = T6.ID_TIPO_CAMPIONE
LEFT JOIN [TIRELLI_40].[dbo].OHEM t7 ON T7.EMPID = T1.ownerid
LEFT JOIN [TIRELLI_40].[dbo].OHEM t8 ON T8.EMPID = T0.owner
left join [TIRELLI_40].[DBO].[coll_campioni_movimentazioni] t9 on t9.tipo_movimentazione='EM' AND t0.[Id_richiesta]=T9.Id_richiesta



where 0 =0 " & par_filtro_owner & par_filtro_id_Campione & par_filtro_n_doc & par_filtro_cliente & filtro_nome_campione & filtro_tipo_campione & "

GROUP BY 
t0.[id_doc_richiesta]
,t0.[Id_richiesta]
      ,t0.[id_campione]
      ,t0.[Insertdate]
      ,t0.[Duedate]
      ,t0.[Owner]
	  ,t8.lastname
	  , t8.firstname
	  ,t0.[Q_tot]
      ,t0.[Q_open]
	  ,t0.[Status]
,T1.[Codice_BP]
,t3.cardname
,T1.[Nome]
,t4.iniziale_sigla
,T1.[Tipo_Campione]
,t4.descrizione
,t1.ownerid
,T7.LASTNAME
,T7.firstname
,t1.updatedate
,t1.insertdate
,t1.immagine
,t1.ubicazione
,t8.email
)
as t10
where t10.Q_open>0
order by t10.[Id_doc_richiesta] desc 

"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(False, cmd_SAP_reader_2("id_doc_richiesta"), cmd_SAP_reader_2("id_richiesta"), cmd_SAP_reader_2("insertdate"), cmd_SAP_reader_2("duedate"), cmd_SAP_reader_2("Cognome_owner"), cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Tipo_nome"), cmd_SAP_reader_2("Iniziale_sigla") & cmd_SAP_reader_2("Nome"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cardname"), cmd_SAP_reader_2("codice_bp"), cmd_SAP_reader_2("Q_TOT"), cmd_SAP_reader_2("Q_open"), cmd_SAP_reader_2("Status"), cmd_SAP_reader_2("Email"))

        Loop


        par_datagridview.ClearSelection()

        cmd_SAP_reader_2.Close()
        Cnn6.Close()
    End Sub

    Private Sub tabpage4_Click(sender As Object, e As EventArgs) Handles TabPage4.Enter

        riempi_movimentazioni_campioni(TextBox10.Text, TextBox14.Text, DataGridView4, TextBox13.Text, TextBox12.Text, TextBox11.Text, TextBox15.Text)
    End Sub

    Sub riempi_movimentazioni_campioni(par_owner As String, par_n_doc As String, par_datagridview As DataGridView, par_cliente As String, par_nome As String, par_tipo As String, par_id_campione As String)

        Dim par_filtro_owner As String = ""
        If par_owner = "" Then

            par_filtro_owner = ""
        Else
            par_filtro_owner = " and concat(t1.lastname,' ',t1.firstname) Like '%%" & par_owner & "%%'  "
        End If

        Dim par_filtro_id_Campione As String = ""
        If par_id_campione = "" Then

            par_filtro_id_Campione = ""
        Else
            par_filtro_id_Campione = " and t0.[id_campione] = '" & par_id_campione & "'  "
        End If

        Dim par_filtro_n_doc As String = ""
        If par_n_doc = "" Then

            par_filtro_n_doc = ""
        Else
            par_filtro_n_doc = " and t0.id_richiesta = '" & par_n_doc & "'  "
        End If

        Dim par_filtro_cliente As String = ""
        If par_cliente = "" Then

            par_filtro_cliente = ""
        Else
            par_filtro_cliente = " and t3.cardname Like '%%" & par_cliente & "%%'   "
        End If

        Dim filtro_nome_campione As String = ""
        If par_nome = "" Then

            filtro_nome_campione = ""
        Else
            filtro_nome_campione = " and  concat(t4.iniziale_sigla,coalesce( T1.[Nome],'')) Like '%%" & par_nome & "%%'   "
        End If

        Dim filtro_tipo_campione As String = ""
        If par_tipo = "" Then

            filtro_tipo_campione = ""
        Else
            filtro_tipo_campione = " and  t4.descrizione Like '%%" & par_tipo & "%%'   "
        End If

        par_datagridview.Rows.Clear()


        Dim Cnn6 As New SqlConnection
        Cnn6.ConnectionString = Homepage.sap_tirelli
        Cnn6.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn6
        CMD_SAP_2.CommandText = "select top 50 
t0.id_movimentazione
,t0.tipo_movimentazione
,t0.id_richiesta
,t0.id_campione
,coalesce(t3.cardname,'') as 'Cliente'
,t0.Insertdate
,concat(t1.lastname,' ',t1.firstname) as 'Owner'
,t0.segno
,t0.q
,t0.mag
,t4.iniziale_sigla
,coalesce(T2.[Tipo_Campione] ,'') AS 'Tipo_Campione'
,t4.descrizione as 'Tipo_nome'
, coalesce( T2.[Nome],'') AS 'Nome'
,CASE WHEN COALESCE(t2.immagine,'')='' then 'N_A.JPG' ELSE t2.immagine END AS 'immagine'


from
[TIRELLI_40].[DBO].[coll_campioni_movimentazioni] t0
left join [TIRELLI_40].[dbo].ohem t1 on t0.owner=t1.empid
LEFT JOIN [TIRELLI_40].[DBO].coll_campioni AS T2 ON t0.id_campione = t2.id_campione
LEFT JOIN [TIRELLISRLDB].[dbo].ocrd t3 ON t3.cardcode = t2.codice_BP 
LEFT JOIN [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t4 ON t4.id_tipo_campione = t2.tipo_campione

where 0 =0 " & par_filtro_owner & par_filtro_id_Campione & par_filtro_n_doc & par_filtro_cliente & filtro_nome_campione & filtro_tipo_campione & "

order by id_movimentazione desc


"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader



        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(cmd_SAP_reader_2("id_movimentazione"), cmd_SAP_reader_2("tipo_movimentazione"), cmd_SAP_reader_2("id_richiesta"), cmd_SAP_reader_2("insertdate"), cmd_SAP_reader_2("Owner"), cmd_SAP_reader_2("id_campione"), cmd_SAP_reader_2("Tipo_nome"), cmd_SAP_reader_2("Iniziale_sigla") & cmd_SAP_reader_2("Nome"), Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2("immagine")), cmd_SAP_reader_2("Cliente"), cmd_SAP_reader_2("Mag"), cmd_SAP_reader_2("Segno"), cmd_SAP_reader_2("Q"))

        Loop


        par_datagridview.ClearSelection()

        cmd_SAP_reader_2.Close()
        Cnn6.Close()
    End Sub






    Public Function ottieni_informazioni_campione(par_id_Campione As Integer)
        Dim dettagli As New Dettagli_Campione()

        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT    T0.[Id_Campione],
CASE WHEN T0.[Codice_BP] IS NULL THEN '' ELSE T0.[Codice_BP] END AS 'Codice_BP',
    coalesce( T0.[Nome],'') AS 'Nome',
t13.iniziale_sigla,
    coalesce(T0.[Descrizione] , '') AS 'Descrizione',
coalesce(T0.[Note] , '') AS 'Note',
    CASE WHEN T0.[Codice_SAP] IS NULL THEN '' ELSE T0.[Codice_SAP] END AS 'Codice_SAP',
    CASE WHEN T0.[Tipo_Campione] IS NULL THEN '' ELSE T0.[Tipo_Campione] END AS 'Tipo_Campione',
	t10.descrizione as 'Tipo_nome',
coalesce(t10.descrizione_inglese,'') as 'Tipo_nome_inglese',
	t0.ownerid,
	CONCAT(T17.LASTNAME,' ',T17.firstname) as 'Owner',
	t0.insertdate,
	t0.updatedate,
    CASE WHEN (t0.immagine is null or t0.immagine ='' ) then 'N_A.JPG' ELSE t0.immagine END AS 'immagine'


 ,CASE WHEN t1.[Altezza] IS NULL THEN 0 ELSE t1.[Altezza] END AS 'Altezza_t1',
    CASE WHEN t1.[Larghezza] IS NULL THEN 0 ELSE t1.[Larghezza] END AS 'Larghezza_t1',
    CASE WHEN t1.[Profondita] IS NULL THEN 0 ELSE t1.[Profondita] END AS 'Profondita_t1',
    CASE WHEN t1.[Diametro_Interno] IS NULL THEN 0 ELSE t1.[Diametro_Interno] END AS 'Diametro_Interno_t1',
  CASE WHEN t1.[Diametro_Esterno] IS NULL THEN 0 ELSE t1.[Diametro_Esterno] END AS 'Diametro_Esterno_t1',
CASE WHEN t1.[Volume] IS NULL THEN 0 ELSE t1.[Volume] END AS 'Volume_t1'


,CASE WHEN t1.[Spazio_Testa] IS NULL THEN 0 ELSE t1.[Spazio_Testa] END AS 'Spazio_Testa_t1',
CASE WHEN t1.[Materiale] IS NULL THEN '' ELSE t1.[Materiale] END AS 'Materiale_t1',
CASE WHEN t1.[Forma] IS NULL THEN '' ELSE t1.[Forma] END AS 'Forma_t1',
CASE WHEN t1.[Sezione] IS NULL THEN '' ELSE t1.[Sezione] END AS 'Sezione_t1'


,CASE WHEN t1.[Superficie] IS NULL THEN '' ELSE t1.[Superficie] END AS 'Superficie_t1',
CASE WHEN t1.[Produttore] IS NULL THEN '' ELSE t1.[Produttore] END AS 'Produttore_t1'

,CASE WHEN t1.[Codice_Produttore] IS NULL THEN '' ELSE t1.[Codice_Produttore] END AS 'Codice_Produttore_t1'

,CASE WHEN t1.[Collo_Centrato] IS NULL THEN '' ELSE t1.[Collo_Centrato]  END AS 'Collo_Centrato_t1'


,CASE WHEN t1.[Tipo_Tappo] IS NULL THEN '' ELSE t1.[Tipo_Tappo] END AS 'Tipo_Tappo_t1'


,CASE WHEN t1.[Filettatura] IS NULL THEN '' ELSE t1.[Filettatura] END AS 'Filettatura_t1',
CASE WHEN t1.[Diametro_Esterno_Fil] IS NULL THEN 0 ELSE t1.[Diametro_Esterno_Fil] END AS 'Diametro_Esterno_Fil_t1',
CASE WHEN t1.[Passo] IS NULL THEN 0 ELSE t1.[Passo] END AS 'Passo_t1',
CASE WHEN t1.[Num_Principi] IS NULL THEN 0 ELSE t1.[Num_Principi] END AS 'Num_Principi_t1'

,CASE WHEN t2.[Altezza] IS NULL THEN 0 ELSE t2.[Altezza] END AS 'Altezza_t2',
CASE WHEN t2.[Larghezza] IS NULL THEN 0 ELSE t2.[Larghezza] END AS 'Larghezza_t2',
CASE WHEN t2.[Profondità] IS NULL THEN 0 ELSE t2.[Profondità] END AS 'Profondità_t2',
CASE WHEN t2.[Diametro_Interno] IS NULL THEN 0 ELSE t2.[Diametro_Interno] END AS 'Diametro_Interno_t2',
CASE WHEN t2.[Fissaggio] IS NULL THEN '' ELSE t2.[Fissaggio] END AS 'Fissaggio_t2',
CASE WHEN t2.[Forma] IS NULL THEN '' ELSE t2.[Forma] END AS 'Forma_t2',
CASE WHEN t2.[Materiale] IS NULL THEN '' ELSE t2.[Materiale] END AS 'Materiale_t2',
CASE WHEN t2.[Superficie] IS NULL THEN '' ELSE t2.[Superficie] END AS 'Superficie_t2',
CASE WHEN t2.[Produttore] IS NULL THEN '' ELSE t2.[Produttore] END AS 'Produttore_t2',
CASE WHEN t2.[Codice_Produttore] IS NULL THEN '' ELSE t2.[Codice_Produttore] END AS 'Codice_Produttore_t2'

-- Campi t3
,CASE WHEN t3.[Altezza] IS NULL THEN 0 ELSE t3.[Altezza] END AS 'Altezza_t3',
CASE WHEN t3.[Larghezza] IS NULL THEN 0 ELSE t3.[Larghezza] END AS 'Larghezza_t3',
CASE WHEN t3.[Profondità] IS NULL THEN 0 ELSE t3.[Profondità] END AS 'Profondità_t3',
CASE WHEN t3.[Diametro_Interno] IS NULL THEN 0 ELSE t3.[Diametro_Interno] END AS 'Diametro_Interno_t3',
CASE WHEN t3.[Vite_Pressione] IS NULL THEN '' ELSE t3.[Vite_Pressione] END AS 'Vite_Pressione_t3',
CASE WHEN t3.[Forma] IS NULL THEN '' ELSE t3.[Forma] END AS 'Forma_t3',
CASE WHEN t3.[Materiale] IS NULL THEN '' ELSE t3.[Materiale] END AS 'Materiale_t3',


   CASE WHEN t4.[A] IS NULL THEN 0 ELSE t4.[A] END AS 'A_t4',
CASE WHEN t4.[B] IS NULL THEN 0 ELSE t4.[B] END AS 'B_t4',
CASE WHEN t4.[C] IS NULL THEN 0 ELSE t4.[C] END AS 'C_t4',
CASE WHEN t4.[D] IS NULL THEN 0 ELSE t4.[D] END AS 'D_t4',
CASE WHEN t4.[Quota_A] IS NULL THEN 0 ELSE t4.[Quota_A] END AS 'Quota_A_t4',
CASE WHEN t4.[Quota_B] IS NULL THEN 0 ELSE t4.[Quota_B] END AS 'Quota_B_t4',
CASE WHEN t4.[Quota_C] IS NULL THEN 0 ELSE t4.[Quota_C] END AS 'Quota_C_t4',
CASE WHEN t4.[Quota_D] IS NULL THEN 0 ELSE t4.[Quota_D] END AS 'Quota_D_t4',
CASE WHEN t4.[Quota_E] IS NULL THEN 0 ELSE t4.[Quota_E] END AS 'Quota_E_t4',
CASE WHEN t4.[Quota_F] IS NULL THEN 0 ELSE t4.[Quota_F] END AS 'Quota_F_t4',
CASE WHEN t4.[Quota_L] IS NULL THEN 0 ELSE t4.[Quota_L] END AS 'Quota_L_t4',
CASE WHEN t4.[SP] IS NULL THEN 0 ELSE t4.[SP] END AS 'SP_t4',
CASE WHEN t4.[Materiale] IS NULL THEN '' ELSE t4.[Materiale] END AS 'Materiale_t4',
CASE WHEN t4.[Tipologia] IS NULL THEN '' ELSE t4.[Tipologia] END AS 'Tipologia_t4',
CASE WHEN t4.[Superficie] IS NULL THEN '' ELSE t4.[Superficie] END AS 'Superficie_t4',
CASE WHEN t4.[Produttore] IS NULL THEN '' ELSE t4.[Produttore] END AS 'Produttore_t4',
CASE WHEN t4.[cod_produttore] IS NULL THEN '' ELSE t4.[cod_produttore] END AS 'Codice_Produttore_t4',
CASE WHEN t4.[Fissaggio] IS NULL THEN '' ELSE t4.[Fissaggio] END AS 'Fissaggio_t4',
CASE WHEN t4.[Ghiera] IS NULL THEN '' ELSE t4.[Ghiera] END AS 'Ghiera_t4',
CASE WHEN t4.[Copritappo] IS NULL THEN '' ELSE t4.[Copritappo] END AS 'Copritappo_t4'

-- Campi t5
,CASE WHEN t5.[Altezza] IS NULL THEN 0 ELSE t5.[Altezza] END AS 'Altezza_t5',
CASE WHEN t5.[Larghezza] IS NULL THEN 0 ELSE t5.[Larghezza] END AS 'Larghezza_t5',
CASE WHEN t5.[Trasparenza] IS NULL THEN '' ELSE t5.[Trasparenza] END AS 'Trasparenza_t5',
CASE WHEN t5.[Forma] IS NULL THEN '' ELSE t5.[Forma] END AS 'Forma_t5',
CASE WHEN t5.[Diametro_Esterno_Bobina] IS NULL THEN 0 ELSE t5.[Diametro_Esterno_Bobina] END AS 'Diametro_Esterno_Bobina_t5',
CASE WHEN t5.[Diametro_Interno_Bobina] IS NULL THEN 0 ELSE t5.[Diametro_Interno_Bobina] END AS 'Diametro_Interno_Bobina_t5',
CASE WHEN t5.[Avvolgimento_Bobina] IS NULL THEN '' ELSE t5.[Avvolgimento_Bobina] END AS 'Avvolgimento_Bobina_t5',
CASE WHEN t5.[Materiale] IS NULL THEN '' ELSE t5.[Materiale] END AS 'Materiale_t5',

CASE WHEN t6.[A] IS NULL THEN 0 ELSE t6.[A] END AS 'A_t6',
CASE WHEN t6.[B] IS NULL THEN 0 ELSE t6.[B] END AS 'B_t6',
CASE WHEN t6.[Quota_S] IS NULL THEN 0 ELSE t6.[Quota_S] END AS 'Quota_S_t6',
CASE WHEN t6.[Quota_H] IS NULL THEN 0 ELSE t6.[Quota_H] END AS 'Quota_H_t6',
CASE WHEN t6.[Quota_L] IS NULL THEN 0 ELSE t6.[Quota_L] END AS 'Quota_L_t6',
CASE WHEN t6.[Quota_W] IS NULL THEN 0 ELSE t6.[Quota_W] END AS 'Quota_W_t6',
CASE WHEN t6.[Quota_V] IS NULL THEN 0 ELSE t6.[Quota_V] END AS 'Quota_V_t6',
CASE WHEN t6.[Pressione/Vite] IS NULL THEN '' ELSE t6.[Pressione/Vite] END AS 'Pressione/Vite_t6',
CASE WHEN t6.[Produttore] IS NULL THEN '' ELSE t6.[Produttore] END AS 'Produttore_t6',
CASE WHEN t6.[Codice_produttore] IS NULL THEN '' ELSE t6.[Codice_produttore] END AS 'Codice_Produttore_t6',
CASE WHEN t6.[Materiale] IS NULL THEN '' ELSE t6.[Materiale] END AS 'Materiale_t6',
CASE WHEN t6.[SP] IS NULL THEN 0 ELSE t6.[SP] END AS 'SP_t6',
CASE WHEN t6.[T] IS NULL THEN 0 ELSE t6.[T] END AS 'T_t6',
CASE WHEN t6.[Fissaggio] IS NULL THEN '' ELSE t6.[Fissaggio] END AS 'Fissaggio_t6',
CASE WHEN t6.[Ghiera] IS NULL THEN '' ELSE t6.[Ghiera] END AS 'Ghiera_t6',
CASE WHEN t6.[Grileltto] IS NULL THEN '' ELSE t6.[Grileltto] END AS 'Grileltto_t6',
CASE WHEN t6.[Protezione] IS NULL THEN '' ELSE t6.[Protezione] END AS 'Protezione_t6',
CASE WHEN t6.[Note] IS NULL THEN '' ELSE t6.[Note] END AS 'Note_t6',
CASE WHEN t6.[Cannuccia] IS NULL THEN '' ELSE t6.[Cannuccia] END AS 'Cannuccia_t6',

-- Campi t7
CASE WHEN t7.[Densita] IS NULL THEN 0 ELSE t7.[Densita] END AS 'Densita_t7',
CASE WHEN t7.[Viscosita_Dinamica] IS NULL THEN 0 ELSE t7.[Viscosita_Dinamica] END AS 'Viscosita_Dinamica_t7',
CASE WHEN t7.[Conducibilita_Elettrica] IS NULL THEN 0 ELSE t7.[Conducibilita_Elettrica] END AS 'Conducibilita_Elettrica_t7',
CASE WHEN t7.[Categoria] IS NULL THEN '' ELSE t7.[Categoria] END AS 'Categoria_t7',
CASE WHEN t7.[Infiammabile] IS NULL THEN '' ELSE t7.[Infiammabile] END AS 'Infiammabile_t7',
CASE WHEN t7.[Nome_Commerciale] IS NULL THEN '' ELSE t7.[Nome_Commerciale] END AS 'Nome_Commerciale_t7',
CASE WHEN t7.[Viscosità_Cinematica] IS NULL THEN 0 ELSE t7.[Viscosità_Cinematica] END AS 'Viscosità_Cinematica_t7',
CASE WHEN t7.[Corrosivo] IS NULL THEN '' ELSE t7.[Corrosivo] END AS 'Corrosivo_t7',
CASE WHEN t7.[Nocivo/Tossico] IS NULL THEN '' ELSE t7.[Nocivo/Tossico] END AS 'Nocivo/Tossico_t7',
CASE WHEN t7.[Note] IS NULL THEN '' ELSE t7.[Note] END AS 'Note_t7',

-- Campi t8
CASE WHEN t8.[Larghezza] IS NULL THEN 0 ELSE t8.[Larghezza] END AS 'Larghezza_t8',
CASE WHEN t8.[Diametro_Fulcro] IS NULL THEN 0 ELSE t8.[Diametro_Fulcro] END AS 'Diametro_Fulcro_t8',
CASE WHEN t8.[Materiale] IS NULL THEN '' ELSE t8.[Materiale] END AS 'Materiale_t8',
CASE WHEN t8.[Temperatura_Saldatura] IS NULL THEN 0 ELSE t8.[Temperatura_Saldatura] END AS 'Temperatura_Saldatura_t8',
CASE WHEN t8.[Diametro_Esterno] IS NULL THEN 0 ELSE t8.[Diametro_Esterno] END AS 'Diametro_Esterno_t8',

 CASE WHEN t9.[Altezza] IS NULL THEN 0 ELSE t9.[Altezza] END AS 'Altezza_t9',
CASE WHEN t9.[Larghezza] IS NULL THEN 0 ELSE t9.[Larghezza] END AS 'Larghezza_t9',
CASE WHEN t9.[Profondità] IS NULL THEN 0 ELSE t9.[Profondità] END AS 'Profondità_t9',
CASE WHEN t9.[Diametro_Interno] IS NULL THEN 0 ELSE t9.[Diametro_Interno] END AS 'Diametro_Interno_t9',
CASE WHEN t9.[Fissaggio] IS NULL THEN '' ELSE t9.[Fissaggio] END AS 'Fissaggio_t9',
CASE WHEN t9.[Forma] IS NULL THEN '' ELSE t9.[Forma] END AS 'Forma_t9',
CASE WHEN t9.[Materiale] IS NULL THEN '' ELSE t9.[Materiale] END AS 'Materiale_t9',
CASE WHEN t9.[Superficie] IS NULL THEN '' ELSE t9.[Superficie] END AS 'Superficie_t9',
CASE WHEN t9.[Produttore] IS NULL THEN '' ELSE t9.[Produttore] END AS 'Produttore_t9',
CASE WHEN t9.[Codice_produttore] IS NULL THEN '' ELSE t9.[Codice_produttore] END AS 'Codice_Produttore_t9',

-- Campi t10
CASE WHEN t10.iniziale_sigla IS NULL THEN '' ELSE t10.iniziale_sigla END AS 'iniziale_sigla',

-- Campi t11
CASE WHEN t11.onhand IS NULL THEN 0 ELSE cast(t11.onhand as integer) END AS 'onhand',
CASE WHEN t11.u_ubicazione IS NULL THEN '' ELSE t11.u_ubicazione END AS 'u_ubicazione',

-- Campi t12
coalesce( t12.cardname,'') AS 'cardname',

t13.immagine_descrizione

,T14.CARDNAME
,case when t15.cardcode is null then '' else t15.cardcode end as 'codice_bp_principale'
,case when t15.cardname is null then '' else t15.cardname end as 'Cliente_principale'
,coalesce(t17.email,'') as 'Email'

FROM [TIRELLI_40].[DBO].coll_campioni AS T0
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_flaconi] t1 ON t0.id_campione = t1.codice_campione AND t0.tipo_campione = 100
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_tappi] t2 ON t0.id_campione = t2.codice_campione AND t0.tipo_campione = 101
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_sottotappi] t3 ON t0.id_campione = t3.codice_campione AND t0.tipo_campione = 102
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_pompette] t4 ON t0.id_campione = t4.codice_campione AND t0.tipo_campione = 103
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_etichette] t5 ON t0.id_campione = t5.codice_campione AND t0.tipo_campione = 104
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_trigger] t6 ON t0.id_campione = t6.codice_campione AND t0.tipo_campione = 105
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_prodotti] t7 ON t0.id_campione = t7.codice_campione AND t0.tipo_campione = 106
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_film] t8 ON t0.id_campione = t8.codice_campione AND t0.tipo_campione = 107
LEFT JOIN [TIRELLI_40].[DBO].[coll_campioni_copritappi] t9 ON t0.id_campione = t9.codice_campione AND t0.tipo_campione = 108
LEFT JOIN [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t10 ON t10.Id_Tipo_Campione = t0.Tipo_Campione
LEFT JOIN [TIRELLISRLDB].[dbo].oitm t11 ON t11.itemcode = t0.codice_sap 
LEFT JOIN [TIRELLISRLDB].[dbo].ocrd t12 ON t12.cardcode = t0.codice_BP 
LEFT JOIN [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t13 ON t13.id_tipo_campione = t0.tipo_campione
LEFT JOIN [TIRELLISRLDB].[dbo].OCRD T14 ON CAST(T14.CARDCODE AS VARCHAR) = CAST(t0.codice_bp AS VARCHAR) 
LEFT JOIN [TIRELLISRLDB].[dbo].ocrd t15 ON CAST(t15.u_bp_riferimento AS VARCHAR) = CAST(t14.cardcode AS VARCHAR) 
LEFT JOIN [TIRELLI_40].[DBO].COLL_TIPO_CAMPIONE t16 ON t0.TIPO_campione = T16.ID_TIPO_CAMPIONE
LEFT JOIN [TIRELLI_40].[dbo].OHEM t17 ON T17.EMPID = T0.ownerid


WHERE t0.id_campione=" & par_id_Campione & "

        order by T14.CARDNAME ,t10.INIZIALE_SIGLA ,  cast(T0.NOME as integer)"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() = True Then
            ' dettagli.Descrizione = cmd_SAP_reader("itemname")
            dettagli.tipo_nome = cmd_SAP_reader_2("tipo_nome")
            dettagli.tipo_nome_inglese = cmd_SAP_reader_2("tipo_nome_inglese")
            dettagli.nome = cmd_SAP_reader_2("Nome")
            dettagli.immagine = cmd_SAP_reader_2("Immagine")
            dettagli.cardname = cmd_SAP_reader_2("Cardname")
            dettagli.codice_bp = cmd_SAP_reader_2("codice_bp")
            dettagli.iniziale_sigla = cmd_SAP_reader_2("iniziale_sigla")
            dettagli.Descrizione = cmd_SAP_reader_2("Descrizione")
            dettagli.Note = cmd_SAP_reader_2("Note")
            dettagli.Email = cmd_SAP_reader_2("Email")




        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()





        Return dettagli
    End Function

    Public Function check_richieste_campione_aperte(par_id_Campione As Integer)

        Dim risposta As String = "N"

        Dim Cnn1 As New SqlConnection

        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT TOP 1 [Id_richiesta]
      ,[id_campione]
      ,[Insertdate]
      ,[Duedate]
      ,[Owner]
      ,[Q_tot]
      ,[Q_open]
      ,[Status]
  FROM [TIRELLI_40].[DBO].[coll_campioni_richieste]
where status='O' and [id_campione] = " & par_id_Campione & " "

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() = True Then

            risposta = "Y"

        Else
            risposta = "N"
        End If
        cmd_SAP_reader_2.Close()
        Cnn1.Close()


        Return risposta
    End Function

    Public Class Dettagli_Campione
        Public tipo_nome As String
        Public tipo_nome_inglese As String
        Public nome As String
        Public immagine As String
        Public cardname As String
        Public codice_bp As String
        Public iniziale_sigla As String
        Public Descrizione As String
        Public Note As String
        Public Email As String




    End Class
    Private Sub DataGridView1_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseDown
        If e.Button = MouseButtons.Left AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If ModifierKeys = Keys.Shift AndAlso startIndex >= 0 Then
                'Se è premuto Shift, cambia il flag per le righe comprese tra startIndex ed e.RowIndex
                Dim endIndex As Integer = e.RowIndex
                Dim minIndex As Integer = Math.Min(startIndex, endIndex) + 1
                Dim maxIndex As Integer = Math.Max(startIndex, endIndex) - 1

                For i As Integer = minIndex To maxIndex
                    DataGridView1.Rows(i).SetValues(True)
                Next i
            Else
                '  Altrimenti, imposta startIndex alla riga corrente
                startIndex = e.RowIndex
            End If
        End If
    End Sub

    Private Sub DataGridView_1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        ' Controlla se il tasto Shift è stato premuto
        isShiftKeyDown = (e.KeyCode = Keys.ShiftKey)
    End Sub

    Private Sub DataGridView_1_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyUp
        ' Controlla se il tasto Shift è stato rilasciato
        isShiftKeyDown = (e.KeyCode <> Keys.ShiftKey)
    End Sub

    Private Sub DataGridView3_CellMouseDown(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView3.CellMouseDown
        If e.Button = MouseButtons.Left AndAlso e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If ModifierKeys = Keys.Shift AndAlso startIndex >= 0 Then
                'Se è premuto Shift, cambia il flag per le righe comprese tra startIndex ed e.RowIndex
                Dim endIndex As Integer = e.RowIndex
                Dim minIndex As Integer = Math.Min(startIndex, endIndex) + 1
                Dim maxIndex As Integer = Math.Max(startIndex, endIndex) - 1

                For i As Integer = minIndex To maxIndex
                    DataGridView3.Rows(i).SetValues(True)
                Next i
            Else
                '  Altrimenti, imposta startIndex alla riga corrente
                startIndex = e.RowIndex
            End If
        End If
    End Sub

    Private Sub DataGridView_3_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView3.KeyDown
        ' Controlla se il tasto Shift è stato premuto
        isShiftKeyDown = (e.KeyCode = Keys.ShiftKey)
    End Sub

    Private Sub DataGridView_3_KeyUp(sender As Object, e As KeyEventArgs) Handles DataGridView3.KeyUp
        ' Controlla se il tasto Shift è stato rilasciato
        isShiftKeyDown = (e.KeyCode <> Keys.ShiftKey)
    End Sub

    Public Function trova_id_doc_richiesta_campione()

        Dim n_doc As Integer = 1

        Dim Cnn_Campioni As New SqlConnection
        Cnn_Campioni.ConnectionString = Homepage.sap_tirelli
        Cnn_Campioni.Open()

        Dim Cmd_Campioni As New SqlCommand
        Dim Cmd_Campioni_Reader As SqlDataReader

        Cmd_Campioni.Connection = Cnn_Campioni
        Cmd_Campioni.CommandText = "SELECT coalesce(max(coalesce([Id_doc_richiesta],0))+1,1) as 'N'
FROM [TIRELLI_40].[DBO].[coll_campioni_richieste] "

        Cmd_Campioni_Reader = Cmd_Campioni.ExecuteReader

        If Cmd_Campioni_Reader.Read() Then
            n_doc = Cmd_Campioni_Reader("N")
        Else
            n_doc = 1

        End If
        Cmd_Campioni_Reader.Close()
        Cnn_Campioni.Close()

        Return n_doc

    End Function

    Sub elimina_richiesta_campioni(par_id_richiesta As Integer)
        Dim Cnn6 As New SqlConnection
        Cnn6.ConnectionString = Homepage.sap_tirelli
        Cnn6.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn6
        CMD_SAP.CommandText = "DELETE [TIRELLI_40].[DBO].[coll_campioni_richieste] WHERE id_richiesta= " & par_id_richiesta & ""



        ' Esegui il comando
        CMD_SAP.ExecuteNonQuery()
        Cnn6.Close()
    End Sub

    Sub crea_richiesta_campioni(par_id_doc As Integer, par_id_campione As Integer, par_owner As Integer, par_Q As Decimal, par_consegna As Date)
        Dim Cnn6 As New SqlConnection
        Cnn6.ConnectionString = Homepage.sap_tirelli
        Cnn6.Open()
        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = Cnn6

        CMD_SAP.CommandText = "INSERT INTO [TIRELLI_40].[DBO].[coll_campioni_richieste]
        ([Id_doc_richiesta] 
,[id_campione]
        ,[Insertdate]
        ,[Duedate]
        ,[Owner]
        ,[Q_tot]
        ,[Q_open]
        ,[Status])
        VALUES
        (@id_doc_richiesta,@id_campione, getdate(), @Duedate, @Owner, @Q_tot, @Q_open, 'O')"

        ' Aggiungi i parametri con i valori
        CMD_SAP.Parameters.AddWithValue("@id_doc_richiesta", par_id_doc)
        CMD_SAP.Parameters.AddWithValue("@id_campione", par_id_campione)
        CMD_SAP.Parameters.AddWithValue("@Duedate", par_consegna)
        CMD_SAP.Parameters.AddWithValue("@Owner", par_owner)
        CMD_SAP.Parameters.AddWithValue("@Q_tot", par_Q)
        CMD_SAP.Parameters.AddWithValue("@Q_open", par_Q)

        ' Esegui il comando
        CMD_SAP.ExecuteNonQuery()
        Cnn6.Close()
    End Sub




    Private Sub Button3_Click_2(sender As Object, e As EventArgs) Handles Button3.Click

        Dim par_datagridview_provenienza As DataGridView = DataGridView1
        Dim par_datagridview_destinazione As DataGridView = DataGridView2
        ' Itera attraverso le righe della DataGridView "datagridview_odp"
        For Each row As DataGridViewRow In par_datagridview_provenienza.Rows
            ' Verifica se la cella della colonna "seleziona" è flaggata
            If CBool(row.Cells("seleziona").Value) = True Then
                ' Crea una nuova riga nella DataGridView "datagridview1"
                Dim index As Integer = par_datagridview_destinazione.Rows.Add()

                ' Copia i valori dalle colonne necessarie
                par_datagridview_destinazione.Rows(index).Cells("campione_").Value = row.Cells("campione").Value
                par_datagridview_destinazione.Rows(index).Cells("tipo_").Value = ottieni_informazioni_campione(row.Cells("campione").Value).tipo_nome
                par_datagridview_destinazione.Rows(index).Cells("nome_").Value = ottieni_informazioni_campione(row.Cells("campione").Value).iniziale_sigla & ottieni_informazioni_campione(row.Cells("campione").Value).Nome
                par_datagridview_destinazione.Rows(index).Cells("Immagine_").Value = Image.FromFile(Homepage.Percorso_immagini & ottieni_informazioni_campione(row.Cells("campione").Value).immagine)

                par_datagridview_destinazione.Rows(index).Cells("Cliente_").Value = ottieni_informazioni_campione(row.Cells("campione").Value).Cardname
                par_datagridview_destinazione.Rows(index).Cells("Q").Value = 0
                row.Cells("seleziona").Value = False

            End If
        Next
        par_datagridview_provenienza.ClearSelection()
    End Sub

    Private Sub Button5_Click_1(sender As Object, e As EventArgs) Handles Button5.Click
        Dim par_datagridview As DataGridView = DataGridView2
        Dim i As Integer = 0
        Dim id_richiesta_campione As Integer = trova_id_doc_richiesta_campione()
        If DateTimePicker4.Value.Date <> DateTime.Today Then


            While i < par_datagridview.Rows.Count
                Dim row As DataGridViewRow = par_datagridview.Rows(i)

                If row.Cells("Q").Value > 0 Then




                    '  If check_richieste_campione_aperte(row.Cells("Campione_").Value) = "N" Then

                    crea_richiesta_campioni(id_richiesta_campione, row.Cells("Campione_").Value, Homepage.ID_SALVATO, row.Cells("Q").Value, DateTimePicker4.Value)
                    ' Rimuovi la riga e NON incrementare il contatore
                    par_datagridview.Rows.RemoveAt(i)
                    '  Else
                    ' MsgBox("Esistono già richieste aperte per il campione " & row.Cells("Nome_").Value)
                    ' Incrementa il contatore solo se la riga non viene eliminata
                    ' i += 1
                    '  End If

                Else
                    MsgBox("Il campione " & row.Cells("Nome_").Value & " ha quantità <=0")
                    ' Incrementa il contatore solo se la riga non viene eliminata

                    i += 1
                End If
            End While
            filtra_datagridview_richieste()
            MsgBox("Richieste inserite con successo")
        Else
            MessageBox.Show("La data di consegna dei campioni deve essere diversa da oggi.", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If

    End Sub



    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting

        Dim par_datagridview As DataGridView = DataGridView2
        If par_datagridview.Rows(e.RowIndex).Cells(columnName:="Q").Value > 0 Then



            par_datagridview.Rows(e.RowIndex).Cells(columnName:="Q").Style.BackColor = Color.Lime
        Else
            par_datagridview.Rows(e.RowIndex).Cells(columnName:="Q").Style.BackColor = Color.LightYellow
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If MessageBox.Show($"Sei sicuro di voler eliminare questa richiesta campioni?", "Elimina richiesta", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then


            Dim par_datagridview As DataGridView = DataGridView3
            Dim i As Integer = 0




            While i < par_datagridview.Rows.Count
                Dim row As DataGridViewRow = par_datagridview.Rows(i)

                If row.Cells("Seleziona_").Value = True Then


                    '  If check_richieste_campione_aperte(row.Cells("Campione_").Value) = "N" Then

                    elimina_richiesta_campioni(row.Cells("Id_richiesta").Value)
                    ' Rimuovi la riga e NON incrementare il contatore
                    par_datagridview.Rows.RemoveAt(i)
                    '  Else
                    ' MsgBox("Esistono già richieste aperte per il campione " & row.Cells("Nome_").Value)
                    ' Incrementa il contatore solo se la riga non viene eliminata
                    ' i += 1
                    '  End If

                Else
                    i += 1
                End If
            End While
            filtra_datagridview_richieste()
            MsgBox("Richieste eliminate con successo")
        End If
    End Sub





    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form_nuovo_campione.Show()
        Form_nuovo_campione.inizializza_form()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        prepara_etichette_excel(Homepage.percorso_server & "00-Tirelli 4.0\T4.0vb\Eseguibili\Layout documenti\Campioni.xlsx", "Documento", ComboBox1.Text)
    End Sub

    Sub prepara_etichette_excel(par_percorso_file As String, par_nome_foglio As String, par_lingua As String)
        Dim appExcel As New Excel.Application
        Dim workbook As Excel.Workbook
        Dim tipo_campione_nome As String

        ' Apri il file Excel
        workbook = appExcel.Workbooks.Open(par_percorso_file)
        appExcel.Visible = True
        Dim colonna_traduzioni As Integer
        If par_lingua = "Ita" Then
            colonna_traduzioni = 1
        ElseIf par_lingua = "Eng" Then
            colonna_traduzioni = 2
        ElseIf par_lingua = "Fr" Then
            colonna_traduzioni = 3

        ElseIf par_lingua = "Spa" Then
            colonna_traduzioni = 4

        Else
            colonna_traduzioni = 2
        End If

        Dim testo_base As String
        Dim testo_grassetto As String
        Dim cella As Excel.Range

        'cliente




        Dim lunghezza_base As Long
        Dim lunghezza_grassetto As Long


        Dim par_datagridview As DataGridView = DataGridView3

        Dim contatore As Integer = 0
        Dim prima_riga As Integer = 12
        Dim par_prima_colonna As Integer = 2


        For Each row As DataGridViewRow In par_datagridview.Rows
            ' Verifica se la cella della colonna "seleziona" è flaggata
            If CBool(row.Cells("seleziona_").Value) = True Then
                Dim c As Integer = 0

                If contatore = 0 Then
                    ' Valori iniziali
                    testo_base = workbook.Sheets("Traduzioni").Cells(2, colonna_traduzioni).Value

                    testo_grassetto = ottieni_informazioni_campione(row.Cells("ID_campione_").Value).cardname
                    'testo_destinatario = trova_dettagli_documento(par_numero_documento, par_testata_sap, par_righe_sap).destinatario_fattura

                    ' Determina lunghezze
                    lunghezza_base = Len(testo_base) + 2 ' Include ": " e vbCrLf
                    lunghezza_grassetto = Len(testo_grassetto)

                    ' Imposta il valore della cella
                    cella = workbook.Sheets(par_nome_foglio).Range("B3")
                    cella.Value = testo_base & vbCrLf & testo_grassetto & vbCrLf

                    ' Applica il grassetto solo alla parte desiderata
                    cella.Characters(Start:=lunghezza_base + 1, Length:=lunghezza_grassetto + 2).Font.Bold = True


                    'Indirizzo di consegna
                    testo_base = workbook.Sheets("Traduzioni").Cells(3, colonna_traduzioni).Value & vbCrLf & workbook.Sheets("Traduzioni").Cells(15, colonna_traduzioni).Value
                    cella = workbook.Sheets(par_nome_foglio).Range("B4")
                    cella.Value = testo_base
                    'cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

                    'Richiesta campioni N°
                    testo_base = workbook.Sheets("Traduzioni").Cells(16, colonna_traduzioni).Value
                    testo_grassetto = row.Cells("N_doc").Value
                    cella = workbook.Sheets(par_nome_foglio).Range("B2")
                    cella.Value = testo_base & ": " & testo_grassetto
                    cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

                    'Data 
                    testo_base = workbook.Sheets("Traduzioni").Cells(17, colonna_traduzioni).Value
                    testo_grassetto = Format(Now, "dd/MM/yyyy")
                    cella = workbook.Sheets(par_nome_foglio).Range("F2")
                    cella.Value = testo_base & ": " & testo_grassetto
                    cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True

                    'matricole
                    testo_base = workbook.Sheets("Traduzioni").Cells(4, colonna_traduzioni).Value
                    cella = workbook.Sheets(par_nome_foglio).Range("F3")
                    cella.Value = testo_base

                    'Progetto
                    testo_base = workbook.Sheets("Traduzioni").Cells(5, colonna_traduzioni).Value
                    cella = workbook.Sheets(par_nome_foglio).Range("F4")
                    cella.Value = testo_base

                    'Ns contatto
                    testo_base = workbook.Sheets("Traduzioni").Cells(6, colonna_traduzioni).Value & row.Cells("Mail").Value
                    cella = workbook.Sheets(par_nome_foglio).Range("F8")
                    cella.Value = testo_base

                    'riga
                    testo_base = workbook.Sheets("Traduzioni").Cells(7, colonna_traduzioni).Value
                    cella = workbook.Sheets(par_nome_foglio).Range("B" & prima_riga - 1)
                    cella.Value = testo_base

                    'Id_campione
                    testo_base = workbook.Sheets("Traduzioni").Cells(8, colonna_traduzioni).Value
                    cella = workbook.Sheets(par_nome_foglio).Range("C" & prima_riga - 1)
                    cella.Value = testo_base

                    'Nome
                    testo_base = workbook.Sheets("Traduzioni").Cells(9, colonna_traduzioni).Value
                    cella = workbook.Sheets(par_nome_foglio).Range("D" & prima_riga - 1)
                    cella.Value = testo_base

                    'Immagine
                    testo_base = workbook.Sheets("Traduzioni").Cells(10, colonna_traduzioni).Value
                    cella = workbook.Sheets(par_nome_foglio).Range("E" & prima_riga - 1)
                    cella.Value = testo_base

                    'Q TOT
                    testo_base = workbook.Sheets("Traduzioni").Cells(11, colonna_traduzioni).Value
                    cella = workbook.Sheets(par_nome_foglio).Range("F" & prima_riga - 1)
                    cella.Value = testo_base

                    'Q Open
                    testo_base = workbook.Sheets("Traduzioni").Cells(12, colonna_traduzioni).Value
                    cella = workbook.Sheets(par_nome_foglio).Range("G" & prima_riga - 1)
                    cella.Value = testo_base

                    'Data consegna
                    testo_base = workbook.Sheets("Traduzioni").Cells(13, colonna_traduzioni).Value
                    cella = workbook.Sheets(par_nome_foglio).Range("H" & prima_riga - 1)
                    cella.Value = testo_base

                    'Descrizione
                    testo_base = workbook.Sheets("Traduzioni").Cells(14, colonna_traduzioni).Value
                    cella = workbook.Sheets(par_nome_foglio).Range("I" & prima_riga - 1)
                    cella.Value = testo_base

                End If

                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).Value = contatore + 1

                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter


                c += 1
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).Value = row.Cells("ID_campione_").Value
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                c += 1
                If par_lingua = "Ita" Then
                    tipo_campione_nome = ottieni_informazioni_campione(row.Cells("ID_campione_").Value).tipo_nome
                Else
                    tipo_campione_nome = ottieni_informazioni_campione(row.Cells("ID_campione_").Value).tipo_nome_inglese
                End If


                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).Value = tipo_campione_nome & " " & ottieni_informazioni_campione(row.Cells("ID_campione_").Value).iniziale_sigla & ottieni_informazioni_campione(row.Cells("ID_campione_").Value).Nome
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                c += 1
                Dim immaginePath As String = Homepage.Percorso_immagini & ottieni_informazioni_campione(row.Cells("ID_campione_").Value).immagine

                If IO.File.Exists(immaginePath) Then
                    ' Calcola la posizione della cella in Excel
                    Dim cell As Excel.Range = workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c)
                    Dim cellLeft As Double = cell.Left
                    Dim cellTop As Double = cell.Top
                    Dim cellWidth As Double = cell.Width
                    Dim cellHeight As Double = cell.Height

                    ' Ottieni le dimensioni originali dell'immagine
                    Dim immagine As Image = Image.FromFile(immaginePath)
                    Dim imgWidth As Double = immagine.Width
                    Dim imgHeight As Double = immagine.Height
                    immagine.Dispose()

                    ' Calcola le nuove dimensioni mantenendo le proporzioni
                    Dim scaleWidth As Double = cellWidth / imgWidth
                    Dim scaleHeight As Double = cellHeight / imgHeight
                    Dim scale As Double = Math.Min(scaleWidth, scaleHeight) ' Usa il rapporto più piccolo per mantenere le proporzioni

                    Dim newWidth As Double = imgWidth * scale
                    Dim newHeight As Double = imgHeight * scale

                    ' Calcola gli offset per centrare l'immagine nella cella
                    Dim leftOffset As Double = (cellWidth - newWidth) / 2
                    Dim topOffset As Double = (cellHeight - newHeight) / 2

                    ' Inserisci l'immagine centrata
                    workbook.Sheets(par_nome_foglio).Shapes.AddPicture(immaginePath,
        Microsoft.Office.Core.MsoTriState.msoFalse,
        Microsoft.Office.Core.MsoTriState.msoCTrue,
        cellLeft + leftOffset, cellTop + topOffset, newWidth, newHeight)
                Else
                    ' MsgBox("Immagine non trovata: " & immaginePath, MsgBoxStyle.Critical)
                End If
                c += 1

                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).Value = row.Cells("Q_tot").Value
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter

                c += 1

                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).Value = row.Cells("Q_open").Value
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                c += 1
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).Value = row.Cells("cONSEGNA").Value
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                c += 1
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).Value = ottieni_informazioni_campione(row.Cells("ID_campione_").Value).Descrizione
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
                workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore, par_prima_colonna + c).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter
                c += 1
                row.Cells("seleziona_").Value = False
                contatore += 1



            End If
        Next

        ' Aggiungi bordi a tutte le celle e al perimetro della tabella
        Dim rangeTable As Excel.Range = workbook.Sheets(par_nome_foglio).Range(
    workbook.Sheets(par_nome_foglio).Cells(prima_riga, par_prima_colonna),
    workbook.Sheets(par_nome_foglio).Cells(prima_riga + contatore - 1, par_prima_colonna + 7))

        ' Bordi esterni
        With rangeTable.Borders(Excel.XlBordersIndex.xlEdgeLeft)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With rangeTable.Borders(Excel.XlBordersIndex.xlEdgeTop)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With rangeTable.Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With rangeTable.Borders(Excel.XlBordersIndex.xlEdgeRight)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
        End With

        ' Bordi interni
        With rangeTable.Borders(Excel.XlBordersIndex.xlInsideHorizontal)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
        End With
        With rangeTable.Borders(Excel.XlBordersIndex.xlInsideVertical)
            .LineStyle = Excel.XlLineStyle.xlContinuous
            .Weight = Excel.XlBorderWeight.xlThin
        End With




        'Data

        testo_base = workbook.Sheets("Traduzioni").Cells(7, colonna_traduzioni).Value
        testo_grassetto = Format(Now, "dd/MM/yyyy")
        cella = workbook.Sheets(par_nome_foglio).Range("G3")
        cella.Value = testo_base & ": " & testo_grassetto
        cella.Characters(Start:=Len(testo_base) + 2, Length:=Len(cella.Value) - (Len(testo_base) + 2) + 9).Font.Bold = True




    End Sub

    Function EsportaFogliInPDF(workbook As Excel.Workbook, par_nome_foglio As String, par_percorso_file As String) As String
        Dim fogliDaStampare As New List(Of Excel.Worksheet)
        Dim pdfFilePath As String = ""
        Dim sheetGTCENG As Excel.Worksheet = Nothing
        Dim sheetGTCFR As Excel.Worksheet = Nothing

        Try


            ' Aggiungi i fogli da includere nel PDF (escludendo "traduzioni")
            Dim sheet1 As Excel.Worksheet = workbook.Sheets(par_nome_foglio)


            fogliDaStampare.Add(sheet1)


            ' Determina il nome del file PDF
            Dim basePdfFilePath As String = Homepage.percorso_offerte_vendita & "PROVA" & "_" & "123" & ".pdf"
            Dim version As Integer = 1

            pdfFilePath = basePdfFilePath
            While IO.File.Exists(pdfFilePath)
                pdfFilePath = Homepage.percorso_offerte_vendita & "PROVA" & "_" & "123" & "_" & version.ToString("D2") & ".pdf"
                version += 1
            End While

            ' Esporta i fogli selezionati in PDF
            workbook.ExportAsFixedFormat(Type:=Excel.XlFixedFormatType.xlTypePDF,
                                      Filename:=pdfFilePath,
                                      Quality:=Excel.XlFixedFormatQuality.xlQualityStandard,
                                      IncludeDocProperties:=True,
                                      IgnorePrintAreas:=False,
                                      OpenAfterPublish:=True) ' Apri il PDF dopo l'esportazione

            ' Determina il nome del file Excel
            Dim excelFilePath As String = Homepage.percorso_offerte_vendita & "PROVA" & "_" & "123" & ".xlsx"
            version = 1
            While IO.File.Exists(excelFilePath)
                excelFilePath = Homepage.percorso_offerte_vendita & "PROVA" & "_" & "123" & "_" & version.ToString("D2") & ".xlsx"
                version += 1
            End While

            ' Salva l'Excel nel percorso specificato
            workbook.SaveAs(excelFilePath)

        Catch ex As Exception
            MsgBox("Errore durante l'esportazione in PDF: " & ex.Message)
        Finally
            ' Ripristina la visibilità delle schede "GTC ENG" e "GTC FR"
            If sheetGTCENG IsNot Nothing Then sheetGTCENG.Visible = Excel.XlSheetVisibility.xlSheetVisible
            If sheetGTCFR IsNot Nothing Then sheetGTCFR.Visible = Excel.XlSheetVisibility.xlSheetVisible

            ' Rilascia le risorse
            For Each sheet As Excel.Worksheet In fogliDaStampare
                ReleaseObject(sheet)
            Next
        End Try

        Return pdfFilePath
    End Function




    ' Funzione per rilasciare le risorse COM


    ' Funzione per rilasciare le risorse COM
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim par_datagrdiview As DataGridView = DataGridView1
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = par_datagrdiview.Columns.IndexOf(Immagine) Then



                Form_campione_visualizza.id_campione = par_datagrdiview.Rows(e.RowIndex).Cells(columnName:="Campione").Value
                Form_campione_visualizza.Show()
                Form_campione_visualizza.inizializza_form()

            End If


        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Dim par_datagrdiview As DataGridView = DataGridView2
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = par_datagrdiview.Columns.IndexOf(Immagine_) Then



                Form_campione_visualizza.id_campione = par_datagrdiview.Rows(e.RowIndex).Cells(columnName:="Campione_").Value
                Form_campione_visualizza.Show()
                Form_campione_visualizza.inizializza_form()

            End If


        End If
    End Sub



    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        Dim par_datagrdiview As DataGridView = DataGridView3
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = par_datagrdiview.Columns.IndexOf(Immagine__) Then



                Form_campione_visualizza.id_campione = par_datagrdiview.Rows(e.RowIndex).Cells(columnName:="ID_Campione_").Value
                Form_campione_visualizza.Show()
                Form_campione_visualizza.inizializza_form()

            End If


        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        Movimento_campioni.Show()


        Dim par_datagridview_provenienza As DataGridView = DataGridView2
        Dim par_datagridview_destinazione As DataGridView = Movimento_campioni.DataGridView1
        ' Itera attraverso le righe della DataGridView "datagridview_odp"
        For Each row As DataGridViewRow In par_datagridview_provenienza.Rows
            ' Verifica se la cella della colonna "seleziona" è flaggata
            If row.Cells("Q").Value > 0 Then
                ' Crea una nuova riga nella DataGridView "datagridview1"
                Dim index As Integer = par_datagridview_destinazione.Rows.Add()
                ' Copia i valori dalle colonne necessarie
                par_datagridview_destinazione.Rows(index).Cells("sel").Value = False
                par_datagridview_destinazione.Rows(index).Cells("id_Campione").Value = row.Cells("campione_").Value
                par_datagridview_destinazione.Rows(index).Cells("tipo").Value = ottieni_informazioni_campione(row.Cells("campione_").Value).tipo_nome
                par_datagridview_destinazione.Rows(index).Cells("Descrizione").Value = ottieni_informazioni_campione(row.Cells("campione_").Value).iniziale_sigla & ottieni_informazioni_campione(row.Cells("campione_").Value).Nome
                par_datagridview_destinazione.Rows(index).Cells("Immagine").Value = Image.FromFile(Homepage.Percorso_immagini & ottieni_informazioni_campione(row.Cells("campione_").Value).immagine)
                par_datagridview_destinazione.Rows(index).Cells("Cliente").Value = ottieni_informazioni_campione(row.Cells("campione_").Value).Cardname
                par_datagridview_destinazione.Rows(index).Cells("Q_trasf").Value = row.Cells("Q").Value
                ' row.Cells("seleziona").Value = False
            Else
                MsgBox("Il campione " & ottieni_informazioni_campione(row.Cells("campione_").Value).iniziale_sigla & ottieni_informazioni_campione(row.Cells("campione_").Value).Nome & " ha quantità =0")

            End If
        Next
        par_datagridview_provenienza.ClearSelection()

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click

        Movimento_campioni.Show()
        'TableLayoutPanel6.Visible = False


        Dim par_datagridview_provenienza As DataGridView = DataGridView3
        Dim par_datagridview_destinazione As DataGridView = Movimento_campioni.DataGridView1
        ' Itera attraverso le righe della DataGridView "datagridview_odp"
        For Each row As DataGridViewRow In par_datagridview_provenienza.Rows
            ' Verifica se la cella della colonna "seleziona" è flaggata
            If row.Cells("Seleziona_").Value = True Then
                ' Crea una nuova riga nella DataGridView "datagridview1"
                Dim index As Integer = par_datagridview_destinazione.Rows.Add()

                ' Copia i valori dalle colonne necessarie
                par_datagridview_destinazione.Rows(index).Cells("sel").Value = False
                par_datagridview_destinazione.Rows(index).Cells("id_Campione").Value = row.Cells("id_Campione_").Value
                par_datagridview_destinazione.Rows(index).Cells("tipo").Value = ottieni_informazioni_campione(row.Cells("id_Campione_").Value).tipo_nome
                par_datagridview_destinazione.Rows(index).Cells("Descrizione").Value = ottieni_informazioni_campione(row.Cells("id_Campione_").Value).iniziale_sigla & ottieni_informazioni_campione(row.Cells("id_Campione_").Value).Nome
                par_datagridview_destinazione.Rows(index).Cells("Immagine").Value = Image.FromFile(Homepage.Percorso_immagini & ottieni_informazioni_campione(row.Cells("id_Campione_").Value).immagine)
                par_datagridview_destinazione.Rows(index).Cells("Cliente").Value = ottieni_informazioni_campione(row.Cells("id_Campione_").Value).Cardname
                par_datagridview_destinazione.Rows(index).Cells("Q_trasf").Value = row.Cells("Q_Open").Value
                par_datagridview_destinazione.Rows(index).Cells("ID_richiesta").Value = row.Cells("Id_richiesta").Value

            End If


        Next
        Movimento_campioni.ComboBox5.SelectedIndex = 1
        par_datagridview_provenienza.ClearSelection()

    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged

    End Sub



    Private Sub DataGridView4_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView4.CellClick
        Dim par_datagrdiview As DataGridView = DataGridView4
        If e.RowIndex >= 0 Then

            If e.ColumnIndex = par_datagrdiview.Columns.IndexOf(Immagine_camp) Then


                Form_campione_visualizza.id_campione = par_datagrdiview.Rows(e.RowIndex).Cells(columnName:="Id_Campione__").Value
                Form_campione_visualizza.Show()
                Form_campione_visualizza.inizializza_form()

            End If


        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        filtra_datagridview_richieste()
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        filtra_datagridview_campioni()
    End Sub
End Class