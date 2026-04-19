Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Windows.Documents
Imports System.Windows.Media.Media3D
Imports System.DirectoryServices.ActiveDirectory
Imports System.Reflection.Emit


Public Class Form_campione_visualizza
    Public id_campione As Integer
    Public Immagine_Esistente As String
    Public percorso_immagine As String
    Public codice_bp_SELEZIONATO As String
    Public codice_bp_jgal_SELEZIONATO As String
    Public immagine_caricata As Integer = 0
    Public tipo_campione As Integer
    Public Elenco_Tipo_Campioni(1000) As Integer
    Public Stringa_Immagine As String
    Public Sel_Stampante As New PrintDialog
    Public Stampante_Selezionata As Boolean
    Private blocco_tab = 0
    Public altezza_Scontrino As Integer
    Public larghezza_scontrino As Integer
    Public numero_combinazioni As Integer
    Private num_collaudati As Integer
    Public cliente_cambiato As Boolean = False

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
    End Sub

    Private Sub Form_campione_visualizza_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        riempi_combobox_tipo_campione()

    End Sub
    Sub inizializza_form()


        ' 1. cliente_cambiato = False
        cliente_cambiato = False


        ' 2. blocco_tab = 0
        blocco_tab = 0


        ' 3. compila_scheda_campione(id_campione)
        compila_scheda_campione(id_campione)


        ' 4. Scheda_tecnica.riempi_datagridview_campioni
        Scheda_tecnica.riempi_datagridview_campioni(DataGridView3, codice_bp_SELEZIONATO, "", "", Homepage.Percorso_immagini, Homepage.sap_tirelli)


        ' 6. riempi_datagridview_combinazioni(id_campione)
        riempi_datagridview_combinazioni(id_campione)

        Form_gestione_campioni.riempi_movimentazioni_campioni("", "", DataGridView4, "", "", "", id_campione)




        ' 7. Label3.Text = id_campione
        Label3.Text = "W" & id_campione.ToString().PadLeft(5, "0"c)


        ' 8. blocco_tab = 1
        blocco_tab = 1


    End Sub
    Sub compila_scheda_campione(par_codice_campione As String)

        Dim Cnn_Campioni As New SqlConnection
        Cnn_Campioni.ConnectionString = Homepage.sap_tirelli
        Cnn_Campioni.Open()
        Dim Cmd_Campioni As New SqlCommand
        Dim Cmd_Campioni_Reader As SqlDataReader

        Cmd_Campioni.Connection = Cnn_Campioni
        Cmd_Campioni.CommandText = "SELECT       T0.[Id_Campione],
CASE WHEN T0.[Codice_BP] IS NULL THEN '' ELSE T0.[Codice_BP] END AS 'Codice_BP',
    CASE WHEN T0.[Nome] IS NULL THEN '' ELSE T0.[Nome] END AS 'Nome',
    CASE WHEN T0.[Descrizione] IS NULL THEN '' ELSE T0.[Descrizione] END AS 'Descrizione',
    coalesce(T0.[Codice_SAP],'') AS 'Codice_SAP',
    CASE WHEN T0.[Tipo_Campione] IS NULL THEN '' ELSE T0.[Tipo_Campione] END AS 'Tipo_Campione',
    coalesce(case when t0.immagine ='' then 'N_A.JPG' else t0.immagine end  ,'N_A.JPG') AS 'immagine'
,coalesce(t0.insertdate,'') as 'insertdate'
,coalesce(t0.updatedate,'') as 'updatedate'
,coalesce(t0.ownerid,0) as 'ownerid'
,coalesce(concat(t14.lastname,' ',t14.firstname),'') as 'Ownername'
,coalesce(t0.note,'') as 'Note'


 ,CASE WHEN t1.[Altezza] IS NULL THEN 0 ELSE t1.[Altezza] END AS 'Altezza_t1',
    CASE WHEN t1.[Larghezza] IS NULL THEN 0 ELSE t1.[Larghezza] END AS 'Larghezza_t1',
    CASE WHEN t1.[Profondita] IS NULL THEN 0 ELSE t1.[Profondita] END AS 'Profondita_t1',
    CASE WHEN t1.[Diametro_Interno] IS NULL THEN 0 ELSE t1.[Diametro_Interno] END AS 'Diametro_Interno_t1',
  CASE WHEN t1.[Diametro_Esterno] IS NULL THEN 0 ELSE t1.[Diametro_Esterno] END AS 'Diametro_Esterno_t1',
CASE WHEN t1.[Volume] IS NULL THEN 0 ELSE t1.[Volume] END AS 'Volume_t1',
CASE WHEN t1.[Spazio_Testa] IS NULL THEN 0 ELSE t1.[Spazio_Testa] END AS 'Spazio_Testa_t1',
CASE WHEN t1.[Materiale] IS NULL THEN '' ELSE t1.[Materiale] END AS 'Materiale_t1',
CASE WHEN t1.[Forma] IS NULL THEN '' ELSE t1.[Forma] END AS 'Forma_t1',
CASE WHEN t1.[Sezione] IS NULL THEN '' ELSE t1.[Sezione] END AS 'Sezione_t1',
CASE WHEN t1.[Superficie] IS NULL THEN '' ELSE t1.[Superficie] END AS 'Superficie_t1',
CASE WHEN t1.[Produttore] IS NULL THEN '' ELSE t1.[Produttore] END AS 'Produttore_t1',
CASE WHEN t1.[Codice_Produttore] IS NULL THEN '' ELSE t1.[Codice_Produttore] END AS 'Codice_Produttore_t1',
CASE WHEN t1.[Collo_Centrato] IS NULL THEN '' ELSE t1.[Collo_Centrato] END AS 'Collo_Centrato_t1',
CASE WHEN t1.[Tipo_Tappo] IS NULL THEN '' ELSE t1.[Tipo_Tappo] END AS 'Tipo_Tappo_t1',
CASE WHEN t1.[Filettatura] IS NULL THEN '' ELSE t1.[Filettatura] END AS 'Filettatura_t1',
CASE WHEN t1.[Diametro_Esterno_Fil] IS NULL THEN 0 ELSE t1.[Diametro_Esterno_Fil] END AS 'Diametro_Esterno_Fil_t1',
CASE WHEN t1.[Passo] IS NULL THEN 0 ELSE t1.[Passo] END AS 'Passo_t1',
CASE WHEN t1.[Num_Principi] IS NULL THEN 0 ELSE t1.[Num_Principi] END AS 'Num_Principi_t1',

CASE WHEN t2.[Altezza] IS NULL THEN 0 ELSE t2.[Altezza] END AS 'Altezza_t2',
CASE WHEN t2.[Larghezza] IS NULL THEN 0 ELSE t2.[Larghezza] END AS 'Larghezza_t2',
CASE WHEN t2.[Profondità] IS NULL THEN 0 ELSE t2.[Profondità] END AS 'Profondità_t2',
CASE WHEN t2.[Diametro_Interno] IS NULL THEN 0 ELSE t2.[Diametro_Interno] END AS 'Diametro_Interno_t2',
CASE WHEN t2.[Fissaggio] IS NULL THEN '' ELSE t2.[Fissaggio] END AS 'Fissaggio_t2',
CASE WHEN t2.[Forma] IS NULL THEN '' ELSE t2.[Forma] END AS 'Forma_t2',
CASE WHEN t2.[Materiale] IS NULL THEN '' ELSE t2.[Materiale] END AS 'Materiale_t2',
CASE WHEN t2.[Superficie] IS NULL THEN '' ELSE t2.[Superficie] END AS 'Superficie_t2',
CASE WHEN t2.[Produttore] IS NULL THEN '' ELSE t2.[Produttore] END AS 'Produttore_t2',
CASE WHEN t2.[Codice_Produttore] IS NULL THEN '' ELSE t2.[Codice_Produttore] END AS 'Codice_Produttore_t2',

-- Campi t3
CASE WHEN t3.[Altezza] IS NULL THEN 0 ELSE t3.[Altezza] END AS 'Altezza_t3',
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
CASE WHEN t4.[Copritappo] IS NULL THEN '' ELSE t4.[Copritappo] END AS 'Copritappo_t4',

-- Campi t5
CASE WHEN t5.[Altezza] IS NULL THEN 0 ELSE t5.[Altezza] END AS 'Altezza_t5',
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

 CASE WHEN t15.[Altezza] IS NULL THEN 0 ELSE t15.[Altezza] END AS 'Altezza_t15',
    CASE WHEN t15.[Larghezza] IS NULL THEN 0 ELSE t15.[Larghezza] END AS 'Larghezza_t15',
    CASE WHEN t15.[Profondita] IS NULL THEN 0 ELSE t15.[Profondita] END AS 'Profondita_t15',

CASE WHEN t15.[Volume] IS NULL THEN 0 ELSE t15.[Volume] END AS 'Volume_t15',

CASE WHEN t15.[Materiale] IS NULL THEN '' ELSE t15.[Materiale] END AS 'Materiale_t15',
CASE WHEN t15.[Forma] IS NULL THEN '' ELSE t15.[Forma] END AS 'Forma_t15',
CASE WHEN t15.[Sezione] IS NULL THEN '' ELSE t15.[Sezione] END AS 'Sezione_t15',
CASE WHEN t15.[Superficie] IS NULL THEN '' ELSE t15.[Superficie] END AS 'Superficie_t15',
CASE WHEN t15.[Produttore] IS NULL THEN '' ELSE t15.[Produttore] END AS 'Produttore_t15',
CASE WHEN t15.[Codice_Produttore] IS NULL THEN '' ELSE t15.[Codice_Produttore] END AS 'Codice_Produttore_t15',


-- Campi t10
CASE WHEN t10.iniziale_sigla IS NULL THEN '' ELSE t10.iniziale_sigla END AS 'iniziale_sigla',

-- Campi t11
CASE WHEN t11.onhand IS NULL THEN 0 ELSE cast(t11.onhand as integer) END AS 'onhand',
CASE WHEN t11.u_ubicazione IS NULL THEN '' ELSE t11.u_ubicazione END AS 'u_ubicazione',

-- Campi t12
CASE WHEN t12.cardname IS NULL THEN '' ELSE t12.cardname END AS 'cardname',

t13.immagine_descrizione

FROM [Tirelli_40].[dbo].[coll_campioni] AS T0
left join [Tirelli_40].[dbo].[coll_campioni_flaconi] t1 on t0.id_campione=t1.codice_campione and t0.tipo_campione=100
left join [Tirelli_40].[dbo].[coll_campioni_tappi] t2 on t0.id_campione=t2.codice_campione and t0.tipo_campione=101
left join [Tirelli_40].[dbo].[Coll_campioni_sottotappi] t3 on t0.id_campione=t3.codice_campione and t0.tipo_campione=102
left join [Tirelli_40].[dbo].[Coll_campioni_pompette] t4 on t0.id_campione=t4.codice_campione and t0.tipo_campione=103
left join [Tirelli_40].[dbo].[Coll_campioni_etichette] t5 on t0.id_campione=t5.codice_campione and t0.tipo_campione=104
left join [Tirelli_40].[dbo].[Coll_campioni_trigger] t6 on t0.id_campione=t6.codice_campione and t0.tipo_campione=105
left join [Tirelli_40].[dbo].[Coll_campioni_prodotti] t7 on t0.id_campione=t7.codice_campione and t0.tipo_campione=106
left join [Tirelli_40].[dbo].[Coll_campioni_film] t8 on t0.id_campione=t8.codice_campione and t0.tipo_campione=107
left join [Tirelli_40].[dbo].[Coll_campioni_copritappi] t9 on t0.id_campione=t9.codice_campione and t0.tipo_campione=108
left join [Tirelli_40].[dbo].[COLL_Tipo_Campione] t10 on t10.Id_Tipo_Campione=t0.Tipo_Campione
left join [TIRELLISRLDB].[dbo].oitm t11 on t11.itemcode=t0.codice_sap
left join [TIRELLISRLDB].[dbo].ocrd t12 on t12.cardcode=t0.codice_BP
left join [Tirelli_40].[dbo].coll_tipo_campione t13 on t13.id_tipo_campione=t0.tipo_campione
left join [TIRELLI_40].[dbo].OHEM t14 on t14.empid=t0.ownerid
left join [Tirelli_40].[dbo].[Coll_campioni_scatole] t15 on t0.id_campione=t15.codice_campione and t0.tipo_campione=109


WHERE t0.id_campione='" & par_codice_campione & "'"
        Cmd_Campioni_Reader = Cmd_Campioni.ExecuteReader

        Cmd_Campioni_Reader.Read()





        Txt_nome.Text = Cmd_Campioni_Reader("nome").ToString()

        Label4.Text = Cmd_Campioni_Reader("ownername").ToString()
        Label5.Text = Cmd_Campioni_Reader("insertdate")
        Label6.Text = Cmd_Campioni_Reader("updatedate")

        Txt_Sigla.Text = Cmd_Campioni_Reader("iniziale_sigla").ToString()

        If Cmd_Campioni_Reader("codice_sap").ToString() <> "" Then
            TextBox114.Text = Cmd_Campioni_Reader("codice_sap").ToString()

        Else

            TextBox114.Text = "W" & id_campione.ToString().PadLeft(5, "0"c)

        End If


        RichTextBox1.Text = Cmd_Campioni_Reader("Note").ToString()



        Txt_descrizione.Text = Cmd_Campioni_Reader("descrizione").ToString()



        TextBox35.Text = Cmd_Campioni_Reader("u_ubicazione").ToString()



        Label1.Text = Cmd_Campioni_Reader("cardname").ToString()


        codice_bp_SELEZIONATO = Cmd_Campioni_Reader("codice_BP")


        tipo_campione = Cmd_Campioni_Reader("tipo_campione").ToString()

        LinkLabel1.Text = Homepage.Percorso_immagini & Cmd_Campioni_Reader("Immagine").ToString
        Stringa_Immagine = Cmd_Campioni_Reader("Immagine").ToString
        riempi_giacenze_campione(par_codice_campione, DataGridView2)

        If Cmd_Campioni_Reader("Immagine_Descrizione").ToString.Length > 1 Then
            Dim MyImage As Bitmap
            Img_Descrizione.SizeMode = PictureBoxSizeMode.Zoom
            Try
                MyImage = New Bitmap(Homepage.Percorso_immagini & Cmd_Campioni_Reader("Immagine_Descrizione").ToString)

            Catch ex As Exception
                MsgBox("Impossibile Aprire l'Immagine d'esempio Selezionata")
            End Try
            Img_Descrizione.Image = CType(MyImage, Image)
        Else
            Img_Descrizione.Image = Nothing
        End If



        Combo_tipo_campione.SelectedIndex = tipo_campione - 99
        blocco_tab = 0
        TabControl1.SelectedIndex = tipo_campione - 100
        blocco_tab = 1



        If tipo_campione = 100 Then
            TextBox3.Text = If(Cmd_Campioni_Reader("altezza_t1") = "0", "", Cmd_Campioni_Reader("altezza_t1").ToString())
            TextBox4.Text = If(Cmd_Campioni_Reader("larghezza_t1") = "0", "", Cmd_Campioni_Reader("larghezza_t1").ToString())
            TextBox5.Text = If(Cmd_Campioni_Reader("profondita_t1") = "0", "", Cmd_Campioni_Reader("profondita_t1").ToString())
            TextBox6.Text = If(Cmd_Campioni_Reader("diametro_interno_t1") = "0", "", Cmd_Campioni_Reader("diametro_interno_t1").ToString())
            TextBox7.Text = If(Cmd_Campioni_Reader("diametro_esterno_t1") = "0", "", Cmd_Campioni_Reader("diametro_esterno_t1").ToString())
            TextBox8.Text = If(Cmd_Campioni_Reader("volume_t1") = "0", "", Cmd_Campioni_Reader("volume_t1").ToString())
            TextBox9.Text = If(Cmd_Campioni_Reader("spazio_testa_t1") = "0", "", Cmd_Campioni_Reader("spazio_testa_t1").ToString())
            TextBox10.Text = If(Cmd_Campioni_Reader("materiale_t1") = "0", "", Cmd_Campioni_Reader("materiale_t1").ToString())
            TextBox11.Text = If(Cmd_Campioni_Reader("forma_t1") = "0", "", Cmd_Campioni_Reader("forma_t1").ToString())
            TextBox12.Text = If(Cmd_Campioni_Reader("sezione_t1") = "0", "", Cmd_Campioni_Reader("sezione_t1").ToString())
            TextBox13.Text = If(Cmd_Campioni_Reader("superficie_t1") = "0", "", Cmd_Campioni_Reader("superficie_t1").ToString())
            TextBox14.Text = If(Cmd_Campioni_Reader("produttore_t1") = "0", "", Cmd_Campioni_Reader("produttore_t1").ToString())
            TextBox15.Text = If(Cmd_Campioni_Reader("codice_produttore_t1") = "0", "", Cmd_Campioni_Reader("codice_produttore_t1").ToString())
            TextBox16.Text = If(Cmd_Campioni_Reader("collo_centrato_t1") = "0", "", Cmd_Campioni_Reader("collo_centrato_t1").ToString())
            TextBox19.Text = If(Cmd_Campioni_Reader("tipo_tappo_t1") = "0", "", Cmd_Campioni_Reader("tipo_tappo_t1").ToString())
            TextBox20.Text = If(Cmd_Campioni_Reader("filettatura_t1") = "0", "", Cmd_Campioni_Reader("filettatura_t1").ToString())
            TextBox21.Text = If(Cmd_Campioni_Reader("diametro_esterno_fil_t1") = "0", "", Cmd_Campioni_Reader("diametro_esterno_fil_t1").ToString())
            TextBox22.Text = If(Cmd_Campioni_Reader("passo_t1") = "0", "", Cmd_Campioni_Reader("passo_t1").ToString())
            TextBox23.Text = If(Cmd_Campioni_Reader("num_principi_t1") = "0", "", Cmd_Campioni_Reader("num_principi_t1").ToString())



        ElseIf tipo_campione = 101 Then
            TextBox17.Text = If(Cmd_Campioni_Reader("altezza_t2") = "0", "", Cmd_Campioni_Reader("altezza_t2").ToString())
            TextBox18.Text = If(Cmd_Campioni_Reader("larghezza_t2") = "0", "", Cmd_Campioni_Reader("larghezza_t2").ToString())
            TextBox1.Text = If(Cmd_Campioni_Reader("profondità_t2") = "0", "", Cmd_Campioni_Reader("profondità_t2").ToString())
            TextBox2.Text = If(Cmd_Campioni_Reader("diametro_interno_t2") = "0", "", Cmd_Campioni_Reader("diametro_interno_t2").ToString())
            TextBox24.Text = If(Cmd_Campioni_Reader("Fissaggio_t2") = "0", "", Cmd_Campioni_Reader("Fissaggio_t2").ToString())
            TextBox25.Text = If(Cmd_Campioni_Reader("Forma_t2") = "0", "", Cmd_Campioni_Reader("Forma_t2").ToString())
            TextBox26.Text = If(Cmd_Campioni_Reader("Materiale_t2") = "0", "", Cmd_Campioni_Reader("Materiale_t2").ToString())
            TextBox27.Text = If(Cmd_Campioni_Reader("Superficie_t2") = "0", "", Cmd_Campioni_Reader("Superficie_t2").ToString())
            TextBox28.Text = If(Cmd_Campioni_Reader("Produttore_t2") = "0", "", Cmd_Campioni_Reader("Produttore_t2").ToString())
            TextBox29.Text = If(Cmd_Campioni_Reader("Codice_produttore_t2") = "0", "", Cmd_Campioni_Reader("Codice_produttore_t2").ToString())

        ElseIf tipo_campione = 102 Then
            TextBox39.Text = If(Cmd_Campioni_Reader("altezza_t3") = "0", "", Cmd_Campioni_Reader("altezza_t3").ToString())
            TextBox38.Text = If(Cmd_Campioni_Reader("larghezza_t3") = "0", "", Cmd_Campioni_Reader("larghezza_t3").ToString())
            TextBox37.Text = If(Cmd_Campioni_Reader("profondità_t3") = "0", "", Cmd_Campioni_Reader("profondità_t3").ToString())
            TextBox36.Text = If(Cmd_Campioni_Reader("diametro_interno_t3") = "0", "", Cmd_Campioni_Reader("diametro_interno_t3").ToString())
            ComboBox6.Text = If(Cmd_Campioni_Reader("vite_pressione_t3") = "0", "", Cmd_Campioni_Reader("vite_pressione_t3").ToString())
            TextBox34.Text = If(Cmd_Campioni_Reader("Forma_t3") = "0", "", Cmd_Campioni_Reader("Forma_t3").ToString())
            TextBox33.Text = If(Cmd_Campioni_Reader("Materiale_t3") = "0", "", Cmd_Campioni_Reader("Materiale_t3").ToString())

        ElseIf tipo_campione = 103 Then

            TextBox30.Text = If(Cmd_Campioni_Reader("A_t4") = "0", "", Cmd_Campioni_Reader("A_t4").ToString())
            TextBox31.Text = If(Cmd_Campioni_Reader("B_t4") = "0", "", Cmd_Campioni_Reader("B_t4").ToString())
            TextBox32.Text = If(Cmd_Campioni_Reader("C_t4") = "0", "", Cmd_Campioni_Reader("C_t4").ToString())
            TextBox40.Text = If(Cmd_Campioni_Reader("D_t4") = "0", "", Cmd_Campioni_Reader("D_t4").ToString())
            TextBox41.Text = If(Cmd_Campioni_Reader("Quota_A_t4") = "0", "", Cmd_Campioni_Reader("Quota_A_t4").ToString())
            TextBox42.Text = If(Cmd_Campioni_Reader("Quota_B_t4") = "0", "", Cmd_Campioni_Reader("Quota_B_t4").ToString())
            TextBox43.Text = If(Cmd_Campioni_Reader("Quota_C_t4") = "0", "", Cmd_Campioni_Reader("Quota_C_t4").ToString())
            TextBox44.Text = If(Cmd_Campioni_Reader("Quota_D_t4") = "0", "", Cmd_Campioni_Reader("Quota_D_t4").ToString())
            TextBox45.Text = If(Cmd_Campioni_Reader("Quota_E_t4") = "0", "", Cmd_Campioni_Reader("Quota_E_t4").ToString())
            TextBox46.Text = If(Cmd_Campioni_Reader("Quota_F_t4") = "0", "", Cmd_Campioni_Reader("Quota_F_t4").ToString())
            TextBox47.Text = If(Cmd_Campioni_Reader("Quota_L_t4") = "0", "", Cmd_Campioni_Reader("Quota_L_t4").ToString())
            TextBox48.Text = If(Cmd_Campioni_Reader("SP_t4") = "0", "", Cmd_Campioni_Reader("SP_t4").ToString())
            TextBox49.Text = If(Cmd_Campioni_Reader("Materiale_t4") = "0", "", Cmd_Campioni_Reader("Materiale_t4").ToString())
            TextBox50.Text = If(Cmd_Campioni_Reader("Tipologia_t4") = "0", "", Cmd_Campioni_Reader("Tipologia_t4").ToString())
            TextBox51.Text = If(Cmd_Campioni_Reader("Superficie_t4") = "0", "", Cmd_Campioni_Reader("Superficie_t4").ToString())
            TextBox52.Text = If(Cmd_Campioni_Reader("Produttore_t4") = "0", "", Cmd_Campioni_Reader("Produttore_t4").ToString())
            TextBox53.Text = If(Cmd_Campioni_Reader("codice_produttore_t4") = "0", "", Cmd_Campioni_Reader("codice_produttore_t4").ToString())
            TextBox54.Text = If(Cmd_Campioni_Reader("Fissaggio_t4") = "0", "", Cmd_Campioni_Reader("Fissaggio_t4").ToString())
            TextBox55.Text = If(Cmd_Campioni_Reader("Ghiera_t4") = "0", "", Cmd_Campioni_Reader("Ghiera_t4").ToString())
            TextBox56.Text = If(Cmd_Campioni_Reader("Copritappo_t4") = "0", "", Cmd_Campioni_Reader("Copritappo_t4").ToString())




        ElseIf tipo_campione = 104 Then
            TextBox75.Text = If(Cmd_Campioni_Reader("Altezza_t5") = "0", "", Cmd_Campioni_Reader("Altezza_t5").ToString())
            TextBox74.Text = If(Cmd_Campioni_Reader("larghezza_t5") = "0", "", Cmd_Campioni_Reader("larghezza_t5").ToString())
            ComboBox1.Text = If(Cmd_Campioni_Reader("Trasparenza_t5") = "0", "", Cmd_Campioni_Reader("Trasparenza_t5").ToString())
            TextBox72.Text = If(Cmd_Campioni_Reader("forma_t5") = "0", "", Cmd_Campioni_Reader("forma_t5").ToString())
            TextBox71.Text = If(Cmd_Campioni_Reader("diametro_esterno_bobina_t5") = "0", "", Cmd_Campioni_Reader("diametro_esterno_bobina_t5").ToString())
            TextBox70.Text = If(Cmd_Campioni_Reader("diametro_interno_bobina_t5") = "0", "", Cmd_Campioni_Reader("diametro_interno_bobina_t5").ToString())
            TextBox69.Text = If(Cmd_Campioni_Reader("Avvolgimento_bobina_t5") = "0", "", Cmd_Campioni_Reader("Avvolgimento_bobina_t5").ToString())
            TextBox68.Text = If(Cmd_Campioni_Reader("materiale_t5") = "0", "", Cmd_Campioni_Reader("materiale_t5").ToString())



        ElseIf tipo_campione = 105 Then
            TextBox78.Text = If(Cmd_Campioni_Reader("A_t6") = "0", "", Cmd_Campioni_Reader("A_t6").ToString())
            TextBox79.Text = If(Cmd_Campioni_Reader("B_t6") = "0", "", Cmd_Campioni_Reader("B_t6").ToString())
            TextBox80.Text = If(Cmd_Campioni_Reader("Quota_S_t6") = "0", "", Cmd_Campioni_Reader("Quota_S_t6").ToString())
            TextBox81.Text = If(Cmd_Campioni_Reader("Quota_H_t6") = "0", "", Cmd_Campioni_Reader("Quota_H_t6").ToString())
            TextBox82.Text = If(Cmd_Campioni_Reader("Quota_L_t6") = "0", "", Cmd_Campioni_Reader("Quota_L_t6").ToString())
            TextBox83.Text = If(Cmd_Campioni_Reader("Quota_W_t6") = "0", "", Cmd_Campioni_Reader("Quota_W_t6").ToString())
            TextBox58.Text = If(Cmd_Campioni_Reader("Quota_V_t6") = "0", "", Cmd_Campioni_Reader("Quota_V_t6").ToString())
            ComboBox2.Text = If(Cmd_Campioni_Reader("Pressione/Vite_t6") = "0", "", Cmd_Campioni_Reader("Pressione/Vite_t6").ToString())
            TextBox76.Text = If(Cmd_Campioni_Reader("Produttore_t6") = "0", "", Cmd_Campioni_Reader("Produttore_t6").ToString())
            TextBox73.Text = If(Cmd_Campioni_Reader("Codice_produttore_t6") = "0", "", Cmd_Campioni_Reader("Codice_produttore_t6").ToString())
            TextBox67.Text = If(Cmd_Campioni_Reader("Materiale_t6") = "0", "", Cmd_Campioni_Reader("Materiale_t6").ToString())
            TextBox66.Text = If(Cmd_Campioni_Reader("SP_t6") = "0", "", Cmd_Campioni_Reader("SP_t6").ToString())
            TextBox65.Text = If(Cmd_Campioni_Reader("T_t6") = "0", "", Cmd_Campioni_Reader("T_t6").ToString())
            TextBox64.Text = If(Cmd_Campioni_Reader("Fissaggio_t6") = "0", "", Cmd_Campioni_Reader("Fissaggio_t6").ToString())
            TextBox63.Text = If(Cmd_Campioni_Reader("Ghiera_t6") = "0", "", Cmd_Campioni_Reader("Ghiera_t6").ToString())
            TextBox62.Text = If(Cmd_Campioni_Reader("Grileltto_t6") = "0", "", Cmd_Campioni_Reader("Grileltto_t6").ToString())
            TextBox61.Text = If(Cmd_Campioni_Reader("Protezione_t6") = "0", "", Cmd_Campioni_Reader("Protezione_t6").ToString())
            TextBox60.Text = If(Cmd_Campioni_Reader("Note_t6") = "0", "", Cmd_Campioni_Reader("Note_t6").ToString())
            TextBox59.Text = If(Cmd_Campioni_Reader("Cannuccia_t6") = "0", "", Cmd_Campioni_Reader("Cannuccia_t6").ToString())


        ElseIf tipo_campione = 106 Then
            TextBox57.Text = If(Cmd_Campioni_Reader("densita_t7") = "0", "", Cmd_Campioni_Reader("densita_t7").ToString())
            TextBox77.Text = If(Cmd_Campioni_Reader("viscosita_dinamica_t7") = "0", "", Cmd_Campioni_Reader("viscosita_dinamica_t7").ToString())
            TextBox100.Text = If(Cmd_Campioni_Reader("conducibilita_elettrica_t7") = "0", "", Cmd_Campioni_Reader("conducibilita_elettrica_t7").ToString())
            TextBox99.Text = If(Cmd_Campioni_Reader("categoria_t7") = "0", "", Cmd_Campioni_Reader("categoria_t7").ToString())
            ComboBox3.Text = If(Cmd_Campioni_Reader("infiammabile_t7") = "0", "", Cmd_Campioni_Reader("infiammabile_t7").ToString())
            TextBox84.Text = If(Cmd_Campioni_Reader("nome_commerciale_t7") = "0", "", Cmd_Campioni_Reader("nome_commerciale_t7").ToString())
            TextBox97.Text = If(Cmd_Campioni_Reader("viscosità_cinematica_t7") = "0", "", Cmd_Campioni_Reader("viscosità_cinematica_t7").ToString())
            ComboBox4.Text = If(Cmd_Campioni_Reader("Corrosivo_t7") = "0", "", Cmd_Campioni_Reader("Corrosivo_t7").ToString())
            ComboBox5.Text = If(Cmd_Campioni_Reader("Nocivo/tossico_t7") = "0", "", Cmd_Campioni_Reader("Nocivo/tossico_t7").ToString())


        ElseIf tipo_campione = 107 Then

            TextBox89.Text = If(Cmd_Campioni_Reader("larghezza_t8") = "0", "", Cmd_Campioni_Reader("larghezza_t8").ToString())
            TextBox90.Text = If(Cmd_Campioni_Reader("diametro_fulcro_t8") = "0", "", Cmd_Campioni_Reader("diametro_fulcro_t8").ToString())
            TextBox85.Text = If(Cmd_Campioni_Reader("materiale_t8") = "0", "", Cmd_Campioni_Reader("materiale_t8").ToString())
            TextBox88.Text = If(Cmd_Campioni_Reader("temperatura_saldatura_t8") = "0", "", Cmd_Campioni_Reader("temperatura_saldatura_t8").ToString())
            TextBox91.Text = If(Cmd_Campioni_Reader("Diametro_esterno_t8") = "0", "", Cmd_Campioni_Reader("Diametro_esterno_t8").ToString())


        ElseIf tipo_campione = 108 Then
            TextBox103.Text = If(Cmd_Campioni_Reader("altezza_t9") = "0", "", Cmd_Campioni_Reader("altezza_t9").ToString())
            TextBox102.Text = If(Cmd_Campioni_Reader("larghezza_t9") = "0", "", Cmd_Campioni_Reader("larghezza_t9").ToString())
            TextBox101.Text = If(Cmd_Campioni_Reader("profondità_t9") = "0", "", Cmd_Campioni_Reader("profondità_t9").ToString())
            TextBox98.Text = If(Cmd_Campioni_Reader("diametro_interno_t9") = "0", "", Cmd_Campioni_Reader("diametro_interno_t9").ToString())
            TextBox96.Text = If(Cmd_Campioni_Reader("Fissaggio_t9") = "0", "", Cmd_Campioni_Reader("Fissaggio_t9").ToString())
            TextBox95.Text = If(Cmd_Campioni_Reader("Forma_t9") = "0", "", Cmd_Campioni_Reader("Forma_t9").ToString())
            TextBox94.Text = If(Cmd_Campioni_Reader("Materiale_t9") = "0", "", Cmd_Campioni_Reader("Materiale_t9").ToString())
            TextBox93.Text = If(Cmd_Campioni_Reader("Superficie_t9") = "0", "", Cmd_Campioni_Reader("Superficie_t9").ToString())
            TextBox92.Text = If(Cmd_Campioni_Reader("Produttore_t9") = "0", "", Cmd_Campioni_Reader("Produttore_t9").ToString())
            TextBox86.Text = If(Cmd_Campioni_Reader("Codice_produttore_t9") = "0", "", Cmd_Campioni_Reader("Codice_produttore_t9").ToString())

        ElseIf tipo_campione = 109 Then
            TextBox112.Text = If(Cmd_Campioni_Reader("altezza_t15") = "0", "", Cmd_Campioni_Reader("altezza_t15").ToString())
            TextBox111.Text = If(Cmd_Campioni_Reader("larghezza_t15") = "0", "", Cmd_Campioni_Reader("larghezza_t15").ToString())
            TextBox110.Text = If(Cmd_Campioni_Reader("profondita_t15") = "0", "", Cmd_Campioni_Reader("profondita_t15").ToString())

            TextBox113.Text = If(Cmd_Campioni_Reader("volume_t15") = "0", "", Cmd_Campioni_Reader("volume_t15").ToString())

            TextBox109.Text = If(Cmd_Campioni_Reader("materiale_t15") = "0", "", Cmd_Campioni_Reader("materiale_t15").ToString())
            TextBox108.Text = If(Cmd_Campioni_Reader("forma_t15") = "0", "", Cmd_Campioni_Reader("forma_t15").ToString())
            TextBox107.Text = If(Cmd_Campioni_Reader("sezione_t15") = "0", "", Cmd_Campioni_Reader("sezione_t15").ToString())
            TextBox106.Text = If(Cmd_Campioni_Reader("superficie_t15") = "0", "", Cmd_Campioni_Reader("superficie_t15").ToString())
            TextBox105.Text = If(Cmd_Campioni_Reader("produttore_t15") = "0", "", Cmd_Campioni_Reader("produttore_t15").ToString())
            TextBox104.Text = If(Cmd_Campioni_Reader("codice_produttore_t15") = "0", "", Cmd_Campioni_Reader("codice_produttore_t15").ToString())





        End If

        If Cmd_Campioni_Reader("Immagine") <> "N_A.JPG" Then
            Immagine_Esistente = Cmd_Campioni_Reader("Immagine")
            If Cmd_Campioni_Reader("Immagine").ToString.Length > 1 Then
                Dim MyImage As Bitmap
                Picture_Campione.SizeMode = PictureBoxSizeMode.Zoom
                Try
                    MyImage = New Bitmap(Homepage.Percorso_immagini & Cmd_Campioni_Reader("Immagine").ToString)
                Catch ex As Exception
                    MsgBox("Impossibile Aprire l'Immagine Selezionata")
                    Immagine_Esistente = ""
                End Try

                Picture_Campione.Image = CType(MyImage, Image)
                percorso_immagine = Homepage.Percorso_immagini & Cmd_Campioni_Reader("Immagine").ToString

                ' Button2.Visible = False
                'Button4.Visible = True

            Else
                '  Button2.Visible = True
                '  Button4.Visible = False
            End If
        Else
            ' Button2.Visible = True
            '  Button4.Visible = False
        End If

        Cnn_Campioni.Close()

        riempi_datagridview_combinazioni(id_campione)

        Form_gestione_campioni.riempi_movimentazioni_campioni("", "", DataGridView4, "", "", "", id_campione)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Picture_Campione.SizeMode = PictureBoxSizeMode.Zoom
        Picture_Campione.Image = Clipboard.GetImage

        If Picture_Campione.Image IsNot Nothing Then
            immagine_caricata = 1



            ' Aggiorna il database
            aggiorna_immagine(id_campione, id_campione, 1)

            MsgBox("Immagine aggiornata con successo")
        Else
            MsgBox("Nessuna immagine trovata negli appunti")
        End If

        Scheda_tecnica.riempi_datagridview_campioni(DataGridView3, codice_bp_SELEZIONATO, "", "", Homepage.Percorso_immagini, Homepage.sap_tirelli)
        Scheda_tecnica.riempi_datagridview_campioni(Scheda_tecnica.DataGridView3, Scheda_tecnica.codice_bp_campione, Scheda_tecnica.bp_code, Scheda_tecnica.final_bp_code, Homepage.Percorso_immagini, Homepage.sap_tirelli)
        Scheda_tecnica.riempi_datagridview_combinazioni(Scheda_tecnica.DataGridView1, Scheda_tecnica.codice_commessa, Homepage.sap_tirelli)

    End Sub

    Sub riempi_combobox_tipo_campione()
        Combo_tipo_campione.Items.Clear()
        Dim Indice As Integer

        Dim Cnn_Tipo As New SqlConnection


        Cnn_Tipo.ConnectionString = Homepage.sap_tirelli
        Cnn_Tipo.Open()

        Dim Cmd_Tipo As New SqlCommand
        Dim Cmd_Tipo_Reader As SqlDataReader

        Indice = 0
        Cmd_Tipo.Connection = Cnn_Tipo
        Cmd_Tipo.CommandText = " SELECT * FROM [Tirelli_40].[dbo].COLL_Tipo_Campione
ORDER BY Id_Tipo_Campione"
        Cmd_Tipo_Reader = Cmd_Tipo.ExecuteReader
        Combo_tipo_campione.Items.Add("")
        Indice = Indice + 1
        Do While Cmd_Tipo_Reader.Read()
            Combo_tipo_campione.Items.Add(Cmd_Tipo_Reader("Descrizione"))
            Elenco_Tipo_Campioni(Indice) = Cmd_Tipo_Reader("Id_Tipo_Campione")
            Indice = Indice + 1
        Loop

        Cmd_Tipo_Reader.Close()
        Cnn_Tipo.Close()

    End Sub

    Private Sub TextBox_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress, TextBox98.KeyPress, TextBox97.KeyPress, TextBox91.KeyPress, TextBox9.KeyPress, TextBox89.KeyPress, TextBox88.KeyPress, TextBox83.KeyPress, TextBox82.KeyPress, TextBox81.KeyPress, TextBox80.KeyPress, TextBox8.KeyPress, TextBox79.KeyPress, TextBox78.KeyPress, TextBox77.KeyPress, TextBox75.KeyPress, TextBox74.KeyPress, TextBox71.KeyPress, TextBox70.KeyPress, TextBox7.KeyPress, TextBox66.KeyPress, TextBox65.KeyPress, TextBox6.KeyPress, TextBox58.KeyPress, TextBox57.KeyPress, TextBox48.KeyPress, TextBox47.KeyPress, TextBox46.KeyPress, TextBox45.KeyPress, TextBox44.KeyPress, TextBox43.KeyPress, TextBox42.KeyPress, TextBox41.KeyPress, TextBox40.KeyPress, TextBox4.KeyPress, TextBox39.KeyPress, TextBox38.KeyPress, TextBox37.KeyPress, TextBox36.KeyPress, TextBox32.KeyPress, TextBox31.KeyPress, TextBox30.KeyPress, TextBox23.KeyPress, TextBox22.KeyPress, TextBox21.KeyPress, TextBox2.KeyPress, TextBox18.KeyPress, TextBox17.KeyPress, TextBox103.KeyPress, TextBox102.KeyPress, TextBox101.KeyPress, TextBox100.KeyPress, TextBox1.KeyPress, TextBox112.KeyPress, TextBox111.KeyPress, TextBox110.KeyPress, TextBox113.KeyPress
        ' Consenti solo numeri interi e il punto decimale
        Dim textBox As TextBox = DirectCast(sender, TextBox)

        If (Not Char.IsDigit(e.KeyChar) AndAlso e.KeyChar <> "." AndAlso e.KeyChar <> ControlChars.Back) Then
            e.Handled = True
        End If

        ' Consenti solo un punto decimale
        If e.KeyChar = "." AndAlso textBox.Text.Contains(".") Then
            e.Handled = True
        End If
    End Sub



    Public Function controllo_esistenza_combinazione_Campione(par_CAMPIONE As Integer)


        Dim controllo As String = "OK"
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "SELECT  [Id_Combinazione]
      ,[Commessa]
      ,[Campione_1]
      ,[Campione_2]
      
    
  FROM [Tirelli_40].[dbo].[COLL_Combinazioni]

   WHERE CAMPIONE_1=" & par_CAMPIONE & " OR CAMPIONE_2=" & par_CAMPIONE & " OR CAMPIONE_3=" & par_CAMPIONE & " OR CAMPIONE_4=" & par_CAMPIONE & " OR CAMPIONE_5=" & par_CAMPIONE & " OR CAMPIONE_6=" & par_CAMPIONE & " OR CAMPIONE_7=" & par_CAMPIONE & " OR CAMPIONE_8=" & par_CAMPIONE & " OR CAMPIONE_9=" & par_CAMPIONE & " OR CAMPIONE_10=" & par_CAMPIONE & "
"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        ' Dim contatore As Integer = 0

        If cmd_SAP_reader_2.Read() Then


            controllo = "NO"

        End If


        cmd_SAP_reader_2.Close()
        Cnn1.Close()


        Return controllo

    End Function


    Sub aggiorna_immagine(par_id_campione As Integer, par_immagine As String, par_nuova_immagine As Integer)

        Dim immagine_import As String = ""

        ' Costruzione del percorso base
        Dim baseFilePath As String = Path.Combine(Homepage.Percorso_immagini, par_immagine & ".jpg")
        immagine_import = par_immagine & ".jpg"
        If File.Exists(baseFilePath) Then
            immagine_import = baseFilePath

            Dim version As Integer = 1
            Dim newPath As String

            Do
                newPath = Path.Combine(Homepage.Percorso_immagini, par_immagine & "_" & version.ToString("D2") & ".jpg")
                If Not File.Exists(newPath) Then
                    immagine_import = par_immagine & "_" & version.ToString("D2") & ".jpg"
                    Exit Do
                End If
                version += 1
            Loop

            'Else
            '    MessageBox.Show("File immagine non trovato: " & baseFilePath)
            '    Exit Sub
        End If
        If par_nuova_immagine > 0 Then
            Picture_Campione.Image.Save(Homepage.Percorso_immagini & immagine_import)
        Else
            immagine_import = ""
        End If

        ' Connessione e aggiornamento
        Using Cnn_Campioni As New SqlConnection(Homepage.sap_tirelli)
            Using Cmd_Campioni As New SqlCommand("UPDATE [Tirelli_40].[dbo].[COLL_Campioni] SET Immagine = @immagine WHERE Id_Campione = @id", Cnn_Campioni)

                Cmd_Campioni.Parameters.AddWithValue("@immagine", immagine_import)
                Cmd_Campioni.Parameters.AddWithValue("@id", par_id_campione)

                Try
                    Cnn_Campioni.Open()
                    Cmd_Campioni.ExecuteNonQuery()
                Catch ex As Exception
                    MessageBox.Show("Errore durante l'aggiornamento del database: " & ex.Message)
                End Try

            End Using
        End Using

    End Sub


    Sub aggiorna_dati_generici_campione(par_id_campione)
        Dim Cnn_Campioni As New SqlConnection
        Cnn_Campioni.ConnectionString = Homepage.sap_tirelli
        Cnn_Campioni.Open()

        Dim Cmd_Campioni As New SqlCommand

        'Stringa_Immagine = par_id_campione & ".jpg"

        'If immagine_caricata = 1 Then
        '    '

        '    If File.Exists(Homepage.Percorso_immagini & Stringa_Immagine) Then

        '        Dim percorsoCompleto As String = Homepage.Percorso_immagini & Stringa_Immagine

        '        File.Delete(Homepage.Percorso_immagini & Stringa_Immagine)
        '        Picture_Campione.Image.Save(Homepage.Percorso_immagini & Stringa_Immagine)
        '        ' Button2.Visible = False
        '    Else
        '        Console.WriteLine(Homepage.Percorso_immagini & Stringa_Immagine & ".jpg")
        '        Picture_Campione.Image.Save(Homepage.Percorso_immagini & Stringa_Immagine)
        '        ' Button2.Visible = True
        '    End If




        'Else

        '    If Picture_Campione.Image IsNot Nothing Then

        '        Try
        '            Picture_Campione.Image.Save(Homepage.Percorso_immagini & Stringa_Immagine)
        '        Catch ex As Exception

        '        End Try



        '        'Button2.Visible = False
        '    Else
        '        Stringa_Immagine = ""
        '        ' Button2.Visible = True
        '    End If


        'End If

        ' Costruisci la query SQL per aggiornare i dati nel database
        Cmd_Campioni.Connection = Cnn_Campioni
        Cmd_Campioni.CommandText = "UPDATE [Tirelli_40].[dbo].COLL_Campioni
                                SET Codice_BP = '" & codice_bp_SELEZIONATO & "',
                                    Nome = '" & Txt_nome.Text & "',
                                    Descrizione = '" & Replace(Txt_descrizione.Text, "'", " ") & "',
                                    Tipo_Campione = " & Elenco_Tipo_Campioni(Combo_tipo_campione.SelectedIndex) & ",
                                    updatedate = GETDATE(),
                                    ownerid = '" & Homepage.ID_SALVATO & "',
                                    note = '" & RichTextBox1.Text & "'
                                WHERE Id_Campione = " & par_id_campione

        ' Esegui la query SQL
        Try
            Cmd_Campioni.ExecuteNonQuery()
        Catch ex As Exception
            MessageBox.Show("Errore durante l'aggiornamento del database: " & ex.Message)
        End Try

        ' Chiudi la connessione al database
        Cnn_Campioni.Close()

        ' Reset del flag immagine_caricata

    End Sub



    Public Function aggiungi_revisione_file(par_percorso_server As String, par_stringa As String, par_estensione_file As String)

        Dim contatore As Integer = 1

        Dim stringa As String


        stringa = par_stringa & contatore

        If File.Exists(par_percorso_server & stringa & par_estensione_file) Then
            stringa = aggiungi_revisione_file(par_percorso_server, stringa, par_estensione_file)
        End If

        stringa = stringa & par_estensione_file

        Return stringa

    End Function

    Sub elimina_Campione(par_id_campione As Integer)
        Dim Cnn_Campioni As New SqlConnection
        Cnn_Campioni.ConnectionString = Homepage.sap_tirelli
        Cnn_Campioni.Open()

        Dim Cmd_Campioni As New SqlCommand
        Cmd_Campioni.Connection = Cnn_Campioni
        Cmd_Campioni.CommandText = "delete [Tirelli_40].[DBO].COLL_Campioni

WHERE Id_Campione = " & par_id_campione & ""

        Cmd_Campioni.ExecuteNonQuery()


        Cmd_Campioni.CommandText = "delete [Tirelli_40].[DBO].COLL_Campioni_flaconi

WHERE codice_Campione = " & par_id_campione & ""

        Cmd_Campioni.ExecuteNonQuery()

        Cmd_Campioni.CommandText = "delete [Tirelli_40].[DBO].COLL_Campioni_tappi

WHERE codice_Campione = " & par_id_campione & ""

        Cmd_Campioni.ExecuteNonQuery()

        Cmd_Campioni.CommandText = "delete [Tirelli_40].[DBO].COLL_Campioni_sottotappi

WHERE codice_Campione = " & par_id_campione & ""

        Cmd_Campioni.ExecuteNonQuery()

        Cmd_Campioni.CommandText = "delete [Tirelli_40].[DBO].COLL_Campioni_copritappi

WHERE codice_Campione = " & par_id_campione & ""

        Cmd_Campioni.ExecuteNonQuery()

        Cmd_Campioni.CommandText = "delete [Tirelli_40].[DBO].COLL_Campioni_film

WHERE codice_Campione = " & par_id_campione & ""

        Cmd_Campioni.ExecuteNonQuery()

        Cmd_Campioni.CommandText = "delete [Tirelli_40].[DBO].COLL_Campioni_prodotti

WHERE codice_Campione = " & par_id_campione & ""

        Cmd_Campioni.ExecuteNonQuery()

        Cmd_Campioni.CommandText = "delete [Tirelli_40].[DBO].COLL_Campioni_etichette

WHERE codice_Campione = " & par_id_campione & ""


        Cmd_Campioni.ExecuteNonQuery()


        Cnn_Campioni.Close()
    End Sub





    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        immagine_caricata = 0
        ' Controlla se c'è un'immagine assegnata
        If Picture_Campione.Image IsNot Nothing Then
            ' Rilascia le risorse dell'immagine
            Picture_Campione.Image.Dispose()
            ' Imposta l'immagine a Nothing
            Picture_Campione.Image = Nothing
        End If
        aggiorna_immagine(id_campione, "", 0)
        MsgBox("Immagine cancellata con successo")
        ' Ripristina eventuali altri controlli (ad esempio il LinkLabel)
        LinkLabel1.Text = ""
        Scheda_tecnica.riempi_datagridview_campioni(DataGridView3, codice_bp_SELEZIONATO, "", "", Homepage.Percorso_immagini, Homepage.sap_tirelli)
        Scheda_tecnica.riempi_datagridview_campioni(Scheda_tecnica.DataGridView3, Scheda_tecnica.codice_bp_campione, Scheda_tecnica.bp_code, Scheda_tecnica.final_bp_code, Homepage.Percorso_immagini, Homepage.sap_tirelli)
        Scheda_tecnica.riempi_datagridview_combinazioni(Scheda_tecnica.DataGridView1, Scheda_tecnica.codice_commessa, Homepage.sap_tirelli)

    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Form_nuovo_campione.check_univocità_campione(Txt_nome.Text, Elenco_Tipo_Campioni(Combo_tipo_campione.SelectedIndex), codice_bp_SELEZIONATO, codice_bp_jgal_SELEZIONATO, id_campione)

        If Form_nuovo_campione.Blocco_univocità = "Y" Then
            MsgBox("Questo campione per questo cliente risulta già")
            Return
        End If
        aggiorna_dati_generici_campione(id_campione)

        If Combo_tipo_campione.SelectedIndex = 1 Then

            aggiorna_flacone(id_campione)

        ElseIf Combo_tipo_campione.SelectedIndex = 2 Then
            aggiorna_tappo(id_campione)
        ElseIf Combo_tipo_campione.SelectedIndex = 3 Then

            aggiorna_sottotappo(id_campione)

        ElseIf Combo_tipo_campione.SelectedIndex = 4 Then
            aggiorna_pompetta(id_campione)
        ElseIf Combo_tipo_campione.SelectedIndex = 5 Then
            aggiorna_etichetta(id_campione)
        ElseIf Combo_tipo_campione.SelectedIndex = 6 Then
            aggiorna_trigger(id_campione)
        ElseIf Combo_tipo_campione.SelectedIndex = 7 Then
            aggiorna_prodotto(id_campione)
        ElseIf Combo_tipo_campione.SelectedIndex = 8 Then
            aggiorna_film(id_campione)
        ElseIf Combo_tipo_campione.SelectedIndex = 9 Then
            aggiorna_copritappo(id_campione)
        ElseIf Combo_tipo_campione.SelectedIndex = 10 Then
            aggiorna_scatola(id_campione)
        End If

        MsgBox("Dati aggiornati con successo")



        Scheda_tecnica.riempi_datagridview_campioni(DataGridView3, codice_bp_SELEZIONATO, "", "", Homepage.Percorso_immagini, Homepage.sap_tirelli)
        Scheda_tecnica.riempi_datagridview_campioni(Scheda_tecnica.DataGridView3, Scheda_tecnica.codice_bp_campione, Scheda_tecnica.bp_code, Scheda_tecnica.final_bp_code, Homepage.Percorso_immagini, Homepage.sap_tirelli)
        Scheda_tecnica.riempi_datagridview_combinazioni(Scheda_tecnica.DataGridView1, Scheda_tecnica.codice_commessa, Homepage.sap_tirelli)

    End Sub

    Sub aggiorna_flacone(par_id_campione As Integer)

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN


        CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].Coll_campioni_flaconi
SET altezza = " & If(String.IsNullOrEmpty(TextBox3.Text), "NULL", Replace(TextBox3.Text, ",", ".")) & ",
    larghezza = " & If(String.IsNullOrEmpty(TextBox4.Text), "NULL", Replace(TextBox4.Text, ",", ".")) & ",
    profondita = " & If(String.IsNullOrEmpty(TextBox5.Text), "NULL", Replace(TextBox5.Text, ",", ".")) & ",
    diametro_interno = " & If(String.IsNullOrEmpty(TextBox6.Text), "NULL", Replace(TextBox6.Text, ",", ".")) & ",
    diametro_esterno = " & If(String.IsNullOrEmpty(TextBox7.Text), "NULL", Replace(TextBox7.Text, ",", ".")) & ",
    volume = " & If(String.IsNullOrEmpty(TextBox8.Text), "NULL", Replace(TextBox8.Text, ",", ".")) & ",
    spazio_testa = " & If(String.IsNullOrEmpty(TextBox9.Text), "NULL", Replace(TextBox9.Text, ",", ".")) & ",
    materiale = " & If(String.IsNullOrEmpty(TextBox10.Text), "NULL", "'" & Replace(TextBox10.Text, ",", ".") & "'") & ",
    forma = " & If(String.IsNullOrEmpty(TextBox11.Text), "NULL", "'" & Replace(TextBox11.Text, ",", ".") & "'") & ",
    sezione = " & If(String.IsNullOrEmpty(TextBox12.Text), "NULL", "'" & Replace(TextBox12.Text, ",", ".") & "'") & ",
    superficie = " & If(String.IsNullOrEmpty(TextBox13.Text), "NULL", "'" & Replace(TextBox13.Text, ",", ".") & "'") & ",
    produttore = " & If(String.IsNullOrEmpty(TextBox14.Text), "NULL", "'" & Replace(TextBox14.Text, ",", ".") & "'") & ",
    codice_produttore = " & If(String.IsNullOrEmpty(TextBox15.Text), "NULL", "'" & Replace(TextBox15.Text, ",", ".") & "'") & ",
    collo_centrato = " & If(String.IsNullOrEmpty(TextBox16.Text), "NULL", "'" & Replace(TextBox16.Text, ",", ".") & "'") & ",
    tipo_tappo = " & If(String.IsNullOrEmpty(TextBox19.Text), "NULL", "'" & Replace(TextBox19.Text, ",", ".") & "'") & ",
    filettatura = " & If(String.IsNullOrEmpty(TextBox20.Text), "NULL", "'" & Replace(TextBox20.Text, ",", ".") & "'") & ",
    diametro_esterno_fil = " & If(String.IsNullOrEmpty(TextBox21.Text), "NULL", Replace(TextBox21.Text, ",", ".")) & ",
    passo = " & If(String.IsNullOrEmpty(TextBox22.Text), "NULL", Replace(TextBox22.Text, ",", ".")) & ",
    num_principi = " & If(String.IsNullOrEmpty(TextBox23.Text), "NULL", Replace(TextBox23.Text, ",", ".")) & "
WHERE codice_campione = " & par_id_campione & ""


        CMD_SAP.ExecuteNonQuery()





        CNN.Close()
    End Sub

    Sub aggiorna_tappo(par_id_campione As Integer)
        Dim CNN As New SqlConnection

        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN


        CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].Coll_campioni_tappi
SET
    altezza = " & If(String.IsNullOrEmpty(Replace(TextBox17.Text, ",", ".")), "NULL", Replace(TextBox17.Text, ",", ".")) & ",
    larghezza = " & If(String.IsNullOrEmpty(Replace(TextBox18.Text, ",", ".")), "NULL", Replace(TextBox18.Text, ",", ".")) & ",
    profondità = " & If(String.IsNullOrEmpty(Replace(TextBox1.Text, ",", ".")), "NULL", Replace(TextBox1.Text, ",", ".")) & ",
    diametro_interno = " & If(String.IsNullOrEmpty(Replace(TextBox2.Text, ",", ".")), "NULL", Replace(TextBox2.Text, ",", ".")) & ",
    Fissaggio = " & If(String.IsNullOrEmpty(Replace(TextBox24.Text, ",", ".")), "NULL", "'" & Replace(TextBox24.Text, ",", ".") & "'") & ",
    Forma = " & If(String.IsNullOrEmpty(Replace(TextBox25.Text, ",", ".")), "NULL", "'" & Replace(TextBox25.Text, ",", ".") & "'") & ",
    Materiale = " & If(String.IsNullOrEmpty(Replace(TextBox26.Text, ",", ".")), "NULL", "'" & Replace(TextBox26.Text, ",", ".") & "'") & ",
    Superficie = " & If(String.IsNullOrEmpty(Replace(TextBox27.Text, ",", ".")), "NULL", "'" & Replace(TextBox27.Text, ",", ".") & "'") & ",
    Produttore = " & If(String.IsNullOrEmpty(Replace(TextBox28.Text, ",", ".")), "NULL", "'" & Replace(TextBox28.Text, ",", ".") & "'") & ",
    Codice_produttore = " & If(String.IsNullOrEmpty(Replace(TextBox29.Text, ",", ".")), "NULL", "'" & Replace(TextBox29.Text, ",", ".")) & "
        WHERE codice_campione = " & par_id_campione & ""


        CMD_SAP.ExecuteNonQuery()


        CNN.Close()
    End Sub
    Sub aggiorna_scatola(par_id_campione As Integer)

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN


        CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].Coll_campioni_scatole
SET altezza = " & If(String.IsNullOrEmpty(TextBox112.Text), "NULL", Replace(TextBox112.Text, ",", ".")) & ",
    larghezza = " & If(String.IsNullOrEmpty(TextBox111.Text), "NULL", Replace(TextBox111.Text, ",", ".")) & ",
    profondita = " & If(String.IsNullOrEmpty(TextBox110.Text), "NULL", Replace(TextBox110.Text, ",", ".")) & ",
    volume = " & If(String.IsNullOrEmpty(TextBox113.Text), "NULL", Replace(TextBox113.Text, ",", ".")) & ",
    materiale = " & If(String.IsNullOrEmpty(TextBox109.Text), "NULL", "'" & Replace(TextBox109.Text, ",", ".") & "'") & ",
    forma = " & If(String.IsNullOrEmpty(TextBox108.Text), "NULL", "'" & Replace(TextBox108.Text, ",", ".") & "'") & ",
    sezione = " & If(String.IsNullOrEmpty(TextBox107.Text), "NULL", "'" & Replace(TextBox107.Text, ",", ".") & "'") & ",
    superficie = " & If(String.IsNullOrEmpty(TextBox106.Text), "NULL", "'" & Replace(TextBox106.Text, ",", ".") & "'") & ",
    produttore = " & If(String.IsNullOrEmpty(TextBox105.Text), "NULL", "'" & Replace(TextBox105.Text, ",", ".") & "'") & ",
    codice_produttore = " & If(String.IsNullOrEmpty(TextBox104.Text), "NULL", "'" & Replace(TextBox104.Text, ",", ".") & "'") & "
  
WHERE codice_campione = " & par_id_campione & ""


        CMD_SAP.ExecuteNonQuery()

        CNN.Close()
    End Sub
    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting
        If blocco_tab = 1 Then
            e.Cancel = True ' Impedisce il cambio di scheda
        End If


    End Sub

    Sub Lista_registrazioni(par_codice_SAP As String)


    End Sub 'Inserisco le risorse nella combo box

    Sub aggiorna_sottotappo(par_id_campione As Integer)

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN


        CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].Coll_campioni_sottotappi
SET
    altezza = " & If(String.IsNullOrEmpty(Replace(TextBox39.Text, ",", ".")), "NULL", Replace(TextBox39.Text, ",", ".")) & ",
    larghezza = " & If(String.IsNullOrEmpty(Replace(TextBox38.Text, ",", ".")), "NULL", Replace(TextBox38.Text, ",", ".")) & ",
    profondità = " & If(String.IsNullOrEmpty(Replace(TextBox37.Text, ",", ".")), "NULL", Replace(TextBox37.Text, ",", ".")) & ",
    diametro_interno = " & If(String.IsNullOrEmpty(Replace(TextBox36.Text, ",", ".")), "NULL", Replace(TextBox36.Text, ",", ".")) & ",
    vite_pressione = '" & ComboBox6.Text & "',
    Forma = " & If(String.IsNullOrEmpty(Replace(TextBox34.Text, ",", ".")), "NULL", "'" & Replace(TextBox34.Text, ",", ".") & "'") & ",
    Materiale = " & If(String.IsNullOrEmpty(Replace(TextBox33.Text, ",", ".")), "NULL", "'" & Replace(TextBox33.Text, ",", ".") & "'") & "
        WHERE codice_campione = " & par_id_campione & ""


        CMD_SAP.ExecuteNonQuery()


        CNN.Close()
    End Sub

    Sub aggiorna_pompetta(par_id_campione As Integer)

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN


        CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].[Coll_campioni_pompette]
SET
    A = " & If(String.IsNullOrEmpty(Replace(TextBox30.Text, ",", ".")), "NULL", Replace(TextBox30.Text, ",", ".")) & ",
    B = " & If(String.IsNullOrEmpty(Replace(TextBox31.Text, ",", ".")), "NULL", Replace(TextBox31.Text, ",", ".")) & ",
    C = " & If(String.IsNullOrEmpty(Replace(TextBox32.Text, ",", ".")), "NULL", Replace(TextBox32.Text, ",", ".")) & ",
    D = " & If(String.IsNullOrEmpty(Replace(TextBox40.Text, ",", ".")), "NULL", Replace(TextBox40.Text, ",", ".")) & ",
    Quota_A = " & If(String.IsNullOrEmpty(Replace(TextBox41.Text, ",", ".")), "NULL", Replace(TextBox41.Text, ",", ".")) & ",
    Quota_B = " & If(String.IsNullOrEmpty(Replace(TextBox42.Text, ",", ".")), "NULL", Replace(TextBox42.Text, ",", ".")) & ",
    Quota_C = " & If(String.IsNullOrEmpty(Replace(TextBox43.Text, ",", ".")), "NULL", Replace(TextBox43.Text, ",", ".")) & ",
    Quota_D = " & If(String.IsNullOrEmpty(Replace(TextBox44.Text, ",", ".")), "NULL", Replace(TextBox44.Text, ",", ".")) & ",
    Quota_E = " & If(String.IsNullOrEmpty(Replace(TextBox45.Text, ",", ".")), "NULL", Replace(TextBox45.Text, ",", ".")) & ",
    Quota_F = " & If(String.IsNullOrEmpty(Replace(TextBox46.Text, ",", ".")), "NULL", Replace(TextBox46.Text, ",", ".")) & ",
    Quota_L = " & If(String.IsNullOrEmpty(Replace(TextBox47.Text, ",", ".")), "NULL", Replace(TextBox47.Text, ",", ".")) & ",
    SP = " & If(String.IsNullOrEmpty(Replace(TextBox48.Text, ",", ".")), "NULL", Replace(TextBox48.Text, ",", ".")) & ",
    Materiale = " & If(String.IsNullOrEmpty(Replace(TextBox49.Text, ",", ".")), "NULL", "'" & Replace(TextBox49.Text, ",", ".") & "'") & ",
    Tipologia = " & If(String.IsNullOrEmpty(Replace(TextBox50.Text, ",", ".")), "NULL", "'" & Replace(TextBox50.Text, ",", ".") & "'") & ",
    Superficie = " & If(String.IsNullOrEmpty(Replace(TextBox51.Text, ",", ".")), "NULL", "'" & Replace(TextBox51.Text, ",", ".") & "'") & ",
    Produttore = " & If(String.IsNullOrEmpty(Replace(TextBox52.Text, ",", ".")), "NULL", "'" & Replace(TextBox52.Text, ",", ".") & "'") & ",
    cod_produttore = " & If(String.IsNullOrEmpty(Replace(TextBox53.Text, ",", ".")), "NULL", "'" & Replace(TextBox53.Text, ",", ".") & "'") & ",
    Fissaggio = " & If(String.IsNullOrEmpty(Replace(TextBox54.Text, ",", ".")), "NULL", "'" & Replace(TextBox54.Text, ",", ".") & "'") & ",
    Ghiera = " & If(String.IsNullOrEmpty(Replace(TextBox55.Text, ",", ".")), "NULL", "'" & Replace(TextBox55.Text, ",", ".") & "'") & ",
    Copritappo = " & If(String.IsNullOrEmpty(Replace(TextBox56.Text, ",", ".")), "NULL", "'" & Replace(TextBox56.Text, ",", ".") & "'") & "
        WHERE codice_campione = " & par_id_campione & ""


        CMD_SAP.ExecuteNonQuery()


        CNN.Close()
    End Sub

    Sub aggiorna_etichetta(par_id_campione As Integer)

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN


        CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].[Coll_campioni_etichette]
SET
    Altezza = " & If(String.IsNullOrEmpty(Replace(TextBox75.Text, ",", ".")), "NULL", Replace(TextBox75.Text, ",", ".")) & ",
    larghezza = " & If(String.IsNullOrEmpty(Replace(TextBox74.Text, ",", ".")), "NULL", Replace(TextBox74.Text, ",", ".")) & ",
    Trasparenza = '" & ComboBox1.Text & "',
    forma = " & If(String.IsNullOrEmpty(Replace(TextBox72.Text, ",", ".")), "NULL", "'" & Replace(TextBox72.Text, ",", ".") & "'") & ",
    diametro_esterno_bobina = " & If(String.IsNullOrEmpty(Replace(TextBox71.Text, ",", ".")), "NULL", Replace(TextBox71.Text, ",", ".")) & ",
    diametro_interno_bobina = " & If(String.IsNullOrEmpty(Replace(TextBox70.Text, ",", ".")), "NULL", Replace(TextBox70.Text, ",", ".")) & ",
    Avvolgimento_bobina = " & If(String.IsNullOrEmpty(Replace(TextBox69.Text, ",", ".")), "NULL", "'" & Replace(TextBox69.Text, ",", ".") & "'") & ",
    materiale = " & If(String.IsNullOrEmpty(Replace(TextBox68.Text, ",", ".")), "NULL", "'" & Replace(TextBox68.Text, ",", ".") & "'") & "
        WHERE codice_campione = " & par_id_campione & ""


        CMD_SAP.ExecuteNonQuery()


        CNN.Close()
    End Sub

    Sub aggiorna_trigger(par_id_campione As Integer)

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN


        CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].[Coll_campioni_trigger]
SET
    A = " & If(String.IsNullOrEmpty(Replace(TextBox78.Text, ",", ".")), "NULL", Replace(TextBox78.Text, ",", ".")) & ",
    B = " & If(String.IsNullOrEmpty(Replace(TextBox79.Text, ",", ".")), "NULL", Replace(TextBox79.Text, ",", ".")) & ",
    Quota_S = " & If(String.IsNullOrEmpty(Replace(TextBox80.Text, ",", ".")), "NULL", Replace(TextBox80.Text, ",", ".")) & ",
    Quota_H = " & If(String.IsNullOrEmpty(Replace(TextBox81.Text, ",", ".")), "NULL", Replace(TextBox81.Text, ",", ".")) & ",
    Quota_L = " & If(String.IsNullOrEmpty(Replace(TextBox82.Text, ",", ".")), "NULL", Replace(TextBox82.Text, ",", ".")) & ",
    Quota_W = " & If(String.IsNullOrEmpty(Replace(TextBox83.Text, ",", ".")), "NULL", Replace(TextBox83.Text, ",", ".")) & ",
    Quota_v = " & If(String.IsNullOrEmpty(Replace(TextBox58.Text, ",", ".")), "NULL", Replace(TextBox58.Text, ",", ".")) & ",
    [Pressione/Vite] = '" & ComboBox2.Text & "',
    produttore = " & If(String.IsNullOrEmpty(Replace(TextBox76.Text, ",", ".")), "NULL", "'" & Replace(TextBox76.Text, ",", ".") & "'") & " ,
    Codice_produttore = " & If(String.IsNullOrEmpty(Replace(TextBox73.Text, ",", ".")), "NULL", "'" & Replace(TextBox73.Text, ",", ".") & "'") & " ,
    Materiale = " & If(String.IsNullOrEmpty(Replace(TextBox67.Text, ",", ".")), "NULL", "'" & Replace(TextBox67.Text, ",", ".") & "'") & "  ,
    SP = " & If(String.IsNullOrEmpty(Replace(TextBox66.Text, ",", ".")), "NULL", "'" & Replace(TextBox66.Text, ",", ".") & "'") & "  ,
    T = " & If(String.IsNullOrEmpty(Replace(TextBox65.Text, ",", ".")), "NULL", Replace(TextBox65.Text, ",", ".")) & ",
    Fissaggio = " & If(String.IsNullOrEmpty(Replace(TextBox64.Text, ",", ".")), "NULL", "'" & Replace(TextBox64.Text, ",", ".") & "'") & ",
    Ghiera = " & If(String.IsNullOrEmpty(Replace(TextBox63.Text, ",", ".")), "NULL", "'" & Replace(TextBox63.Text, ",", ".") & "'") & "  ,
    Grileltto = " & If(String.IsNullOrEmpty(Replace(TextBox62.Text, ",", ".")), "NULL", "'" & Replace(TextBox62.Text, ",", ".") & "'") & " ,
    Protezione = " & If(String.IsNullOrEmpty(Replace(TextBox61.Text, ",", ".")), "NULL", "'" & Replace(TextBox61.Text, ",", ".") & "'") & " ,
    Note = " & If(String.IsNullOrEmpty(Replace(TextBox60.Text, ",", ".")), "NULL", "'" & Replace(TextBox60.Text, ",", ".") & "'") & " ,
    Cannuccia = " & If(String.IsNullOrEmpty(Replace(TextBox59.Text, ",", ".")), "NULL", "'" & Replace(TextBox59.Text, ",", ".") & "'") & " 
WHERE codice_campione = " & par_id_campione & ""



        CMD_SAP.ExecuteNonQuery()


        CNN.Close()
    End Sub

    Sub aggiorna_prodotto(par_id_campione As Integer)

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN


        CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].[Coll_campioni_prodotti]
SET
    densita = " & If(String.IsNullOrEmpty(Replace(TextBox57.Text, ",", ".")), "NULL", Replace(TextBox57.Text, ",", ".")) & ",
    viscosita_dinamica = " & If(String.IsNullOrEmpty(Replace(TextBox77.Text, ",", ".")), "NULL", Replace(TextBox77.Text, ",", ".")) & ",
    conducibilita_elettrica = " & If(String.IsNullOrEmpty(Replace(TextBox100.Text, ",", ".")), "NULL", Replace(TextBox100.Text, ",", ".")) & ",
    categoria = " & If(String.IsNullOrEmpty(Replace(TextBox99.Text, ",", ".")), "NULL", "'" & Replace(TextBox99.Text, ",", ".") & "'") & ",
    infiammabile = '" & ComboBox3.Text & "',
    nome_commerciale = " & If(String.IsNullOrEmpty(Replace(TextBox84.Text, ",", ".")), "NULL", "'" & Replace(TextBox84.Text, ",", ".") & "'") & ",
    viscosità_cinematica = " & If(String.IsNullOrEmpty(Replace(TextBox97.Text, ",", ".")), "NULL", Replace(TextBox97.Text, ",", ".")) & ",
    Corrosivo = '" & ComboBox4.Text & "',
    [Nocivo/tossico] = '" & ComboBox5.Text & "',
    Note = " & If(String.IsNullOrEmpty(Replace(TextBox87.Text, ",", ".")), "NULL", "'" & Replace(TextBox87.Text, ",", ".") & "'") & "
WHERE codice_campione = " & par_id_campione & ""


        CMD_SAP.ExecuteNonQuery()


        CNN.Close()
    End Sub

    Sub aggiorna_film(par_id_campione As Integer)

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN


        CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].[Coll_campioni_film]
SET
    larghezza = " & If(String.IsNullOrEmpty(Replace(TextBox89.Text, ",", ".")), "NULL", Replace(TextBox89.Text, ",", ".")) & ",
    diametro_fulcro = " & If(String.IsNullOrEmpty(Replace(TextBox90.Text, ",", ".")), "NULL", Replace(TextBox90.Text, ",", ".")) & ",
   materiale = " & If(String.IsNullOrEmpty(Replace(TextBox85.Text, ",", ".")), "NULL", "'" & Replace(TextBox85.Text, ",", ".") & "'") & ",
    temperatura_saldatura = " & If(String.IsNullOrEmpty(Replace(TextBox88.Text, ",", ".")), "NULL", Replace(TextBox88.Text, ",", ".")) & ",
    diametro_esterno = " & If(String.IsNullOrEmpty(Replace(TextBox91.Text, ",", ".")), "NULL", Replace(TextBox91.Text, ",", ".")) & "
WHERE codice_campione = " & par_id_campione & ""


        CMD_SAP.ExecuteNonQuery()


        CNN.Close()
    End Sub

    Sub aggiorna_copritappo(par_id_campione As Integer)

        Dim CNN As New SqlConnection

        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        CMD_SAP.Connection = CNN



        CMD_SAP.CommandText = "UPDATE [Tirelli_40].[dbo].Coll_campioni_copritappi
SET
    altezza = " & If(String.IsNullOrEmpty(Replace(TextBox103.Text, ",", ".")), "NULL", Replace(TextBox103.Text, ",", ".")) & ",
    larghezza = " & If(String.IsNullOrEmpty(Replace(TextBox102.Text, ",", ".")), "NULL", Replace(TextBox102.Text, ",", ".")) & ",
    profondità = " & If(String.IsNullOrEmpty(Replace(TextBox101.Text, ",", ".")), "NULL", Replace(TextBox101.Text, ",", ".")) & ",
    diametro_interno = " & If(String.IsNullOrEmpty(Replace(TextBox98.Text, ",", ".")), "NULL", Replace(TextBox98.Text, ",", ".")) & ",
   fissaggio = " & If(String.IsNullOrEmpty(Replace(TextBox96.Text, ",", ".")), "NULL", "'" & Replace(TextBox96.Text, ",", ".") & "'") & ",
    forma = " & If(String.IsNullOrEmpty(Replace(TextBox95.Text, ",", ".")), "NULL", "'" & Replace(TextBox95.Text, ",", ".") & "'") & ",
   materiale = " & If(String.IsNullOrEmpty(Replace(TextBox94.Text, ",", ".")), "NULL", "'" & Replace(TextBox94.Text, ",", ".") & "'") & ",
   superficie = " & If(String.IsNullOrEmpty(Replace(TextBox93.Text, ",", ".")), "NULL", "'" & Replace(TextBox93.Text, ",", ".") & "'") & ",
produttore = " & If(String.IsNullOrEmpty(Replace(TextBox92.Text, ",", ".")), "NULL", "'" & Replace(TextBox92.Text, ",", ".") & "'") & ",
 codice_produttore = " & If(String.IsNullOrEmpty(Replace(TextBox86.Text, ",", ".")), "NULL", "'" & Replace(TextBox86.Text, ",", ".") & "'") & "
        WHERE codice_campione = " & par_id_campione & ""


        CMD_SAP.ExecuteNonQuery()


        CNN.Close()
    End Sub





    Sub riempi_datagridview_combinazioni(par_id_campione As Integer)
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
        CMD_SAP_2.CommandText = "declare @id_campione as integer

set @id_campione=" & par_id_campione & "

SELECT t21.itemcode,t21.itemname, t21.u_final_customer_name,
t0.id_combinazione,t0.vel_richiesta, t0.campione_1, t11.INIZIALE_SIGLA + T1.NOME   as 'Nome_1',t1.immagine as 'Immagine_1', t0.campione_2, t12.INIZIALE_SIGLA + T2.NOME  as 'Nome_2', t2.immagine as 'Immagine_2', t0.campione_3,t13.INIZIALE_SIGLA + T3.NOME  as 'Nome_3',t3.immagine as 'immagine_3', t0.campione_4,t14.INIZIALE_SIGLA + T4.NOME  as 'Nome_4',t4.immagine as 'immagine_4', t0.campione_5,t15.INIZIALE_SIGLA + T5.NOME  as 'Nome_5',t5.immagine as 'immagine_5', t0.campione_6,t16.INIZIALE_SIGLA + T6.NOME  as 'Nome_6' ,t6.immagine as 'immagine_6', t0.campione_7, t17.INIZIALE_SIGLA + T7.NOME  as 'Nome_7',t7.immagine as 'immagine_7', t0.campione_8,t18.INIZIALE_SIGLA + T8.NOME  as 'Nome_8',t8.immagine as 'immagine_8', t0.campione_9,t19.INIZIALE_SIGLA + T9.NOME  as 'Nome_9',t9.immagine as 'immagine_9', t0.campione_10
,t20.INIZIALE_SIGLA + T10.NOME  as 'Nome_10',t10.immagine as 'immagine_10', 
case when t21.u_progetto is null then '' else t21.u_progetto end as 'u_progetto'

FROM [Tirelli_40].[dbo].COLL_COMBINAZIONI t0
left join [Tirelli_40].[dbo].coll_campioni t1 on t0.campione_1=t1.id_campione
left join [Tirelli_40].[dbo].coll_campioni t2 on t0.campione_2=t2.id_campione
left join [Tirelli_40].[dbo].coll_campioni t3 on t0.campione_3=t3.id_campione
left join [Tirelli_40].[dbo].coll_campioni t4 on t0.campione_4=t4.id_campione
left join [Tirelli_40].[dbo].coll_campioni t5 on t0.campione_5=t5.id_campione
left join [Tirelli_40].[dbo].coll_campioni t6 on t0.campione_6=t6.id_campione
left join [Tirelli_40].[dbo].coll_campioni t7 on t0.campione_7=t7.id_campione
left join [Tirelli_40].[dbo].coll_campioni t8 on t0.campione_8=t8.id_campione
left join [Tirelli_40].[dbo].coll_campioni t9 on t0.campione_9=t9.id_campione
left join [Tirelli_40].[dbo].coll_campioni t10 on t0.campione_10=t10.id_campione

left join [Tirelli_40].[dbo].COLL_TIPO_CAMPIONE t11 on t1.TIPO_campione= T11.ID_TIPO_CAMPIONE
left join [Tirelli_40].[dbo].COLL_TIPO_CAMPIONE t12 on t2.TIPO_campione= T12.ID_TIPO_CAMPIONE
left join [Tirelli_40].[dbo].COLL_TIPO_CAMPIONE t13 on t3.TIPO_campione= T13.ID_TIPO_CAMPIONE
left join [Tirelli_40].[dbo].COLL_TIPO_CAMPIONE t14 on t4.TIPO_campione= T14.ID_TIPO_CAMPIONE
left join [Tirelli_40].[dbo].COLL_TIPO_CAMPIONE t15 on t5.TIPO_campione= T15.ID_TIPO_CAMPIONE
left join [Tirelli_40].[dbo].COLL_TIPO_CAMPIONE t16 on t6.TIPO_campione= T16.ID_TIPO_CAMPIONE
left join [Tirelli_40].[dbo].COLL_TIPO_CAMPIONE t17 on t7.TIPO_campione= T17.ID_TIPO_CAMPIONE
left join [Tirelli_40].[dbo].COLL_TIPO_CAMPIONE t18 on t8.TIPO_campione= T18.ID_TIPO_CAMPIONE
left join [Tirelli_40].[dbo].COLL_TIPO_CAMPIONE t19 on t9.TIPO_campione= T19.ID_TIPO_CAMPIONE
left join [Tirelli_40].[dbo].COLL_TIPO_CAMPIONE t20 on t10.TIPO_campione= T20.ID_TIPO_CAMPIONE

left join [TIRELLISRLDB].[DBO].oitm t21 on t21.itemcode=t0.Commessa

where t0.campione_1=@id_campione
or t0.campione_2=@id_campione
or t0.campione_3=@id_campione
or t0.campione_4=@id_campione
or t0.campione_5=@id_campione
or t0.campione_6=@id_campione
or t0.campione_7=@id_campione 
or t0.campione_8=@id_campione 
or t0.campione_9=@id_campione
or t0.campione_10=@id_campione
order by t0.id_combinazione"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader
        Dim contatore As Integer = 0
        Dim campioni = Enumerable.Range(1, 10).Select(Function(i) $"campione_{i}").ToArray()
        Dim immagini = Enumerable.Range(1, 10).Select(Function(i) $"immagine_{i}").ToArray()
        Dim nomi = Enumerable.Range(1, 10).Select(Function(i) $"nome_{i}").ToArray()

        Do While cmd_SAP_reader_2.Read()
            DataGridView1.Columns("vel_richiesta").Visible = True

            ' Aggiunta di base della riga
            DataGridView1.Rows.Add(
        cmd_SAP_reader_2("itemcode"),
        cmd_SAP_reader_2("itemname"),
        cmd_SAP_reader_2("u_final_customer_name"),
        cmd_SAP_reader_2("u_progetto"),
        cmd_SAP_reader_2("id_combinazione"),
        cmd_SAP_reader_2("vel_richiesta")
    )

            ' Impostazione valori per i campioni
            For Each campione As String In campioni
                If Not IsDBNull(cmd_SAP_reader_2(campione)) Then
                    DataGridView1.Rows(contatore).Cells(campione).Value = cmd_SAP_reader_2(campione)
                End If
            Next

            ' Impostazione immagini
            For Each immagine As String In immagini
                If Not IsDBNull(cmd_SAP_reader_2(immagine)) Then
                    Try
                        DataGridView1.Rows(contatore).Cells(immagine).Value = Image.FromFile(Homepage.Percorso_immagini & cmd_SAP_reader_2(immagine))
                        DataGridView1.Columns(immagine).Visible = True
                    Catch ex As Exception
                        ' Log dell'errore, se necessario
                    End Try
                End If
            Next

            ' Impostazione nomi
            For Each nome As String In nomi
                If Not IsDBNull(cmd_SAP_reader_2(nome)) Then
                    DataGridView1.Rows(contatore).Cells(nome).Value = cmd_SAP_reader_2(nome)
                    DataGridView1.Columns(nome).Visible = True
                End If
            Next

            contatore += 1
        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        DataGridView1.ClearSelection()

    End Sub


    Sub riempi_giacenze_campione(par_id_campione As Integer, par_datagridview As DataGridView)
        par_datagridview.Rows.Clear()


        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()


        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = Cnn1
        CMD_SAP_2.CommandText = "declare @id_campione as integer

set @id_campione=" & par_id_campione & "

select mag, sum(q_in) as 'Q'
from
[TIRELLI_40].[DBO].[coll_campioni_GIACENZE]
where id_campione=@id_campione
group by mag"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        Do While cmd_SAP_reader_2.Read()

            par_datagridview.Rows.Add(
        cmd_SAP_reader_2("mag"),
        cmd_SAP_reader_2("q"))

        Loop

        cmd_SAP_reader_2.Close()
        Cnn1.Close()

        par_datagridview.ClearSelection()

    End Sub

    Private Sub Img_Descrizione_Click(sender As Object, e As EventArgs) Handles Img_Descrizione.Click
        Form_Zoom.Show()
        Form_Zoom.Picture_Zoom.Image = Img_Descrizione.Image


    End Sub

    Private Sub Picture_Campione_Click(sender As Object, e As EventArgs) Handles Picture_Campione.Click
        Try
            Process.Start(LinkLabel1.Text)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Try
            Process.Start(LinkLabel1.Text)
        Catch ex As Exception
            MsgBox(LinkLabel1.Text)
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If MessageBox.Show($"Sei sicuro di voler eliminare per sempre questo campione? Prima di farlo ricordati di aver eliminato tutte le combinazioni che lo contengono", "Elimina campione", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            If controllo_esistenza_combinazione_Campione(id_campione) = "OK" Then
                elimina_Campione(id_campione)
                MsgBox("Campione eliminato con successo")

                Scheda_tecnica.riempi_datagridview_campioni(Scheda_tecnica.DataGridView3, Scheda_tecnica.codice_bp_campione, Scheda_tecnica.bp_code, Scheda_tecnica.final_bp_code, Homepage.Percorso_immagini, Homepage.sap_tirelli)
                Scheda_tecnica.riempi_datagridview_combinazioni(Scheda_tecnica.DataGridView1, Scheda_tecnica.codice_commessa, Homepage.sap_tirelli)
                Me.Close()
            Else
                MsgBox("Non è possibile eliminare il campione, eliminare prima le combinazioni che coinvolgono il campione")
            End If


        End If
    End Sub



    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then


            If DataGridView1.Columns.IndexOf(Immagine_1) Or DataGridView1.Columns.IndexOf(Immagine_2) Or DataGridView1.Columns.IndexOf(Immagine_3) Or DataGridView1.Columns.IndexOf(Immagine_4) Or DataGridView1.Columns.IndexOf(Immagine_5) Or DataGridView1.Columns.IndexOf(Immagine_6) Or DataGridView1.Columns.IndexOf(Immagine_7) Or DataGridView1.Columns.IndexOf(Immagine_8) Or DataGridView1.Columns.IndexOf(Immagine_9) Or DataGridView1.Columns.IndexOf(Immagine_10) Then


                id_campione = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex - 2).Value
                Show()
                inizializza_form()




            End If
        End If
    End Sub

    Private Sub Combo_tipo_campione_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Combo_tipo_campione.SelectedIndexChanged

    End Sub



    Private Sub DataGridView3_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellClick
        id_campione = DataGridView3.Rows(e.RowIndex).Cells(columnName:="Campione_").Value
        compila_scheda_campione(id_campione)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Business_partner.Show()
        Business_partner.Provenienza = "Form_campione_visualizza"
    End Sub

    Private Sub Txt_nome_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Txt_nome.KeyPress
        ' Verifica se il tasto premuto è un numero o il tasto backspace
        If Not Char.IsControl(e.KeyChar) AndAlso Not Char.IsDigit(e.KeyChar) Then
            ' Annulla l'input se non è un numero
            e.Handled = True

            MsgBox("Utilizzare sol numeri interi")
        End If
    End Sub

    Private Sub TableLayoutPanel2_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel2.Paint

    End Sub

    Private Sub Button7_Click_1(sender As Object, e As EventArgs) Handles Button7.Click
        Movimento_campioni.Show()




        Dim par_datagridview_destinazione As DataGridView = Movimento_campioni.DataGridView1
        ' Itera attraverso le righe della DataGridView "datagridview_odp"

        ' Crea una nuova riga nella DataGridView "datagridview1"
        Dim index As Integer = par_datagridview_destinazione.Rows.Add()

        ' Copia i valori dalle colonne necessarie
        par_datagridview_destinazione.Rows(index).Cells("sel").Value = False
        par_datagridview_destinazione.Rows(index).Cells("id_Campione").Value = id_campione
        par_datagridview_destinazione.Rows(index).Cells("tipo").Value = Form_gestione_campioni.ottieni_informazioni_campione(id_campione).tipo_nome
        par_datagridview_destinazione.Rows(index).Cells("Descrizione").Value = Form_gestione_campioni.ottieni_informazioni_campione(id_campione).iniziale_sigla & Form_gestione_campioni.ottieni_informazioni_campione(id_campione).Nome
        par_datagridview_destinazione.Rows(index).Cells("Immagine").Value = Image.FromFile(Homepage.Percorso_immagini & Form_gestione_campioni.ottieni_informazioni_campione(id_campione).immagine)
                par_datagridview_destinazione.Rows(index).Cells("Cliente").Value = Form_gestione_campioni.ottieni_informazioni_campione(id_campione).Cardname
        par_datagridview_destinazione.Rows(index).Cells("Q_trasf").Value = 1
        par_datagridview_destinazione.Rows(index).Cells("ID_richiesta").Value = 0





        Movimento_campioni.ComboBox5.SelectedIndex = 1

    End Sub

    Private Sub DataGridView3_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView3.CellContentClick

    End Sub
End Class