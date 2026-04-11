Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib
Imports System.Threading
Imports System.Windows.Documents
Imports System.Reflection.Emit
Imports Microsoft.Office.Interop.Excel
Imports System.IO.Compression


Public Class Acquisti


    Public testo As String = "Manca" & vbCrLf & ""
    Public Elenco_dipendenti(1000) As String
    Public elenco_fasi(100) As String
    Public CODICEDIP As String
    Public documento_sap_testata As String = "OPQT"
    Public documento_sap_righe As String = "PQT1"
    Public visualizzazione As String = "Documenti"
    Public riga As Integer
    Public codice_fase_inserimento As String
    Public stato_padre As String
    Public stato_figlio As String
    Public distinta_base As String
    Public visorder_doc As Integer
    Public docentry_doc As Integer
    Public Series As String
    Public pindicator As String
    Public Versionnum As String
    Public JrnlMemo As String
    Public docentry_1 As Integer
    Public docnum_1 As Integer
    Public itemname As String
    Public Magazzino_dest As String
    Public suppcatnum As String

    Public Cardcode As String
    Public docnum_OC As String
    Public cardname As String


    Public stato_Accettato As String

    Public nome_documento As String
    Private invia_mail As String
    Public percorso_cartella As String
    Public filtro_dipendente As String

    'mail
    Public objOutlook As Object
    Public objMail As Object

    Public docentry_ODP As Integer
    Public docnum_ODP As Integer

    Public riga_odp As String = 0
    Public max_riga_odp As Integer = 0


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        trova_dato_da_excel_par_importazione_odp(Homepage.percorso_acquisti & "\Lancio_ODP.xlsx", "Lancio_ODP", 2, 200)
        MsgBox("Importazione avvenuta con successo")
    End Sub

    Sub ordine_acquisto()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()
        Dim contatore As Integer = 0
        Dim contatore_pdf As Integer = 1
        Dim contatore_dxf As Integer = 1
        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "SELECT t1.itemcode as 'Codice', t2.U_disegno as 'Disegno' FROM OPOR T0  INNER JOIN POR1 T1 ON T0.[DocEntry] = T1.[DocEntry] INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] WHERE T0.[DocNum] ='" & TextBox1.Text & "'"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            contatore = 0


            Try


                If File.Exists(Homepage.percorso_disegni_generico & "PDF\" & cmd_SAP_reader("Disegno") & ".PDF") Then


                    My.Computer.FileSystem.CopyFile(Homepage.percorso_disegni_generico & "PDF\" & cmd_SAP_reader("Disegno") & ".PDF",
    Homepage.percorso_OFF & TextBox1.Text & "\" & cmd_SAP_reader("Disegno") & ".PDF", overwrite:=True)
                Else
                    Do While File.Exists(Homepage.percorso_disegni_generico & "PDF\" & "_foglio_" & contatore_pdf & ".PDF")
                        My.Computer.FileSystem.CopyFile(Homepage.percorso_disegni_generico & "PDF\" & cmd_SAP_reader("Disegno") & "_foglio_" & contatore_pdf & ".PDF",
    Homepage.percorso_OFF & TextBox1.Text & "\" & cmd_SAP_reader("Disegno") & "_foglio_" & contatore_pdf & ".PDF", overwrite:=True)
                        contatore_pdf = contatore_pdf + 1
                    Loop
                End If

            Catch ex As Exception

                testo = testo & "PDF Codice " & cmd_SAP_reader("Codice") & " disegno " & cmd_SAP_reader("Disegno") & " " & vbCrLf & ""
            End Try


            If File.Exists(Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & ".DXF") Then
                My.Computer.FileSystem.CopyFile(Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & ".DXF",
    Homepage.percorso_ODA & TextBox1.Text & "\" & cmd_SAP_reader("Disegno") & ".DXF", overwrite:=True)
                contatore = contatore + 1
            End If

            If File.Exists(Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & "_Foglio.DXF") Then

                My.Computer.FileSystem.CopyFile(
    Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & "_Foglio.DXF",
    Homepage.percorso_DXF & TextBox1.Text & "\" & cmd_SAP_reader("Disegno") & "_Foglio.DXF", overwrite:=True)
                contatore = contatore + 1
            End If

            Do While File.Exists(Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & "_Foglio_" & contatore_dxf & ".DXF")

                My.Computer.FileSystem.CopyFile(
    Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & "_Foglio_" & contatore_dxf & ".DXF",
    Homepage.percorso_OFF & TextBox1.Text & "\" & cmd_SAP_reader("Disegno") & "_Foglio_" & contatore_dxf & ".DXF", overwrite:=True)
                contatore = contatore + 1
                contatore_dxf = contatore_dxf + 1
            Loop

            If File.Exists(Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & "_sviluppo.DXF") Then

                My.Computer.FileSystem.CopyFile(
    Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & "_sviluppo.DXF",
    Homepage.percorso_ODA & TextBox1.Text & "\" & cmd_SAP_reader("Disegno") & "_sviluppo.DXF", overwrite:=True)
                'contatore = contatore + 1
            End If

            If File.Exists(Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & "_esecutivo.DXF") Then

                My.Computer.FileSystem.CopyFile(
    Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & "_esecutivo.DXF",
    Homepage.percorso_ODA & TextBox1.Text & "\" & cmd_SAP_reader("Disegno") & "_esecutivoo.DXF", overwrite:=True)
                'contatore = contatore + 1
            End If


            If contatore = 0 Then

                testo = testo & "DXF Codice " & cmd_SAP_reader("Codice") & " disegno " & cmd_SAP_reader("Disegno") & " " & vbCrLf & ""

            End If




        Loop
        cmd_SAP_reader.Close()
        CNN.Close()
    End Sub






    Sub esporta_documenti()
        ' Percorso della cartella di esportazione
        Dim percorso_cartella As String = Homepage.percorso_disegni & nome_documento & "_" & TextBox1.Text & "\"

        ' Percorso del file mancanti.ini (fuori dalla cartella)
        Dim percorso_file_ini As String = Homepage.percorso_disegni & "mancante_" & nome_documento & "_" & TextBox1.Text & ".ini"

        ' Creazione della cartella se non esiste
        If Not Directory.Exists(percorso_cartella) Then
            Directory.CreateDirectory(percorso_cartella)
        End If

        ' Esporta i documenti
        Pesca_DXF_PDF(Homepage.percorso_disegni, Homepage.percorso_disegni_generico, nome_documento, TextBox1.Text, documento_sap_testata, documento_sap_righe)

        ' Controlla se la cartella ha contenuto prima di creare lo ZIP
        If Directory.GetFiles(percorso_cartella).Length = 0 And Directory.GetDirectories(percorso_cartella).Length = 0 Then
            MsgBox("La cartella è vuota. Nessun file verrà compresso.", MsgBoxStyle.Exclamation, "Attenzione")
        Else
            ' Percorso del file ZIP
            Dim percorso_zip As String = percorso_cartella.TrimEnd("\"c) & ".zip"

            ' Se il file ZIP esiste già, lo elimina
            If File.Exists(percorso_zip) Then File.Delete(percorso_zip)

            ' Crea l'archivio ZIP della cartella
            ZipFile.CreateFromDirectory(percorso_cartella, percorso_zip)
        End If

        ' Crea il file mancanti.ini fuori dalla cartella compressa
        Using File_INI_Stream As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(percorso_file_ini, False)
            File_INI_Stream.WriteLine(testo)
        End Using

        ' Se non si deve inviare la mail, apre la cartella
        If invia_mail = "NO" Then
            Process.Start("explorer.exe", percorso_cartella)
        End If

        ' Apre il file mancanti.ini (fuori dalla cartella compressa)
        Process.Start("notepad.exe", percorso_file_ini)

        ' Reset del testo
        testo = "Manca" & vbCrLf
    End Sub



    Private Sub Button_disegno_Click(sender As Object, e As EventArgs)
        If Directory.Exists(Homepage.percorso_disegni & "Disegno_" & TextBox1.Text & "") Then

            MsgBox("L'ordine è già stato esportato")

        Else
            Directory.CreateDirectory(Homepage.percorso_disegni & "Disegno_" & TextBox1.Text & "\")

            Disegno()


            Dim File_INI_Stream As StreamWriter
            File_INI_Stream = My.Computer.FileSystem.OpenTextFileWriter(Homepage.percorso_disegni & "Disegno_" & TextBox1.Text & "\mancante.ini", False)
            File_INI_Stream.WriteLine(testo)

            File_INI_Stream.Close()
            File_INI_Stream = Nothing

            Process.Start(Homepage.percorso_disegni & "Disegno_" & TextBox1.Text & "\")
            Process.Start(Homepage.percorso_disegni & "Disegno_" & TextBox1.Text & "\mancante.ini")
            testo = "Manca" & vbCrLf & ""
        End If


    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

        Magazzino.visualizza_disegno(TextBox1.Text)


    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs) Handles PictureBox2.Click
        Try
            Process.Start(Homepage.percorso_DXF & TextBox1.Text & ".DXF")
        Catch ex As Exception
            MsgBox("Il disegno " & TextBox1.Text & " non è ancora stato processato")
        End Try
    End Sub

    Public Sub Aggiorna_INI()
        Dim File_INI_Stream As StreamWriter
        File_INI_Stream = My.Computer.FileSystem.OpenTextFileWriter(Homepage.percorso_ODA & TextBox1.Text & "\mancante.ini", False)
        File_INI_Stream.WriteLine(testo)

        File_INI_Stream.Close()
        File_INI_Stream = Nothing
    End Sub

    Sub Pesca_DXF_PDF(par_destinazione_Cartelle As String, par_sorgente_disegni As String, par_nome_documento As String, par_numero_documento As String, par_documento_Sap_testata As String, par_documento_sap_righe As String)


        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()
        Dim contatore As Integer = 0
        Dim contatore_pdf As Integer = 1
        Dim contatore_dxf As Integer = 1
        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "Select t1.itemcode As 'Codice', t2.U_disegno as 'Disegno' 
FROM " & par_documento_Sap_testata & " T0  INNER JOIN " & par_documento_sap_righe & " T1 ON T0.[DocEntry] = T1.[DocEntry] INNER JOIN OITM T2 ON T1.[ItemCode] = T2.[ItemCode] WHERE T0.[DocNum] ='" & TextBox1.Text & "'"
        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            contatore = 0

            Try
                If File.Exists(par_sorgente_disegni & "PDF\" & cmd_SAP_reader("Disegno") & ".PDF") Then


                    My.Computer.FileSystem.CopyFile(par_sorgente_disegni & "PDF\" & cmd_SAP_reader("Disegno") & ".PDF",
    par_destinazione_Cartelle & par_nome_documento & "_" & par_numero_documento & "\" & cmd_SAP_reader("Disegno") & ".PDF", overwrite:=True)
                Else
                    Do While File.Exists(par_sorgente_disegni & "PDF\" & cmd_SAP_reader("Disegno") & "_foglio_" & contatore_pdf & ".PDF")
                        My.Computer.FileSystem.CopyFile(par_sorgente_disegni & "PDF\" & cmd_SAP_reader("Disegno") & "_foglio_" & contatore_pdf & ".PDF",
    par_destinazione_Cartelle & par_nome_documento & "_" & par_numero_documento & "\" & cmd_SAP_reader("Disegno") & "_foglio_" & contatore_pdf & ".PDF", overwrite:=True)
                        contatore_pdf = contatore_pdf + 1
                    Loop
                End If
            Catch ex As Exception

                testo = testo & "PDF Codice " & cmd_SAP_reader("Codice") & " disegno " & cmd_SAP_reader("Disegno") & " " & vbCrLf & ""
            End Try



            If File.Exists(par_sorgente_disegni & "DXF\" & cmd_SAP_reader("Disegno") & ".DXF") Then
                My.Computer.FileSystem.CopyFile(par_sorgente_disegni & "DXF\" & cmd_SAP_reader("Disegno") & ".DXF",
    par_destinazione_Cartelle & par_nome_documento & "_" & par_numero_documento & "\" & cmd_SAP_reader("Disegno") & ".DXF", overwrite:=True)

                contatore = contatore + 1

            End If

            If File.Exists(Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & "_Foglio.DXF") Then

                My.Computer.FileSystem.CopyFile(
    par_sorgente_disegni & "DXF\" & cmd_SAP_reader("Disegno") & "_Foglio.DXF",
   par_destinazione_Cartelle & par_nome_documento & "_" & par_numero_documento & "\" & cmd_SAP_reader("Disegno") & "_Foglio.DXF", overwrite:=True)
                contatore = contatore + 1
            End If

            Do While File.Exists(Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & "_Foglio_" & contatore_dxf & ".DXF")

                My.Computer.FileSystem.CopyFile(
    par_sorgente_disegni & "DXF\" & cmd_SAP_reader("Disegno") & "_Foglio_" & contatore_dxf & ".DXF",
    par_destinazione_Cartelle & par_nome_documento & "_" & par_numero_documento & "\" & cmd_SAP_reader("Disegno") & "_Foglio_" & contatore_dxf & ".DXF", overwrite:=True)
                contatore = contatore + 1
                contatore_dxf = contatore_dxf + 1
            Loop



            If File.Exists(Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & "_sviluppo.DXF") Then

                My.Computer.FileSystem.CopyFile(
    par_sorgente_disegni & "DXF\" & cmd_SAP_reader("Disegno") & "_sviluppo.DXF",
    par_destinazione_Cartelle & par_nome_documento & "_" & par_numero_documento & "\" & cmd_SAP_reader("Disegno") & "_sviluppo.DXF", overwrite:=True)

            End If

            If File.Exists(Homepage.percorso_DXF & cmd_SAP_reader("Disegno") & "_esecutivo.DXF") Then

                My.Computer.FileSystem.CopyFile(
     par_sorgente_disegni & "DXF\" & cmd_SAP_reader("Disegno") & "_esecutivo.DXF",
    par_destinazione_Cartelle & par_nome_documento & "_" & par_numero_documento & "\" & cmd_SAP_reader("Disegno") & "_esecutivo.DXF", overwrite:=True)

            End If

            If contatore = 0 Then

                testo = testo & "DXF Codice " & cmd_SAP_reader("Codice") & " disegno " & cmd_SAP_reader("Disegno") & " " & vbCrLf & ""

            End If


        Loop
        cmd_SAP_reader.Close()
        CNN.Close()
    End Sub

    Sub trova_disegni_codice(par_codice_disegno As String, par_destinazione_Cartelle As String, par_sorgente_disegni As String)

        If File.Exists(par_sorgente_disegni & "PDF\" & par_codice_disegno & ".PDF") Then


            My.Computer.FileSystem.CopyFile(par_sorgente_disegni & "PDF\" & par_codice_disegno & ".PDF",
par_destinazione_Cartelle & "\" & par_codice_disegno & ".PDF", overwrite:=True)
        Else
            Dim contatore_pdf As Integer = 1

            Do While File.Exists(par_sorgente_disegni & "PDF\" & par_codice_disegno & "_foglio_" & contatore_pdf & ".PDF")
                My.Computer.FileSystem.CopyFile(par_sorgente_disegni & "PDF\" & par_codice_disegno & "_foglio_" & contatore_pdf & ".PDF",
par_destinazione_Cartelle & "\" & par_codice_disegno & "_foglio_" & contatore_pdf & ".PDF", overwrite:=True)
                contatore_pdf += 1
            Loop
        End If




        If File.Exists(par_sorgente_disegni & "DXF\" & par_codice_disegno & ".DXF") Then
            My.Computer.FileSystem.CopyFile(par_sorgente_disegni & "DXF\" & par_codice_disegno & ".DXF",
par_destinazione_Cartelle & "\" & par_codice_disegno & ".DXF", overwrite:=True)



        End If

        If File.Exists(Homepage.percorso_DXF & par_codice_disegno & "_Foglio.DXF") Then

            My.Computer.FileSystem.CopyFile(
par_sorgente_disegni & "DXF\" & par_codice_disegno & "_Foglio.DXF",
par_destinazione_Cartelle & "\" & par_codice_disegno & "_Foglio.DXF", overwrite:=True)

        End If

        Dim contatore_dxf As Integer = 0

        Do While File.Exists(Homepage.percorso_DXF & par_codice_disegno & "_Foglio_" & contatore_dxf & ".DXF")

            My.Computer.FileSystem.CopyFile(
par_sorgente_disegni & "DXF\" & par_codice_disegno & "_Foglio_" & contatore_dxf & ".DXF",
par_destinazione_Cartelle & "\" & par_codice_disegno & "_Foglio_" & contatore_dxf & ".DXF", overwrite:=True)

            contatore_dxf = contatore_dxf + 1
        Loop



        If File.Exists(Homepage.percorso_DXF & par_codice_disegno & "_sviluppo.DXF") Then

            My.Computer.FileSystem.CopyFile(
par_sorgente_disegni & "DXF\" & par_codice_disegno & "_sviluppo.DXF",
par_destinazione_Cartelle & "\" & par_codice_disegno & "_sviluppo.DXF", overwrite:=True)

        End If

        If File.Exists(Homepage.percorso_DXF & par_codice_disegno & "_esecutivo.DXF") Then

            My.Computer.FileSystem.CopyFile(
 par_sorgente_disegni & "DXF\" & par_codice_disegno & "_esecutivo.DXF",
par_destinazione_Cartelle & "\" & par_codice_disegno & "_esecutivo.DXF", overwrite:=True)

        End If
    End Sub
    Sub Disegno()


        Try
            My.Computer.FileSystem.CopyFile(
    Homepage.percorso_disegni_generico & "PDF\" & TextBox1.Text & ".PDF",
    Homepage.percorso_disegni & "Disegno_" & TextBox1.Text & "\" & TextBox1.Text & ".PDF", overwrite:=True)

        Catch ex As Exception

            testo = testo & "PDF Disegno " & TextBox1.Text & " " & vbCrLf & ""
        End Try

        Try
            My.Computer.FileSystem.CopyFile(
    Homepage.percorso_DXF & TextBox1.Text & ".DXF",
    Homepage.percorso_disegni & "Disegno_" & TextBox1.Text & "\" & TextBox1.Text & ".DXF", overwrite:=True)


        Catch ex As Exception

            testo = testo & "DXF Disegno " & TextBox1.Text & " " & vbCrLf & ""
        End Try

    End Sub





    Private Sub Acquisti_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ComboBox1.Text = "RDO"
        Me.BackColor = Homepage.colore_sfondo

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Pianificazione.max_docentry_docnum()
        MsgBox("Numeratori riparati")
    End Sub

    Sub Inserimento_dipendenti()
        ComboBox3.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT T0.[empID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome', T1.[name] 
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code inner join [TIRELLI_40].[DBO].COLL_Reparti t2 on (t2.sap_id_reparto =t1.code or t2.sap_id_reparto_2 =t1.code)   
where t0.active='Y'  and t2.id_reparto='" & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).codice_reparto & "'  order by T0.[lastName] + ' ' + T0.[firstName] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        ComboBox3.Items.Add("")
        Indice = Indice + 1
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            ComboBox3.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        CNN.Close()

    End Sub 'Inserisco le risorse nella combo box

    Sub lista_documenti(par_datagridview As DataGridView)
        par_datagridview.Rows.Clear()
        Dim n_righe As Integer = 0
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT TOP 100 t0.docentry,T0.[DocNum], T0.[DocDate], T0.[CardCode], T0.[CardName], T0.[DocStatus], t1.lastname,  T0.[DocTotal] , case when T0.[u_rs_arubastatus] ='100' then 'SI' else 'NO' end as 'Invio_Mail'
FROM " & documento_sap_testata & " T0 left join [TIRELLI_40].[dbo].ohem t1 on t0.ownercode=t1.empid
WHERE T0.[DocNum] Like '%%" & TextBox2.Text & "%%' and  T0.[CardCode] Like '%%" & TextBox3.Text & "%%'  and T0.[CardName] Like '%%" & TextBox4.Text & "%%'  
and T0.[DocStatus] Like '%%" & ComboBox2.Text & "%%' and coalesce(T0.[u_rs_arubastatus],'') Like '%%" & TextBox9.Text & "%%' " & filtro_dipendente & "

order by t0.docentry DESC"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()
            par_datagridview.Rows.Add(cmd_SAP_reader("docentry"), cmd_SAP_reader("DocNum"), cmd_SAP_reader("lastname"), cmd_SAP_reader("DocDate"), cmd_SAP_reader("Cardname"), cmd_SAP_reader("DocTotal"), cmd_SAP_reader("Invio_Mail"))
            n_righe += 1
        Loop
        cmd_SAP_reader.Close()
        CNN.Close()
        par_datagridview.ClearSelection()
        Label5.Text = n_righe
    End Sub

    Sub lista_righe(PAR_DATAGRIDVIEW As DataGridView, par_docnum As String, par_cardcode As String, par_cardname As String, par_docstatus As String, par_codicedip As String, par_itemcode As String, par_itemname As String, checkbox_Scaduti As Boolean)
        Dim par_filtro_scaduti As String = ""
        If checkbox_Scaduti = True Then
            par_filtro_scaduti = " And T2.SHIPDATE<=getdate() "
        Else
            par_filtro_scaduti = ""
        End If
        Dim n_righe As Integer = 0

        PAR_DATAGRIDVIEW.Rows.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT t0.docentry,t2.visorder,T0.[DocNum], T0.[DocDate], case when '" & documento_sap_righe & "' ='POR1' THEN T2.SHIPDATE ELSE T2.[pqtreqdate] END AS 'pqtreqdate'  ,case when t2.u_accettato is null then '-' else t2.u_accettato end as 'u_accettato', T0.[CardCode], T0.[CardName], T0.[DocStatus], t1.lastname, t2.itemcode, t3.ITEMNAME, case when t2.u_disegno is null then '' else t2.u_disegno end as 'u_disegno', T2.PQTREQQTY, T2.PRICEBEFDI,T2.DISCPRCNT, T2.[lineTotal] , case when t2.u_prg_azs_commessa is null then '' else t2.u_prg_azs_commessa end as 'u_prg_azs_commessa' , case when t2.u_fase is null then '' else t2.u_fase end as 'u_fase'
FROM " & documento_sap_testata & " T0 left join [TIRELLI_40].[dbo].ohem t1 on t0.ownercode=t1.empid
inner join " & documento_sap_righe & " t2 on t2.docentry=t0.docentry
LEFT JOIN OITM T3 ON T3.ITEMCODE=T2.ITEMCODE

WHERE t2.linestatus <>'C' and T0.[DocNum] Like '%%" & par_docnum & "%%' and  T0.[CardCode] Like '%%" & par_cardcode & "%%'  and T0.[CardName] Like '%%" & par_cardname & "%%'  and T0.[DocStatus] Like '%%" & par_docstatus & "%%' and t0.ownercode Like '%%" & par_codicedip & "%%' and T2.[itemcode] Like '%%" & par_itemcode & "%%' and T3.[itemname] Like '%%" & par_itemname & "%%' " & par_filtro_scaduti & "
order by t0.docentry DESC, t2.visorder"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        Do While cmd_SAP_reader.Read()
            PAR_DATAGRIDVIEW.Rows.Add(cmd_SAP_reader("docentry"), cmd_SAP_reader("visorder"), cmd_SAP_reader("DocNum"), cmd_SAP_reader("lastname"), cmd_SAP_reader("DocDate"), cmd_SAP_reader("pqtreqdate"), cmd_SAP_reader("u_accettato"), cmd_SAP_reader("Cardname"), cmd_SAP_reader("itemcode"), cmd_SAP_reader("itemname"), cmd_SAP_reader("u_disegno"), cmd_SAP_reader("PQTREQQTY"), cmd_SAP_reader("PRicebefdi"), cmd_SAP_reader("Discprcnt"), cmd_SAP_reader("linetotal"), cmd_SAP_reader("u_prg_azs_commessa"), cmd_SAP_reader("U_FASE"))
            n_righe += 1
        Loop
        cmd_SAP_reader.Close()
        CNN.Close()
        PAR_DATAGRIDVIEW.ClearSelection()
        Label5.Text = n_righe
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        'If visualizzazione = "Documenti" Then
        '    lista_documenti(DataGridView1)

        'ElseIf visualizzazione = "Righe" Then

        '    filtra_righe()
        'End If
    End Sub

    Sub filtra_righe()
        lista_righe(DataGridView2, TextBox2.Text, TextBox3.Text, TextBox4.Text, ComboBox2.Text, CODICEDIP, TextBox5.Text, TextBox6.Text, CheckBox2.Checked)
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        'If visualizzazione = "Documenti" Then
        '    lista_documenti(DataGridView1)
        'ElseIf visualizzazione = "Righe" Then
        '    filtra_righe()
        'End If
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        'If visualizzazione = "Documenti" Then
        '    lista_documenti(DataGridView1)
        'ElseIf visualizzazione = "Righe" Then
        '    filtra_righe()
        'End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        'Try
        '    CODICEDIP = Elenco_dipendenti(ComboBox3.SelectedIndex)
        'Catch ex As Exception

        'End Try
        'If ComboBox3.SelectedIndex > 0 Then
        '    filtro_dipendente = "and t0.ownercode = " & CODICEDIP & ""
        'Else
        '    filtro_dipendente = ""
        'End If
        'If visualizzazione = "Documenti" Then
        '    lista_documenti(DataGridView1)
        'ElseIf visualizzazione = "Righe" Then
        '    filtra_righe()
        'End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        'If ComboBox1.Text = "RDO" Then
        '    nome_documento = "Richiesta_di_offerta"
        '    documento_sap_testata = "OPQT"
        '    documento_sap_righe = "PQT1"
        'ElseIf ComboBox1.Text = "OA" Then
        '    nome_documento = "Ordine_acquisto"
        '    documento_sap_testata = "OPOR"
        '    documento_sap_righe = "POR1"

        'ElseIf ComboBox1.Text = "DDT" Then
        '    nome_documento = "Trasferimento"
        '    documento_sap_testata = "OWTR"
        '    documento_sap_righe = "WTR1"

        'ElseIf ComboBox1.Text = "RESO A FORNITORE" Then
        '    nome_documento = "Reso"
        '    documento_sap_testata = "ORPD"
        '    documento_sap_righe = "RPD1"
        'End If
        'If visualizzazione = "Documenti" Then
        '    lista_documenti(DataGridView1)
        'ElseIf visualizzazione = "Righe" Then
        '    filtra_righe()
        'End If

    End Sub

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Enter
        'visualizzazione = "Documenti"
        'lista_documenti(DataGridView1)
        'TableLayoutPanel10.Visible = False
    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Enter
        visualizzazione = "Righe"
        filtra_righe()
        TableLayoutPanel10.Visible = True
    End Sub


    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.RowIndex >= 0 Then

            riga = e.RowIndex
            dettagli_riga()


            If e.ColumnIndex = DataGridView2.Columns.IndexOf(Codice) Then

                Magazzino.Show()
                Magazzino.TextBox2.Text = DataGridView2.Rows(riga).Cells(columnName:="Codice").Value
                Magazzino.Refresh()


            ElseIf e.ColumnIndex = DataGridView2.Columns.IndexOf(Accettato) Then
                If DataGridView2.Rows(riga).Cells(columnName:="Accettato").Value = "-" Then
                    stato_Accettato = "Y"
                ElseIf DataGridView2.Rows(riga).Cells(columnName:="Accettato").Value = "Y" Then
                    stato_Accettato = "N"
                ElseIf DataGridView2.Rows(riga).Cells(columnName:="Accettato").Value = "N" Then
                    stato_Accettato = "-"
                End If
                DataGridView2.Rows(riga).Cells(columnName:="Accettato").Value = stato_Accettato
                accetta_offerta()

            ElseIf e.ColumnIndex = DataGridView2.Columns.IndexOf(Dis) Then
                If File.Exists(Homepage.percorso_disegni_generico & "PDF\" & DataGridView2.Rows(riga).Cells(columnName:="Dis").Value & ".PDF") Then
                    Process.Start(Homepage.percorso_disegni_generico & "PDF\" & DataGridView2.Rows(riga).Cells(columnName:="Dis").Value & ".PDF")
                Else
                    MsgBox("PDF non presente")
                End If
            End If

        End If
    End Sub

    Sub accetta_offerta()
        Dim CNN5 As New SqlConnection
        CNN5.ConnectionString = Homepage.sap_tirelli
        CNN5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = CNN5
        CMD_SAP_5.CommandText = "update pqt1 set pqt1.u_accettato='" & stato_Accettato & "' where pqt1.docentry='" & DataGridView2.Rows(riga).Cells(columnName:="Docentry_column").Value & "' and pqt1.visorder='" & DataGridView2.Rows(riga).Cells(columnName:="Visorder").Value & "'  "
        CMD_SAP_5.ExecuteNonQuery()


        CNN5.Close()

    End Sub

    Sub elimina_riga()
        Dim CNN5 As New SqlConnection
        CNN5.ConnectionString = Homepage.sap_tirelli
        CNN5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = CNN5
        CMD_SAP_5.CommandText = "delete pqt1 where pqt1.docentry='" & DataGridView2.Rows(riga).Cells(columnName:="Docentry_column").Value & "' and pqt1.visorder='" & DataGridView2.Rows(riga).Cells(columnName:="Visorder").Value & "'  "
        CMD_SAP_5.ExecuteNonQuery()


        CNN5.Close()

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        filtra_righe()
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        filtra_righe()
    End Sub

    Sub dettagli_riga()
        docentry_doc = DataGridView2.Rows(riga).Cells(columnName:="Docentry_column").Value
        visorder_doc = DataGridView2.Rows(riga).Cells(columnName:="Visorder").Value
        Label8.Text = DataGridView2.Rows(riga).Cells(columnName:="Data_cons").Value
        Label1.Text = DataGridView2.Rows(riga).Cells(columnName:="Codice").Value
        Label2.Text = DataGridView2.Rows(riga).Cells(columnName:="Descrizione").Value
        Button6.Text = DataGridView2.Rows(riga).Cells(columnName:="Dis").Value

        If Not DataGridView2.Rows(riga).Cells(columnName:="Quantità").Value Is System.DBNull.Value Then
            TextBox7.Text = Math.Round(DataGridView2.Rows(riga).Cells(columnName:="Quantità").Value, 2)
        Else
            TextBox7.Text = 0
        End If

        TextBox8.Text = DataGridView2.Rows(riga).Cells(columnName:="Commessa").Value

        ' ComboBox4.Text = DataGridView2.Rows(riga).Cells(columnName:="Nome_fase").Value

        If DataGridView2.Rows(riga).Cells("NOME_fase").Value.ToString().StartsWith("PREM") Then
            ComboBox4.SelectedIndex = 0
        ElseIf DataGridView2.Rows(riga).Cells("NOME_fase").Value.ToString().StartsWith("MONT") Then
            ComboBox4.SelectedIndex = 1

        ElseIf DataGridView2.Rows(riga).Cells("NOME_fase").Value.ToString().StartsWith("FOR") Then
            ComboBox4.SelectedIndex = 2
        ElseIf DataGridView2.Rows(riga).Cells("NOME_fase").Value.ToString().StartsWith("COL") Then
            ComboBox4.SelectedIndex = 3
        ElseIf DataGridView2.Rows(riga).Cells("NOME_fase").Value.ToString().StartsWith("CON") Then
            ComboBox4.SelectedIndex = 4
        ElseIf DataGridView2.Rows(riga).Cells("NOME_fase").Value.ToString().StartsWith("RIC") Then
            ComboBox4.SelectedIndex = 5

        Else
            ComboBox4.SelectedIndex = -1
        End If



        'If DataGridView2.Rows(riga).Cells(columnName:="NOME_fase").Value = "PREMONT" Then

        '    codice_fase_inserimento = "P01501"
        'ElseIf DataGridView2.Rows(riga).Cells(columnName:="NOME_fase").Value = "MONTAG" Then
        '    codice_fase_inserimento = "P02001"
        'ElseIf DataGridView2.Rows(riga).Cells(columnName:="NOME_fase").Value = "COLLAU" Then
        '    codice_fase_inserimento = "P04001"
        'ElseIf DataGridView2.Rows(riga).Cells(columnName:="NOME_fase").Value = "CONSEGNA" Then
        '    codice_fase_inserimento = "P05001"
        'End If
        anagrafiche_min_disp(DataGridView2.Rows(riga).Cells(columnName:="Codice").Value)

    End Sub

    Sub anagrafiche_min_disp(par_codice_sap As String)
        Label6.Text = disponibilità(par_codice_sap)
        Label3.Text = Magazzino.OttieniDettagliAnagrafica(par_codice_sap).Minimo
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

    Public Sub Inserimento_fasi(par_combobox As ComboBox)

        par_combobox.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT T0.[Code], T0.[Name] FROM [dbo].[@FASE]  T0 ORDER BY T0.CODE"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            elenco_fasi(Indice) = cmd_SAP_reader("Code")
            par_combobox.Items.Add(cmd_SAP_reader("Name"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        CNN.Close()

    End Sub

    Public Function disponibilità(par_itemcode As String)
        Dim disponibile As Integer
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT CASE WHEN T2.CODE IS NULL THEN 'N' ELSE 'Y' END AS 'DB'
, CASE WHEN T1.MINLEVEL IS NULL THEN 0 ELSE T1.MINLEVEL END AS 'MINLEVEL'
, CASE WHEN T1.MINORDRQTY IS NULL THEN 0 ELSE T1.MINORDRQTY END AS 'MINORDRQTY'
,sum(case when T0.[OnHand] is null then 0 else T0.[OnHand] end- case when T0.[iscommited] is null then 0 else T0.[iscommited] end+ case when T0.[onorder] is null then 0 else T0.[onorder] end) as 'Disponibile' 
FROM OITW T0 inner join oitm t1 on t0.itemcode=t1.itemcode
LEFT JOIN OITT T2 ON T2.CODE=T0.ITEMCODE
WHERE T0.[ItemCode] ='" & par_itemcode & "'
GROUP BY T1.MINLEVEL, T1.MINORDRQTY, T2.CODE"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            disponibile = Math.Round(cmd_SAP_reader("Disponibile"), 2)
            ' Label6.Text = Math.Round(cmd_SAP_reader("Disponibile"), 2)


        End If
        cmd_SAP_reader.Close()
        CNN.Close()
        Return disponibile
    End Function

    Public Function check_doppio_codice(par_docentry As String)
        Dim doppio_codice As Boolean = False
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT 
    t1.itemcode, 
    COUNT(*) AS Occorrenze,
    SUM(t1.PlannedQty) AS TotalPlannedQty, 
    SUM(t1.BaseQty) AS TotalBaseQty
FROM owor t0 
INNER JOIN wor1 t1 ON t0.docentry = t1.docentry
WHERE t0.docentry = " & par_docentry & "
GROUP BY t1.itemcode
HAVING COUNT(*) > 1
ORDER BY t1.itemcode;"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then
            doppio_codice = True
        Else

            doppio_codice = False

        End If
        cmd_SAP_reader.Close()
        CNN.Close()
        Return doppio_codice
    End Function

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Try
            codice_fase_inserimento = elenco_fasi(ComboBox4.SelectedIndex)
        Catch ex As Exception

        End Try

    End Sub

    Sub CHECK_distinta_base()
        stato_padre = Nothing
        distinta_base = Nothing
        stato_figlio = Nothing
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP_7 As New SqlCommand
        Dim cmd_SAP_reader_7 As SqlDataReader
        CMD_SAP_7.Connection = CNN

        CMD_SAP_7.CommandText = "SELECT T0.ITEMCODE, case when T1.CODE is null then 'null' else t1.code end as 'code',  T0.VALIDFOR as 'Padre', t0.frozenfor, T3.frozenfor as 'Figlio'

FROM OITM T0 LEFT JOIN OITT T1 ON T0.ITEMCODE=T1.CODE 
LEFT JOIN ITT1 T2 ON T2.FATHER=T0.ITEMCODE
LEFT JOIN OITM T3 ON T3.ITEMCODE=T2.CODE

WHERE T0.[ITEMCODE]= '" & Label1.Text & "' 
GROUP BY T0.ITEMCODE,  T1.CODE, T0.VALIDFOR, T3.VALIDFOR, T3.frozenfor, T0.frozenfor"

        cmd_SAP_reader_7 = CMD_SAP_7.ExecuteReader

        Do While cmd_SAP_reader_7.Read()

            If cmd_SAP_reader_7("frozenfor") = "Y" Then
                stato_padre = "N"

            End If

            If cmd_SAP_reader_7("CODE") = "null" Then
                distinta_base = "N"

            End If
            Try
                If cmd_SAP_reader_7("figlio") = "Y" Then
                    stato_figlio = "N"

                End If
            Catch ex As Exception

            End Try


        Loop
        cmd_SAP_reader_7.Close()
        CNN.Close()

    End Sub

    Sub INFO_documento_odp()

        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP_7 As New SqlCommand
        Dim cmd_SAP_reader_7 As SqlDataReader
        CMD_SAP_7.Connection = CNN

        CMD_SAP_7.CommandText = "SELECT T0.[ItemName] as 'Itemname'
, CASE WHEN T0.DFLTWH IS NULL THEN '01' ELSE T0.DFLTWH END as 'Magazzino'
, CASE WHEN t0.suppcatnum is null then'' else t0.suppcatnum end as 'suppcatnum' 
                        From OITM T0 inner join oitt t1 on t0.itemcode=t1.code where t0.itemcode='" & Label1.Text & "'"

        cmd_SAP_reader_7 = CMD_SAP_7.ExecuteReader

        If cmd_SAP_reader_7.Read() Then
            itemname = cmd_SAP_reader_7("Itemname")

            If ComboBox5.Text = "INT" Then
                Magazzino_dest = "CAP2"
            ElseIf ComboBox5.Text = "ASSEMBL" Then
                Magazzino_dest = "02"
            ElseIf ComboBox5.Text = "INT_SALD" Then
                Magazzino_dest = "01"
            Else
                Magazzino_dest = cmd_SAP_reader_7("magazzino")
            End If

            suppcatnum = cmd_SAP_reader_7("suppcatnum")
        End If


        cmd_SAP_reader_7.Close()
        CNN.Close()

    End Sub



    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'Dim answer As Integer
        'If Button5.BackColor = Color.Red Then
        '    MsgBox("Prima creare una distinta base")

        'Else
        '    If ComboBox4.SelectedIndex < 0 Then
        '        MsgBox("scegliere una fase")
        '    Else

        '        If Label1.Text = Nothing Or Label1.Text = "" Then
        '            MsgBox("Non risulta alcun articolo")
        '        Else

        '            If disponibilità(Label1.Text) >= 0 And TextBox8.Text <> "STOCK" And TextBox8.Text <> "SCORTA" Then
        '                answer = MsgBox("Il codice risulta disponibile e non risulta essere ordinato a STOCK. Confermare l'ordine?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
        '                If answer = vbYes Then


        '                    If ComboBox5.Text = "INT" Then
        '                        Magazzino_dest = "CAP2"
        '                    ElseIf ComboBox5.Text = "ASSEMBL" Then
        '                        Magazzino_dest = "02"
        '                    ElseIf ComboBox5.Text = "INT_SALD" Then
        '                        Magazzino_dest = "01"
        '                    End If



        '                    procedura_lancio_odp(Label1.Text, ComboBox5.Text, codice_fase_inserimento, Label8.Text, Label8.Text, TextBox8.Text, Cliente_relativo_alla_commessa(TextBox8.Text).Cardname, TextBox7.Text, Cliente_relativo_alla_commessa(TextBox8.Text).DocNum, Cliente_relativo_alla_commessa(TextBox8.Text).Cardcode, ComboBox7.Text)


        '                    anagrafiche_min_disp(DataGridView2.Rows(riga).Cells(columnName:="Codice").Value)
        '                    CHIUDI_RIGA_DOCUMENTO()
        '                    pulizia()
        '                    If visualizzazione = "Documenti" Then
        '                        lista_documenti(DataGridView1)
        '                    ElseIf visualizzazione = "Righe" Then
        '                        filtra_righe()
        '                    End If
        '                    MsgBox("Ordine lanciato con successo")

        '                Else
        '                    MsgBox("Ordine annullato")

        '                End If

        '            ElseIf TextBox7.Text < -Label6.Text Then
        '                answer = MsgBox("Il codice risulta ordinato per una minore quantità del necessario. Confermare la quantità?", vbQuestion + vbYesNo + vbDefaultButton2, "Message Box Title")
        '                If answer = vbYes Then

        '                    If Label1.Text = Nothing Or Label1.Text = "" Then
        '                        MsgBox("Risulta un errore nel codice da lanciare, non risulta selezionato nessun codice")
        '                    Else
        '                        If ComboBox5.Text = "INT" Then
        '                            Magazzino_dest = "CAP2"
        '                        ElseIf ComboBox5.Text = "ASSEMBL" Then
        '                            Magazzino_dest = "02"
        '                        ElseIf ComboBox5.Text = "INT_SALD" Then
        '                            Magazzino_dest = "01"
        '                        End If
        '                        Cliente_relativo_alla_commessa(TextBox8.Text)
        '                        procedura_lancio_odp(Label1.Text, ComboBox5.Text, codice_fase_inserimento, Label8.Text, Label8.Text, TextBox8.Text, Cliente_relativo_alla_commessa(TextBox8.Text).Cardname, TextBox7.Text, Cliente_relativo_alla_commessa(TextBox8.Text).DocNum, Cliente_relativo_alla_commessa(TextBox8.Text).Cardcode, ComboBox7.Text)
        '                        CHIUDI_RIGA_DOCUMENTO()
        '                        pulizia()
        '                        MsgBox("Ordine lanciato con successo")
        '                    End If



        '                    anagrafiche_min_disp(DataGridView2.Rows(riga).Cells(columnName:="Codice").Value)
        '                    If visualizzazione = "Documenti" Then
        '                        lista_documenti(DataGridView1)
        '                    ElseIf visualizzazione = "Righe" Then
        '                        filtra_righe()
        '                    End If

        '                Else
        '                    MsgBox("Ordine annullato")

        '                End If


        '            Else


        '                If ComboBox5.Text = "INT" Then
        '                    Magazzino_dest = "CAP2"
        '                ElseIf ComboBox5.Text = "ASSEMBL" Then
        '                    Magazzino_dest = "02"
        '                ElseIf ComboBox5.Text = "INT_SALD" Then
        '                    Magazzino_dest = "01"
        '                End If

        '                'Cliente_relativo_alla_commessa(TextBox8.Text)
        '                procedura_lancio_odp(Label1.Text, ComboBox5.Text, codice_fase_inserimento, Label8.Text, Label8.Text, TextBox8.Text, Cliente_relativo_alla_commessa(TextBox8.Text).Cardname, TextBox7.Text, Cliente_relativo_alla_commessa(TextBox8.Text).DocNum, Cliente_relativo_alla_commessa(TextBox8.Text).Cardcode, ComboBox7.Text)

        '                anagrafiche_min_disp(Label1.Text)
        '                CHIUDI_RIGA_DOCUMENTO()
        '                pulizia()
        '                MsgBox("Ordine lanciato con successo")
        '                If visualizzazione = "Documenti" Then
        '                    lista_documenti(DataGridView1)
        '                ElseIf visualizzazione = "Righe" Then
        '                    filtra_righe()
        '                End If

        '            End If
        '        End If
        '    End If
        'End If
    End Sub



    Sub INFO_odp_precedente()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "Select  t11.docentry, t11.docnum,t11.series as 'series',t11.pindicator as 'pindicator',case when t11.versionnum is null then '' else t11.versionnum end as 'Versionnum', case when t11.JrnlMemo is null then '' else t11.JrnlMemo end as 'JrnlMemo' 
from
(
SELECT max(t0.docentry) as 'Max_docentry', max(t0.docnum) as 'Max_docnum'
from owor t0 
)
as t10
inner join owor t11 on t10.max_docentry=t11.docentry"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            Series = cmd_SAP_reader("series")
            pindicator = cmd_SAP_reader("pindicator")
            Versionnum = cmd_SAP_reader("versionnum")
            JrnlMemo = cmd_SAP_reader("JrnlMemo")
            docentry_1 = cmd_SAP_reader("docentry")
            docnum_1 = cmd_SAP_reader("docnum")

        End If
        cmd_SAP_reader.Close()
        CNN.Close()


    End Sub




    Public Structure ClienteInfo
        Public Cardcode As String
        Public DocNum As String
        Public Cardname As String
        ' Add more fields as needed
    End Structure

    Public Function Cliente_relativo_alla_commessa(par_commessa) As ClienteInfo
        Dim clienteInfo As New ClienteInfo()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN

        If par_commessa.Length >= 3 AndAlso par_commessa.Substring(0, 3) = "CDS" Then
            CMD_SAP.CommandText = "SELECT T0.[customer] as 'cardcode',
T0.[custmrName] as 'cardname',
COALESCE(T1.[DocNum],0) AS 'DOCNUM' 
FROM OSCL T0
left join ordr t1 on substring('" & par_commessa & "',4,5)=substring(t1.u_matrcds,4,5)
WHERE T0.[callID] =substring(cast('" & par_commessa & "' as varchar),4,5)"

        ElseIf par_commessa.Length >= 1 AndAlso par_commessa.Substring(0, 1) = "_" Then
            CMD_SAP.CommandText = "select coalesce(t1.cardname,t0.cardname) as 'Cardname',
t0.cardcode as 'Cardcode',
COALESCE(T0.[DocNum],0) AS 'DOCNUM'
from ordr t0 left join ocrd t1 on t0.u_codicebp=t1.cardcode
where t0.docnum=substring(cast('" & par_commessa & "' as varchar),2,5)"
        Else
            CMD_SAP.CommandText = "SELECT COALESCE(T1.[DocNum],0) AS 'DOCNUM', 
case when t1.u_codicebp is null then t1.cardcode else t1.u_codicebp end as 'cardcode',
case when t2.cardname is null then t1.cardname else t2.cardname end as 'cardname'
FROM RDR1 T0  INNER JOIN ORDR T1 ON T0.[DocEntry] = T1.[DocEntry] 
left join ocrd t2 on t2.cardcode=t1.u_codicebp 
WHERE T0.[ItemCode] ='" & par_commessa & "'
ORDER BY T0.[DocEntry] DESC"
        End If




        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then
            clienteInfo.Cardcode = cmd_SAP_reader("cardcode")
            clienteInfo.DocNum = cmd_SAP_reader("docnum")
            clienteInfo.Cardname = Replace(cmd_SAP_reader("cardname"), "'", " ")
            ' Add more assignments for additional fields
        End If

        cmd_SAP_reader.Close()
        CNN.Close()

        Return clienteInfo
    End Function

    '    Public Function acquisto_BRB_O_TIRELLI(par_docnum As String, par_tabella_testata As String, par_tabella_righe As String)
    '        Dim regola_distribuzione As String = "NO"
    '        Dim CNN As New SqlConnection
    '        CNN.ConnectionString = Homepage.sap_tirelli
    '        cnn.Open()

    '        Dim CMD_SAP As New SqlCommand
    '        Dim cmd_SAP_reader As SqlDataReader
    '        CMD_SAP.Connection = cnn

    '        CMD_SAP.CommandText = "SELECT t1.itemcode, COALESCE(t1.ocrcode,'') as 'ocrcode' 
    'FROM " & par_tabella_testata & " T0  INNER JOIN " & par_tabella_righe & " T1 ON T0.[DocEntry] = T1.[DocEntry] 
    'WHERE T0.[DocNum] ='" & par_docnum & "' and t1.ocrcode<>'' and t1.ocrcode is not null"

    '        cmd_SAP_reader = CMD_SAP.ExecuteReader
    '        If cmd_SAP_reader.Read() = True Then
    '            If cmd_SAP_reader("ocrcode") = "BRB01" Then
    '                regola_distribuzione = "BRB"

    '            Else
    '                regola_distribuzione = "NO"

    '            End If
    '        End If

    '        cmd_SAP_reader.Close()
    '        cnn.Close()

    '        Return regola_distribuzione

    '    End Function

    Sub AGGIUSTA_NUMERATORE()
        Dim CNN5 As New SqlConnection
        CNN5.ConnectionString = Homepage.sap_tirelli
        CNN5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = CNN5
        CMD_SAP_5.CommandText = "UPDATE T0 SET T0.AUTOKEY= A.MAX_DOCENTRY+1 FROM ONNM T0, 

(SELECT T10.MAX_DOCNUM, MAX_DOCENTRY, T11.SERIES
FROM
(
SELECT MAX(T0.DOCNUM) AS 'MAX_DOCNUM' , MAX(T0.DOCENTRY) AS 'MAX_DOCENTRY' FROM OWOR T0 
)
AS T10 INNER JOIN OWOR T11 ON T10.MAX_DOCNUM=T11.DOCNUM) A WHERE T0.OBJECTCODE='202'

UPDATE T0 SET T0.NEXTNUMBER= A.MAX_DOCNUM+1 FROM NNM1 T0 INNER JOIN 

(SELECT T10.MAX_DOCNUM, MAX_DOCENTRY, T11.SERIES
FROM
(
SELECT MAX(T0.DOCNUM) AS 'MAX_DOCNUM' , MAX(T0.DOCENTRY) AS 'MAX_DOCENTRY' FROM OWOR T0 
)
AS T10 INNER JOIN OWOR T11 ON T10.MAX_DOCNUM=T11.DOCNUM) A ON A.SERIES=T0.SERIES WHERE T0.OBJECTCODE=202"
        CMD_SAP_5.ExecuteNonQuery()


        CNN5.Close()

    End Sub






    Sub aggiusta_CONFERMATO(PAR_ITEMCODE As String)
        Try

            Dim CNN As New SqlConnection

            CNN.ConnectionString = Homepage.sap_tirelli

            CNN.Open()

            Dim CMD_SAP_7 As New SqlCommand

            CMD_SAP_7.Connection = CNN

            CMD_SAP_7.CommandText = "update t41 set t41.iscommited=t40.confermati
from
(

SELECT t10.itemcode,sum(t10.confermati) AS 'CONFERMATI', t10.MAG
FROM
(
SELECT T0.ITEMCODE, 0 AS 'CONFERMATI', T0.WHSCODE AS 'MAG'
FROM OITW T0
WHERE (T0.ISCOMMITED>0 OR T0.ISCOMMITED<0) and t0.itemcode='" & PAR_ITEMCODE & "'
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[PlannedQty]) AS 'CONFERMATI', T0.[wareHouse] AS 'mag'
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE (T1.[STATUS] ='P' OR  T1.[STATUS] ='R') AND T1.[CmpltQty]< T1.[PlannedQty]  and t0.itemcode='" & PAR_ITEMCODE & "'
GROUP BY T0.[ItemCode],T0.[wareHouse]

UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.[WhsCode] 
FROM RDR1 T0  INNER JOIN ORDR T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T1.[DocStatus] ='O' AND T0.[OpenQty]>0 and t0.itemcode='" & PAR_ITEMCODE & "'
GROUP BY 
T0.[ItemCode],  T0.[WhsCode]
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.FROMWHSCOD 
FROM WTQ1 T0  INNER JOIN OWTQ T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T0.[OpenQty] >0 AND  T1.[DocStatus] ='O' and t0.itemcode='" & PAR_ITEMCODE & "'
GROUP BY 
T0.[ItemCode], T0.FROMWHSCOD
)
AS T10
group by t10.itemcode, t10.MAG

)
as t40 inner join oitw t41 on t41.itemcode=t40.itemcode and t41.whscode=t40.mag"

            CMD_SAP_7.ExecuteNonQuery()

            CNN.Close()
        Catch ex As Exception
            MsgBox("Attenzione al confermato")
        End Try
    End Sub



    Sub aggiusta_ORDINATO(PAR_ITEMCODE As String)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP_6 As New SqlCommand

        CMD_SAP_6.Connection = CNN

        CMD_SAP_6.CommandText = "update t41 set t41.onorder=t40.ORDINATI
from
(
SELECT t10.itemcode,sum(t10.Ordinati) AS 'ORDINATI', t10.MAG
FROM
(
SELECT T0.ITEMCODE, 0 AS 'ORDINATI', T0.WHSCODE AS 'MAG'
FROM OITW T0
WHERE (T0.ONORDER>0 OR T0.ONORDER<0) AND T0.ITEMCODE='" & PAR_ITEMCODE & "'
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[PlannedQty]) AS 'CONFERMATI', T0.[wareHouse] AS 'mag'
FROM OWOR T0   WHERE (T0.[STATUS] ='P' OR  T0.[STATUS] ='R') aND T0.ITEMCODE='" & PAR_ITEMCODE & "'
GROUP BY T0.[ItemCode],T0.[wareHouse]

UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.[WhsCode] 
FROM POR1 T0  INNER JOIN OPOR T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T1.[DocStatus] ='O' AND T0.[OpenQty]>0 AND T0.ITEMCODE='" & PAR_ITEMCODE & "'
GROUP BY 
T0.[ItemCode],  T0.[WhsCode]
UNION ALL
SELECT T0.[ItemCode], SUM(T0.[OpenQty]), T0.WHSCODE
FROM WTQ1 T0  INNER JOIN OWTQ T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE T0.[OpenQty] >0 AND  T1.[DocStatus] ='O' AND T0.ITEMCODE='" & PAR_ITEMCODE & "'
GROUP BY 
T0.[ItemCode], T0.WHSCODE
)
AS T10
group by t10.itemcode, t10.MAG
)
as t40 inner join oitw t41 on t41.itemcode=t40.itemcode and t41.whscode=t40.MAG"

        CMD_SAP_6.ExecuteNonQuery()

        CNN.Close()
    End Sub

    Sub aggiusta_risorse(par_docentry As Integer)
        Dim CNN4 As New SqlConnection
        CNN4.ConnectionString = Homepage.sap_tirelli
        CNN4.Open()

        Dim CMD_SAP_4 As New SqlCommand

        CMD_SAP_4.Connection = CNN4

        CMD_SAP_4.CommandText = "update wor1 set wor1.resalloc='F' where wor1.docentry=" & par_docentry & " and wor1.itemtype=290"

        CMD_SAP_4.ExecuteNonQuery()

        CNN4.Close()

    End Sub

    Sub aggiusta_codici_doppi(par_docentry As Integer)
        Dim CNN4 As New SqlConnection(Homepage.sap_tirelli)
        CNN4.Open()

        Dim CMD_SAP_4 As New SqlCommand
        CMD_SAP_4.Connection = CNN4

        ' Creiamo una tabella temporanea per salvare i dati aggregati
        CMD_SAP_4.CommandText = "
        IF OBJECT_ID('tempdb..#TempWOR1') IS NOT NULL DROP TABLE #TempWOR1;
        CREATE TABLE #TempWOR1 (
            DocEntry INT,
            LineNum INT,
            ItemCode NVARCHAR(50),
            ItemName NVARCHAR(1000),
            VisOrder INT,
            TotalBaseQty DECIMAL(18,6),
            TotalAdditQty DECIMAL(18,6),
            TotalPlannedQty DECIMAL(18,6),
            Warehouse NVARCHAR(50),
            ItemType NVARCHAR(10),
            IssueType NVARCHAR(10),
            StartDate DATE,
            EndDate DATE,
            TotalIssuedQty DECIMAL(18,6),
            TotalCompTotal DECIMAL(18,6),
            TotalPickQty DECIMAL(18,6),
            TotalBaseQtyNum DECIMAL(18,6),
            TotalBaseQtyDen DECIMAL(18,6),
            TotalReleaseQty DECIMAL(18,6),
            UomEntry INT,
            UomCode NVARCHAR(50),
            LineText NVARCHAR(MAX),
            TotalWIP_QtaDaTrasf DECIMAL(18,6)
        );

        WITH CTE AS (
            SELECT 
                t1.DocEntry, 
                t1.LineNum, 
                t1.ItemCode, 
                t1.ItemName, 
                t1.VisOrder, 
                t1.Warehouse, 
                t1.ItemType, 
                t1.IssueType, 
                t1.StartDate, 
                t1.EndDate, 
                t1.UomEntry, 
                t1.UomCode, 
                t1.LineText, 
                ROW_NUMBER() OVER (PARTITION BY t1.ItemCode ORDER BY t1.LineNum) AS RowNum
            FROM WOR1 t1
            WHERE t1.DocEntry = " & par_docentry & "
        )
        INSERT INTO #TempWOR1
        SELECT 
            t1.DocEntry, 
            (SELECT LineNum FROM CTE WHERE CTE.ItemCode = t1.ItemCode AND CTE.RowNum = 1),
            t1.ItemCode, 
            (SELECT ItemName FROM CTE WHERE CTE.ItemCode = t1.ItemCode AND CTE.RowNum = 1),
            (SELECT VisOrder FROM CTE WHERE CTE.ItemCode = t1.ItemCode AND CTE.RowNum = 1),
            SUM(t1.BaseQty), 
            SUM(t1.AdditQty), 
            SUM(t1.PlannedQty), 
            (SELECT Warehouse FROM CTE WHERE CTE.ItemCode = t1.ItemCode AND CTE.RowNum = 1),
            (SELECT ItemType FROM CTE WHERE CTE.ItemCode = t1.ItemCode AND CTE.RowNum = 1),
            (SELECT IssueType FROM CTE WHERE CTE.ItemCode = t1.ItemCode AND CTE.RowNum = 1),
            (SELECT StartDate FROM CTE WHERE CTE.ItemCode = t1.ItemCode AND CTE.RowNum = 1),
            (SELECT EndDate FROM CTE WHERE CTE.ItemCode = t1.ItemCode AND CTE.RowNum = 1),
            SUM(t1.IssuedQty), 
            SUM(t1.CompTotal), 
            SUM(t1.PickQty), 
            SUM(t1.BaseQtyNum), 
            SUM(t1.BaseQtyDen), 
            SUM(t1.ReleaseQty), 
            (SELECT UomEntry FROM CTE WHERE CTE.ItemCode = t1.ItemCode AND CTE.RowNum = 1),
            (SELECT UomCode FROM CTE WHERE CTE.ItemCode = t1.ItemCode AND CTE.RowNum = 1),
            (SELECT LineText FROM CTE WHERE CTE.ItemCode = t1.ItemCode AND CTE.RowNum = 1),
            SUM(t1.U_PRG_WIP_QtaDaTrasf)
        FROM WOR1 t1
        WHERE t1.DocEntry = " & par_docentry & "
        GROUP BY t1.ItemCode, t1.DocEntry;

        -- Cancelliamo le righe esistenti
        DELETE FROM WOR1 WHERE DocEntry = " & par_docentry & ";

        -- Reinseriamo i dati consolidati
        INSERT INTO WOR1 
        (DocEntry, LineNum, ItemCode, ItemName, VisOrder, BaseQty, AdditQty, PlannedQty, Warehouse, ItemType, IssueType, 
         StartDate, EndDate, IssuedQty, CompTotal, PickQty, BaseQtyNum, BaseQtyDen, ReleaseQty, UomEntry, UomCode, LineText, U_PRG_WIP_QtaDaTrasf)
        SELECT * FROM #TempWOR1;
    "

        ' Eseguiamo la query
        CMD_SAP_4.ExecuteNonQuery()

        ' Chiudiamo la connessione
        CNN4.Close()
    End Sub

    Sub aggiusta_testi(par_docentry_odp As Integer)
        Dim CNN5 As New SqlConnection
        CNN5.ConnectionString = Homepage.sap_tirelli
        CNN5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = CNN5

        CMD_SAP_5.CommandText = "update wor1 set wor1.issuetype='', wor1.startdate= null, wor1.enddate= null where wor1.docentry=" & docentry_ODP & " and wor1.itemtype='-18'"

        CMD_SAP_5.ExecuteNonQuery()

        CNN5.Close()

    End Sub

    Sub CHIUDI_RIGA_DOCUMENTO()
        Dim CNN5 As New SqlConnection

        CNN5.ConnectionString = Homepage.sap_tirelli
        CNN5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = CNN5
        CMD_SAP_5.CommandText = "update " & documento_sap_righe & " set linestatus='C' where docentry=" & docentry_doc & " and visorder = " & visorder_doc & ""
        CMD_SAP_5.ExecuteNonQuery()


        CNN5.Close()

    End Sub

    Sub pulizia()
        docentry_doc = Nothing
        visorder_doc = Nothing
        Label1.Text = ""
        Label2.Text = ""
        Button6.Text = ""
        TextBox7.Text = 0
        TextBox8.Text = ""
        ComboBox4.SelectedIndex = -1
        Label7.Text = ""
        Label6.Text = ""
        Label3.Text = ""
        Label4.Text = ""

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Distinta_base_form.Show()

        Distinta_base_form.TextBox1.Text = Label1.Text
    End Sub

    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting

        If DataGridView2.Rows(e.RowIndex).Cells(columnName:="Accettato").Value = "Y" Then
            DataGridView2.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Lime
        ElseIf DataGridView2.Rows(e.RowIndex).Cells(columnName:="Accettato").Value = "N" Then
            DataGridView2.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.White
        End If

    End Sub



    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If File.Exists(Homepage.percorso_disegni_generico & "PDF\" & Button6.Text & ".PDF") Then
            Process.Start(Homepage.percorso_disegni_generico & "PDF\" & Button6.Text & ".PDF")
        Else
            MsgBox("PDF non presente")
        End If
    End Sub



    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        elimina_riga()
        filtra_righe()
    End Sub



    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex >= 0 Then


            TextBox1.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Docnum").Value

            If e.ColumnIndex = DataGridView1.Columns.IndexOf(Docnum) Then


                Form_nuova_offerta.Show()

                Form_nuova_offerta.TextBox10.Text = DataGridView1.Rows(e.RowIndex).Cells(columnName:="Docnum").Value
                Form_nuova_offerta.tipo_offerta = "Visualizzazione"
                Form_nuova_offerta.inizializzazione_form(DataGridView1.Rows(e.RowIndex).Cells(columnName:="Docnum").Value, documento_sap_testata, documento_sap_righe, nome_documento)




            End If



        End If
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        invia_mail = "NO"
        esporta_documenti()
    End Sub



    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        GENERA_PDF_DOC(TextBox1.Text)

    End Sub

    Sub GENERA_PDF_DOC(par_docnum As Integer)



        If documento_sap_testata = "OPQT" Then
            Layout_documenti.nome_documento_SAP = "Richiesta_di_offerta"
            Layout_documenti.documento_SAP = "OPQT"
            Layout_documenti.righe_SAP = "PQT1"
        ElseIf documento_sap_testata = "OPOR" Then
            Layout_documenti.nome_documento_SAP = "Ordine_acquisto"
            Layout_documenti.documento_SAP = "OPOR"
            Layout_documenti.righe_SAP = "POR1"
        End If
        Layout_documenti.docnum = par_docnum

        If Layout_documenti.nome_documento_SAP = "Ordine_acquisto" Then
            Layout_documenti.Informazioni_documento_acquisto(Layout_documenti.docnum, Layout_documenti.documento_SAP, Layout_documenti.righe_SAP)
            Dim destinazione_azienda_OA As String

            'If acquisto_BRB_O_TIRELLI(TextBox1.Text, Layout_documenti.documento_SAP, Layout_documenti.righe_SAP) = "BRB" Then
            '    destinazione_azienda_OA = "BRB"
            'Else
            '    destinazione_azienda_OA = Homepage.azienda
            'End If


            Layout_documenti.trova_word_base(Layout_documenti.Lingua, Layout_documenti.documento_SAP, Layout_documenti.garanzia, Layout_documenti.nome_documento_SAP)
            Layout_documenti.Genera_documento_ACQUISTO()
        ElseIf Layout_documenti.nome_documento_SAP = "Richiesta_di_offerta" Then
            Layout_documenti.Informazioni_documento_acquisto(Layout_documenti.docnum, Layout_documenti.documento_SAP, Layout_documenti.righe_SAP)
            Dim destinazione_azienda_rdo As String

            'If acquisto_BRB_O_TIRELLI(TextBox1.Text, Layout_documenti.documento_SAP, Layout_documenti.righe_SAP) = "BRB" Then
            '    destinazione_azienda_rdo = "BRB"
            'Else
            '    destinazione_azienda_rdo = Homepage.azienda
            'End If

            Layout_documenti.trova_word_base(Layout_documenti.Lingua, Layout_documenti.documento_SAP, Layout_documenti.garanzia, Layout_documenti.nome_documento_SAP)
            Layout_documenti.Genera_documento_ACQUISTO()

        Else

            Layout_documenti.Informazioni_documento(TextBox1.Text)
            Layout_documenti.trova_word_base(Layout_documenti.Lingua, Layout_documenti.documento_SAP, Layout_documenti.garanzia, Layout_documenti.nome_documento_SAP)
            Layout_documenti.Genera_documento("")
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        invia_mail = "SI"
        esporta_documenti()
        GENERA_PDF_DOC(TextBox1.Text)
        InviaEmailConAllegato()
        aggiorna_invia_mail(TextBox1.Text, documento_sap_testata)
        Beep()

        If visualizzazione = "Documenti" Then
            lista_documenti(DataGridView1)

        End If
    End Sub

    Sub aggiorna_invia_mail(par_docnum As Integer, par_testata_documento As String)

        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli

        Cnn.Open()

        Dim Cmd_SAP As New SqlCommand

        Cmd_SAP.Connection = Cnn
        Cmd_SAP.CommandText = "UPDATE " & par_testata_documento & "
SET U_rs_arubastatus='100' WHERE DOCNUM=" & par_docnum & ""
        Cmd_SAP.ExecuteNonQuery()


        Cnn.Close()

    End Sub

    '    Sub InviaEmailConAllegato()


    '        '
    '        Dim strSubject As String
    '        Dim strBody As String
    '        Dim strAttachmentPath As String
    '        Dim strImagePath As String
    '        Dim intImageWidth As Integer
    '        Dim strFileName As String
    '        Dim strSenderEmailAddress As String

    '        If Homepage.azienda = "Tirelli" Then
    '            strSenderEmailAddress = "acquisti@tirelli.net"
    '        ElseIf Homepage.azienda = "4LIFE" Then

    '            strSenderEmailAddress = "acquisti@4lifemachinery.com"
    '        End If


    '        intImageWidth = 200 ' Imposta la larghezza dell'immagine a 400 pixel


    '        strImagePath = Homepage.logo_azienda




    '        'Imposta i valori dei campi email, oggetto, corpo e percorso dell'allegato
    '        '  strEmail = Layout_documenti.e_mail_contatto

    '        If nome_documento = "Ordine_acquisto" Then


    '            strSubject = "Ordine di acquisto NR° " & TextBox1.Text & ""

    '            strBody = "<font face='Century Gothic' size='3'>Buongiorno, <br> ecco in allegato ordine di acquisto NR° " & TextBox1.Text & " ed eventuali relativi disegni  "


    '            'If acquisto_BRB_O_TIRELLI(TextBox1.Text, Layout_documenti.documento_SAP, Layout_documenti.righe_SAP) = "BRB" Then
    '            '    strBody = strBody & "<font face='Century Gothic' size='3'<br>Destinazione: 46047 <b>PORTO MANTOVANO </b></font>(MN) IT, Via Industria, 20/22"
    '            'Else
    '            '    strBody = strBody & "<font face='Century Gothic' size='3'<br>Destinazione: 46045 <b>MARMIROLO </b></font>(MN) IT, Via Veronesi, 3"
    '            'End If


    '            strBody = strBody & "<font face='Century Gothic' size='3'<br>Destinazione: 46045 <b>MARMIROLO </b></font>(MN) IT, Via Veronesi, 3"



    '            strBody = strBody & "<font face='Century Gothic' size='3'<br>" & Layout_documenti.alert
    '            strBody = strBody & "<br> Rispettare le date di consegna, il materiale verrà accettato con un anticipo massimo di 10 giorni. <br>  Inviare conferme d'ordine ed eventuali comunicazioni ad " & strSenderEmailAddress & "<br><br>"

    '        ElseIf nome_documento = "Richiesta_di_offerta" Then
    '            strSubject = "Richiesta di offerta NR° " & TextBox1.Text & ""
    '            strBody = "<font face='Century Gothic' size='3'>Buongiorno, <br> ecco in allegato richiesta di offerta NR° " & TextBox1.Text & " ed eventuali relativi disegni. "
    '            strBody = strBody & "<font face='Century Gothic' size='3'><br> Vi chiediamo vostra migliore offerta di prezzo e consegna"

    '            'If acquisto_BRB_O_TIRELLI(TextBox1.Text, Layout_documenti.documento_SAP, Layout_documenti.righe_SAP) = "BRB" Then
    '            '    strBody = strBody & "<font face='Century Gothic' size='3'<br>Destinazione: 46047 <b>PORTO MANTOVANO </b></font>(MN) IT, Via Industria, 20/22"
    '            'Else
    '            '    strBody = strBody & "<font face='Century Gothic' size='3'<br>Destinazione: 46045 <b>MARMIROLO </b></font>(MN) IT, Via Veronesi, 3"

    '            'End If


    '            strBody = strBody & "<font face='Century Gothic' size='3'<br>Destinazione: 46045 <b>MARMIROLO </b></font>(MN) IT, Via Veronesi, 3"


    '            strBody = strBody & "<font face='Century Gothic' size='3'><br>" & Layout_documenti.alert
    '            strBody = strBody & "<font face='Century Gothic' size='3'><br> Inviare offerte ed eventuali comunicazioni ad " & strSenderEmailAddress & "<br><br>"
    '        End If
    '        strBody = strBody & "<br><font face='Century Gothic' size='3'><b>" & Layout_documenti.Compilatore & "</b></font>"
    '        strBody = strBody & "<br><font face='Century Gothic' size='3'>Ufficio acquisti"
    '        strBody = strBody & "<br> "
    '        strBody = strBody & "<br><img src='" & strImagePath & "' width='" & intImageWidth & "'><br>"
    '        If Homepage.azienda = "Tirelli" Then
    '            'If acquisto_BRB_O_TIRELLI(TextBox1.Text, Layout_documenti.documento_SAP, Layout_documenti.righe_SAP) = "BRB" Then
    '            '    strBody = strBody & "<br><font face='Century Gothic' size='3'>Tirelli SRL LABELLING DIVISION"
    '            'Else
    '            '    strBody = strBody & "<br><font face='Century Gothic' size='3'>Tirelli SRL"
    '            'End If


    '            strBody = strBody & "<br><font face='Century Gothic' size='3'>Tirelli SRL"


    '            strBody = strBody & "<br><font face='Century Gothic' size='3'>Via Vittorio Veronesi, 1 - 46045 Marmirolo (MN) - ITALY"
    '            strBody = strBody & "<br><font face='Century Gothic' size='3'>Tel. +39 0376 396 820 / 387 048"
    '            strBody = strBody & "<br><font face='Century Gothic' size='3'>www.tirelli.net"
    '            strBody = strBody & "<br><font face='Century Gothic' size='3'>>R.I. - C.F. e P.IVA IT01905710206"

    '        Else Homepage.azienda = "4LIFE"

    '            strBody = strBody & "<br><font face='Century Gothic' size='3'>4 LIFE</font>"
    '            strBody = strBody & "<br><font face='Century Gothic' size='3'>Via Progresso, 9 - 46047 Porto Mantovano (MN) - ITALY</font>"
    '            strBody = strBody & "<br><font face='Century Gothic' size='3'>Tel. + 39 0376 396820</font>"
    '            strBody = strBody & "<br><font face='Century Gothic' size='3'><a href='http://www.4lifemachinery.com' target='_blank'>www.4lifemachinery.com</a></font>"
    '            strBody = strBody & "<br><font face='Century Gothic' size='3'>R.I. - C.F. e P.IVA IT02694950201</font>"
    '            strBody = strBody & "<br><br><font face='Century Gothic' size='2'>Si rende noto che le informazioni contenute nella presente missiva sono strettamente riservate. Per maggiori informazioni, si rinvia alla pagina web www.4lifemachinery.com/privacy-policy ; se non siete i destinatari della presente, Vi preghiamo di darcene immediata notizia per telefono allo 0376 396820 o via e-mail all’indirizzo del mittente e di distruggere il messaggio nonché gli allegati. 
    'We hereby notify that the information in this message are confidential. For further information, please visit our Web site at www.4lifemachinery.com/en/privacy-policy; If you are not the addressees above mentioned, please contact us immediately by phone, telephone number +39 0376 396820, or by email at the sender’s address and destroy the message as well as its attachments
    '"
    '        End If
    '        strAttachmentPath = Layout_documenti.percorso_documento_PDF

    '        'Crea un oggetto Outlook e una nuova email
    '        objOutlook = CreateObject("Outlook.Application")
    '        objMail = objOutlook.CreateItem(0)
    '        trova_destinatari()
    '        'Imposta i campi della nuova email
    '        With objMail
    '            '.To = strEmail
    '            .Subject = strSubject
    '            .HTMLBody = strBody
    '            .Display 'Apre la mail in anteprima
    '            .Attachments.Add(strAttachmentPath)
    '            .Attachments.Add(Homepage.percorso_server & "00-Tirelli 4.0\Report\GTC_ACQUISTO\GTC_ACQUISTO_REV_02_2022AGO01 - TIRELLI.pdf")
    '            If CheckBox1.Checked = True Then

    '            Else
    '                .SentOnBehalfOfName = strSenderEmailAddress
    '            End If

    '            strFileName = Dir(percorso_cartella & "\*.*")

    '            Do While strFileName <> ""

    '                If LCase(Split(strFileName, ".")(UBound(Split(strFileName, ".")))) <> "ini" Then
    '                    .Attachments.Add(percorso_cartella & "\" & strFileName)

    '                End If
    '                strFileName = Dir()
    '            Loop


    '        End With

    '        'Rilascia gli oggetti creati
    '        objMail = Nothing
    '        objOutlook = Nothing


    '    End Sub

    Sub InviaEmailConAllegato()
        Dim strSubject As String
        Dim strBody As String
        Dim strAttachmentPath As String
        Dim strImagePath As String
        Dim intImageWidth As Integer
        Dim strFileName As String
        Dim strSenderEmailAddress As String
        Dim objOutlook As Object
        Dim objMail As Object
        Dim percorso_zip As String = Homepage.percorso_disegni & nome_documento & "_" & TextBox1.Text & ".zip"

        ' Determina l'indirizzo email del mittente
        If Homepage.azienda = "Tirelli" Then
            strSenderEmailAddress = "acquisti@tirelli.net"
        ElseIf Homepage.azienda = "4LIFE" Then
            strSenderEmailAddress = "acquisti@4lifemachinery.com"
        End If

        intImageWidth = 200 ' Imposta la larghezza dell'immagine
        strImagePath = Homepage.logo_azienda

        ' Imposta l'oggetto e il corpo dell'email
        If nome_documento = "Ordine_acquisto" Then
            strSubject = "Ordine di acquisto NR° " & TextBox1.Text
            strBody = "<font face='Century Gothic' size='3'>Buongiorno,<br> ecco in allegato ordine di acquisto NR° " & TextBox1.Text & " ed eventuali relativi disegni."
            strBody &= "<br>Destinazione: 46045 <b>MARMIROLO </b>(MN) IT, Via Veronesi, 3"
            strBody &= "<br> Rispettare le date di consegna, il materiale verrà accettato con un anticipo massimo di 10 giorni."
            strBody &= "<br> Inviare conferme d'ordine ed eventuali comunicazioni ad " & strSenderEmailAddress & "<br><br>"
        ElseIf nome_documento = "Richiesta_di_offerta" Then
            strSubject = "Richiesta di offerta NR° " & TextBox1.Text
            strBody = "<font face='Century Gothic' size='3'>Buongiorno,<br> ecco in allegato richiesta di offerta NR° " & TextBox1.Text & " ed eventuali relativi disegni."
            strBody &= "<br> Vi chiediamo vostra migliore offerta di prezzo e consegna."
            strBody &= "<br>Destinazione: 46045 <b>MARMIROLO </b>(MN) IT, Via Veronesi, 3"
            strBody &= "<br> Inviare offerte ed eventuali comunicazioni ad " & strSenderEmailAddress & "<br><br>"
        End If

        strBody &= "<br><b>" & Layout_documenti.Compilatore & "</b><br>Ufficio acquisti"
        strBody &= "<br><img src='" & strImagePath & "' width='" & intImageWidth & "'><br>"

        ' Firma aziendale
        If Homepage.azienda = "Tirelli" Then
            strBody &= "<br>Tirelli SRL"
            strBody &= "<br>Via Vittorio Veronesi, 1 - 46045 Marmirolo (MN) - ITALY"
            strBody &= "<br>Tel. +39 0376 396 820 / 387 048"
            strBody &= "<br>www.tirelli.net"
            strBody &= "<br>R.I. - C.F. e P.IVA IT01905710206"
        ElseIf Homepage.azienda = "4LIFE" Then
            strBody &= "<br>4 LIFE"
            strBody &= "<br>Via Progresso, 9 - 46047 Porto Mantovano (MN) - ITALY"
            strBody &= "<br>Tel. +39 0376 396820"
            strBody &= "<br><a href='http://www.4lifemachinery.com' target='_blank'>www.4lifemachinery.com</a>"
            strBody &= "<br>R.I. - C.F. e P.IVA IT02694950201"
        End If

        ' Crea l'oggetto Outlook e l'email
        objOutlook = CreateObject("Outlook.Application")
        objMail = objOutlook.CreateItem(0)
        trova_destinatari(objMail)

        With objMail
            .Subject = strSubject
            .HTMLBody = strBody
            .Display ' Apre la mail in anteprima
            .Attachments.Add(Layout_documenti.percorso_documento_PDF)
            .Attachments.Add(Homepage.percorso_server & "00-Tirelli 4.0\Report\GTC_ACQUISTO\GTC_ACQUISTO_REV_02_2022AGO01 - TIRELLI.pdf")

            ' Aggiunge lo ZIP come allegato
            If File.Exists(percorso_zip) Then
                .Attachments.Add(percorso_zip)
            End If

            ' Se il mittente non deve essere modificato manualmente
            If Not CheckBox1.Checked Then
                .SentOnBehalfOfName = strSenderEmailAddress
            End If
        End With

        ' Rilascia gli oggetti
        objMail = Nothing
        objOutlook = Nothing
    End Sub


    Sub InviaReportConAllegato(par_allegato As String, emailList As List(Of String))
        Dim strSubject As String
        Dim strBody As String
        Dim strImagePath As String
        Dim intImageWidth As Integer
        Dim strSenderEmailAddress As String
        Dim objOutlook As Object
        Dim objMail As Object

        ' Se non ci sono destinatari, esce dalla sub
        If emailList Is Nothing OrElse emailList.Count = 0 Then
            MsgBox("Nessun destinatario trovato.", vbExclamation, "Errore invio email")
            Exit Sub
        End If

        ' Seleziona il mittente in base all'azienda

        strSenderEmailAddress = "acquisti@tirelli.net"


        intImageWidth = 200 ' Imposta la larghezza dell'immagine
        strImagePath = Homepage.logo_azienda

        ' Composizione del corpo dell'email
        strSubject = "Report Mensile - Puntualità delle Consegne"
        strBody = "<font face='Century Gothic' size='3'>Gentile Fornitore, <br><br> in allegato troverete il report mensile relativo alla puntualità delle consegne degli ordini. Vi invitiamo a prenderne visione e ad adottare eventuali azioni correttive per migliorare le vostre performance in termini di tempistiche di consegna."
        strBody &= "<br><br>Desideriamo inoltre sottolineare che i risultati di questo report, insieme ad altri parametri di valutazione, influenzano le nostre decisioni in merito alla gestione dei fornitori e alle future collaborazioni."
        strBody &= "<br><br>Certi della vostra collaborazione, restiamo a disposizione per eventuali chiarimenti ad acquisti@tirelli.net"
        strBody &= "<br><br>Cordiali saluti"
        ' strBody &= "<br><br>Destinazione: 46045 <b>MARMIROLO </b>(MN) IT, Via Veronesi, 3"
        ' strBody &= "<br>" & Layout_documenti.alert
        '   strBody &= "<br> Rispettare le date di consegna, il materiale verrà accettato con un anticipo massimo di 10 giorni."
        ' strBody &= "<br> Inviare conferme d'ordine ed eventuali comunicazioni a: " & strSenderEmailAddress & "<br><br>"
        strBody &= "<br><b>" & Layout_documenti.Compilatore & "</b>"
        strBody &= "<br><b>Ufficio acquisti</b>"
        strBody &= "<br><br><img src='" & strImagePath & "' width='" & intImageWidth & "'><br>"
        strBody &= "<br>Tirelli SRL"
        strBody &= "<br>Via Vittorio Veronesi, 1 - 46045 Marmirolo (MN) - ITALY"
        strBody &= "<br>Tel. +39 0376 396 820 / 387 048"
        strBody &= "<br>www.tirelli.net"
        strBody &= "<br>R.I. - C.F. e P.IVA IT01905710206"

        ' Crea un oggetto Outlook e una nuova email
        objOutlook = CreateObject("Outlook.Application")
        objMail = objOutlook.CreateItem(0)

        ' Imposta i destinatari con le email lette dal file Excel
        objMail.To = String.Join(";", emailList) ' Unisce gli indirizzi con ";"
        ' Aggiunge Giovanni Tirelli in CCN
        objMail.BCC = "giovanni.tirelli@tirelli.net"

        ' Imposta i campi della nuova email
        With objMail
            .Subject = strSubject
            .HTMLBody = strBody
            .Display ' Apre la mail in anteprima

            ' Aggiunge l'allegato, se presente
            If Not String.IsNullOrEmpty(par_allegato) Then
                .Attachments.Add(par_allegato)
                .Attachments.Add("\\tirfs01\06-Procurement\Documenti vari di lavoro\Report fornitori\Report puntualità\File di servizio\Istruzioni di lettura report.pdf")
            End If

            ' Imposta il mittente se non si sta inviando da un account personale
            If Not CheckBox1.Checked Then
                .SentOnBehalfOfName = strSenderEmailAddress
            End If
        End With

        ' Rilascia gli oggetti Outlook
        objMail = Nothing
        objOutlook = Nothing
    End Sub




    Sub trova_destinatari(ByRef objMail As Object)
        Dim c As String = ""
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT t1.e_maill 
                           FROM " & documento_sap_testata & " t0 
                           INNER JOIN [TIRELLISRLDB].DBO.[ocpr] t1 ON t0.cardcode = t1.cardcode
                           WHERE t0.docnum = @docnum 
                           AND t1.u_riceve_mail = 'Y' 
                           AND (t1.e_maill IS NOT NULL AND t1.e_maill <> '')"

        CMD_SAP.Parameters.AddWithValue("@docnum", TextBox1.Text)

        cmd_SAP_reader = CMD_SAP.ExecuteReader()

        While cmd_SAP_reader.Read()
            If Not IsDBNull(cmd_SAP_reader("e_maill")) Then
                c &= cmd_SAP_reader("e_maill") & "; "
            End If
        End While

        objMail.To = Trim(c) ' Rimuove eventuali spazi vuoti alla fine
        CNN.Close()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        aggiorna_4_life()
        MsgBox("Database 4LIFE aggiornato con successo")
    End Sub

    Sub aggiorna_4_life()
        Dim CNN5 As New SqlConnection
        CNN5.ConnectionString = Homepage.sap_tirelli
        CNN5.Open()

        Dim CMD_SAP_5 As New SqlCommand

        CMD_SAP_5.Connection = CNN5
        CMD_SAP_5.CommandText = "DECLARE @cols AS NVARCHAR(MAX)
DECLARE @cols_insert AS NVARCHAR(MAX)
DECLARE @table_name AS NVARCHAR(MAX)
DECLARE @sql AS NVARCHAR(MAX)

set @table_name='OWHS'
SET @cols = (
        SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @sql = N'INSERT INTO [4LIFE].DBO.[OWHS] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[OWHS] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[OWHS] T1 ON T0.[WHSCODE]=T1.[WHSCODE] ' +
           N'WHERE T1.[WHSCODE] IS NULL '
EXECUTE sp_executesql @sql


set @table_name='OCRG'
SET @cols = (
        SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @sql = N'INSERT INTO [4LIFE].DBO.[OCRG] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[OCRG] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[OCRG] T1 ON T0.[Groupcode]=T1.[Groupcode] ' +
           N'WHERE T1.[groupcode] IS NULL '
EXECUTE sp_executesql @sql


set @table_name='OSHP'
SET @cols = (
        SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @sql = N'INSERT INTO [4LIFE].DBO.[OSHP] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[OSHP] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[OSHP] T1 ON T0.[TrnspCode]=T1.[TrnspCode] ' +
           N'WHERE T1.[TrnspCode] IS NULL '
EXECUTE sp_executesql @sql

set @table_name='ocrd'
SET @cols = (
        SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @sql = N'INSERT INTO [4LIFE].DBO.[OCRD] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[OCRD] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[OCRD] T1 ON T0.[CARDCODE]=T1.[CARDCODE] ' +
           N'WHERE T1.[CARDCODE] IS NULL '
EXECUTE sp_executesql @sql

set @table_name='crd1'
SET @cols = (
        SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @sql = N'INSERT INTO [4LIFE].DBO.[crd1] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[crd1] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[crd1] T1 ON T0.[CARDCODE]=T1.[CARDCODE] ' +
           N'WHERE T1.[CARDCODE] IS NULL '
EXECUTE sp_executesql @sql

SET @sql = N' UPDATE [4LIFE].DBO.[OCRD] SET SERIES=''1'''
EXECUTE sp_executesql @sql


set @table_name='OMRC'
SET @cols = (
        SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @sql = N'INSERT INTO [4LIFE].DBO.[OMRC] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[OMRC] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[OMRC] T1 ON T0.[FIRMCODE]=T1.[FIRMCODE] ' +
           N'WHERE T1.[FIRMCODE] IS NULL '
EXECUTE sp_executesql @sql

set @table_name='OVTG'
SET @cols = (
        SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @sql = N'INSERT INTO [4LIFE].DBO.[OVTG] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[OVTG] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[OVTG] T1 ON T0.[CODE]=T1.[CODE] ' +
           N'WHERE T1.[CODE] IS NULL '
EXECUTE sp_executesql @sql

set @table_name='OPLN'
SET @cols = (
        SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @sql = N'INSERT INTO [4LIFE].DBO.[OPLN] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[OPLN] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[OPLN] T1 ON T0.[LISTNUM]=T1.[LISTNUM] ' +
           N'WHERE T1.[LISTNUM] IS NULL '
EXECUTE sp_executesql @sql

	set @table_name='OITB'
SET @cols = (
        SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @sql = N'INSERT INTO [4LIFE].DBO.[OITB] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[OITB] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[OITB] T1 ON T0.[ItmsGrpCod]=T1.[ItmsGrpCod] ' +
           N'WHERE T1.[ItmsGrpCod] IS NULL '
EXECUTE sp_executesql @sql


set @table_name='itm1'
SET @cols = (
    SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @sql = N'INSERT INTO [4LIFE].DBO.[ITM1] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[ITM1] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[ITM1] T1 ON T0.ITEMCODE=T1.ITEMCODE ' +
		   N'LEFT JOIN [TIRELLISRLDB].DBO.[oitm] T2 ON T0.ITEMCODE=T2.ITEMCODE ' +
           N'WHERE T1.ITEMCODE IS NULL and t2.itemtype<>''F'''

		   
		  

	EXECUTE sp_executesql @sql


	set @table_name='OITW'
SET @cols = (
        SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @sql = N' declare @mag01 AS NVARCHAR(MAX)
	set @mag01 =01 
	INSERT INTO [4LIFE].DBO.[OITW] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[OITW] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[OITW] T1 ON T0.ITEMCODE=T1.ITEMCODE ' +
		   N'LEFT JOIN [4LIFE].DBO.[oitm] T2 ON T0.ITEMCODE=T2.ITEMCODE ' +
		   N'LEFT JOIN [TIRELLISRLDB].DBO.[oitm] T3 ON T0.ITEMCODE=T3.ITEMCODE ' +
           N'WHERE T1.ITEMCODE IS NULL and t3.itemtype<>''F'''
EXECUTE sp_executesql @sql

set @table_name='ODSC'
SET @cols = (
    SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	
SET @sql = N'INSERT INTO [4LIFE].DBO.[ODSC] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[ODSC] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[ODSC] T1 ON T0.BANKCODE=T1.BANKCODE ' +
           N'WHERE T1.BANKCODE IS NULL' 

EXECUTE sp_executesql @sql

set @table_name='OCTG'
SET @cols = (
    SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	
SET @sql = N'INSERT INTO [4LIFE].DBO.[OCTG] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[OCTG] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[OCTG] T1 ON T0.GROUPNUM=T1.GROUPNUM ' +
           N'WHERE T1.GROUPNUM IS NULL' 

EXECUTE sp_executesql @sql



set @table_name='OITM'
SET @cols = (
    SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	
SET @sql = N'INSERT INTO [4LIFE].DBO.[OITM] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[OITM] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[OITM] T1 ON T0.ITEMCODE=T1.ITEMCODE ' +
           N'WHERE T1.ITEMCODE IS NULL and t0.itemtype<>''F''' 

EXECUTE sp_executesql @sql

set @table_name='ORSC'
SET @cols = (
    SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	
SET @sql = N'INSERT INTO [4LIFE].DBO.[ORSC] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[ORSC] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[ORSC] T1 ON T0.VISRESCODE=T1.VISRESCODE ' +
           N'WHERE T1.VISRESCODE IS NULL ' 

EXECUTE sp_executesql @sql

DELETE A
FROM [4life].DBO.[oitt] A
INNER JOIN [TIRELLISRLDB].DBO.[oitm] B ON A.code = B.itemcode
WHERE B.UPDATEDATE >= GETDATE() - 100

EXECUTE sp_executesql @sql

DELETE A
FROM [4life].DBO.[itt1] A
INNER JOIN [TIRELLISRLDB].DBO.[oitm] B ON A.father = B.itemcode
WHERE B.UPDATEDATE >= GETDATE() - 100

EXECUTE sp_executesql @sql

DELETE A
FROM [4life].DBO.[oitt] A
INNER JOIN [TIRELLISRLDB].DBO.[oitt] B ON A.code = B.code
WHERE B.UPDATEDATE >= GETDATE() - 100

EXECUTE sp_executesql @sql

DELETE A
FROM [4life].DBO.[itt1] A
INNER JOIN [TIRELLISRLDB].DBO.[oitt] B ON A.father = B.code
WHERE B.UPDATEDATE >= GETDATE() - 100

EXECUTE sp_executesql @sql


--EXECUTE sp_executesql @sql

delete t1
from [TIRELLISRLDB].DBO.[ITT1] T0 
inner join [4LIFE].DBO.[ITT1] T1 ON T0.FATHER=T1.FATHER AND T0.CODE<>T1.CODE AND T0.CHILDNUM=T1.CHILDNUM 
INNER JOIN [TIRELLISRLDB].DBO.[OITT] T2 ON T0.FATHER=T2.CODE
where t1.code is null

EXECUTE sp_executesql @sql

delete t1
from [TIRELLISRLDB].DBO.[ITT1] T0 
inner join [4LIFE].DBO.[ITT1] T1 ON T0.FATHER=T1.FATHER AND T0.CODE<>T1.CODE AND T0.visorder=T1.visorder 
INNER JOIN [TIRELLISRLDB].DBO.[OITT] T2 ON T0.FATHER=T2.CODE


set @table_name='OITT'
SET @cols = (
    SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	
SET @sql = N'INSERT INTO [4LIFE].DBO.[OITT] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[OITT] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[OITT] T1 ON T0.CODE=T1.CODE ' +
           N'WHERE T1.CODE IS NULL ' 

EXECUTE sp_executesql @sql


set @table_name='ITT1'
SET @cols = (
    SELECT STUFF((
        SELECT ', ' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	SET @cols_insert = (
    SELECT STUFF((
        SELECT ', t0.' + QUOTENAME(COLUMN_NAME)
        FROM INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_NAME = @table_name
        FOR XML PATH(''), TYPE).value('.', 'NVARCHAR(MAX)'), 1, 2, '')
    )
	
SET @sql = N'INSERT INTO [4LIFE].DBO.[ITT1] (' + @cols + ') ' +
           N'SELECT ' + @cols_insert + ' ' +
           N'FROM [TIRELLISRLDB].DBO.[ITT1] T0 ' +
           N'LEFT JOIN [4LIFE].DBO.[ITT1] T1 ON T0.CODE=T1.CODE AND T0.FATHER=T1.FATHER ' +
           N'WHERE T1.CODE IS NULL and t0.code is not null ' 

EXECUTE sp_executesql @sql

SET @sql = N'UPDATE  T0 SET T0.DFLTWH =''01'' from [4LIFE].DBO.[OITM] t0 WHERE (SUBSTRING(T0.ITEMCODE,1,1)=''0'' OR SUBSTRING(T0.ITEMCODE,1,1)=''c'' OR SUBSTRING(T0.ITEMCODE,1,1)=''D'' OR SUBSTRING(T0.ITEMCODE,1,1)=''F'') and t0.createdate>=getdate()-30' 

EXECUTE sp_executesql @sql

update t1 set t1.Warehouse='01' 
from [TIRELLISRLDB].DBO.[ITT1] T1
where t1.Warehouse<>'01' AND T1.Type='4'


EXECUTE sp_executesql @sql

UPDATE T0 SET T0.SERIES=3 FROM [4LIFE].DBO.OITM T0 WHERE 
(SUBSTRING(T0.[ItemCode],1,1)='0' OR SUBSTRING(T0.[ItemCode],1,1)='C' OR SUBSTRING(T0.[ItemCode],1,1)='D' OR SUBSTRING(T0.[ItemCode],1,1)='M' OR SUBSTRING(T0.[ItemCode],1,1)='F') AND T0.SERIES<>3
EXECUTE sp_executesql @sql"
        CMD_SAP_5.ExecuteNonQuery()



        CNN5.Close()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        MRP.Show()

    End Sub


    Sub procedura_lancio_odp(par_itemcode As String, par_produzione As String, par_fase As String, par_data_inizio As String, par_data_fine As String, par_commessa As String, par_cliente As String, par_quantità As Decimal, par_ordine_cliente As String, par_bp_code As String, par_magazzino_destinazione As String)

        par_itemcode = par_itemcode.ToUpper()



        Dim blocco As String = "OK"

        If validita_distinta(par_itemcode, par_produzione) = "OK" Then
        Else
            blocco = validita_distinta(par_itemcode, par_produzione)
        End If

        If check_non_duplicazione_impegni_con_anticipi(par_itemcode, par_commessa) = "OK" Then
        Else
            blocco = check_non_duplicazione_impegni_con_anticipi(par_itemcode, par_commessa)
        End If


        If check_che_non_ci_siano_codici_doppi_nelle_righe(par_itemcode) = "OK" Then
        Else
            blocco = check_che_non_ci_siano_codici_doppi_nelle_righe(par_itemcode)
        End If

        If blocco = "OK" Then

            ' Dim magazzino_destinazione As String = trova_magazzino_destinazione_ODP(par_commessa, par_produzione, Homepage.UTENTE_sap_SALVATO)

            Dim ultimo_progressivo_Commessa As Integer = 0

            If par_produzione = "ASSEMBL" Or par_produzione = "EST" Then
                ultimo_progressivo_Commessa = progressivo_commessa(par_commessa)
            Else
                ultimo_progressivo_Commessa = 0
            End If
            max_docentry_docnum()
            INFO_odp_precedente()
            'If controllo_codice_già_lanciato_per_matricola(par_itemcode, par_commessa) = 0 Then
            insert_into_owor(docentry_ODP, docnum_ODP, par_itemcode, Magazzino.OttieniDettagliAnagrafica(par_itemcode).Descrizione, par_quantità, par_commessa, par_fase, par_cliente, par_data_inizio, par_data_fine, par_magazzino_destinazione, suppcatnum, JrnlMemo, pindicator, Versionnum, par_produzione, Series, par_ordine_cliente, par_bp_code, ultimo_progressivo_Commessa, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)
            'Else
            '    aumenta_quantità_owor(controllo_codice_già_lanciato_per_matricola(par_itemcode, par_commessa), par_quantità)

            'End If

            aggiusta_ORDINATO(par_itemcode)
            aggiusta_righe_odp_confermato_tot_ordinato_tot(par_itemcode)
            IMPORTA_DB(docentry_ODP, docnum_ODP, par_itemcode, Magazzino.OttieniDettagliAnagrafica(par_itemcode).Descrizione, par_quantità, par_commessa, par_fase, par_cliente, par_data_inizio, par_data_fine, par_magazzino_destinazione, suppcatnum, JrnlMemo, pindicator, Versionnum, par_produzione, Series, par_ordine_cliente, par_bp_code, ultimo_progressivo_Commessa, 0)

            If check_doppio_codice(docentry_ODP + 1) = True Then
                aggiusta_codici_doppi(docentry_ODP + 1)

            End If
            max_docentry_docnum()


        Else
                MsgBox(blocco)
        End If

    End Sub

    Public Function controllo_codice_già_lanciato_per_matricola(par_codice As String, commessa As String)


        Dim docentry As Integer = 0

        Dim Cnn_Ticket As New SqlConnection

        Cnn_Ticket.ConnectionString = Homepage.sap_tirelli
        Cnn_Ticket.Open()
        Dim Cmd_Ticket As New SqlCommand
        Dim Reader_Ticket As SqlDataReader
        Cmd_Ticket.Connection = Cnn_Ticket

        Cmd_Ticket.CommandText = "select coalesce(t0.docentry,0) as 'docentry'
from owor t0 
where
t0.itemcode='" & par_codice & "' and t0.u_prg_azs_commessa ='" & commessa & "' and t0.status='P'"

        Reader_Ticket = Cmd_Ticket.ExecuteReader()

        If Reader_Ticket.Read() Then


            docentry = Reader_Ticket("docnum")

        Else docentry = 0

        End If


        Reader_Ticket.Close()
        Cnn_Ticket.Close()
        Return docentry

    End Function

    Sub aggiusta_righe_odp_confermato_tot_ordinato_tot(par_itemcode As String)
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP_6 As New SqlCommand

        CMD_SAP_6.Connection = CNN

        CMD_SAP_6.CommandText = "UPDATE T23 SET T23.U_CONFERMATO_TOT=T20.CONF, T23.U_ORDINATO_TOT=T20.ORD, T23.U_DISPONIIBILETOT=T20.MAG-T20.CONF+T20.ORD

FROM
(
SELECT T10.ITEMCODE,SUM(T11.ONHAND) AS 'MAG', SUM(T11.ISCOMMITED) AS 'CONF', SUM(T11.ONORDER) AS 'ORD'
FROM
(
SELECT T0.[ItemCode] 
FROM WOR1 T0  INNER JOIN OWOR T1 ON T0.[DocEntry] = T1.[DocEntry] 
WHERE (T1.[Status] ='P' OR  T1.[Status] ='R') AND T0.ITEMTYPE=4  and t0.itemcode='" & par_itemcode & "'
GROUP BY T0.[ItemCode]
)
AS T10 INNER JOIN OITW T11 ON T11.ITEMCODE=T10.ITEMCODE
GROUP BY T10.ITEMCODE
)
AS T20 INNER JOIN WOR1 T21 ON T20.ITEMCODE=T21.ITEMCODE
INNER JOIN OWOR T22 ON T22.DOCENTRY=T21.DOCENTRY
INNER JOIN WOR1 T23 ON T23.DOCENTRY=T22.DOCENTRY AND T23.ITEMCODE=T20.ITEMCODE

WHERE T22.STATUS='P' OR T22.STATUS='R'"

        CMD_SAP_6.ExecuteNonQuery()

        CNN.Close()
    End Sub

    Public Function validita_distinta(par_itemcode As String, par_produzione As String)
        Dim CNN As New SqlConnection
        Dim risposta As String = "OK"
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()
        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "SELECT coalesce(t2.validfor,'')  as 'validita_padre'
, t2.u_gestione_magazzino
, coalesce(T1.VALIDFOR,'') 'Valido'
, t1.itemcode as 'Codice' 
,t1.ItmsGrpCod
,substring(t1.itemcode,1,1) as 'Prima_lettera'
FROM ITT1 T0 
INNER JOIN OITM T1 ON T0.CODE=T1.ITEMCODE 
inner join oitm t2 on t2.itemcode=t0.father 
WHERE T0.[Father]= '" & par_itemcode & "' "
        'AND (T1.VALIDFOR='N' or t2.validfor='N' or t2.u_gestione_magazzino='ESAURIMENTO')

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            If cmd_SAP_reader("Validita_padre") = "N" Then
                risposta = "Il codice " & par_itemcode & " Risulta inattivo"

            End If

            If cmd_SAP_reader("Valido") = "N" Then
                risposta = "Il codice figlio " & cmd_SAP_reader("codice") & " della distinta " & par_itemcode & " Risulta inattivo"

            End If
            If cmd_SAP_reader("u_gestione_magazzino") = "ESAURIMENTO" Then
                risposta = "Il codice " & par_itemcode & " Risulta in ESAURIMENTO guardare |dati anagrafici articolo| gestione magazzino"
            End If

            If cmd_SAP_reader("ItmsGrpCod") = "121" And cmd_SAP_reader("prima_lettera") = "C" And par_produzione = "ASSEMBL" Then
                risposta = "Nella distinta " & par_itemcode & " è presente il codice " & cmd_SAP_reader("codice") & " materia prima. non può essere presente in ODP Assemblaggio"
            End If
        Else
            risposta = "Non è presente distinta base del codice " & par_itemcode
        End If

        cmd_SAP_reader.Close()
        CNN.Close()

        Return risposta

    End Function

    Public Function check_non_duplicazione_impegni_con_anticipi(par_itemcode As String, par_commessa As String)

        Dim risposta As String = "OK"
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN6

        CMD_SAP.CommandText = "select t1.itemcode, t0.docnum, t0.prodname, t0.u_prg_azs_commessa 
from owor t0 inner join wor1 t1 on t0.docentry=t1.docentry 
where (t0.status='P' or t0.status='R') and (t0.U_PRG_AZS_Commessa='" & par_commessa & "' or substring(t0.U_PRG_AZS_Commessa,1,1)<>'M')

and t0.prodname Like 'ANTICIP%%' and substring(t1.itemcode,1,1)<>'L' and t1.itemcode='" & par_itemcode & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            '   risposta = " La distinta " & cmd_SAP_reader("father") & " contiene il codice " & cmd_SAP_reader("Code") & " già contenuto nell ordine di anticipo " & cmd_SAP_reader("docnum") & " "

            risposta = " Il codice " & cmd_SAP_reader("itemcode") & " è già contenuto nell' ODP di anticipo  " & cmd_SAP_reader("docnum") & " " & cmd_SAP_reader("prodname") & " Rimuoverlo da quell odp per inserirlo dove si vuole "



        End If
        cmd_SAP_reader.Close()
        CNN6.Close()

        Return risposta
    End Function

    Public Function check_che_non_ci_siano_codici_doppi_nelle_righe(par_itemcode As String)
        Dim risposta As String = "OK"
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN6

        CMD_SAP.CommandText = "select t10.code,t10.n
from
(
select t0.code, count(t0.code) as 'N'
from
itt1 t0 where t0.father='" & par_itemcode & "' and t0.type=4
group by t0.code
)
as t10
where t10.n>1
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            risposta = " La distinta " & par_itemcode & " contiene il codice " & cmd_SAP_reader("Code") & " in due differenti righe, unirle in una unica e poi lanciare l ODP nuovamente "



        End If
        cmd_SAP_reader.Close()
        CNN6.Close()
        Return risposta

    End Function

    Public Function trova_magazzino_destinazione_ODP(par_commessa As String, par_produzione As String, par_utente_sap As String)
        Dim risposta As String = ""
        Dim CNN2 As New SqlConnection
        CNN2.ConnectionString = Homepage.sap_tirelli
        CNN2.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = CNN2
        CMD_SAP_2.CommandText = "SELECT TOP(1)
    CASE
WHEN  '" & par_produzione & "' = 'B_INT' THEN '01'
        WHEN  '" & par_produzione & "' = 'ASSEMBL' THEN '02'
       WHEN  SUBSTRING('" & par_produzione & "', 1, 3) = 'INT' THEN 'CAP2'
        WHEN ( '" & par_commessa & "'='BSCORTA' OR '" & par_commessa & "'='BSCTOCK')  AND '" & par_produzione & "' = 'ASSEMBL' THEN '02'
        WHEN ( '" & par_commessa & "'='BSCORTA' OR '" & par_commessa & "'='BSCTOCK') AND SUBSTRING('" & par_produzione & "', 1, 3) = 'INT' THEN 'CAP2'
WHEN SUBSTRING('" & par_produzione & "', 1, 3) ='INT' THEN 'CAP2'
ELSE '02'  
 
    END as 'Magazzino'
                            
FROM 
[TIRELLI_40].[dbo].ohem t1 LEFT JOIN
 rdr1 t2 ON t2.itemcode = CAST('" & par_commessa & "' AS VARCHAR)
 left join ordr t3 ON T3.docentry=t2.docentry and T3.[DocStatus]<>'C'
WHERE 
t1.userid = '" & par_utente_sap & "'
ORDER BY T3.DOCENTRY DESC"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() = True Then
            ' itemname = cmd_SAP_reader_2("Itemname")
            risposta = cmd_SAP_reader_2("Magazzino")
            ' suppcatnum = cmd_SAP_reader_2("suppcatnum")



        End If
        cmd_SAP_reader_2.Close()
        CNN2.Close()
        Return risposta
    End Function

    Public Function trova_magazzino_destinazione_OC(PAR_OC As String)
        Dim risposta As String = ""
        Dim CNN2 As New SqlConnection
        CNN2.ConnectionString = Homepage.sap_tirelli
        CNN2.Open()

        Dim CMD_SAP_2 As New SqlCommand
        Dim cmd_SAP_reader_2 As SqlDataReader


        CMD_SAP_2.Connection = CNN2
        CMD_SAP_2.CommandText = "select top 1 
case when
t1.ocrcode ='BRB01' then 'CAP2'
ELSE
'CAP2'
END as 'magazzino'
from ordr t0 inner join rdr1 t1 on t0.docentry=t1.docentry
where cast(t0.docnum as varchar)='" & PAR_OC & "'
order by t1.linenum"

        cmd_SAP_reader_2 = CMD_SAP_2.ExecuteReader


        If cmd_SAP_reader_2.Read() = True Then

            risposta = cmd_SAP_reader_2("Magazzino")




        End If
        cmd_SAP_reader_2.Close()
        CNN2.Close()
        Return risposta
    End Function

    Sub insert_into_owor(par_docentry_odp As Integer, par_docnum_odp As Integer, par_itemcode As String, par_itemname As String, par_quantità As String, par_commessa As String, par_fase As String, par_cliente As String, par_data_inizio As String, par_data_fine As String, par_magazzino_destinazione As String, par_suppcatnum As String, par_JrnlMemo As String, par_pindicator As String, par_versionnum As String, par_produzione As String, par_series As String, par_ordine_cliente As String, par_bp_code As String, par_ultimo_progressivo_Commessa As Integer, par_usersign As Integer)

        par_itemname = Replace(par_itemname, "'", " ")
        par_quantità = Replace(par_quantità, ",", ".")
        par_commessa = Replace(par_commessa, "'", "")
        par_cliente = Replace(par_cliente, "'", " ")
        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand

        CMD_SAP_1.Connection = Cnn1
        CMD_SAP_1.CommandText = "INSERT INTO OWOR
(OWOR.DOCENTRY,OWOR.DOCNUM,OWOR.ITEMCODE,OWOR.PRODNAME, owor.plannedqty, OWOR.[U_PRG_AZS_Commessa], owor.U_fase, OWOR.U_UTILIZZ,OWOR.STARTDATE,OWOR.DUEDATE,owor.cmpltqty,owor.rjctqty,owor.postdate,owor.warehouse,owor.uom,owor.JrnlMemo,owor.pindicator,owor.uomentry,owor.updalloc,owor.versionnum, owor.U_produzione,owor.series, owor.usersign, OWOR.originnum, owor.cardcode,owor.U_Progressivo_commessa,owor.createdate) 
                        values (" & par_docentry_odp & "+1,'" & par_docnum_odp & "'+1, '" & par_itemcode & "','" & par_itemname & "','" & par_quantità & "','" & par_commessa & "','" & par_fase & "','" & par_cliente & "',CONVERT(DATETIME, '" & par_data_inizio & "',103),CONVERT(DATETIME, '" & par_data_fine & "',103),0,0,getdate(),'" & par_magazzino_destinazione & "','" & par_suppcatnum & "','" & par_JrnlMemo & "','" & par_pindicator & "','-1','C','" & par_versionnum & "','" & par_produzione & "','" & par_series & "','" & par_usersign & "', '" & par_ordine_cliente & "', '" & par_bp_code & "'," & par_ultimo_progressivo_Commessa & ",getdate() )"
        CMD_SAP_1.ExecuteNonQuery()

        Cnn1.Close()
    End Sub

    Sub aumenta_quantità_owor(par_docentry_odp As Integer, par_quantità As String)


        par_quantità = Replace(par_quantità, ",", ".")
        Dim vecchia_quantita As String = Replace(vecchia_quantita_odp(par_docentry_odp), ",", ".")

        Dim Cnn1 As New SqlConnection
        Cnn1.ConnectionString = Homepage.sap_tirelli
        Cnn1.Open()

        Dim CMD_SAP_1 As New SqlCommand

        CMD_SAP_1.Connection = Cnn1
        CMD_SAP_1.CommandText = "update OWOR set owor.plannedqty=owor.plannedqty+" & par_quantità & "
        where owor.docentry=" & par_docentry_odp & ""


        CMD_SAP_1.ExecuteNonQuery()

        CMD_SAP_1.CommandText = "update wor1 set wor1.[PlannedQty]=wor1.[PlannedQty]*((" & vecchia_quantita & "+" & par_quantità & ")/" & vecchia_quantita & ")
,wor1.[U_PRG_WIP_QTADATRASF]=wor1.[U_PRG_WIP_QTADATRASF]*((" & vecchia_quantita & "+" & par_quantità & ")/" & vecchia_quantita & ")

        where wor1.docentry=" & par_docentry_odp & ""


        CMD_SAP_1.ExecuteNonQuery()

        Cnn1.Close()
    End Sub

    Public Function vecchia_quantita_odp(par_docentry_odp As Integer)
        Dim vecchia_quantita As String = "0"
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN6

        CMD_SAP.CommandText = "select t0.plannedqty
from owor t0 
where t0.docentry= " & par_docentry_odp & ""



        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            vecchia_quantita = cmd_SAP_reader("plannedqty")


        End If
        cmd_SAP_reader.Close()
        CNN6.Close()
        Return vecchia_quantita

    End Function



    Sub max_docentry_docnum()


        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN6

        CMD_SAP.CommandText = "select max(t0.docentry) as 'Docentry',max(t0.docnum) as 'Docnum' from owor t0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader
        If cmd_SAP_reader.Read() = True Then

            docentry_ODP = cmd_SAP_reader("docentry")
            docnum_ODP = cmd_SAP_reader("docnum")


        End If
        cmd_SAP_reader.Close()
        CNN6.Close()
        AGGIUSTA_NUMERATORE()

    End Sub

    Public Function progressivo_commessa(par_COMMESSA)
        Dim progressivo As Integer = 0
        Dim CNN6 As New SqlConnection
        CNN6.ConnectionString = Homepage.sap_tirelli
        CNN6.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN6

        CMD_SAP.CommandText = "SELECT MAX(coalesce(T0.U_Progressivo_commessa,0)) AS 'Progressivo_commessa'
FROM OWOR T0
WHERE T0.U_PRG_AZS_Commessa='" & par_COMMESSA & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() = True Then

            If Not cmd_SAP_reader("Progressivo_commessa") Is System.DBNull.Value Then
                progressivo = cmd_SAP_reader("Progressivo_commessa") + 1
            Else
                progressivo = 1
            End If
        Else
            progressivo = 1


        End If
        cmd_SAP_reader.Close()
        CNN6.Close()

        Return progressivo
    End Function

    Sub IMPORTA_DB(par_docentry_odp As Integer, par_docnum_odp As Integer, par_itemcode As String, par_itemname As String, par_quantità As Decimal, par_commessa As String, par_fase As String, par_cliente As String, par_data_inizio As String, par_data_fine As String, par_magazzino_destinazione As String, par_suppcatnum As String, par_JrnlMemo As String, par_pindicator As String, par_versionnum As String, par_produzione As String, par_series As String, par_ordine_cliente As String, par_bp_code As String, par_ultimo_progressivo_Commessa As Integer, par_riga_utilizzata As Integer)

        Dim ItemcodeDB As String
        Dim DescrizioneDB As String
        Dim QuantitàDB As Decimal
        ' Dim VisorderDB As Integer
        ' Dim maxvisorder As Integer
        Dim MagazzinoDB As String
        Dim TypeDB As Integer
        Dim AttrezzaggioDB As Decimal
        Dim TestoDB As String
        Dim Resallocdb As String
        Dim UomcodeDB As String
        Dim UomentryDB As String
        Dim IssueMthd As String

        Dim DatrasferireDB As String


        Dim CNN4 As New SqlConnection
        CNN4.ConnectionString = Homepage.sap_tirelli
        CNN4.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN4



        CMD_SAP.CommandText = "SELECT 
coalesce(T0.[Code],'') as 'ItemcodeDB'
, CASE WHEN t1.itemname IS NULL THEN'' ELSE T1.ITEMNAME END as 'DescrizioneDB'
, T0.[Quantity]/T2.[Qauntity] as 'QuantitàDB'
, T0.[VisOrder] as 'VisorderDB'
, case when t0.warehouse is null then '01' else T0.[Warehouse] end as 'MagazzinoDB'
, T0.Type as 'TypeDB'
, case when T0.AddQuantit is null then 0 else T0.AddQuantit end as 'AttrezzaggioDB'
, case when T0.LineText is null then '' else t0.linetext end as 'TestoDB'
, case when t0.type=4 then 'null' else 'F' end as'Resallocdb'
, case when t0.type=4 then '-1' else '' end as'UOMENTRYDB'
, case when t0.type=4 then 'Manuale' else '' end as'UomcodeDB'
, case when (substring(t0.code,1,1)='0' or substring(t0.code,1,1)='C' or substring(t0.code,1,1)='D' or substring(t0.code,1,1)='F') then t0.quantity else 0 end as 'DatrasferireDB' 
, coalesce(t1.phantom,'') as 'Pseudoarticolo'
, coalesce(t1.u_prg_tir_Explosion,'Y') as 'Esplosione_distinta'
, T0.[IssueMthd]
from itt1 t0 left join oitm t1 on t0.code=t1.itemcode
left join oitt t2 on t0.father=t2.code 
 WHERE T0.[Father] ='" & par_itemcode & "' and T0.Type<>'-18'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            ItemcodeDB = cmd_SAP_reader("ItemcodeDB")
            DescrizioneDB = cmd_SAP_reader("DescrizioneDB")
            QuantitàDB = cmd_SAP_reader("QuantitàDB")

            MagazzinoDB = cmd_SAP_reader("MagazzinoDB")
            TypeDB = cmd_SAP_reader("TypeDB")
            AttrezzaggioDB = cmd_SAP_reader("AttrezzaggioDB")
            TestoDB = cmd_SAP_reader("TestoDB")
            If cmd_SAP_reader("resallocdb") = "null" Then
                Resallocdb = Nothing
            Else
                Resallocdb = cmd_SAP_reader("resallocdb")
            End If
            UomentryDB = cmd_SAP_reader("UomentryDB")
            UomcodeDB = cmd_SAP_reader("UomcodeDB")
            DatrasferireDB = cmd_SAP_reader("DatrasferireDB")
            IssueMthd = cmd_SAP_reader("IssueMthd")

            If TypeDB = 4 Then
                If par_magazzino_destinazione = "B02" Or par_magazzino_destinazione = "BCAP2" Or par_magazzino_destinazione = "B01" Then
                    MagazzinoDB = "B01"
                End If
            End If

            If (cmd_SAP_reader("Pseudoarticolo") = "Y" Or cmd_SAP_reader("Esplosione_distinta") = "N") Then
                If MessageBox.Show($"Il codice " & ItemcodeDB & " " & DescrizioneDB & " è FANTASMA , Vuoi inserire direttamente i figli? ", "Inserisci figli", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

                    If par_riga_utilizzata < max_riga_odp Then
                        par_riga_utilizzata = max_riga_odp + 1
                    End If
                    IMPORTA_DB(par_docentry_odp, par_docnum_odp, ItemcodeDB, Magazzino.OttieniDettagliAnagrafica(ItemcodeDB).Descrizione, par_quantità * QuantitàDB, par_commessa, par_fase, par_cliente, par_data_inizio, par_data_fine, par_magazzino_destinazione, par_suppcatnum, par_JrnlMemo, par_pindicator, par_versionnum, par_produzione, par_series, par_ordine_cliente, par_bp_code, par_ultimo_progressivo_Commessa, par_riga_utilizzata)

                Else

                    Dim Cnn1 As New SqlConnection
                    Cnn1.ConnectionString = Homepage.sap_tirelli
                    Cnn1.Open()

                    Dim CMD_SAP_1 As New SqlCommand

                    CMD_SAP_1.Connection = Cnn1
                    If TypeDB = 4 Then

                        CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType,  wor1.[StartDate], wor1.[EndDate], wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf])
                                            VALUES(" & par_docentry_odp & "+1, " & par_riga_utilizzata & " ,'" & ItemcodeDB & "','" & DescrizioneDB & "'," & par_riga_utilizzata & ", " & Replace(QuantitàDB, ",", ".") & " ,'" & Replace(AttrezzaggioDB, ",", ".") & " ', " & Replace(AttrezzaggioDB, ",", ".") & " + (" & Replace(QuantitàDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & "),'" & MagazzinoDB & "'," & TypeDB & ",'" & IssueMthd & "', CONVERT(DATETIME, '" & par_data_inizio & "',103), CONVERT(DATETIME, '" & par_data_fine & "',103),0,0,0," & Replace(QuantitàDB, ",", ".") & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & Replace(DatrasferireDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & ")"

                    ElseIf TypeDB = -18 Then

                        CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType, wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf])
                                            VALUES(" & par_docentry_odp & "+1, " & par_riga_utilizzata & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & par_riga_utilizzata & ", " & Replace(QuantitàDB, ",", ".") & " ,'" & Replace(AttrezzaggioDB, ",", ".") & " ', " & Replace(AttrezzaggioDB, ",", ".") & " + (" & Replace(QuantitàDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & "),'" & MagazzinoDB & "'," & TypeDB & ",'B',0,0,0," & Replace(QuantitàDB, ",", ".") & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & Replace(DatrasferireDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & ")"

                    Else
                        CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType,  wor1.[StartDate], wor1.[EndDate], wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf],WOR1.RESALLOC)
                                            VALUES(" & par_docentry_odp & "+1, " & par_riga_utilizzata & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & par_riga_utilizzata & ", " & Replace(QuantitàDB, ",", ".") & " ,'" & Replace(AttrezzaggioDB, ",", ".") & " ', " & Replace(AttrezzaggioDB, ",", ".") & " + (" & Replace(QuantitàDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & "),'" & MagazzinoDB & "'," & TypeDB & ",'B', CONVERT(DATETIME, '" & par_data_inizio & "',103), CONVERT(DATETIME, '" & par_data_fine & "',103),0,0,0," & Replace(QuantitàDB, ",", ".") & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & Replace(DatrasferireDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & ",'" & Resallocdb & "')"

                    End If

                    CMD_SAP_1.ExecuteNonQuery()
                    Cnn1.Close()
                    aggiusta_CONFERMATO(ItemcodeDB)
                    aggiusta_righe_odp_confermato_tot_ordinato_tot(ItemcodeDB)
                    par_riga_utilizzata = par_riga_utilizzata + 1

                    If par_riga_utilizzata >= max_riga_odp Then
                        max_riga_odp = par_riga_utilizzata
                    End If

                End If




            Else


                If par_riga_utilizzata < max_riga_odp Then
                    par_riga_utilizzata = max_riga_odp
                End If

                Dim Cnn1 As New SqlConnection
                Cnn1.ConnectionString = Homepage.sap_tirelli
                Cnn1.Open()

                Dim CMD_SAP_1 As New SqlCommand

                CMD_SAP_1.Connection = Cnn1
                If TypeDB = 4 Then

                    CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType,  wor1.[StartDate], wor1.[EndDate], wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf])
                                            VALUES(" & par_docentry_odp & "+1, " & par_riga_utilizzata & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & par_riga_utilizzata & ", " & Replace(QuantitàDB, ",", ".") & " ,'" & Replace(AttrezzaggioDB, ",", ".") & " ', " & Replace(AttrezzaggioDB, ",", ".") & " + (" & Replace(QuantitàDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & "),'" & MagazzinoDB & "'," & TypeDB & ",'" & IssueMthd & "', CONVERT(DATETIME, '" & par_data_inizio & "',103), CONVERT(DATETIME, '" & par_data_fine & "',103),0,0,0," & Replace(QuantitàDB, ",", ".") & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & Replace(DatrasferireDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & ")"

                ElseIf TypeDB = -18 Then

                    CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType, wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf])
                                            VALUES(" & par_docentry_odp & "+1, " & par_riga_utilizzata & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & par_riga_utilizzata & ", " & Replace(QuantitàDB, ",", ".") & " ,'" & Replace(AttrezzaggioDB, ",", ".") & " ', " & Replace(AttrezzaggioDB, ",", ".") & " + (" & Replace(QuantitàDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & "),'" & MagazzinoDB & "'," & TypeDB & ",'B',0,0,0," & Replace(QuantitàDB, ",", ".") & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & Replace(DatrasferireDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & ")"

                Else
                    CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType,  wor1.[StartDate], wor1.[EndDate], wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf],WOR1.RESALLOC)
                                            VALUES(" & par_docentry_odp & "+1, " & par_riga_utilizzata & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & par_riga_utilizzata & ", " & Replace(QuantitàDB, ",", ".") & " ,'" & Replace(AttrezzaggioDB, ",", ".") & " ', " & Replace(AttrezzaggioDB, ",", ".") & " + (" & Replace(QuantitàDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & "),'" & MagazzinoDB & "'," & TypeDB & ",'B', CONVERT(DATETIME, '" & par_data_inizio & "',103), CONVERT(DATETIME, '" & par_data_fine & "',103),0,0,0," & Replace(QuantitàDB, ",", ".") & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & Replace(DatrasferireDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & ",'" & Resallocdb & "')"

                End If

                CMD_SAP_1.ExecuteNonQuery()
                Cnn1.Close()
                aggiusta_CONFERMATO(ItemcodeDB)
                aggiusta_righe_odp_confermato_tot_ordinato_tot(ItemcodeDB)
                par_riga_utilizzata = par_riga_utilizzata + 1

                If par_riga_utilizzata >= max_riga_odp Then
                    max_riga_odp = par_riga_utilizzata
                End If

            End If





        Loop


        cmd_SAP_reader.Close()
        CNN4.Close()

        aggiusta_risorse(par_docentry_odp)
        aggiusta_testi(par_docentry_odp)






        'riga_odp = 0
        ' inserisci_righe_odp(par_docentry_odp, par_docnum_odp, par_itemcode, par_itemname, par_quantità, par_commessa, par_fase, par_cliente, par_data_inizio, par_data_fine, par_magazzino_destinazione, par_suppcatnum, par_JrnlMemo, par_pindicator, par_versionnum, par_produzione, par_series, par_ordine_cliente, par_bp_code, par_ultimo_progressivo_Commessa, riga_odp)
    End Sub

    Sub inserisci_righe_odp(par_docentry_odp As Integer, par_docnum_odp As Integer, par_itemcode As String, par_itemname As String, par_quantità As Decimal, par_commessa As String, par_fase As String, par_cliente As String, par_data_inizio As String, par_data_fine As String, par_magazzino_destinazione As String, par_suppcatnum As String, par_JrnlMemo As String, par_pindicator As String, par_versionnum As String, par_produzione As String, par_series As String, par_ordine_cliente As String, par_bp_code As String, par_ultimo_progressivo_Commessa As Integer, par_riga_utilizzata As Integer)


        '        Dim ItemcodeDB As String
        '        Dim DescrizioneDB As String
        '        Dim QuantitàDB As Decimal
        '        ' Dim VisorderDB As Integer
        '        ' Dim maxvisorder As Integer
        '        Dim MagazzinoDB As String
        '        Dim TypeDB As Integer
        '        Dim AttrezzaggioDB As Decimal
        '        Dim TestoDB As String
        '        Dim Resallocdb As String
        '        Dim UomcodeDB As String
        '        Dim UomentryDB As String

        '        Dim DatrasferireDB As String


        '        Dim CNN4 As New SqlConnection
        '        CNN4.ConnectionString = Homepage.sap_tirelli
        '        CNN4.Open()

        '        Dim CMD_SAP As New SqlCommand
        '        Dim cmd_SAP_reader As SqlDataReader
        '        CMD_SAP.Connection = CNN4



        '        CMD_SAP.CommandText = "SELECT 
        'coalesce(T0.[Code],'') as 'ItemcodeDB'
        ', CASE WHEN t1.itemname IS NULL THEN'' ELSE T1.ITEMNAME END as 'DescrizioneDB'
        ', T0.[Quantity]/T2.[Qauntity] as 'QuantitàDB'
        ', T0.[VisOrder] as 'VisorderDB'
        ', case when t0.warehouse is null then '' else T0.[Warehouse] end as 'MagazzinoDB'
        ', T0.Type as 'TypeDB'
        ', case when T0.AddQuantit is null then 0 else T0.AddQuantit end as 'AttrezzaggioDB'
        ', case when T0.LineText is null then '' else t0.linetext end as 'TestoDB'
        ', case when t0.type=4 then 'null' else 'F' end as'Resallocdb'
        ', case when t0.type=4 then '-1' else '' end as'UOMENTRYDB'
        ', case when t0.type=4 then 'Manuale' else '' end as'UomcodeDB'
        ', case when (substring(t0.code,1,1)='0' or substring(t0.code,1,1)='C' or substring(t0.code,1,1)='D' or substring(t0.code,1,1)='F') then t0.quantity else 0 end as 'DatrasferireDB' 
        ', coalesce(t1.phantom,'') as 'Pseudoarticolo'
        ', coalesce(t1.u_prg_tir_Explosion,'Y') as 'Esplosione_distinta'
        'from itt1 t0 left join oitm t1 on t0.code=t1.itemcode
        'left join oitt t2 on t0.father=t2.code 
        ' WHERE T0.[Father] ='" & par_itemcode & "' and T0.Type<>'-18'"

        '        cmd_SAP_reader = CMD_SAP.ExecuteReader

        '        Do While cmd_SAP_reader.Read()

        '            ItemcodeDB = cmd_SAP_reader("ItemcodeDB")
        '            DescrizioneDB = cmd_SAP_reader("DescrizioneDB")
        '            QuantitàDB = cmd_SAP_reader("QuantitàDB")

        '            MagazzinoDB = cmd_SAP_reader("MagazzinoDB")
        '            TypeDB = cmd_SAP_reader("TypeDB")
        '            AttrezzaggioDB = cmd_SAP_reader("AttrezzaggioDB")
        '            TestoDB = cmd_SAP_reader("TestoDB")
        '            If cmd_SAP_reader("resallocdb") = "null" Then
        '                Resallocdb = Nothing
        '            Else
        '                Resallocdb = cmd_SAP_reader("resallocdb")
        '            End If
        '            UomentryDB = cmd_SAP_reader("UomentryDB")
        '            UomcodeDB = cmd_SAP_reader("UomcodeDB")
        '            DatrasferireDB = cmd_SAP_reader("DatrasferireDB")


        '            If TypeDB = 4 Then
        '                If par_magazzino_destinazione = "B02" Or par_magazzino_destinazione = "BCAP2" Or par_magazzino_destinazione = "B01" Then
        '                    MagazzinoDB = "B01"
        '                End If
        '            End If

        '            If (cmd_SAP_reader("Pseudoarticolo") = "Y" Or cmd_SAP_reader("Esplosione_distinta") = "N") Then
        '                If MessageBox.Show($"Il codice " & ItemcodeDB & " " & DescrizioneDB & " è FANTASMA , Vuoi inserire direttamente i figli? ", "Inserisci figli", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

        '                    IMPORTA_DB(par_docentry_odp, par_docnum_odp, ItemcodeDB, Magazzino.OttieniDettagliAnagrafica(ItemcodeDB).Descrizione, par_quantità * QuantitàDB, par_commessa, par_fase, par_cliente, par_data_inizio, par_data_fine, par_magazzino_destinazione, par_suppcatnum, par_JrnlMemo, par_pindicator, par_versionnum, par_produzione, par_series, par_ordine_cliente, par_bp_code, par_ultimo_progressivo_Commessa, par_riga_utilizzata)

        '                Else

        '                    Dim Cnn1 As New SqlConnection
        '                    Cnn1.ConnectionString = Homepage.sap_tirelli
        '                    Cnn1.Open()

        '                    Dim CMD_SAP_1 As New SqlCommand

        '                    CMD_SAP_1.Connection = Cnn1
        '                    If TypeDB = 4 Then

        '                        CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType,  wor1.[StartDate], wor1.[EndDate], wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf])
        '                                            VALUES(" & par_docentry_odp & "+1, " & par_riga_utilizzata & " ,'" & ItemcodeDB & "','" & DescrizioneDB & "'," & par_riga_utilizzata & ", " & Replace(QuantitàDB, ",", ".") & " ,'" & Replace(AttrezzaggioDB, ",", ".") & " ', " & Replace(AttrezzaggioDB, ",", ".") & " + (" & Replace(QuantitàDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & "),'" & MagazzinoDB & "'," & TypeDB & ",'B', CONVERT(DATETIME, '" & par_data_inizio & "',103), CONVERT(DATETIME, '" & par_data_fine & "',103),0,0,0," & Replace(QuantitàDB, ",", ".") & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & Replace(DatrasferireDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & ")"

        '                    ElseIf TypeDB = -18 Then

        '                        CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType, wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf])
        '                                            VALUES(" & par_docentry_odp & "+1, " & par_riga_utilizzata & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & par_riga_utilizzata & ", " & Replace(QuantitàDB, ",", ".") & " ,'" & Replace(AttrezzaggioDB, ",", ".") & " ', " & Replace(AttrezzaggioDB, ",", ".") & " + (" & Replace(QuantitàDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & "),'" & MagazzinoDB & "'," & TypeDB & ",'B',0,0,0," & Replace(QuantitàDB, ",", ".") & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & Replace(DatrasferireDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & ")"

        '                    Else
        '                        CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType,  wor1.[StartDate], wor1.[EndDate], wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf],WOR1.RESALLOC)
        '                                            VALUES(" & par_docentry_odp & "+1, " & par_riga_utilizzata & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & par_riga_utilizzata & ", " & Replace(QuantitàDB, ",", ".") & " ,'" & Replace(AttrezzaggioDB, ",", ".") & " ', " & Replace(AttrezzaggioDB, ",", ".") & " + (" & Replace(QuantitàDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & "),'" & MagazzinoDB & "'," & TypeDB & ",'B', CONVERT(DATETIME, '" & par_data_inizio & "',103), CONVERT(DATETIME, '" & par_data_fine & "',103),0,0,0," & Replace(QuantitàDB, ",", ".") & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & Replace(DatrasferireDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & ",'" & Resallocdb & "')"

        '                    End If

        '                    CMD_SAP_1.ExecuteNonQuery()
        '                    Cnn1.Close()
        '                    aggiusta_CONFERMATO(ItemcodeDB)
        '                    aggiusta_righe_odp_confermato_tot_ordinato_tot(ItemcodeDB)
        '                    par_riga_utilizzata = par_riga_utilizzata + 1
        '                End If




        '            Else

        '                Dim Cnn1 As New SqlConnection
        '                Cnn1.ConnectionString = Homepage.sap_tirelli
        '                Cnn1.Open()

        '                Dim CMD_SAP_1 As New SqlCommand

        '                CMD_SAP_1.Connection = Cnn1
        '                If TypeDB = 4 Then

        '                    CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType,  wor1.[StartDate], wor1.[EndDate], wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf])
        '                                            VALUES(" & par_docentry_odp & "+1, " & par_riga_utilizzata & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & par_riga_utilizzata & ", " & Replace(QuantitàDB, ",", ".") & " ,'" & Replace(AttrezzaggioDB, ",", ".") & " ', " & Replace(AttrezzaggioDB, ",", ".") & " + (" & Replace(QuantitàDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & "),'" & MagazzinoDB & "'," & TypeDB & ",'B', CONVERT(DATETIME, '" & par_data_inizio & "',103), CONVERT(DATETIME, '" & par_data_fine & "',103),0,0,0," & Replace(QuantitàDB, ",", ".") & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & Replace(DatrasferireDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & ")"

        '                ElseIf TypeDB = -18 Then

        '                    CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType, wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf])
        '                                            VALUES(" & par_docentry_odp & "+1, " & par_riga_utilizzata & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & par_riga_utilizzata & ", " & Replace(QuantitàDB, ",", ".") & " ,'" & Replace(AttrezzaggioDB, ",", ".") & " ', " & Replace(AttrezzaggioDB, ",", ".") & " + (" & Replace(QuantitàDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & "),'" & MagazzinoDB & "'," & TypeDB & ",'B',0,0,0," & Replace(QuantitàDB, ",", ".") & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & Replace(DatrasferireDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & ")"

        '                Else
        '                    CMD_SAP_1.CommandText = "insert into WOR1 (WOR1.DOCENTRY, WOR1.LINENUM, WOR1.ITEMCODE, wor1.itemname, WOR1.VISORDER, WOR1.[BaseQty], wor1.[AdditQty], WOR1.PLANNEDQTY, wor1.warehouse, wor1.itemtype, wor1.IssueType,  wor1.[StartDate], wor1.[EndDate], wor1.issuedqty,wor1.comptotal,wor1.pickqty,wor1.baseqtynum,wor1.baseqtyden,wor1.releaseqty, wor1.uomentry, wor1.uomcode, wor1.linetext, WOR1.[U_PRG_WIP_QtaDaTrasf],WOR1.RESALLOC)
        '                                            VALUES(" & par_docentry_odp & "+1, " & par_riga_utilizzata & ",'" & ItemcodeDB & "','" & DescrizioneDB & "'," & par_riga_utilizzata & ", " & Replace(QuantitàDB, ",", ".") & " ,'" & Replace(AttrezzaggioDB, ",", ".") & " ', " & Replace(AttrezzaggioDB, ",", ".") & " + (" & Replace(QuantitàDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & "),'" & MagazzinoDB & "'," & TypeDB & ",'B', CONVERT(DATETIME, '" & par_data_inizio & "',103), CONVERT(DATETIME, '" & par_data_fine & "',103),0,0,0," & Replace(QuantitàDB, ",", ".") & ",1,0,' " & UomentryDB & "', '" & UomcodeDB & "', '" & TestoDB & "', " & Replace(DatrasferireDB, ",", ".") & " * " & Replace(par_quantità, ",", ".") & ",'" & Resallocdb & "')"

        '                End If

        '                CMD_SAP_1.ExecuteNonQuery()
        '                Cnn1.Close()
        '                aggiusta_CONFERMATO(ItemcodeDB)
        '                aggiusta_righe_odp_confermato_tot_ordinato_tot(ItemcodeDB)
        '                par_riga_utilizzata = par_riga_utilizzata + 1
        '            End If





        '        Loop


        '        cmd_SAP_reader.Close()
        '        CNN4.Close()

        '        aggiusta_risorse(par_docentry_odp)
        '        aggiusta_testi(par_docentry_odp)


    End Sub



    Sub trova_dato_da_excel_par_importazione_odp(par_percorso_file As String, par_nome_foglio As String, par_riga_inizio As Integer, par_riga_fine As Integer)




        Dim itemcode As String
        Dim produzione As String
        Dim Fase As String
        Dim data_inizio As String
        Dim data_fine As String
        Dim Commessa_odp As String
        Dim Cliente As String
        Dim quantità_odp As String
        Dim Ordine_cliente As String
        Dim BP_code As String
        Dim MAG_DESTINAZIONE As String = ""

        Dim Excel As Excel.Application
        Excel = CreateObject("Excel.application")

        Excel.Workbooks.Open(par_percorso_file)
        Excel.Visible = True


        Do While par_riga_inizio <= par_riga_fine


            If Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value <> Nothing Then
                itemcode = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 1).value
                produzione = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 2).value
                Fase = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 3).value
                data_inizio = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 4).value
                data_fine = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 5).value
                Commessa_odp = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 6).value
                Cliente = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 7).value
                quantità_odp = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 8).value
                Ordine_cliente = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 9).value
                BP_code = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 10).value
                MAG_DESTINAZIONE = Excel.Sheets(par_nome_foglio).Cells(par_riga_inizio, 11).value

                If MAG_DESTINAZIONE = "" Then
                    MsgBox("MEttere nella colonna 11 il magazzino di destinazione")
                    End
                End If

                procedura_lancio_odp(itemcode, produzione, Fase, data_inizio, data_fine, Commessa_odp, Cliente, quantità_odp, Ordine_cliente, BP_code, MAG_DESTINAZIONE)

            End If
            par_riga_inizio = par_riga_inizio + 1

        Loop
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        If TextBox8.Text.StartsWith("M") Then
            ' Azioni se il testo inizia con "M"
            ComboBox7.Text = trova_magazzino_destinazione_ODP(TextBox8.Text, ComboBox5.Text, Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).utente_sap_salvato)

        Else

            If TextBox8.Text.StartsWith("_") Then
                ' Rimuovi il primo carattere "_" dalla stringa
                Dim cleanedText As String = TextBox8.Text.Replace("_", "")
                ComboBox7.Text = trova_magazzino_destinazione_OC(cleanedText)
            Else
                trova_magazzino_destinazione_OC(TextBox8.Text)
            End If



        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Form_stock.Show()
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        Solleciti_OA.Show()
        Solleciti_OA.Owner = Me
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        Dim par_datagridiview As DataGridView = DataGridView1


        If par_datagridiview.Rows(e.RowIndex).Cells(columnName:="Invio_Mail").Value = "SI" Then

            par_datagridiview.Rows(e.RowIndex).Cells(columnName:="Invio_Mail").Style.BackColor = Color.Lime
        Else
            par_datagridiview.Rows(e.RowIndex).Cells(columnName:="Invio_Mail").Style.BackColor = Color.Yellow

        End If

    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        If visualizzazione = "Documenti" Then
            lista_documenti(DataGridView1)
        ElseIf visualizzazione = "Righe" Then
            filtra_righe()
        End If
    End Sub

    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellContentClick

    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        filtra_righe()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        Dim appExcel As New Excel.Application
        Dim workbook As Excel.Workbook = Nothing

        Try
            ' Ottieni tutti i file Excel nella cartella
            Dim files As String() = Directory.GetFiles(Homepage.percorso_acquisti & "Documenti vari di lavoro\Report fornitori\Report puntualità", "*.xls*")

            For Each filePath As String In files
                ' Apri il file Excel
                Apri_file_Excel(filePath, Homepage.percorso_acquisti & "Documenti vari di lavoro\Report fornitori\Report puntualità\report eseguiti\")
            Next

        Catch ex As Exception
            MsgBox("Errore: " & ex.Message)
        Finally
            ' Chiudi Excel se necessario
            If Not workbook Is Nothing Then
                workbook.Close(False)
            End If
            appExcel.Quit()
            ReleaseObject(workbook)
            ReleaseObject(appExcel)
        End Try
    End Sub

    Sub Apri_file_Excel(par_percorso_file As String, par_cartella_destinazione As String)

        Dim appExcel As New Excel.Application
        Dim workbook As Excel.Workbook = Nothing
        Dim pdfFilePath As String
        Dim wsEmail As Excel.Worksheet
        Dim emailList As New List(Of String)
        Dim row As Integer

        Try
            ' Apri il file Excel
            workbook = appExcel.Workbooks.Open(par_percorso_file)
            ' appExcel.Visible = True

            ' 🔹 Aggiorna tutte le QueryTables nei fogli
            For Each ws As Excel.Worksheet In workbook.Sheets
                For Each qt As Excel.QueryTable In ws.QueryTables
                    qt.Refresh(False)
                Next
            Next

            ' 🔹 Aggiorna tutte le connessioni dati (se presenti)
            For Each conn As Excel.WorkbookConnection In workbook.Connections
                conn.OLEDBConnection.BackgroundQuery = False ' Aspetta il completamento
                conn.Refresh()
            Next

            ' 🔹 Aspetta che Excel abbia completato i calcoli
            Do While appExcel.CalculationState <> Excel.XlCalculationState.xlDone
                System.Threading.Thread.Sleep(500) ' Attesa attiva ogni 500ms
            Loop

            ' 🔹 Leggi gli indirizzi email dalla scheda "Email"
            wsEmail = workbook.Sheets("Mail") ' Assumendo che la scheda si chiami esattamente "Email"
            row = 1

            Do While Not String.IsNullOrEmpty(wsEmail.Cells(row, 1).Value)
                emailList.Add(wsEmail.Cells(row, 1).Value)
                row = row + 1
            Loop

            ' Ora esporta in PDF
            pdfFilePath = Esporta_in_pdf(workbook, par_percorso_file, par_cartella_destinazione)

            ' Apri il PDF creato
            If Not String.IsNullOrEmpty(pdfFilePath) Then
                Process.Start(pdfFilePath)
            Else
                MsgBox("Errore nell'esportazione del PDF.")
            End If

        Catch ex As Exception
            MsgBox("Errore: " & ex.Message)

        Finally
            ' Chiudi il workbook e l'applicazione Excel
            If workbook IsNot Nothing Then workbook.Close(False)
            appExcel.Quit()

            ' Rilascia le risorse
            ReleaseObject(workbook)
            ReleaseObject(appExcel)
        End Try

        ' Invia il report con gli allegati agli indirizzi email raccolti
        InviaReportConAllegato(pdfFilePath, emailList)
    End Sub

    Function Esporta_in_pdf(workbook As Excel.Workbook, par_percorso_file As String, par_cartella_destinazione As String) As String
        Dim pdfFilePath As String = ""

        Try
            ' Estrai il nome del file senza estensione
            Dim nomeFileExcel As String = IO.Path.GetFileNameWithoutExtension(par_percorso_file)

            ' Determina il percorso del PDF con lo stesso nome dell'Excel
            Dim basePdfFilePath As String = IO.Path.Combine(par_cartella_destinazione, nomeFileExcel & ".pdf")
            Dim version As Integer = 1

            pdfFilePath = basePdfFilePath
            While IO.File.Exists(pdfFilePath)
                pdfFilePath = IO.Path.Combine(par_cartella_destinazione, nomeFileExcel & "_" & version.ToString("D2") & ".pdf")
                version += 1
            End While

            ' Creiamo una lista di fogli visibili
            Dim fogliDaStampare As New List(Of Excel.Worksheet)

            For Each ws As Excel.Worksheet In workbook.Sheets
                If ws.Visible = Excel.XlSheetVisibility.xlSheetVisible Then
                    ' Adatta l'area di stampa al contenuto
                    ws.PageSetup.Zoom = False
                    ws.PageSetup.FitToPagesWide = 1
                    ws.PageSetup.FitToPagesTall = False
                    ws.PageSetup.Orientation = Excel.XlPageOrientation.xlPortrait

                    ' Imposta l'area di stampa solo se non è già definita
                    If ws.UsedRange.Rows.Count > 0 And ws.UsedRange.Columns.Count > 0 Then
                        ws.PageSetup.PrintArea = ws.UsedRange.Address
                    End If

                    fogliDaStampare.Add(ws)
                End If
            Next

            ' Se non ci sono fogli visibili, interrompe l'esecuzione
            If fogliDaStampare.Count = 0 Then
                MsgBox("Errore: Nessun foglio visibile da esportare.")
                Return ""
            End If

            ' Esporta tutti i fogli visibili in un unico PDF
            workbook.ExportAsFixedFormat(Type:=Excel.XlFixedFormatType.xlTypePDF,
                                     Filename:=pdfFilePath,
                                     Quality:=Excel.XlFixedFormatQuality.xlQualityStandard,
                                     IncludeDocProperties:=True,
                                     IgnorePrintAreas:=False,
                                     OpenAfterPublish:=True)

        Catch ex As Exception
            MsgBox("Errore durante l'esportazione in PDF: " & ex.Message)
        End Try

        Return pdfFilePath
    End Function
    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            ' Ignora l'errore, ma lascia obj = Nothing
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub


End Class
