Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop


Public Class Form_ritiri
    Public Elenco_owner(1000) As String
    Public Indice_Modifica As Integer
    Public Elenco_Dipendenti(1000) As Integer
    Public Nomi_Dipendenti(1000) As String
    Public Num_Dipendenti As Integer
    Private filtro_bp As String
    Private filtro_cardname As String

    Private Sub Form_ritiri_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
        Inserimento_owner(ComboBox1)
        CB_Tipo.Items.Add("Ritiro")
        CB_Tipo.Items.Add("Consegna")
        CB_Tipo.Items.Add("Ritiro e Consegna")
        CB_Tipo.SelectedIndex = 0
        TXT_Ora.Text = "08"
        TXT_Minuti.Text = "00"
        Aggiorna_Grid_Ritiri(DG_Ritiri)
        Cmd_Eseguito.Enabled = False
        TXT_Esecutore.Enabled = False
        Cmd_Annulla.Enabled = False
        Cmd_Aggiorna.Enabled = False
        Compila_Elenco_Dipendenti()

    End Sub

    Sub Inserimento_owner(PAR_COMBOBOX As ComboBox)
        PAR_COMBOBOX.Items.Clear()
        PAR_COMBOBOX.Items.Add("")
        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        CMD_SAP.CommandText = "SELECT T1.[empID] as 'Codice dipendenti', T1.[lastName] + ' ' + T1.[firstName] AS 'Nome'
        FROM [TIRELLI_40].[dbo].OHEM T1 
WHERE T1.ACTIVE='Y'
group by T1.[empID],T1.[lastName] + ' ' + T1.[firstName]
order by T1.[lastName] + ' ' + T1.[firstName]  "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 1
        Do While cmd_SAP_reader.Read()

            Elenco_owner(Indice) = cmd_SAP_reader("Codice dipendenti")

            PAR_COMBOBOX.Items.Add(cmd_SAP_reader("Nome"))
            Try
                If cmd_SAP_reader("Codice dipendenti") = Homepage.ID_SALVATO Then
                    PAR_COMBOBOX.Text = cmd_SAP_reader("Nome")

                End If
            Catch ex As Exception

            End Try

            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        Cnn.Close()

    End Sub 'Inserisco le risorse nella combo box

    Private Sub TXT_Nome_LostFocus(sender As Object, e As EventArgs) Handles TXT_Nome.LostFocus
        If TXT_Nome.Text.Length > 0 Then
            Dim Cnn_Ritiri As New SqlConnection
            Cnn_Ritiri.ConnectionString = Homepage.sap_tirelli
            Cnn_Ritiri.Open()
            Dim Cmd_Ritiri As New SqlCommand
            Dim Reader_Ritiri As SqlDataReader
            Cmd_Ritiri.Connection = Cnn_Ritiri
            Cmd_Ritiri.CommandText = "SELECT T0.[CardCode] as 'Codice'
, T0.[CardName] as 'Nome', case when T0.[Phone1] is null then ' ' else  T0.[Phone1] end as 'Tel',case when T1.[Street] is null then ' ' else  T1.[Street] end as 'Via',case when T1.[City] is null then ' ' else  T1.[City] end as 'Citta',case when T1.[ZipCode] is null then ' ' else  T1.[ZipCode] end as 'CAP'
FROM [TIRELLISRLDB].[dbo].OCRD T0  
INNER JOIN [TIRELLISRLDB].[dbo].CRD1 T1 ON T0.[CardCode] = T1.[CardCode]
WHERE 0=0 " & filtro_bp & filtro_cardname & ""

            Reader_Ritiri = Cmd_Ritiri.ExecuteReader()

            If Reader_Ritiri.Read() Then
                TXT_Nome.Text = Reader_Ritiri("Nome")
                TXT_Codice_BP.Text = Reader_Ritiri("Codice")
                TXT_Via.Text = Reader_Ritiri("Via")
                TXT_Tel.Text = Reader_Ritiri("Tel")
                TXT_Citta.Text = Reader_Ritiri("Citta")
                TXT_Cap.Text = Reader_Ritiri("CAP")
            Else
                TXT_Nome.Text = ""
                TXT_Codice_BP.Text = ""
                TXT_Via.Text = ""
                TXT_Tel.Text = ""
                TXT_Citta.Text = ""
                TXT_Cap.Text = ""
            End If


            Cnn_Ritiri.Close()
        End If
    End Sub

    Private Sub TXT_Codice_BP_LostFocus(sender As Object, e As EventArgs) Handles TXT_Codice_BP.LostFocus
        If TXT_Codice_BP.Text.Length > 0 Then


            Dim Cnn_Ritiri As New SqlConnection
            Cnn_Ritiri.ConnectionString = Homepage.sap_tirelli
            Cnn_Ritiri.Open()
            Dim Cmd_Ritiri As New SqlCommand
            Dim Reader_Ritiri As SqlDataReader
            Cmd_Ritiri.Connection = Cnn_Ritiri
            Cmd_Ritiri.CommandText = "SELECT T0.[CardCode] as 'Codice', 
T0.[CardName] as 'Nome'
, case when T0.[Phone1] is null then ' ' else  T0.[Phone1] end as 'Tel'
,case when T1.[Street] is null then ' ' else  T1.[Street] end as 'Via',case when T1.[City] is null then ' ' else  T1.[City] end as 'Citta',case when T1.[ZipCode] is null then ' ' else  T1.[ZipCode] end as 'CAP' 
FROM [TIRELLISRLDB].[dbo].OCRD T0  INNER JOIN [TIRELLISRLDB].[dbo].CRD1 T1 ON T0.[CardCode] = T1.[CardCode]
WHERE 0=0 " & filtro_bp & filtro_cardname & ""

            Reader_Ritiri = Cmd_Ritiri.ExecuteReader()

            If Reader_Ritiri.Read() Then
                TXT_Nome.Text = Reader_Ritiri("Nome")
                TXT_Codice_BP.Text = Reader_Ritiri("Codice")
                TXT_Via.Text = Reader_Ritiri("Via")
                TXT_Tel.Text = Reader_Ritiri("Tel")
                TXT_Citta.Text = Reader_Ritiri("Citta")
                TXT_Cap.Text = Reader_Ritiri("CAP")
            Else
                TXT_Nome.Text = ""
                TXT_Codice_BP.Text = ""
                TXT_Via.Text = ""
                TXT_Tel.Text = ""
                TXT_Citta.Text = ""
                TXT_Cap.Text = ""
            End If

            Cnn_Ritiri.Close()
        End If
    End Sub


    Private Sub TXT_Esecutore_LostFocus(sender As Object, e As EventArgs) Handles TXT_Esecutore.LostFocus
        Dim i As Integer
        Dim Esito_Ricerca As Integer
        Esito_Ricerca = 0
        For i = 0 To Num_Dipendenti - 1 Step 1
            If Nomi_Dipendenti(i).ToLower Like (TXT_Esecutore.Text.ToLower & "*") Then
                TXT_Esecutore.Text = Nomi_Dipendenti(i)
                Esito_Ricerca = 1
            End If
        Next
        If Esito_Ricerca = 0 Then
            TXT_Esecutore.Text = ""
        End If
    End Sub

    Private Sub Compila_Elenco_Dipendenti()
        Dim Indice As Integer

        Dim Cnn_Dipendenti As New SqlConnection

        Cnn_Dipendenti.ConnectionString = Homepage.sap_tirelli
        Cnn_Dipendenti.Open()

        Dim Cmd_Dipendenti As New SqlCommand
        Dim Reader_Dipendenti As SqlDataReader

        Cmd_Dipendenti.Connection = Cnn_Dipendenti
        Cmd_Dipendenti.CommandText = "select * FROM [TIRELLI_40].[dbo].OHEM WHERE active='Y' ORDER BY lastName, firstName"

        Reader_Dipendenti = Cmd_Dipendenti.ExecuteReader()
        Indice = 0

        Do While Reader_Dipendenti.Read()
            Elenco_Dipendenti(Indice) = Reader_Dipendenti("empID")
            Nomi_Dipendenti(Indice) = (Reader_Dipendenti("lastName") & " " & Reader_Dipendenti("firstName"))
            Indice = Indice + 1
        Loop
        Num_Dipendenti = Indice
        Cnn_Dipendenti.Close()
    End Sub


    Private Sub TXT_Ora_LostFocus(sender As Object, e As EventArgs) Handles TXT_Ora.LostFocus
        If Val(TXT_Ora.Text) < 0 Or Val(TXT_Ora.Text) > 23 Then
            TXT_Ora.Text = "08"
            MsgBox("L'Ora deve essere compresa tra 0 e 23")
        Else
            TXT_Ora.Text = TXT_Ora.Text.ToString.PadLeft(2, "0")
        End If
    End Sub

    Private Sub TXT_Minuti_LostFocus(sender As Object, e As EventArgs) Handles TXT_Minuti.LostFocus
        If Val(TXT_Minuti.Text) < 0 Or Val(TXT_Minuti.Text) > 59 Then
            TXT_Minuti.Text = "00"
            MsgBox("I minuti devono essere compresi tra 0 e 59")
        Else
            TXT_Minuti.Text = TXT_Minuti.Text.ToString.PadLeft(2, "0")
        End If
    End Sub

    Private Sub Cmd_Exit_Click(sender As Object, e As EventArgs) Handles Cmd_Exit.Click
        Me.Close()
    End Sub

    Private Sub Cmd_Inserisci_Click(sender As Object, e As EventArgs) Handles Cmd_Inserisci.Click
        If ComboBox1.SelectedIndex < 0 Then
            MsgBox("Selezionare un autorizzatore")
        Else
            Dim giorno_periodico As Integer = 0
            Dim periodico As Integer = 0
            If CheckBox1.Checked = True Then
                periodico = 1
                giorno_periodico = ComboBox2.Text
            End If
            If Cmd_Inserisci.Text = "Inserisci" Then
                If TXT_Nome.Text.Length = 0 Then
                    MsgBox("Inserire il Business Partner")
                Else

                    Dim dataBase As Date = Data_Ora.Value.Date ' Ottiene solo la data senza ora
                    Dim ore As Integer = Integer.Parse(TXT_Ora.Text) ' Converte il testo in numero intero per l'ora
                    Dim minuti As Integer = Integer.Parse(TXT_Minuti.Text) ' Converte il testo in numero intero per i minuti

                    Dim dataCompleta As DateTime = New DateTime(dataBase.Year, dataBase.Month, dataBase.Day, ore, minuti, 0)

                    Using Cnn_Ritiri As New SqlConnection(Homepage.sap_tirelli)
                        Cnn_Ritiri.Open()

                        Dim Cmd_Ritiri As New SqlCommand("INSERT INTO [TIRELLI_40].[DBO].COLL_Ritiri_Materiale 
                    (Indice, Data_, BP, Tipologia, Note, Eseguito, Periodico,giorno_periodico, Esecutore, Autorizzatore) 
                    VALUES (@Indice, @Data_, @BP, @Tipologia, @Note, 0, @periodico,@giorno_periodico, ''," & Elenco_owner(ComboBox1.SelectedIndex) & ")", Cnn_Ritiri)

                        ' Aggiunta dei parametri
                        Cmd_Ritiri.Parameters.AddWithValue("@Indice", Get_Max_Index())
                        Cmd_Ritiri.Parameters.AddWithValue("@Data_", dataCompleta)
                        Cmd_Ritiri.Parameters.AddWithValue("@BP", TXT_Codice_BP.Text)
                        Cmd_Ritiri.Parameters.AddWithValue("@Tipologia", CB_Tipo.Text)
                        Cmd_Ritiri.Parameters.AddWithValue("@Note", TXT_Note.Text)
                        Cmd_Ritiri.Parameters.AddWithValue("@periodico", periodico)
                        Cmd_Ritiri.Parameters.AddWithValue("@giorno_periodico", giorno_periodico)

                        Cmd_Ritiri.ExecuteNonQuery()
                    End Using

                    Pulisci_Form()
                End If
            End If

            If Cmd_Inserisci.Text = "Nuovo" Then
                Pulisci_Form()
            End If
        End If
    End Sub

    Public Sub Pulisci_Form()
        Cmd_Inserisci.Text = "Inserisci"
        CB_Tipo.SelectedIndex = 0
        TXT_Ora.Text = "08"
        TXT_Minuti.Text = "00"
        TXT_Codice_BP.Text = ""
        TXT_Nome.Text = ""
        TXT_Note.Text = ""
        TXT_Esecutore.Text = ""
        TXT_Via.Text = ""
        TXT_Citta.Text = ""
        TXT_Cap.Text = ""
        TXT_Tel.Text = ""
        Aggiorna_Grid_Ritiri(DG_Ritiri)
        Cmd_Eseguito.Enabled = False
        TXT_Esecutore.Enabled = False
        Cmd_Annulla.Enabled = False
        Cmd_Aggiorna.Enabled = False
        TXT_Esecutore.Enabled = False
        TXT_Codice_BP.Enabled = True
        TXT_Nome.Enabled = True
        Data_Ora.Enabled = True
        TXT_Ora.Enabled = True
        TXT_Minuti.Enabled = True
        TXT_Note.Enabled = True
        CB_Tipo.Enabled = True
        Cmd_Inserisci.Enabled = True
        Data_Ora.Value = Today
        Aggiorna_Grid_Ritiri(DG_Ritiri)
    End Sub
    Public Sub Aggiorna_Grid_Ritiri(par_datagridview As DataGridView)
        par_datagridview.Rows.Clear()

        Dim Val_Data As Long
        Dim Val_Data_Limite As Long = 365 * Today.Year + 31 * Today.Month + Today.Day - 7

        Using Cnn_Ritiri As New SqlConnection(Homepage.sap_tirelli)
            Cnn_Ritiri.Open()

            Dim query As String = "SELECT T2.Indice, T2.Data_, T2.BP, T2.Tipologia, T2.Note, t3.lastname as 'Autorizzatore'
, case when T2.Eseguito ='-1' then 'STORNATO' WHEN T2.Eseguito ='0' THEN 'APERTO' WHEN T2.Eseguito ='1' THEN 'ESEGUITO' end as 'Eseguito' , " &
                              "T2.Periodico,t2.giorno_periodico, T2.Esecutore, T0.CardCode AS 'Codice', T0.CardName AS 'Nome', " &
                              "ISNULL(T0.Phone1, ' ') AS 'Tel', ISNULL(T1.Street, ' ') AS 'Via', " &
                              "ISNULL(T1.City, ' ') AS 'Citta', ISNULL(T1.ZipCode, ' ') AS 'CAP' " &
                              "FROM TIRELLISRLDB.DBO.OCRD T0 WITH (NOLOCK) " &
                              "INNER JOIN TIRELLISRLDB.DBO.CRD1 T1 WITH (NOLOCK) ON T0.CardCode = T1.CardCode " &
                              "INNER JOIN [TIRELLI_40].[DBO].COLL_Ritiri_Materiale T2 WITH (NOLOCK) ON T0.CardCode = T2.BP 
                              left join [TIRELLI_40].[dbo].ohem t3 on t3.empid=t2.autorizzatore " &
                              "WHERE (T1.LineNum IS NULL OR T1.LineNum < 1 ) and (t2.data_>=getdate()-15 or (T2.Periodico=1 and T2.Eseguito <>'-1')) " &
                              "ORDER BY T2.Data_ DESC"

            Using Cmd_Ritiri As New SqlCommand(query, Cnn_Ritiri)
                Using Reader_Ritiri As SqlDataReader = Cmd_Ritiri.ExecuteReader()

                    While Reader_Ritiri.Read()
                        Dim Stringa_BP As String
                        If Reader_Ritiri("Periodico") = 0 Then
                            Stringa_BP = $"{Reader_Ritiri("Nome")} {Reader_Ritiri("Via")} {Reader_Ritiri("CAP")} {Reader_Ritiri("Citta")} {Reader_Ritiri("Tel")}"

                            If Not IsDBNull(Reader_Ritiri("Data_")) Then
                                Dim Data_ As DateTime = Reader_Ritiri("Data_")
                                Dim Val_Data_Local As Integer = 365 * Data_.Year + 31 * Data_.Month + Data_.Day ' Nuovo nome
                                Dim Stringa_Data As String = Data_.ToString("dd/MM/yyyy - HH:mm")

                                Dim rowIndex As Integer = par_datagridview.Rows.Add(Reader_Ritiri("Indice"), Reader_Ritiri("Autorizzatore"), Stringa_Data, Stringa_BP, Reader_Ritiri("Tipologia"), Reader_Ritiri("Eseguito"), Reader_Ritiri("Note"), "Visualizza")
                                Dim row As DataGridViewRow = par_datagridview.Rows(rowIndex)
                                ' End If
                            End If
                        Else
                            Stringa_BP = $"{Reader_Ritiri("Nome")} "
                            Dim giorno_Sett As String
                            If Reader_Ritiri("Giorno_periodico") = 1 Then
                                giorno_Sett = "Lunedì"
                            ElseIf Reader_Ritiri("Giorno_periodico") = 2 Then
                                giorno_Sett = "Martedì"
                            ElseIf Reader_Ritiri("Giorno_periodico") = 3 Then
                                giorno_Sett = "Mercoledì"
                            ElseIf Reader_Ritiri("Giorno_periodico") = 4 Then
                                giorno_Sett = "Giovedì"
                            ElseIf Reader_Ritiri("Giorno_periodico") = 5 Then
                                giorno_Sett = "Venerdì"
                            ElseIf Reader_Ritiri("Giorno_periodico") = 6 Then
                                giorno_Sett = "Sabato"
                            ElseIf Reader_Ritiri("Giorno_periodico") = 7 Then
                                giorno_Sett = "Domenica"
                            End If
                            DataGridView1.Rows.Add(Reader_Ritiri("Indice"), Reader_Ritiri("Autorizzatore"), giorno_Sett, Stringa_BP, Reader_Ritiri("Tipologia"), Reader_Ritiri("Note"))

                        End If
                    End While
                End Using
            End Using
        End Using
    End Sub




    Public Function Get_Max_Index() As Integer
        Dim Risultato As Integer
        Dim Cnn_Ritiri As New SqlConnection
        Cnn_Ritiri.ConnectionString = Homepage.sap_tirelli
        Cnn_Ritiri.Open()
        Dim Cmd_Ritiri As New SqlCommand
        Dim Reader_Ritiri As SqlDataReader
        Cmd_Ritiri.Connection = Cnn_Ritiri
        Cmd_Ritiri.CommandText = "SELECT case when MAX(Indice) is null then '0' else  MAX(Indice) end as 'Massimo' 
FROM [TIRELLI_40].[DBO].COLL_Ritiri_Materiale"

        Reader_Ritiri = Cmd_Ritiri.ExecuteReader()

        If Reader_Ritiri.Read() Then
            Risultato = Reader_Ritiri("Massimo") + 1
        Else
            Risultato = 1
        End If
        Cnn_Ritiri.Close()
        Return Risultato
    End Function


    Private Sub DG_Ritiri_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG_Ritiri.CellClick

        Dim par_datagridview As DataGridView = DG_Ritiri
        If e.RowIndex >= 0 Then
            Sub_Visualizza(DG_Ritiri.Rows(e.RowIndex).Cells("Indice").Value)

        End If


    End Sub

    Private Sub Sub_Visualizza(Indice As Integer)
        Indice_Modifica = Indice
        Dim Cnn_Ritiri As New SqlConnection
        Cnn_Ritiri.ConnectionString = Homepage.sap_tirelli
        Cnn_Ritiri.Open()
        Dim Cmd_Ritiri As New SqlCommand
        Dim Reader_Ritiri As SqlDataReader
        Cmd_Ritiri.Connection = Cnn_Ritiri
        Cmd_Ritiri.CommandText = "SELECT T2.Indice, T2.Data_, T2.BP, T2.Tipologia, T2.Note, T2.Eseguito, T2.Periodico, T2.Esecutore, T1.[LineNum], T0.[CardCode] as 'Codice', T0.[CardName] as 'Nome', " &
                             "CASE WHEN T0.[Phone1] IS NULL THEN ' ' ELSE T0.[Phone1] END AS 'Tel', " &
                             "CASE WHEN T1.[Street] IS NULL THEN ' ' ELSE T1.[Street] END AS 'Via', " &
                             "CASE WHEN T1.[City] IS NULL THEN ' ' ELSE T1.[City] END AS 'Citta', " &
                             "CASE WHEN T1.[ZipCode] IS NULL THEN ' ' ELSE T1.[ZipCode] END AS 'CAP' " &
                             "FROM TIRELLISRLDB.DBO.OCRD T0 " &
                             "INNER JOIN TIRELLISRLDB.DBO.CRD1 T1 ON T0.[CardCode] = T1.[CardCode], [TIRELLI_40].[DBO].COLL_Ritiri_Materiale T2 " &
                             "WHERE T0.[CardCode] = T2.BP AND (T1.LineNum IS NULL OR T1.LineNum < 1) AND T2.Indice = " & Indice_Modifica & " " &
                             "ORDER BY T2.Data_ desc"

        Reader_Ritiri = Cmd_Ritiri.ExecuteReader()

        If Reader_Ritiri.Read() Then
            ' Parsing della data nel formato "dd/MM/yyyy - HH:mm"
            Dim dataRitiro As DateTime
            If DateTime.TryParseExact(Reader_Ritiri("Data_").ToString(), "yyyyMMddHHmm", Nothing, Globalization.DateTimeStyles.None, dataRitiro) Then
                Data_Ora.Value = dataRitiro
                TXT_Ora.Text = dataRitiro.ToString("HH")
                TXT_Minuti.Text = dataRitiro.ToString("mm")
            End If

            TXT_Nome.Text = Reader_Ritiri("Nome")
            TXT_Codice_BP.Text = Reader_Ritiri("Codice")
            TXT_Via.Text = Reader_Ritiri("Via")
            TXT_Tel.Text = Reader_Ritiri("Tel")
            TXT_Citta.Text = Reader_Ritiri("Citta")
            TXT_Cap.Text = Reader_Ritiri("CAP")
            TXT_Note.Text = Reader_Ritiri("Note")
            CB_Tipo.Text = Reader_Ritiri("Tipologia")

            ' Gestione dell'abilitazione dei controlli in base al valore di Eseguito
            If Reader_Ritiri("Eseguito") = -1 Then
                Cmd_Eseguito.Enabled = False
                Cmd_Annulla.Enabled = False
                Cmd_Aggiorna.Enabled = False
                TXT_Esecutore.Text = ""
                TXT_Esecutore.Enabled = False
                TXT_Codice_BP.Enabled = False
                TXT_Nome.Enabled = False
                Data_Ora.Enabled = False
                TXT_Ora.Enabled = False
                TXT_Minuti.Enabled = False
                TXT_Note.Enabled = False
                CB_Tipo.Enabled = False
                Cmd_Inserisci.Enabled = True
                Cmd_Inserisci.Text = "Nuovo"
            End If

            If Reader_Ritiri("Eseguito") = 1 Then
                Cmd_Eseguito.Enabled = False
                Cmd_Annulla.Enabled = False
                Cmd_Aggiorna.Enabled = False
                TXT_Esecutore.Text = Reader_Ritiri("Esecutore")
                TXT_Esecutore.Enabled = False
                TXT_Codice_BP.Enabled = False
                TXT_Nome.Enabled = False
                Data_Ora.Enabled = False
                TXT_Ora.Enabled = False
                TXT_Minuti.Enabled = False
                TXT_Note.Enabled = False
                CB_Tipo.Enabled = False
                Cmd_Inserisci.Enabled = True
                Cmd_Inserisci.Text = "Nuovo"
            End If

            If Reader_Ritiri("Eseguito") = 0 Then
                Cmd_Eseguito.Enabled = True
                Cmd_Annulla.Enabled = True
                Cmd_Aggiorna.Enabled = True
                TXT_Esecutore.Enabled = True
                TXT_Esecutore.Text = ""
                TXT_Codice_BP.Enabled = True
                TXT_Nome.Enabled = True
                Data_Ora.Enabled = True
                TXT_Ora.Enabled = True
                TXT_Minuti.Enabled = True
                TXT_Note.Enabled = True
                CB_Tipo.Enabled = True
                Cmd_Inserisci.Enabled = True
                Cmd_Inserisci.Text = "Nuovo"
            End If
        End If

        Cnn_Ritiri.Close()
    End Sub

    Private Sub Cmd_Annulla_Click(sender As Object, e As EventArgs) Handles Cmd_Annulla.Click
        Using Cnn_Ritiri As New SqlConnection(Homepage.sap_tirelli)
            Cnn_Ritiri.Open()

            Dim Cmd_Ritiri As New SqlCommand("UPDATE [TIRELLI_40].[DBO].COLL_Ritiri_Materiale 
            SET Eseguito = -1 WHERE Indice = @Indice", Cnn_Ritiri)

            ' Aggiunta del parametro
            Cmd_Ritiri.Parameters.AddWithValue("@Indice", Indice_Modifica)

            Cmd_Ritiri.ExecuteNonQuery()
        End Using

        Pulisci_Form()
    End Sub

    Private Sub Cmd_Aggiorna_Click(sender As Object, e As EventArgs) Handles Cmd_Aggiorna.Click
        If TXT_Nome.Text.Length = 0 Then
            MsgBox("Inserire il Business Partner")
        Else
            Dim giorno_periodico As Integer = 0
            Dim periodico As Integer = 0
            If CheckBox1.Checked = True Then
                periodico = 1
                giorno_periodico = ComboBox2.Text
            End If
            Dim dataBase As Date = Data_Ora.Value.Date ' Ottiene solo la data senza ora
            Dim ore As Integer = Integer.Parse(TXT_Ora.Text) ' Converte il testo in numero intero per l'ora
            Dim minuti As Integer = Integer.Parse(TXT_Minuti.Text) ' Converte il testo in numero intero per i minuti

            Dim dataCompleta As DateTime = New DateTime(dataBase.Year, dataBase.Month, dataBase.Day, ore, minuti, 0)


            Using Cnn_Ritiri As New SqlConnection(Homepage.sap_tirelli)
                Cnn_Ritiri.Open()

                Dim Cmd_Ritiri As New SqlCommand("UPDATE [TIRELLI_40].[DBO].COLL_Ritiri_Materiale 
                SET Data_ = @Data_, BP = @BP, Tipologia = @Tipologia, Note = @Note , Periodico=@periodico, giorno_periodico=@giorno_periodico
                WHERE Indice = @Indice", Cnn_Ritiri)

                ' Aggiunta dei parametri
                Cmd_Ritiri.Parameters.AddWithValue("@Data_", dataCompleta)
                Cmd_Ritiri.Parameters.AddWithValue("@BP", TXT_Codice_BP.Text)
                Cmd_Ritiri.Parameters.AddWithValue("@Tipologia", CB_Tipo.Text)
                Cmd_Ritiri.Parameters.AddWithValue("@Note", TXT_Note.Text)
                Cmd_Ritiri.Parameters.AddWithValue("@Indice", Indice_Modifica)
                Cmd_Ritiri.Parameters.AddWithValue("@periodico", periodico)
                Cmd_Ritiri.Parameters.AddWithValue("@giorno_periodico", giorno_periodico)

                Cmd_Ritiri.ExecuteNonQuery()
            End Using

            Aggiorna_Grid_Ritiri(DG_Ritiri)
            Pulisci_Form()
        End If
    End Sub

    Private Sub Cmd_Eseguito_Click(sender As Object, e As EventArgs) Handles Cmd_Eseguito.Click
        If TXT_Esecutore.Text.Length = 0 Then
            MsgBox("Inserire il nome")
        Else
            Using Cnn_Ritiri As New SqlConnection(Homepage.sap_tirelli)
                Cnn_Ritiri.Open()

                Dim Cmd_Ritiri As New SqlCommand("UPDATE [TIRELLI_40].[DBO].COLL_Ritiri_Materiale 
                SET Esecutore = @Esecutore, Eseguito = 1 
                WHERE Indice = @Indice", Cnn_Ritiri)

                ' Aggiunta dei parametri
                Cmd_Ritiri.Parameters.AddWithValue("@Esecutore", TXT_Esecutore.Text)
                Cmd_Ritiri.Parameters.AddWithValue("@Indice", Indice_Modifica)

                Cmd_Ritiri.ExecuteNonQuery()
            End Using

            Aggiorna_Grid_Ritiri(DG_Ritiri)
            Pulisci_Form()
        End If
    End Sub

    Private Sub TXT_Codice_BP_TextChanged(sender As Object, e As EventArgs) Handles TXT_Codice_BP.TextChanged
        If TXT_Codice_BP.Text = "" Then
            filtro_bp = ""
        Else
            filtro_bp = " and t0.cardcode = '" & TXT_Codice_BP.Text & "'"

        End If

    End Sub

    Private Sub TXT_Nome_TextChanged(sender As Object, e As EventArgs) Handles TXT_Nome.TextChanged
        If TXT_Nome.Text = "" Then
            filtro_cardname = ""
        Else
            filtro_cardname = " and T0.[CardName] Like '%" & TXT_Nome.Text & "%'"

        End If

    End Sub

    Private Sub DG_Ritiri_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DG_Ritiri.CellContentClick

    End Sub

    Private Sub DG_Ritiri_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DG_Ritiri.CellFormatting
        Dim par_datagridview As DataGridView = DG_Ritiri
        Dim oggi As Date = Date.Today
        Dim domani As Date = oggi.AddDays(1)

        ' Controllo per evitare errori su righe non valide
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            ' Gestione del colore per la colonna "Stato"
            Dim cellaStato As DataGridViewCell = par_datagridview.Rows(e.RowIndex).Cells("Stato")
            If cellaStato.Value IsNot Nothing Then
                Select Case cellaStato.Value.ToString().ToUpper()
                    Case "ESEGUITO"
                        cellaStato.Style.BackColor = Color.Lime
                    Case "STORNATO"
                        cellaStato.Style.BackColor = Color.Red
                    Case "APERTO"
                        cellaStato.Style.BackColor = Color.LightYellow
                End Select
            End If

            ' Gestione del colore per la colonna "Data"
            Dim cellaData As DataGridViewCell = par_datagridview.Rows(e.RowIndex).Cells("Data")
            If cellaData.Value IsNot Nothing Then
                Dim dataRitiro As DateTime
                ' Prova a parsare la data nel formato "dd/MM/yyyy - HH:mm"
                If DateTime.TryParseExact(cellaData.Value.ToString(), "dd/MM/yyyy - HH:mm", Nothing, Globalization.DateTimeStyles.None, dataRitiro) Then
                    If dataRitiro.Date = oggi Then
                        par_datagridview.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Orange
                    ElseIf dataRitiro.Date = domani Then
                        par_datagridview.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.Yellow
                    End If
                End If
            End If
        End If

    End Sub

    Private Sub Data_Ora_ValueChanged(sender As Object, e As EventArgs) Handles Data_Ora.ValueChanged

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick


    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Dim par_datagridview As DataGridView = DataGridView1
        If e.RowIndex >= 0 Then
            Sub_Visualizza(par_datagridview.Rows(e.RowIndex).Cells("Indice_").Value)

        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        ' Verifica se la colonna è quella dei giorni
        If e.ColumnIndex = DataGridView1.Columns("Giorno").Index Then
            ' Ottieni il giorno della settimana attuale
            Dim giornoCorrente As String = DateTime.Now.ToString("dddd")

            ' Confronta il valore della cella con il giorno corrente
            If e.Value IsNot Nothing AndAlso e.Value.ToString() = giornoCorrente Then
                e.CellStyle.BackColor = Color.Orange ' Imposta il colore arancione
            End If
        End If
    End Sub
End Class