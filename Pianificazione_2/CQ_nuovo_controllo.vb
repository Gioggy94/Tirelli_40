Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib


Public Class CQ_nuovo_controllo

    Public Elenco_dipendenti(1000) As String
    Public Elenco_esito_controllo(1000) As String
    Public codicedip As Integer
    Public esito_controllo As String
    Public codice_sap As String
    Public id As Integer
    Public test_odp As Integer = 0

    Public pezzi_da_controllare As Integer = 0
    Public pezzi_NC As Integer = 0
    Public pezzi_OK As Integer = 0

    Sub Inserimento_dipendenti()
        ComboBox2.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT coalesce(T0.[USERID],'') as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome'
        FROM [TIRELLI_40].[dbo].OHEM T0 left join [TIRELLI_40].[dbo].oudp t1 on T0.[dept]=t1.code where t0.active='Y' and t0.dept=19 order by T0.[lastName]"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_dipendenti(Indice) = cmd_SAP_reader("Codice dipendenti")
            ComboBox2.Items.Add(cmd_SAP_reader("Nome"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()

        ComboBox2.Text = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).cognome & " " & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).nome


    End Sub 'Inserisco le risorse nella combo box


    Sub Inserimento_Esito_controllo()
        ComboBox7.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT fldvalue,descr
FROM [TIRELLISRLDB].[dbo].UFD1
WHERE tableid='oclg' and fieldid=16
order by indexid"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Dim Indice As Integer
        Indice = 0
        Do While cmd_SAP_reader.Read()
            Elenco_esito_controllo(Indice) = cmd_SAP_reader("fldvalue")
            ComboBox7.Items.Add(cmd_SAP_reader("descr"))
            Indice = Indice + 1
        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub 'Inserisco le risorse nella combo box

    Sub Inserimento_imputazione()
        ComboBox3.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT T0.imputazione 
FROM [Tirelli_40].[dbo].cq_imputazioni 
t0 group by T0.imputazione,t0.categoria
order by t0.categoria"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            ComboBox3.Items.Add(cmd_SAP_reader("imputazione"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()


    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        codice_sap = TextBox5.Text


        If Len(TextBox5.Text) >= 6 Then
            Label10.Text = Magazzino.OttieniDettagliAnagrafica(codice_sap).Descrizione
            TextBox6.Text = Magazzino.OttieniDettagliAnagrafica(codice_sap).Disegno
        End If

    End Sub

    Sub trova_ID()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "SELECT  max(case when t0.id is null then 0 else t0.id end )+1 as 'ID' 
from [Tirelli_40].[dbo].cq_nuovo_controllo t0"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True Then

            If Not cmd_SAP_reader("ID") Is System.DBNull.Value Then
                id = cmd_SAP_reader("ID")
            Else
                id = 1
            End If
        Else
            id = 1
        End If

        cnn.Close()
        cmd_SAP_reader.Close()


    End Sub

    Public Function check_emp_nel_database(par_emp As String)
        Dim risultato As Boolean = False
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        CNN.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = CNN

        CMD_SAP.CommandText = "SELECT  t0.emp
        from [Tirelli_40].[dbo].cq_nuovo_controllo t0 where
        t0.emp='" & par_emp & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True And par_emp <> Nothing Then
            risultato = True
        Else
            risultato = False

        End If

        CNN.Close()
        cmd_SAP_reader.Close()
        Return risultato

    End Function

    Sub check_oa_nel_database()

        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "SELECT  t0.oa
        from [Tirelli_40].[dbo].cq_nuovo_controllo t0 where
        t0.oa='" & TextBox9.Text & "' and t0.codice = '" & TextBox5.Text & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True And TextBox9.Text <> Nothing Then
            test_odp = test_odp + 1


        End If

        cnn.Close()
        cmd_SAP_reader.Close()


    End Sub


    Sub check_odp_nelle_attività()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "SELECT t2.baseref
FROM [TIRELLISRLDB].[dbo].OCLG T0 
left join [TIRELLISRLDB].[dbo].oign t1 on t1.docnum= t0.docnum
left join [TIRELLISRLDB].[dbo].ign1 t2 on t1.docentry=t2.docentry and t2.itemcode=t0.u_prg_qlt_itemcode

WHERE T0.[U_Stato] ='O' and t0.cntcttype=10 and t0.doctype=59 and t2.baseref='" & TextBox1.Text & "' ORDER BY T0.[ClgCode] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True And TextBox1.Text <> Nothing Then
            test_odp = test_odp + 1

        End If

        cnn.Close()
        cmd_SAP_reader.Close()


    End Sub

    Sub check_oa_nelle_attività()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader
        CMD_SAP.Connection = cnn

        CMD_SAP.CommandText = "SELECT t2.baseref
FROM [TIRELLISRLDB].[dbo].OCLG T0 
left join [TIRELLISRLDB].[dbo].opdn t1 on t1.docnum= t0.docnum
inner join [TIRELLISRLDB].[dbo].pdn1 t2 on t2.docentry=t1.docentry and t2.itemcode= t0.u_prg_qlt_itemcode

WHERE T0.[U_Stato] ='O' and t0.cntcttype=10 and t0.doctype=20 and t2.baseref='" & TextBox9.Text & "' ORDER BY T0.[ClgCode] "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() = True And TextBox9.Text <> Nothing Then
            test_odp = test_odp + 1

        End If

        cnn.Close()
        cmd_SAP_reader.Close()


    End Sub



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Controllo di corretta compilazione  
        If Label10.Text = "" Then

            MsgBox("Scegliere un codice valido")

        Else

            If TextBox1.Text = "" And TextBox9.Text = "" Then
                MsgBox("Imputare controllo ad un'entrata merce da fornitore o produzione")

            Else


                If pezzi_OK + pezzi_NC = pezzi_da_controllare Then

                    If pezzi_NC > 0 Then



                        If pezzi_NC = Nothing Or ComboBox3.Text = Nothing Or ComboBox4.Text = Nothing Or ComboBox1.Text = Nothing Or ComboBox2.Text = Nothing Or ComboBox6.Text = Nothing Or ComboBox7.Text = Nothing Or RichTextBox1.Text = Nothing Or RichTextBox2.Text = Nothing Then

                            'BISOGNA FARE LE VARIANTI A SECONDA SE SIA CONFORME O MENO

                            MsgBox("il controllo ha stabilito che ci sono delle non conformità ma mancano uno o più dati ")


                        Else
                            ' If ComboBox3.Text = "MU" And TextBox7.Text = Nothing And ComboBox7.Text = "Non_conforme" Then
                            'MsgBox("Non è possibile dare NC a MU se non si sceglie l'autocontrollo")
                            ' Else

                            If check_emp_nel_database(TextBox1.Text) = True Then
                                    MsgBox("L'entrata merce da produzione esiste già")
                                    Return
                                End If
                                check_oa_nel_database()

                                check_odp_nelle_attività()
                                check_oa_nelle_attività()

                                If test_odp > 0 Then
                                    MsgBox("Risulta già un record relativo a questo ordine di produzione o risulta un'attività aperta")

                                Else
                                    Label1.Text = Label1.Text
                                    trova_ID()
                                    inserisci_record()
                                    pulizia_form()

                                ' End If
                            End If


                        End If
                    Else
                        If ComboBox3.Text = Nothing Or ComboBox1.Text = Nothing Or ComboBox2.Text = Nothing Or ComboBox7.Text = Nothing Then
                            MsgBox("il controllo ha stabilito che ci sono delle non conformità ma mancano uno o più dati ")


                        Else
                            If ComboBox3.Text = "MU" And TextBox7.Text = Nothing And ComboBox7.Text = "Non_conforme" Then
                                MsgBox("Non è possibile dare NC a MU se non si sceglie l'autocontrollo")
                            Else
                                If check_emp_nel_database(TextBox1.Text) = True Then
                                    MsgBox("L'entrata merce è già registrata")
                                End If
                                check_oa_nel_database()
                                    check_odp_nelle_attività()
                                    check_oa_nelle_attività()

                                    If test_odp > 0 Then
                                        MsgBox("Risulta già un record relativo a questo ordine di produzione o risulta un'attività aperta")

                                    Else
                                        Label1.Text = Label1.Text
                                        trova_ID()
                                        inserisci_record()
                                        pulizia_form()

                                    End If
                                End If

                            End If


                    End If
                Else

                    MsgBox("Controllare che la quantità controllata sia = alla quantità OK + Quantità NC")

                End If

            End If


        End If
    End Sub

    Sub inserisci_record()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand

        CMD_SAP_7.Connection = CNN
        CMD_SAP_7.CommandText = "insert into 
[Tirelli_40].[dbo].cq_nuovo_controllo
(ID, CODICE, DISEGNO, PZ_CONTR, PZ_NC, PZ_OK, IMPUTAZIONE, CAMPO_DEFINIZIONE_NC,DESCRIZIONE_NC, ESITO_AUTOCONTROLLO, PESO_NC,OSSERVAZIONI_NC, EMP, EMF,autocontrollo, OPERATORE, ZONA_CONTROLLO, DATA, ORA,stato,rilevato,erp)
VALUES ('" & id & "','" & TextBox5.Text & "','" & TextBox6.Text & "','" & pezzi_da_controllare & "','" & pezzi_NC & "','" & pezzi_OK & "', '" & ComboBox3.Text & "', '" & ComboBox4.Text & "','" & ComboBox5.Text & "' , '" & ComboBox7.Text & "', '" & ComboBox6.Text & "','" & RichTextBox1.Text & "','" & TextBox1.Text & "','" & TextBox9.Text & "','" & TextBox7.Text & "'," & codicedip & ",'" & ComboBox1.Text & "',GETDATE(),convert(varchar, getdate(), 108),'In_corso','" & RichTextBox2.Text & "', '" & ComboBox9.Text & "')  "




        CMD_SAP_7.ExecuteNonQuery()
        cnn.Close()
        MsgBox("Registazione effettuata")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Me.Close()

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        pezzi_NC = TextBox3.Text

        'Attivazione Campi NC
        If pezzi_NC = "0" Then

            ComboBox4.Enabled = False
            ComboBox5.Enabled = False
            ComboBox6.Enabled = False
            RichTextBox1.Enabled = False
            ComboBox4.Text = Nothing
            ComboBox5.Text = Nothing
            ComboBox6.Text = Nothing
            ComboBox7.Text = ""
            RichTextBox1.Text = Nothing

            'Impostazione materiale conforme
            pezzi_OK = pezzi_da_controllare
        Else
            If Int(pezzi_NC) > Int(pezzi_da_controllare) Then
                MsgBox("QuantiTà non conforme superiore alla controllata")
                pezzi_NC = Nothing
            Else
                ComboBox4.Enabled = True
                ComboBox5.Enabled = True
                ComboBox6.Enabled = True
                RichTextBox1.Enabled = True
                RichTextBox2.Enabled = True

                'Sottrazione materiale conforme
                pezzi_OK = Int(pezzi_da_controllare) - Int(pezzi_NC)
            End If
        End If

        TextBox4.Text = pezzi_OK
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If File.Exists(Homepage.percorso_disegni_generico & "PDF\"  & TextBox6.Text & ".PDF") Then
            Process.Start(Homepage.percorso_disegni_generico & "PDF\"  & TextBox6.Text & ".PDF")
        Else
            MsgBox("PDF non presente")
        End If
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If File.Exists(Homepage.percorso_disegni_generico & "PDF\"  & TextBox6.Text & ".PDF") Then
            Button3.BackColor = Color.Lime
        Else
            Button3.BackColor = Color.Red
        End If
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.Text = "MU" Then
            GroupBox4.Visible = True
        Else
            GroupBox4.Visible = False
            TextBox7.Text = Nothing
        End If
        Inserimento_definizione()

    End Sub


    Sub Inserimento_definizione()
        ComboBox4.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()


        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = CNN
        CMD_SAP.CommandText = "SELECT T0.categoria
FROM [Tirelli_40].[dbo].cq_imputazioni t0 
where T0.imputazione = '" & ComboBox3.Text & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            ComboBox4.Items.Add(cmd_SAP_reader("categoria"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        CQ_Tabelle.TextBox2.Text = TextBox5.Text
        CQ_Tabelle.Show()
        CQ_Tabelle.Owner = Me
        CQ_Tabelle.riempi_autocontrollo()

    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        codicedip = Elenco_dipendenti(ComboBox2.SelectedIndex)
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs)
        CQ_Tabelle.Show()
        CQ_Tabelle.Owner = Me

    End Sub

    Sub Inserimento_Descrizione_NC()
        ComboBox5.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT descrizione
FROM [TIRELLI_40].[dbo].[CQ_descrizione] where categoria ='" & ComboBox4.Text & "' group by descrizione "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            ComboBox5.Items.Add(cmd_SAP_reader("descrizione"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub



    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        esito_controllo = Elenco_esito_controllo(ComboBox7.SelectedIndex)


        If ComboBox7.Text = "Concesso" Then
            GroupBox25.Visible = True
            GroupBox25.Text = "Chi ha Concesso"
        ElseIf ComboBox7.Text = "Deroga" Then
            GroupBox25.Visible = True
            GroupBox25.Text = "Chi ha Derogato?"
        Else
            GroupBox25.Visible = False
            ComboBox8.Text = ""
        End If


    End Sub

    Sub pulizia_form()
        ComboBox3.Text = Nothing
        ComboBox1.Text = Nothing

        ComboBox7.Text = ""
        TextBox2.Text = 0
        TextBox3.Text = 0
        TextBox4.Text = 0
        pezzi_da_controllare = 0
        pezzi_NC = 0
        pezzi_OK = 0
        ComboBox4.Text = Nothing
        RichTextBox1.Text = Nothing
        RichTextBox2.Text = Nothing
        RichTextBox2.Enabled = False
        ComboBox6.Text = Nothing
        TextBox5.Text = Nothing
        TextBox1.Text = Nothing
        TextBox9.Text = Nothing
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Inserimento_Descrizione_NC()
    End Sub



    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        pezzi_da_controllare = TextBox2.Text
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        pezzi_OK = TextBox4.Text
        pezzi_NC = pezzi_da_controllare - pezzi_OK

        TextBox3.Text = pezzi_NC

    End Sub
End Class