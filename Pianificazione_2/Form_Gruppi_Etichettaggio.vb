Imports System.Data.SqlClient

Public Class Form_Gruppi_Etichettaggio
    Public commessa As String
    Public id As Integer
    Public N As Integer
    Public stato_gruppo As String

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub
    Sub inizializza_form()
        Label1.Text = commessa
        If stato_gruppo = "Nuovo" Then
            Button1.Text = "Aggiungi"
            N = trova_numero_gruppo(commessa) + 1
            Label2.Text = N

        ElseIf stato_gruppo = "Visualizza" Then

            Label2.Text = N
            Button1.Text = "Aggiorna"
            dati_gruppo(id)




        End If
    End Sub
    Private Sub Form_Gruppi_Etichettaggio_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.BackColor = Homepage.colore_sfondo
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If stato_gruppo = "Nuovo" Then
            If ComboBox1.SelectedIndex < 0 Then
                MsgBox("Scegliere tecnologia")
            Else
                inserisci_nuovo_gruppo(commessa, N, ComboBox1.Text, ComboBox2.Text, ComboBox3.Text, ComboBox4.Text, ComboBox5.Text, ComboBox6.Text, RichTextBox1.Text, ComboBox7.Text, TextBox1.Text, ComboBox8.Text, ComboBox9.Text, ComboBox10.Text)
                Scheda_tecnica.trova_gruppi_etichettaggio(Scheda_tecnica.DataGridView2, commessa)
                MsgBox("Gruppo inserito con successo")
                dati_gruppo(id)
                Me.Close()
            End If

        ElseIf stato_gruppo = "Visualizza" Then

            If ComboBox1.SelectedIndex < 0 Then
                MsgBox("Scegliere tecnologia")

            Else

                aggiorna_gruppo(id, ComboBox1.Text, ComboBox2.Text, ComboBox3.Text, ComboBox4.Text, ComboBox5.Text, ComboBox6.Text, RichTextBox1.Text, ComboBox7.Text, TextBox1.Text, ComboBox8.Text, ComboBox9.Text, ComboBox10.Text)
                Scheda_tecnica.trova_gruppi_etichettaggio(Scheda_tecnica.DataGridView2, commessa)
                MsgBox("Gruppo aggiornato con successo")
                dati_gruppo(id)
                Me.Close()
            End If

        End If


    End Sub

    Function trova_numero_gruppo(par_commessa As String)
        Dim numero As Integer = 0


        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        Dim contatore As Integer = 1

        CMD_SAP.CommandText = "select coalesce(sum( case when t0.commessa <>'' then 1 else 0 end),0) as 'N'
  FROM [TIRELLI_40].[DBO].BRB_Gruppi_etichettaggio t0

 where t0.commessa='" & par_commessa & "'
"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() Then
            numero = cmd_SAP_reader("N")
        Else
            numero = 0

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()


        Return numero

    End Function 'Inserisco le risorse nella combo box

    Sub dati_gruppo(par_id As Integer)



        Dim Cnn As New SqlConnection
        Cnn.ConnectionString = Homepage.sap_tirelli
        Cnn.Open()

        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = Cnn
        Dim contatore As Integer = 1

        CMD_SAP.CommandText = "SELECT t0.[ID]
      ,t0.[N]
      ,t0.[Commessa]
      ,t0.[Tecnologia]
, t0.marca
  ,coalesce(t0.[altezza_supporti],0) as 'Altezza_supporti'
      ,coalesce(t0.[Sensore_etichette],'') as 'Sensore_etichette'
      ,coalesce(t0.[Clear_spam],0) as 'Clear_spam'
,coalesce(t0.modello,'') as 'Modello'
,coalesce(t0.note,'') as 'Note'
,coalesce(t0.[tipo_stazione],'') as 'Tipo_stazione'
      ,coalesce(t0.[Zero_Down_Time],'') as 'Zero_down_time'
      ,coalesce(t0.[Lunghezza_slitta],'') as 'Lunghezza_slitta'
,coalesce(t0.[altezza_terminale],0) as 'altezza_terminale'
      ,coalesce(t0.[Ownerid],0) as 'Ownerid'
      ,t0.[Updatedate]
,coalesce(concat(t1.lastname,' ',t1.firstname),'') as 'Ownername'
,coalesce(t0.comunicazione,'') as 'Comunicazione'
  FROM [TIRELLI_40].[DBO].BRB_Gruppi_etichettaggio t0
left join [TIRELLI_40].[dbo].ohem t1 on t0.ownerid=t1.empid
where id = " & par_id & ""

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        If cmd_SAP_reader.Read() Then

            ComboBox1.Text = cmd_SAP_reader("Tecnologia")
            ComboBox2.Text = cmd_SAP_reader("marca")
            ComboBox3.Text = cmd_SAP_reader("altezza_supporti")
            ComboBox4.Text = cmd_SAP_reader("Sensore_etichette")
            ComboBox5.Text = cmd_SAP_reader("Clear_spam")
            ComboBox6.Text = cmd_SAP_reader("Modello")
            ComboBox7.Text = cmd_SAP_reader("tipo_stazione")
            TextBox1.Text = cmd_SAP_reader("Zero_down_time")
            ComboBox8.Text = cmd_SAP_reader("Lunghezza_slitta")
            ComboBox9.Text = cmd_SAP_reader("altezza_terminale")
            Label3.Text = cmd_SAP_reader("Ownername")
            Label4.Text = cmd_SAP_reader("updatedate")
            RichTextBox1.Text = cmd_SAP_reader("note")
            ComboBox10.Text = cmd_SAP_reader("Comunicazione")

        End If
        cmd_SAP_reader.Close()
        Cnn.Close()




    End Sub 'Inserisco le risorse nella combo box

    Sub inserisci_nuovo_gruppo(par_commessa As String, par_N As Integer, par_tecnologia As String, par_marca As String, par_altezza_supporti As String, par_sensore_etichette As String, par_clear_spam As String, par_modello As String, par_note As String, par_tipo_stazione As String, par_zero_down_timpe As String, par_lunghezza_slitta As String, par_altezza_terminale As String, par_comunicazione As String)



        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "INSERT INTO [TIRELLI_40].[DBO].BRB_Gruppi_etichettaggio
           ([N]
      ,[Commessa]
      ,[Tecnologia]
      ,[Marca]
      ,[altezza_supporti]
      ,[Sensore_etichette]
      ,[Clear_spam]
,modello
,note
,[tipo_stazione]
      ,[Zero_Down_Time]
      ,[Lunghezza_slitta]
,[altezza_terminale]
      ,[Ownerid]
      ,[Updatedate]
,comunicazione)
     VALUES
           (" & par_N & "
           ,'" & par_commessa & "'
           ,'" & par_tecnologia & "'
,'" & par_marca & "'
,'" & par_altezza_supporti & "'
,'" & par_sensore_etichette & "'
,'" & par_clear_spam & "'
,'" & par_modello & "'
,'" & par_note & "'
,'" & par_tipo_stazione & "'
,'" & par_zero_down_timpe & "'
,'" & par_lunghezza_slitta & "'
,'" & par_altezza_terminale & "'
,'" & Homepage.ID_SALVATO & "'
,getdate()
,'" & par_comunicazione & "')"


        CMD_SAP_3.ExecuteNonQuery()

        Cnn3.Close()


    End Sub

    Sub aggiorna_gruppo(par_id As Integer, par_tecnologia As String, par_marca As String, par_altezza_supporti As String, par_sensore_etichette As String, par_clear_spam As String, par_modello As String, par_note As String, par_tipo_stazione As String, par_zero_down_time As String, par_lunghezza_slitta As String, par_altezza_terminale As String, par_comunicazione As String)



        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "update [TIRELLI_40].[DBO].BRB_Gruppi_etichettaggio
set [Tecnologia]='" & par_tecnologia & "'
,marca='" & par_marca & "'
 ,[altezza_supporti]='" & par_altezza_supporti & "'
 ,[Sensore_etichette]='" & par_sensore_etichette & "'
 ,[Clear_spam]='" & par_clear_spam & "'
,modello='" & par_modello & "'
,note='" & par_note & "'
,[tipo_stazione]='" & par_tipo_stazione & "'
      ,[Zero_Down_Time]='" & par_zero_down_time & "'
      ,[Lunghezza_slitta]='" & par_lunghezza_slitta & "'
      ,[altezza_terminale]='" & par_altezza_terminale & "'
      ,[Ownerid]='" & Homepage.ID_SALVATO & "'
      ,[Updatedate]=getdate()
      ,comunicazione='" & par_comunicazione & "'

where id = " & par_id & ""


        CMD_SAP_3.ExecuteNonQuery()

        Cnn3.Close()


    End Sub


    Sub aggiorna_numero_gruppi(par_codice_commessa As String)



        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "WITH NumberedRows AS (
    SELECT *, ROW_NUMBER() OVER (ORDER BY [N]) AS RowNum
    FROM [TIRELLI_40].[DBO].BRB_Gruppi_etichettaggio
    WHERE commessa = '" & par_codice_commessa & "'
)
UPDATE NumberedRows
SET [N] = RowNum"


        CMD_SAP_3.ExecuteNonQuery()

        Cnn3.Close()


    End Sub

    Sub elimina_gruppi(par_id As Integer)



        Dim Cnn3 As New SqlConnection
        Cnn3.ConnectionString = Homepage.sap_tirelli
        Cnn3.Open()

        Dim CMD_SAP_3 As New SqlCommand

        CMD_SAP_3.Connection = Cnn3


        CMD_SAP_3.CommandText = "delete [TIRELLI_40].[DBO].BRB_Gruppi_etichettaggio where id = " & par_id & ""


        CMD_SAP_3.ExecuteNonQuery()

        Cnn3.Close()


    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        elimina_gruppi(id)
        aggiorna_numero_gruppi(commessa)

        Scheda_tecnica.trova_gruppi_etichettaggio(Scheda_tecnica.DataGridView2, commessa)
        MsgBox("Gruppo eliminato con successo ")
        Me.Close()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.Text = "Herma" Then
            GroupBox6.Visible = True
        Else
            GroupBox6.Visible = False
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Process.Start("\\tirfs01\tirelli\00-BRB\Herma\Contratto quadro Tirelli Herma.xlsx")
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        Dim par_combobox As ComboBox = ComboBox7
        If par_combobox.SelectedIndex = 1 Then
            Scheda_tecnica.assegna_foto(PictureBox2, Form_Codici_vendita.trova_immagine_codice("Y00147").ToString)
        ElseIf par_combobox.SelectedIndex = 2 Then
            Scheda_tecnica.assegna_foto(PictureBox2, Form_Codici_vendita.trova_immagine_codice("Y00148").ToString)
        ElseIf par_combobox.SelectedIndex = 3 Then
            Scheda_tecnica.assegna_foto(PictureBox2, Form_Codici_vendita.trova_immagine_codice("Y00148").ToString)
        Else PictureBox2.Image = Nothing
        End If
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        Dim par_combobox As ComboBox = ComboBox6
        Dim par_picturebox As PictureBox = PictureBox1
        If par_combobox.SelectedIndex = 1 Then
            Scheda_tecnica.assegna_foto(par_picturebox, Form_Codici_vendita.trova_immagine_codice("Y00051").ToString)
        ElseIf par_combobox.SelectedIndex = 2 Then
            Scheda_tecnica.assegna_foto(par_picturebox, Form_Codici_vendita.trova_immagine_codice("Y00051").ToString)
        ElseIf par_combobox.SelectedIndex = 3 Then
            Scheda_tecnica.assegna_foto(par_picturebox, Form_Codici_vendita.trova_immagine_codice("Y00052").ToString)
        ElseIf par_combobox.SelectedIndex = 4 Then
            Scheda_tecnica.assegna_foto(par_picturebox, Form_Codici_vendita.trova_immagine_codice("Y00053").ToString)
        ElseIf par_combobox.SelectedIndex = 5 Then
            Scheda_tecnica.assegna_foto(par_picturebox, Form_Codici_vendita.trova_immagine_codice("Y00054").ToString)
        Else par_picturebox.Image = Nothing
        End If
    End Sub
End Class