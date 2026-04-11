Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports AxFOXITREADERLib

Public Class CQ_Modulo_operativo
    Public Elenco_dipendenti(1000) As String
    Public Elenco_esito_controllo(10000000) As String
    Public codicedip As Integer
    Public esito_controllo As String

    Sub Inserimento_dipendenti()
        ComboBox2.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.[USERID] as 'Codice dipendenti', T0.[lastName] + ' ' + T0.[firstName] AS 'Nome'
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

        If Homepage.totem = "N" Then
            ComboBox2.Text = Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).cognome & " " & Homepage.trova_Dettagli_dipendente(Homepage.ID_SALVATO).nome
        End If

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
FROM UFD1
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

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.imputazione FROM cq_imputazioni t0 group by T0.imputazione"

        cmd_SAP_reader = CMD_SAP.ExecuteReader



        Do While cmd_SAP_reader.Read()

            ComboBox3.Items.Add(cmd_SAP_reader("imputazione"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub


    Sub Inserimento_definizione()
        ComboBox4.Items.Clear()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.categoria FROM cq_imputazioni t0 where T0.imputazione = '" & ComboBox3.Text & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            ComboBox4.Items.Add(cmd_SAP_reader("categoria"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



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
FROM [TIRELLISRLDB].[dbo].[CQ_descrizione] where categoria ='" & ComboBox4.Text & "' group by descrizione "

        cmd_SAP_reader = CMD_SAP.ExecuteReader

        Do While cmd_SAP_reader.Read()

            ComboBox5.Items.Add(cmd_SAP_reader("descrizione"))

        Loop
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub
    Sub Anagrafica_attività()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = Homepage.sap_tirelli
        cnn.Open()



        Dim CMD_SAP As New SqlCommand
        Dim cmd_SAP_reader As SqlDataReader

        CMD_SAP.Connection = cnn
        CMD_SAP.CommandText = "SELECT T0.u_PRG_QLT_ITEMCODE, T1.ITEMNAME, case when T1.U_DISEGNO is null then '' else t1.u_disegno end as 'u_disegno',case when T0.DOCTYPE = 20 then 'EM' else t0.doctype end as 'DOCTYPE' , T0.DOCNUM, T0.U_PRG_QLT_TCDOCQTY, case when T2.CARDNAME is null then 'MU_TIRELLI' else t2.cardname end as 'Cardname',case when T0.TEL is null then '' else t0.tel end as 'tel', case when T2.E_MAIL is null then '' else t2.E_mail end as 'E_mail'
FROM OCLG T0 LEFT JOIN OITM T1 ON T0.u_PRG_QLT_ITEMCODE=T1.ITEMCODE
left JOIN OCRD T2 ON T0.CARDCODE=T2.CARDCODE
WHERE T0.[ClgCode] ='" & CQ_AttivitaAperte.N_attivita & "'"

        cmd_SAP_reader = CMD_SAP.ExecuteReader


        If cmd_SAP_reader.Read() Then

            Label1.Text = CQ_AttivitaAperte.N_attivita
            Label8.Text = cmd_SAP_reader("u_PRG_QLT_ITEMCODE")
            Label9.Text = cmd_SAP_reader("U_Disegno")
            Label10.Text = cmd_SAP_reader("Itemname")
            Label3.Text = cmd_SAP_reader("Cardname")
            Label4.Text = cmd_SAP_reader("Tel")
            Label5.Text = cmd_SAP_reader("e_mail")
            Label7.Text = Math.Round(cmd_SAP_reader("U_PRG_QLT_TCDOCQTY"))
            Label2.Text = cmd_SAP_reader("DOCNUM")
            Label6.Text = cmd_SAP_reader("DOCType")
            TextBox2.Text = Math.Round(cmd_SAP_reader("U_PRG_QLT_TCDOCQTY"))
            TextBox4.Text = Math.Round(cmd_SAP_reader("U_PRG_QLT_TCDOCQTY"))

        End If
        cmd_SAP_reader.Close()
        cnn.Close()



    End Sub 'Inserisco le risorse nella combo box


    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        codicedip = Elenco_dipendenti(ComboBox2.SelectedIndex)

    End Sub




    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
        CQ_AttivitaAperte.Button3.Enabled = True

        'Rendere inattivi dati per Gestione NC
        Button1.Enabled = True
        Button5.Enabled = False
        Button6.Enabled = False

        'Rendere invisibili dati per Gestione NC
        Button1.Visible = True
        Button5.Visible = False
        Button6.Visible = False



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

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Inserimento_Descrizione_NC()
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

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

        'Attivazione Campi NC
        If TextBox3.Text = Nothing Or TextBox3.Text = "0" Then

            ComboBox4.Enabled = False
            ComboBox5.Enabled = False
            ComboBox6.Enabled = False
            RichTextBox1.Enabled = False
            RichTextBox2.Enabled = False
            ComboBox4.Text = Nothing
            ComboBox5.Text = Nothing
            ComboBox6.Text = Nothing
            ComboBox7.Text = ""
            RichTextBox1.Text = Nothing
            RichTextBox2.Text = Nothing

            'Impostazione materiale conforme
            TextBox4.Text = TextBox2.Text
        Else
            If Int(TextBox3.Text) > Int(TextBox2.Text) Then
                MsgBox("Quantità non conforme superiore alla controllata")
                TextBox3.Text = Nothing
            Else
                ComboBox4.Enabled = True
                ComboBox5.Enabled = True
                ComboBox6.Enabled = True
                RichTextBox1.Enabled = True
                RichTextBox2.Enabled = True

                'Sottrazione materiale conforme
                TextBox4.Text = Int(TextBox2.Text) - Int(TextBox3.Text)
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If File.Exists(Homepage.percorso_disegni_generico & "PDF\"  & Label9.Text & ".PDF") Then
            Process.Start(Homepage.percorso_disegni_generico & "PDF\"  & Label9.Text & ".PDF")
        Else
            MsgBox("PDF non presente")
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If File.Exists(Homepage.percorso_DWF & Label9.Text & ".iam.dwf") Then
            Process.Start(Homepage.percorso_DWF & Label9.Text & ".iam.dwf")
        Else
            MsgBox("3D non presente")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim CNN As New SqlConnection
        'Controllo di corretta compilazione  
        If ComboBox3.Text = "MU" And TextBox7.Text = Nothing Then
            MsgBox("Non è possibile dare NC a MU senza indicare l'autocontrollo")

        Else

            If TextBox4.Text <> TextBox2.Text Then



                If TextBox3.Text = Nothing Or ComboBox3.Text = Nothing Or ComboBox4.Text = Nothing Or ComboBox1.Text = Nothing Or ComboBox2.Text = Nothing Or ComboBox6.Text = Nothing Or ComboBox7.Text = Nothing Or RichTextBox1.Text = Nothing Or RichTextBox2.Text = Nothing Then

                    'BISOGNA FARE LE VARIANTI A SECONDA SE SIA CONFORME O MENO

                    MsgBox("il controllo ha stabilito che ci sono delle non conformità ma mancano uno o più dati ")
                Else




                    If ComboBox7.Text <> "Deroga" Then
                        cnn.ConnectionString = homepage.sap_tirelli
                        cnn.Open()

                        Dim CMD_SAP_7 As New SqlCommand

                        CMD_SAP_7.Connection = cnn
                        CMD_SAP_7.CommandText = "UPDATE OCLG SET OCLG.[U_PRG_QLT_TCCtrlQty]='" & TextBox2.Text & "', OCLG.[U_PRG_QLT_TCNOKQty]='" & TextBox3.Text & "', OCLG.[U_PRG_QLT_TCOKQty]='" & TextBox4.Text & "', OCLG.[U_PRG_QLT_QCNCEmp]= '" & ComboBox3.Text & "', OCLG.[U_Campo_definizione_NC]='" & ComboBox4.Text & "', OCLG.[U_Descrizione_NC]='" & ComboBox5.Text & "', OCLG.[Notes]=CONCAT('" & RichTextBox1.Text & "',' ', '" & RichTextBox2.Text & "' ), OCLG.[U_PRG_QLT_TCResult]='" & esito_controllo & "',   OCLG.[U_Peso_NC]='" & ComboBox6.Text & "' ,OCLG.[U_PRG_QLT_TCTOEmp]='" & ComboBox1.Text & "', OCLG.[AttendUser] = '" & codicedip & "', OCLG.[U_Stato]='W'
 where OCLG.clgcode='" & CQ_AttivitaAperte.N_attivita & "'"
                        CMD_SAP_7.ExecuteNonQuery()
                        cnn.Close()
                        regisrazione_in_nuovo_controllo_cq()
                        MsgBox("Registazione effettuata")
                        pulizia_form()
                    Else
                        If ComboBox7.Text = "Deroga" And Int(TextBox2.Text) = Int(TextBox3.Text) Then
                            cnn.ConnectionString = homepage.sap_tirelli
                            cnn.Open()

                            Dim CMD_SAP_7 As New SqlCommand

                            CMD_SAP_7.Connection = cnn
                            CMD_SAP_7.CommandText = "UPDATE OCLG SET OCLG.[U_PRG_QLT_TCCtrlQty]='" & TextBox2.Text & "', OCLG.[U_PRG_QLT_TCNOKQty]='" & TextBox3.Text & "', OCLG.[U_PRG_QLT_TCOKQty]='" & TextBox4.Text & "', OCLG.[U_PRG_QLT_QCNCEmp]= '" & ComboBox3.Text & "', OCLG.[U_Campo_definizione_NC]='" & ComboBox4.Text & "', OCLG.[U_Descrizione_NC]='" & ComboBox5.Text & "', OCLG.[Notes]=CONCAT('" & RichTextBox1.Text & "',' ', '" & RichTextBox2.Text & "' ), OCLG.[U_PRG_QLT_TCResult]='C',   OCLG.[U_Peso_NC]='" & ComboBox6.Text & "' ,OCLG.[U_PRG_QLT_TCTOEmp]='" & ComboBox1.Text & "', OCLG.[AttendUser] = '" & codicedip & "', OCLG.[U_Stato]='W', OCLG.U_PRG_QLT_TCTORES='AD'
 where OCLG.clgcode='" & CQ_AttivitaAperte.N_attivita & "'"
                            CMD_SAP_7.ExecuteNonQuery()
                            cnn.Close()
                            regisrazione_in_nuovo_controllo_cq()
                            MsgBox("Registazione effettuata")
                            pulizia_form()

                        Else

                            cnn.ConnectionString = homepage.sap_tirelli
                            cnn.Open()

                            Dim CMD_SAP_7 As New SqlCommand

                            CMD_SAP_7.Connection = cnn
                            CMD_SAP_7.CommandText = "UPDATE OCLG SET OCLG.[U_PRG_QLT_TCCtrlQty]='" & TextBox2.Text & "', OCLG.[U_PRG_QLT_TCNOKQty]='" & TextBox3.Text & "', OCLG.[U_PRG_QLT_TCOKQty]='" & TextBox4.Text & "', OCLG.[U_PRG_QLT_QCNCEmp]= '" & ComboBox3.Text & "', OCLG.[U_Campo_definizione_NC]='" & ComboBox4.Text & "', OCLG.[U_Descrizione_NC]='" & ComboBox5.Text & "', OCLG.[Notes]=CONCAT('" & RichTextBox1.Text & "',' ', '" & RichTextBox2.Text & "' ), OCLG.[U_PRG_QLT_TCResult]='C',   OCLG.[U_Peso_NC]='" & ComboBox6.Text & "' ,OCLG.[U_PRG_QLT_TCTOEmp]='" & ComboBox1.Text & "', OCLG.[AttendUser] = '" & codicedip & "', OCLG.[U_Stato]='W',OCLG.U_PRG_QLT_TCTORES='AR'
 where OCLG.clgcode='" & CQ_AttivitaAperte.N_attivita & "'"
                            CMD_SAP_7.ExecuteNonQuery()
                            cnn.Close()
                            regisrazione_in_nuovo_controllo_cq()
                            MsgBox("Registazione effettuata")
                            pulizia_form()


                        End If
                    End If
                End If

            Else

                If ComboBox3.Text = Nothing Or ComboBox1.Text = Nothing Or ComboBox2.Text = Nothing Or ComboBox7.Text = Nothing Then

                    'BISOGNA FARE LE VARIANTI A SECONDA SE SIA CONFORME O MENO

                    MsgBox("il controllo ha stabilito che i pezzi sono conformi ma mancano uno o più dati ")
                Else
                    If ComboBox7.Text <> "Deroga" Then
                        cnn.ConnectionString = homepage.sap_tirelli
                        cnn.Open()

                        Dim CMD_SAP_7 As New SqlCommand

                        CMD_SAP_7.Connection = cnn
                        CMD_SAP_7.CommandText = "UPDATE OCLG SET OCLG.[U_PRG_QLT_TCCtrlQty]='" & TextBox2.Text & "', OCLG.[U_PRG_QLT_TCOKQty]='" & TextBox4.Text & "', OCLG.[U_PRG_QLT_QCNCEmp]= '" & ComboBox3.Text & "', OCLG.[U_Campo_definizione_NC]='" & ComboBox4.Text & "', OCLG.[U_PRG_QLT_TCResult]='" & esito_controllo & "',   OCLG.[U_Peso_NC]='" & ComboBox6.Text & "' ,OCLG.[U_PRG_QLT_TCTOEmp]='" & ComboBox1.Text & "', OCLG.[AttendUser] = '" & codicedip & "', OCLG.[U_Stato]='W'
 where OCLG.clgcode='" & CQ_AttivitaAperte.N_attivita & "'"
                        CMD_SAP_7.ExecuteNonQuery()
                        cnn.Close()
                        regisrazione_in_nuovo_controllo_cq()
                        MsgBox("Registazione effettuata")
                        pulizia_form()
                    Else
                        If ComboBox7.Text = "Deroga" And Int(TextBox2.Text) = Int(TextBox3.Text) Then
                            cnn.ConnectionString = homepage.sap_tirelli
                            cnn.Open()

                            Dim CMD_SAP_7 As New SqlCommand

                            CMD_SAP_7.Connection = cnn
                            CMD_SAP_7.CommandText = "UPDATE OCLG SET OCLG.[U_PRG_QLT_TCCtrlQty]='" & TextBox2.Text & "', OCLG.[U_PRG_QLT_TCNOKQty]='" & TextBox3.Text & "', OCLG.[U_PRG_QLT_TCOKQty]='" & TextBox4.Text & "', OCLG.[U_PRG_QLT_QCNCEmp]= '" & ComboBox3.Text & "', OCLG.[U_Campo_definizione_NC]='" & ComboBox4.Text & "', OCLG.[U_Descrizione_NC]='" & ComboBox5.Text & "', OCLG.[Notes]=CONCAT('" & RichTextBox1.Text & "',' ', '" & RichTextBox2.Text & "' ), OCLG.[U_PRG_QLT_TCResult]='C',   OCLG.[U_Peso_NC]='" & ComboBox6.Text & "' ,OCLG.[U_PRG_QLT_TCTOEmp]='" & ComboBox1.Text & "', OCLG.[AttendUser] = '" & codicedip & "', OCLG.[U_Stato]='W', OCLG.U_PRG_QLT_TCTORES='AD'
 where OCLG.clgcode='" & CQ_AttivitaAperte.N_attivita & "'"
                            CMD_SAP_7.ExecuteNonQuery()
                            cnn.Close()
                            regisrazione_in_nuovo_controllo_cq()
                            MsgBox("Registazione effettuata")
                            pulizia_form()

                        Else

                            cnn.ConnectionString = homepage.sap_tirelli
                            cnn.Open()

                            Dim CMD_SAP_7 As New SqlCommand

                            CMD_SAP_7.Connection = cnn
                            CMD_SAP_7.CommandText = "UPDATE OCLG SET OCLG.[U_PRG_QLT_TCCtrlQty]='" & TextBox2.Text & "', OCLG.[U_PRG_QLT_TCNOKQty]='" & TextBox3.Text & "', OCLG.[U_PRG_QLT_TCOKQty]='" & TextBox4.Text & "', OCLG.[U_PRG_QLT_QCNCEmp]= '" & ComboBox3.Text & "', OCLG.[U_Campo_definizione_NC]='" & ComboBox4.Text & "', OCLG.[U_Descrizione_NC]='" & ComboBox5.Text & "', OCLG.[Notes]=CONCAT('" & RichTextBox1.Text & "',' ', '" & RichTextBox2.Text & "' ), OCLG.[U_PRG_QLT_TCResult]='C',   OCLG.[U_Peso_NC]='" & ComboBox6.Text & "' ,OCLG.[U_PRG_QLT_TCTOEmp]='" & ComboBox1.Text & "', OCLG.[AttendUser] = '" & codicedip & "', OCLG.[U_Stato]='W',OCLG.U_PRG_QLT_TCTORES='AR'
 where OCLG.clgcode='" & CQ_AttivitaAperte.N_attivita & "'"
                            CMD_SAP_7.ExecuteNonQuery()
                            cnn.Close()
                            regisrazione_in_nuovo_controllo_cq()
                            MsgBox("Registazione effettuata")
                            pulizia_form()
                        End If
                    End If
                End If

            End If
        End If
    End Sub

    Sub regisrazione_in_nuovo_controllo_cq()
        CQ_nuovo_controllo.trova_ID()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand

        CMD_SAP_7.Connection = cnn
        CMD_SAP_7.CommandText = "insert into cq_nuovo_controllo (ID, CODICE, DISEGNO, PZ_CONTR, PZ_NC, PZ_OK, IMPUTAZIONE, CAMPO_DEFINIZIONE_NC,descrizione_nc, ESITO_AUTOCONTROLLO, PESO_NC,richiesto,autocontrollo, OPERATORE, ZONA_CONTROLLO, DATA, ORA,stato,attività,rilevato,concedente)
VALUES ('" & CQ_nuovo_controllo.id & "','" & Label8.Text & "','" & Label9.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "', '" & ComboBox3.Text & "', '" & ComboBox4.Text & "','" & ComboBox5.Text & "' , '" & ComboBox7.Text & "', '" & ComboBox6.Text & "','" & RichTextBox1.Text & "','" & TextBox7.Text & "'," & codicedip & ",'" & ComboBox1.Text & "',GETDATE(),convert(varchar, getdate(), 108),'In_corso','" & Label1.Text & "','" & RichTextBox2.Text & "','" & ComboBox8.Text & "')  "


        CMD_SAP_7.ExecuteNonQuery()
        cnn.Close()
    End Sub

    Sub pulizia_form()
        ComboBox3.Text = Nothing
        ComboBox1.Text = Nothing

        ComboBox7.Text = ""
        TextBox2.Text = Nothing
        TextBox3.Text = Nothing
        TextBox4.Text = Nothing
        ComboBox4.Text = Nothing
        RichTextBox1.Text = Nothing
        RichTextBox2.Text = Nothing
        ComboBox6.Text = Nothing
    End Sub

    Sub Cambia_stato_ODP()
        Dim CNN As New SqlConnection
        CNN.ConnectionString = homepage.sap_tirelli
        cnn.Open()

        Dim CMD_SAP_7 As New SqlCommand

        CMD_SAP_7.Connection = cnn

        CMD_SAP_7.CommandText = "insert into OCLG (OCLG.[U_PRG_QLT_TCCtrlQty], OCLG.[U_PRG_QLT_TCNOKQty], OCLG.[U_PRG_QLT_TCOKQty], OCLG.[U_PRG_QLT_QCNCEmp], OCLG.[U_Campo_definizione_NC], OCLG.[U_Descrizione_NC], OCLG.[Notes], OCLG.[U_PRG_QLT_TCResult], OCLG.[U_PRG_QLT_TCTORes], OCLG.[U_Peso_NC], OCLG.[U_PRG_QLT_TCTOEmp], OCLG.[AttendUser], OCLG.[U_Stato])
VALUES ('" & TextBox2.Text & "', '" & TextBox3.Text & "' , '" & TextBox4.Text & "', '" & ComboBox3.Text & "','" & ComboBox4.Text & "','" & ComboBox5.Text & "',concat('" & RichTextBox1.Text & "',' ','" & RichTextBox2.Text & "'),'" & ComboBox7.Text & "','" & ComboBox4.Text & "' )
where OCLG.clgcode='" & CQ_AttivitaAperte.N_attivita & "'"

        CMD_SAP_7.ExecuteNonQuery()

        cnn.Close()


        'Chiusura scheda
        Me.Close()
        'Attivare tasto X
        CQ_AttivitaAperte.Button3.Enabled = True
        'Rendere inattivi dati per Gestione NC
        Button1.Enabled = True
        Button5.Enabled = False
        Button6.Enabled = False

        'Rendere invisibili dati per Gestione NC
        Button1.Visible = True
        Button5.Visible = False
        Button6.Visible = False


    End Sub



    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        CQ_Password.Show()
        CQ_Password.RadioButton2.Checked = True
        CQ_Password.Owner = Me
    End Sub

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
        If TextBox3.TextLength = 0 And KeyAscii = 46 Then
            KeyAscii = 0
        End If

        ' ovviamente se c'è già un punto
        ' non è consentito digitarne altri
        If (KeyAscii = 46) And
            TextBox3.Text.IndexOf(".") > 0 Then
            KeyAscii = 0
        End If

        ' Reimposto il keychar
        e.KeyChar = Chr(KeyAscii)
    End Sub 'permetto solo numeri come input



    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        CQ_Tabelle.TextBox2.Text = label8.text
        CQ_Tabelle.Show()
        CQ_Tabelle.Owner = Me
        CQ_Tabelle.riempi_autocontrollo()
    End Sub

    Private Sub CQ_Modulo_operativo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If ComboBox7.Text = "Concesso" Then
            GroupBox25.Visible = True
        Else
            GroupBox25.Visible = False
            ComboBox8.Text = ""
        End If
    End Sub
End Class