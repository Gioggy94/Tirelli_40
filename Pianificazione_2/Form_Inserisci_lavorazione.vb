Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb

Public Class Form_Inserisci_Lavorazione
    Public Elenco_Dipendenti(1000) As Integer
    Public Nomi_Dipendenti(1000) As String
    Public Num_Dipendenti As Integer
    Public Stringa_Ricerca As String
    Public Codice_Dipendente_Attuale As Integer
    Public Risorsa As String
    Public Tipologia_Lavorazione As String
    Public Id_Chiusura As Integer
    Public Chiusura As Integer


    Private Sub Form_Inserisci_Lavorazione_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Compila_Elenco_Dipendenti()
        Txt_Dipendente.Enabled = False
        Stringa_Ricerca = ""
        Codice_Dipendente_Attuale = -1
        Cmd_Collaudo.Checked = True
        Cmd_Meccanico.Checked = False
        Cmd_Elettrico.Checked = False
        Cmd_Imballo.Checked = False
        Risorsa = "R00524"
        Tipologia_Lavorazione = "5"
        Cmd_Chiudi.Enabled = False
        Cmd_Apri.Enabled = False
        Lbl_Messaggio.Text = "Inserire Cognome e Nome ..."
    End Sub

    Private Sub Compila_Elenco_Dipendenti()
        Dim Indice As Integer

        Dim Cnn_Dipendenti As New SqlConnection

        Cnn_Dipendenti.ConnectionString = homepage.sap_tirelli
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

    Private Sub Ricerca_Nome()
        Cmd_Chiudi.Enabled = False
        Cmd_Apri.Enabled = False
        Lbl_Messaggio.Text = "Inserire Cognome e Nome ..."
        If Stringa_Ricerca.Length > 0 Then
            'Txt_Ricerca.Text = Stringa_Ricerca
            Dim Esito_Ricerca As Integer

            Esito_Ricerca = -1

            Dim i As Integer
            For i = 0 To Num_Dipendenti - 1 Step 1
                If Nomi_Dipendenti(i).ToLower Like (Stringa_Ricerca.ToLower & "*") Then
                    Txt_Dipendente.Text = Nomi_Dipendenti(i)
                    Codice_Dipendente_Attuale = Elenco_Dipendenti(i)
                    Esito_Ricerca = Codice_Dipendente_Attuale
                End If
            Next
            If Esito_Ricerca = -1 Then
                If Stringa_Ricerca.Length > 0 Then
                    Stringa_Ricerca = Stringa_Ricerca.Remove(Stringa_Ricerca.Length - 1)
                    'Txt_Ricerca.Text = Stringa_Ricerca
                    If Stringa_Ricerca.Length = 0 Then
                        Codice_Dipendente_Attuale = -1
                    End If
                Else
                    Stringa_Ricerca = ""
                    'Txt_Ricerca.Text = ""
                    Codice_Dipendente_Attuale = -1
                End If
            End If
        Else
            Codice_Dipendente_Attuale = -1
            Txt_Dipendente.Text = ""
            Stringa_Ricerca = ""
            'Txt_Ricerca.Text = Stringa_Ricerca
        End If
        Txt_Dipendente.SelectionStart = 0
        Txt_Dipendente.SelectionLength = Stringa_Ricerca.Length
    End Sub


    Private Sub Cmd_G_Click(sender As Object, e As EventArgs) Handles Cmd_G.Click
        Stringa_Ricerca = Stringa_Ricerca + "g"
        Ricerca_Nome()
    End Sub

    Private Sub Cmb_Cancella_Click(sender As Object, e As EventArgs) Handles Cmb_Cancella.Click
        If Stringa_Ricerca.Length > 0 Then
            Stringa_Ricerca = Stringa_Ricerca.Remove(Stringa_Ricerca.Length - 1)
        End If
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_Q_Click(sender As Object, e As EventArgs) Handles Cmd_Q.Click
        Stringa_Ricerca = Stringa_Ricerca + "q"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_W_Click(sender As Object, e As EventArgs) Handles Cmd_W.Click
        Stringa_Ricerca = Stringa_Ricerca + "w"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_E_Click(sender As Object, e As EventArgs) Handles Cmd_E.Click
        Stringa_Ricerca = Stringa_Ricerca + "e"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_R_Click(sender As Object, e As EventArgs) Handles Cmd_R.Click
        Stringa_Ricerca = Stringa_Ricerca + "r"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_T_Click(sender As Object, e As EventArgs) Handles Cmd_T.Click
        Stringa_Ricerca = Stringa_Ricerca + "t"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_Y_Click(sender As Object, e As EventArgs) Handles Cmd_Y.Click
        Stringa_Ricerca = Stringa_Ricerca + "y"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_U_Click(sender As Object, e As EventArgs) Handles Cmd_U.Click
        Stringa_Ricerca = Stringa_Ricerca + "u"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_I_Click(sender As Object, e As EventArgs) Handles Cmd_I.Click
        Stringa_Ricerca = Stringa_Ricerca + "i"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_O_Click(sender As Object, e As EventArgs) Handles Cmd_O.Click
        Stringa_Ricerca = Stringa_Ricerca + "o"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_P_Click(sender As Object, e As EventArgs) Handles Cmd_P.Click
        Stringa_Ricerca = Stringa_Ricerca + "p"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_A_Click(sender As Object, e As EventArgs) Handles Cmd_A.Click
        Stringa_Ricerca = Stringa_Ricerca + "a"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_S_Click(sender As Object, e As EventArgs) Handles Cmd_S.Click
        Stringa_Ricerca = Stringa_Ricerca + "s"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_D_Click(sender As Object, e As EventArgs) Handles Cmd_D.Click
        Stringa_Ricerca = Stringa_Ricerca + "d"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_F_Click(sender As Object, e As EventArgs) Handles Cmd_F.Click
        Stringa_Ricerca = Stringa_Ricerca + "f"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_H_Click(sender As Object, e As EventArgs) Handles Cmd_H.Click
        Stringa_Ricerca = Stringa_Ricerca + "h"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_J_Click(sender As Object, e As EventArgs) Handles Cmd_J.Click
        Stringa_Ricerca = Stringa_Ricerca + "j"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_K_Click(sender As Object, e As EventArgs) Handles Cmd_K.Click
        Stringa_Ricerca = Stringa_Ricerca + "k"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_L_Click(sender As Object, e As EventArgs) Handles Cmd_L.Click
        Stringa_Ricerca = Stringa_Ricerca + "l"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_Z_Click(sender As Object, e As EventArgs) Handles Cmd_Z.Click
        Stringa_Ricerca = Stringa_Ricerca + "z"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_X_Click(sender As Object, e As EventArgs) Handles Cmd_X.Click
        Stringa_Ricerca = Stringa_Ricerca + "x"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_C_Click(sender As Object, e As EventArgs) Handles Cmd_C.Click
        Stringa_Ricerca = Stringa_Ricerca + "c"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_V_Click(sender As Object, e As EventArgs) Handles Cmd_V.Click
        Stringa_Ricerca = Stringa_Ricerca + "v"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_B_Click(sender As Object, e As EventArgs) Handles Cmd_B.Click
        Stringa_Ricerca = Stringa_Ricerca + "b"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_N_Click(sender As Object, e As EventArgs) Handles Cmd_N.Click
        Stringa_Ricerca = Stringa_Ricerca + "n"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_M_Click(sender As Object, e As EventArgs) Handles Cmd_M.Click
        Stringa_Ricerca = Stringa_Ricerca + "m"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_Spazio_Click(sender As Object, e As EventArgs) Handles Cmd_Spazio.Click
        Stringa_Ricerca = Stringa_Ricerca + " "
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_Apice_Click(sender As Object, e As EventArgs) Handles Cmd_Apice.Click
        Stringa_Ricerca = Stringa_Ricerca + "'"
        Ricerca_Nome()
    End Sub

    Private Sub Cmd_Collaudo_CheckedChanged(sender As Object, e As EventArgs) Handles Cmd_Collaudo.Click
        Cmd_Collaudo.Checked = True
        Cmd_Meccanico.Checked = False
        Cmd_Elettrico.Checked = False
        Cmd_Imballo.Checked = False
        Risorsa = "R00524"
        Tipologia_Lavorazione = "5"
    End Sub

    Private Sub Cmd_Meccanico_CheckedChanged(sender As Object, e As EventArgs) Handles Cmd_Meccanico.Click
        Cmd_Collaudo.Checked = False
        Cmd_Meccanico.Checked = True
        Cmd_Elettrico.Checked = False
        Cmd_Imballo.Checked = False
        Risorsa = "R00525"
        Tipologia_Lavorazione = "3"
    End Sub

    Private Sub Cmd_Elettrico_CheckedChanged(sender As Object, e As EventArgs) Handles Cmd_Elettrico.Click
        Cmd_Collaudo.Checked = False
        Cmd_Meccanico.Checked = False
        Cmd_Elettrico.Checked = True
        Cmd_Imballo.Checked = False
        Risorsa = "R00530"
        Tipologia_Lavorazione = "2"
    End Sub

    Private Sub Cmd_Imballo_CheckedChanged(sender As Object, e As EventArgs) Handles Cmd_Imballo.Click
        Cmd_Collaudo.Checked = False
        Cmd_Meccanico.Checked = False
        Cmd_Elettrico.Checked = False
        Cmd_Imballo.Checked = True
        Risorsa = "R00545"
        Tipologia_Lavorazione = "6"
    End Sub

    Private Sub Cmd_Invio_Click(sender As Object, e As EventArgs) Handles Cmd_Invio.Click
        If Codice_Dipendente_Attuale <> -1 Then
            Dim Cnn_Lavorazione As New SqlConnection
            Cnn_Lavorazione.ConnectionString = homepage.sap_tirelli
            Cnn_Lavorazione.Open()

            Dim Cmd_Lavorazione As New SqlCommand
            Dim Cmd_Lavorazione_Reader As SqlDataReader

            Cmd_Lavorazione.Connection = Cnn_Lavorazione
            Cmd_Lavorazione.CommandText = " SELECT * FROM MANODOPERA WHERE dipendente=" & Codice_Dipendente_Attuale & " AND (stop IS NULL OR stop='' OR stop='0') AND (consuntivo IS NULL OR consuntivo='' OR consuntivo='0')" 'data=getdate()
            Cmd_Lavorazione_Reader = Cmd_Lavorazione.ExecuteReader
            If Cmd_Lavorazione_Reader.Read() Then
                'Lavorazione Aperta
                Id_Chiusura = Cmd_Lavorazione_Reader("Id")
                Lbl_Messaggio.Text = "Necessario chiudere lavorazione su " & Cmd_Lavorazione_Reader("tipo_Documento") & " " & Cmd_Lavorazione_Reader("docnum") & " ..."
                Cmd_Chiudi.Enabled = True
            Else
                'Nessuna Lavorazione Aperta
                If Chiusura = 1 Then
                    Lbl_Messaggio.Text = "Nessuna Lavorazione da Chiudere ... "
                Else
                    Lbl_Messaggio.Text = "Apertura Lavorazione su Ordine di Produzione " & Lbl_ODP.Text & "  ..."
                    Cmd_Apri.Enabled = True
                End If
            End If
            Cnn_Lavorazione.Close()
        End If
    End Sub

    Private Sub Cmd_Chiudi_Click(sender As Object, e As EventArgs) Handles Cmd_Chiudi.Click
        Dim Cnn_Lavorazione As New SqlConnection
        Cnn_Lavorazione.ConnectionString = homepage.sap_tirelli
        Cnn_Lavorazione.Open()

        Dim Cmd_Lavorazione As New SqlCommand

        Cmd_Lavorazione.Connection = Cnn_Lavorazione
        Cmd_Lavorazione.CommandText = " UPDATE MANODOPERA SET STOP=convert(varchar, getdate(), 108) WHERE id=" & Id_Chiusura
        Cmd_Lavorazione.ExecuteNonQuery()
        Cnn_Lavorazione.Close()
        Cmd_Chiudi.Enabled = False
        If Chiusura = 1 Then
            Lbl_Messaggio.Text = "Lavorazione Chiusa con Successo ... "
        Else
            Lbl_Messaggio.Text = "Apertura Lavorazione su Ordine di Produzione " & Lbl_ODP.Text & "  ..."
            Cmd_Apri.Enabled = True
        End If
    End Sub

    Private Sub Cmd_Apri_Click(sender As Object, e As EventArgs) Handles Cmd_Apri.Click
        Dim Cnn_Lavorazione As New SqlConnection
        Cnn_Lavorazione.ConnectionString = homepage.sap_tirelli
        Cnn_Lavorazione.Open()

        Dim Cmd_Lavorazione As New SqlCommand

        Cmd_Lavorazione.Connection = Cnn_Lavorazione
        Cmd_Lavorazione.CommandText = "INSERT INTO manodopera (id,tipo_documento,docnum,dipendente,risorsa,data,start,consuntivo, tipologia_lavorazione) 
values (" & Trova_ID() & ",'ODP'," & Lbl_ODP.Text & ",'" & Codice_Dipendente_Attuale & "','" & Risorsa & "',getdate(),convert(varchar, getdate(), 108),0,'" & Tipologia_Lavorazione & "')"
        Cmd_Lavorazione.ExecuteNonQuery()
        Cnn_Lavorazione.Close()
        Cmd_Apri.Enabled = False

        Txt_Dipendente.Enabled = False
        Stringa_Ricerca = ""
        Txt_Dipendente.Text = ""
        Codice_Dipendente_Attuale = -1
        Cmd_Collaudo.Checked = True
        Cmd_Meccanico.Checked = False
        Cmd_Elettrico.Checked = False
        Cmd_Imballo.Checked = False
        Risorsa = "R00524"
        Tipologia_Lavorazione = "5"
        Cmd_Chiudi.Enabled = False
        Cmd_Apri.Enabled = False
        Lbl_Messaggio.Text = "Inserire Cognome e Nome ..."
    End Sub

    Function Trova_ID() As Integer
        Dim Cnn_ID As New SqlConnection
        Dim Cmd_ID As New SqlCommand
        Dim Cmd_ID_Reader As SqlDataReader
        Dim Id As Integer

        Cnn_ID.ConnectionString = homepage.sap_tirelli
        Cnn_ID.Open()
        Cmd_ID.Connection = Cnn_ID
        Cmd_ID.CommandText = "select max(id)+1 as 'ID' from manodopera"

        Cmd_ID_Reader = Cmd_ID.ExecuteReader

        If Cmd_ID_Reader.Read() = True Then
            If Not Cmd_ID_Reader("ID") Is System.DBNull.Value Then
                Id = Cmd_ID_Reader("ID")
            Else
                Id = 1
            End If
        End If
        Cnn_ID.Close()
        Return Id
    End Function

    Private Sub Cmd_Esci_Click(sender As Object, e As EventArgs) Handles Cmd_Esci.Click

        Me.Close()
    End Sub
End Class